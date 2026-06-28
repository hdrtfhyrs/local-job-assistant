# -*- coding: utf-8 -*-
"""
智能求职方向参谋（决定"搜什么"——在采集之前先想清楚该投哪些方向）

结合三样东西，用本地 LLM 推理出"为你定制的求职方向 + 该搜的关键词"：
  1. 个人画像        match_profile.json
  2. 本地方向库      career_direction_profile.json（带景德镇本地适配评分）
  3. 本地就业行情    local_market_snapshot.md（联网整理的真实行情）

模型：默认 qwen2.5:14b（更聪明），没装则自动回退 qwen2.5:7b。

输出：
  recommended_directions.json   结构化推荐（含汇总搜索关键词，供采集直接使用）
  求职方向参谋_<时间>.md         人类可读报告

运行：python career_advisor.py [--model qwen2.5:14b] [--top 6]
"""

import argparse
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys

import requests

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OLLAMA_HOST = "http://localhost:11434"
PREFERRED_MODELS = ["qwen2.5:14b", "qwen2.5:7b"]

SYSTEM_PROMPT = (
    "你是一位资深的本地求职顾问，熟悉中国三四线城市（尤其景德镇）的就业市场，"
    "擅长为『学历普通、技术基础一般』的毕业生规划务实、能落地的求职方向。"
    "你会结合求职者的真实技能与硬性底线（如不接受驻场、薪资底线）、以及当地真实的产业与岗位行情，"
    "推荐『既适合 TA、本地又确实有岗位』的方向，并给出能直接拿去招聘网站搜索的关键词。"
    "务实、不画饼、主动规避求职陷阱（传销/电销/驻场/纯体力/夜班）。只输出严格的 JSON。"
)


def configure_console_encoding():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def load_json(path, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"  ⚠️ 读取 {os.path.basename(path)} 失败：{e}")
        return default


def load_text(path, default=""):
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except Exception:
        return default


def installed_models():
    try:
        r = requests.get(f"{OLLAMA_HOST}/api/tags", timeout=5)
        r.raise_for_status()
        return [m.get("name", "") for m in r.json().get("models", [])]
    except Exception:
        return []


def pick_model(requested):
    """优先用 requested；没装则按 PREFERRED 回退；都没有则用第一个已装模型。"""
    inst = installed_models()
    if not inst:
        return requested, []
    for cand in [requested] + PREFERRED_MODELS:
        if not cand:
            continue
        if cand in inst:
            return cand, inst
        base = cand.split(":")[0]
        for m in inst:
            if m.split(":")[0] == base:
                return m, inst
    return inst[0], inst


def build_profile_brief(profile):
    have = "、".join(profile.get("skills", {}).get("have", []))
    grow = "、".join(profile.get("skills", {}).get("want_to_grow", [])) or "（未指定，请你全权判断推荐）"
    city = profile.get("city", "")
    floor = profile.get("salary_floor_k")
    relo = ("不接受长期出差/驻场，优先本地"
            if not profile.get("accept_relocation") else "可接受出差/驻场")
    background = profile.get("background") or "普通本科软件工程，技术基础一般"
    notes = str(profile.get("extra_notes", "")).strip()
    brief = (f"城市：{city}；{relo}；期望月薪下限：{floor}K；"
             f"背景：{background}；已具备技能：{have}；想发展方向：{grow}。")
    if notes:
        brief += f"特别要求/顾虑：{notes}。"
    return brief


def build_directions_brief(career):
    lines = []
    for d in career.get("directions", []):
        aliases = "、".join(d.get("aliases", [])[:4])
        lines.append(
            f"- {d.get('name','')}（本地适配{d.get('local_fit_score','?')}/入门{d.get('entry_score','?')}"
            f"/前景{d.get('outlook_score','?')}）：{str(d.get('why',''))[:50]}；"
            f"可搜：{aliases}")
    return "\n".join(lines)


def build_user_prompt(profile_brief, market, directions_brief, top):
    return (
        f"【求职者画像】\n{profile_brief}\n\n"
        f"【景德镇当前就业行情（真实，务必参考）】\n{market}\n\n"
        f"【本地候选方向库（供参考，可增删、可合并）】\n{directions_brief}\n\n"
        f"请为这位求职者推荐最该投、且本地确实有岗位的 {top} 个求职方向，按优先级从高到低。\n\n"
        "【重要】动手前，先用联网搜索（WebSearch）逐个核实你打算推荐的方向在该城市的"
        "**真实评价**：真实待遇区间、加班/驻场情况、晋升天花板、从业者吐槽的坑。"
        "高效搜索（每个方向搜 1-2 次即可），基于搜到的真实口碑来打分和提醒，别凭印象编。\n\n"
        "每个方向给出能直接拿去 BOSS/智联 搜索的关键词（2-4 个，含同义变体）。\n"
        "只输出一个 JSON 对象，结构如下：\n"
        '{"recommendations": [\n'
        '  {"direction": "方向名称",\n'
        '   "why_fit": "为什么适合TA（结合其技能/底线，20-50字）",\n'
        '   "local_opportunity": "本地机会（结合行情，20-50字）",\n'
        '   "real_review": "联网核实到的真实评价摘要（待遇/加班/坑，30-60字，点明是网上口碑）",\n'
        '   "search_keywords": ["关键词1", "关键词2"],\n'
        '   "priority": "高/中/低",\n'
        '   "learn_first": ["上手前先补的1-3项"]}\n'
        "]}\n"
        "要务实、贴合本地，规避驻场/夜班/纯体力/销售拉人头类方向。最后只输出这个 JSON。"
    )


def call_ollama(model, system, user, timeout=360):
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        "stream": False,
        "format": "json",
        "options": {"temperature": 0.3},
    }
    r = requests.post(f"{OLLAMA_HOST}/api/chat", json=payload, timeout=timeout)
    r.raise_for_status()
    return r.json().get("message", {}).get("content", "")


def find_claude():
    for name in ("claude.cmd", "claude.exe", "claude"):
        p = shutil.which(name)
        if p:
            return p
    fallback = os.path.expandvars(r"%APPDATA%\npm\claude.cmd")
    return fallback if os.path.exists(fallback) else None


def call_claude(prompt, model="opus", timeout=240, allow_web=False):
    """通过 Claude Code 命令行(headless)调用 Claude，返回回答文本。
    allow_web=True 时放开联网工具，让 Claude 能搜真实评价。"""
    exe = find_claude()
    if not exe:
        raise RuntimeError("未找到 claude 命令行（Claude Code CLI）")
    cmd = [exe, "-p", "--output-format", "json", "--model", model]
    if allow_web:
        cmd += ["--allowedTools", "WebSearch", "WebFetch"]
    proc = subprocess.run(cmd, input=prompt, capture_output=True,
                          text=True, encoding="utf-8", errors="replace",
                          timeout=timeout)
    if proc.returncode != 0:
        raise RuntimeError(f"claude 调用失败：{(proc.stderr or proc.stdout or '')[:300]}")
    data = json.loads(proc.stdout)
    if data.get("is_error"):
        raise RuntimeError(f"claude 返回错误：{str(data.get('result', ''))[:200]}")
    return data.get("result", "")


def coerce(raw):
    obj = None
    try:
        obj = json.loads(raw)
    except Exception:
        m = re.search(r"\{.*\}", raw or "", re.S)
        if m:
            try:
                obj = json.loads(m.group(0))
            except Exception:
                obj = None
    if not isinstance(obj, dict):
        return []
    recs = obj.get("recommendations") or obj.get("directions") or []
    return recs if isinstance(recs, list) else []


def normalize(rec):
    def s(v):
        return str(v).strip() if v is not None else ""
    kws = rec.get("search_keywords") or rec.get("keywords") or []
    if isinstance(kws, str):
        kws = [k.strip() for k in re.split(r"[，,、/]", kws) if k.strip()]
    kws = [k for k in (s(k) for k in kws) if k][:4]
    learn = rec.get("learn_first") or []
    if isinstance(learn, str):
        learn = [x.strip() for x in re.split(r"[，,、/]", learn) if x.strip()]
    pr = s(rec.get("priority")) or "中"
    pr = "高" if "高" in pr else ("低" if "低" in pr else "中")
    return {
        "direction": s(rec.get("direction")) or s(rec.get("name")),
        "why_fit": s(rec.get("why_fit")) or s(rec.get("why")),
        "local_opportunity": s(rec.get("local_opportunity")) or s(rec.get("local")),
        "real_review": s(rec.get("real_review")) or s(rec.get("review")),
        "search_keywords": kws,
        "priority": pr,
        "learn_first": [s(x) for x in learn if s(x)][:3],
    }


def build_report(model, recs):
    order = {"高": 0, "中": 1, "低": 2}
    recs = sorted(recs, key=lambda r: order.get(r["priority"], 1))
    stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    lines = [
        "# 求职方向参谋（为你定制）",
        f"> 模型：{model}　|　生成：{stamp}　|　依据：你的画像 + 景德镇真实就业行情",
        "",
    ]
    for i, r in enumerate(recs, 1):
        lines.append(f"## {i}. {r['direction']}　【优先级：{r['priority']}】")
        if r["why_fit"]:
            lines.append(f"- **为什么适合你**：{r['why_fit']}")
        if r["local_opportunity"]:
            lines.append(f"- **本地机会**：{r['local_opportunity']}")
        if r["search_keywords"]:
            lines.append("- **去招聘网站搜这些词**：" + "　".join(f"`{k}`" for k in r["search_keywords"]))
        if r["learn_first"]:
            lines.append(f"- **上手前先补**：{('、'.join(r['learn_first']))}")
        lines.append("")
    all_kw = dedup_keywords(recs)
    lines.append("## 建议直接采集的搜索词（已汇总去重）")
    lines.append("、".join(all_kw))
    lines.append("")
    lines.append("> 用法：把上面这串词复制到工作台「岗位关键词」框，然后点「采集岗位」。")
    return "\n".join(lines)


def dedup_keywords(recs, limit=12):
    seen, out = set(), []
    for r in recs:
        for k in r.get("search_keywords", []):
            if k and k not in seen:
                seen.add(k)
                out.append(k)
    return out[:limit]


def main(argv=None):
    configure_console_encoding()
    ap = argparse.ArgumentParser(description="智能求职方向参谋")
    ap.add_argument("--backend", default="claude", choices=["claude", "ollama"],
                    help="用哪个大脑：claude=最智能(走订阅额度)；ollama=本地免费")
    ap.add_argument("--model", default="qwen2.5:14b", help="Ollama 模型名（ollama 后端用）")
    ap.add_argument("--claude-model", default="opus", help="Claude 模型（claude 后端用）")
    ap.add_argument("--top", type=int, default=6, help="推荐多少个方向")
    args = ap.parse_args(argv)

    profile = load_json(os.path.join(BASE_DIR, "match_profile.json"), {})
    career = load_json(os.path.join(BASE_DIR, "career_direction_profile.json"), {})
    market = load_text(os.path.join(BASE_DIR, "local_market_snapshot.md"),
                       "（暂无本地行情快照）")
    if not profile:
        print("  ✗ 缺少 match_profile.json。")
        return 1

    profile_brief = build_profile_brief(profile)
    directions_brief = build_directions_brief(career)
    user = build_user_prompt(profile_brief, market, directions_brief, args.top)

    backend = args.backend
    model = ""
    if backend == "claude":
        try:
            print(f"  方向参谋后端：Claude（{args.claude_model}）+ 联网核实评价，目标：推荐 {args.top} 个方向")
            print("  正在联网核实各方向真实评价，请耐心等（约 3-6 分钟）…\n")
            raw = call_claude(SYSTEM_PROMPT + "\n\n" + user, model=args.claude_model,
                              timeout=600, allow_web=True)
            model = f"Claude ({args.claude_model})"
        except Exception as e:
            print(f"  ⚠️ Claude 调用失败（{e}），自动回退本地 Ollama。")
            backend = "ollama"
    if backend == "ollama":
        m, inst = pick_model(args.model)
        if not inst:
            print(f"  ✗ 连不上本地 Ollama（{OLLAMA_HOST}）。请确认 Ollama 已启动。")
            return 2
        print(f"  方向参谋后端：本地 {m}，目标：推荐 {args.top} 个方向\n")
        try:
            raw = call_ollama(m, SYSTEM_PROMPT, user)
            model = m
        except Exception as e:
            print(f"  ✗ 调用模型失败：{e}")
            return 3

    recs = [normalize(r) for r in coerce(raw) if isinstance(r, dict)]
    recs = [r for r in recs if r["direction"]]
    if not recs:
        print("  ✗ 模型没有给出有效推荐，原始输出片段：")
        print("   ", (raw or "")[:300])
        return 4

    all_kw = dedup_keywords(recs)
    out_json = os.path.join(BASE_DIR, "recommended_directions.json")
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump({"生成时间": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                   "model": model, "recommendations": recs,
                   "all_keywords": all_kw}, f, ensure_ascii=False, indent=2)

    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_md = os.path.join(BASE_DIR, f"求职方向参谋_{stamp}.md")
    with open(out_md, "w", encoding="utf-8") as f:
        f.write(build_report(model, recs))

    order = {"高": 0, "中": 1, "低": 2}
    for r in sorted(recs, key=lambda x: order.get(x["priority"], 1)):
        print(f"    【{r['priority']}】{r['direction']}  →  搜：{('、'.join(r['search_keywords']))}")
    print(f"\n  ✅ 方向参谋完成，共 {len(recs)} 个方向。")
    print(f"     报告：{out_md}")
    print(f"     数据：{out_json}")
    print(f"  建议采集关键词：{('、'.join(all_kw))}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
