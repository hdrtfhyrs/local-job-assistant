# -*- coding: utf-8 -*-
"""
AI 精排（本地 qwen 或 Claude，全自动，不再手动交接）——支持两种大脑

  --backend ollama  本地 qwen 逐条精排，免费离线（默认）
  --backend claude  调 Claude Code 命令行，一次性批量精排，质量最高、走订阅额度

读 analysis_inbox/pending_jobs.json → 写 analysis_inbox/analysis_result.json，
后续 merge_analysis.py 合并成最终推荐表。

运行：python local_ai_matcher.py [--backend ollama|claude] [--model qwen2.5:7b]
                                  [--claude-model sonnet] [--limit N] [--retries 2]
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
INBOX = os.path.join(BASE_DIR, "analysis_inbox")
OLLAMA_HOST = "http://localhost:11434"
DEFAULT_MODEL = "qwen3:8b"
VALID_VERDICTS = ["强烈推荐", "可投", "观望", "不建议"]

SYSTEM_PROMPT = (
    "你是一位资深求职顾问，专门帮『技术基础一般的普通本科毕业生』在小城市找靠谱的过渡岗位。"
    "你会结合求职者画像，对岗位做语义精排，只输出严格的 JSON。"
    "请特别警惕并给低分的坑：传销/拉人头（如『组建团队』『日入千元』『商家拓展』）、"
    "电销话术（『底薪+提成+流水』）、纯体力/户外/夜班倒班岗、计件或兼职（元/时、元/天）、"
    "职责不明的『高薪诚招助理』、需要特种资质的岗位（如变电/电工/高压）。"
    "【重点】务必结合『公司所属行业』判断岗位的真实要求：很多岗位名称看着对口(如『运维』『技术支持』『工程师』『助理』)，"
    "但在特定行业里其实需要求职者画像中【不具备】的专业技能/资质/对口专业——"
    "例如软件背景的人去做环保/电力/化工/机械等行业的『设备运维/技术支持』(实为现场设备/工艺岗)，"
    "或需要会计证/设计能力/医护资质/施工经验等。凡是岗位真实要求与求职者画像专业不对口的，"
    "即使名称沾边也要判『不建议』(最多『观望』)，并在 risk 里写明『行业/专业不对口：需要XX，求职者不具备』。"
    "薪资明显高于本地同类(如普通入门岗开十几K)的，按虚标/销售提成处理给低分，并在 risk 点明；"
    "若岗位信息显示 HR『久未活跃』(几个月前活跃、已招满未下架)，投了多半石沉大海，要降分并在 risk 提示『HR久未活跃，疑似僵尸号』；"
    "用户明确希望上班能接触/学习 AI 或数字化、能成长，这类岗位在靠谱前提下适当加分。"
    "真正契合的是：技能或发展方向对口、薪资达标且合理、本地不驻场、HR近期活跃、能入门能积累的岗位。"
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


def build_profile_brief(profile):
    have = "、".join(profile.get("skills", {}).get("have", []))
    grow = "、".join(profile.get("skills", {}).get("want_to_grow", []))
    city = profile.get("city", "")
    floor = profile.get("salary_floor_k")
    relo = ("不接受长期出差/驻场，优先本地"
            if not profile.get("accept_relocation") else "可接受出差/驻场")
    ceiling = profile.get("salary_ceiling_k")
    ceiling_note = (f"(本地入门岗正常3-6K，下限超过约{ceiling}K的要按高薪虚标存疑)" if ceiling else "")
    return (f"城市：{city}；{relo}；期望月薪下限：{floor}K{ceiling_note}；"
            f"已具备技能：{have}；想发展方向：{grow}；"
            f"背景：普通本科软件工程，技术基础一般，想找能入门、能积累、能学AI/数字化、"
            f"坚决不要重体力/夜班倒班/纯销售、工作时长别太长、"
            f"最好结合景德镇陶瓷/电商/文旅特色的过渡岗。")


def build_user_prompt(profile_brief, job):
    return (
        f"【求职者画像】\n{profile_brief}\n\n"
        f"【待评估岗位】\n"
        f"岗位名称：{job.get('岗位名称', '')}\n"
        f"公司：{job.get('公司名称', '')}\n"
        f"薪资：{job.get('薪资范围', '')}\n"
        f"地点：{job.get('工作地点', '')}\n"
        f"经验：{job.get('经验要求', '')}\n"
        f"学历：{job.get('学历要求', '')}\n"
        f"HR活跃近况：{job.get('活跃度', '') or '（未知）'}\n"
        f"摘要：{(str(job.get('原始摘要', '')) or '（无）')[:400]}\n\n"
        "请评估这个岗位与求职者的契合度，只输出一个 JSON 对象，字段如下：\n"
        '{"match_score": 0到100的整数, "reason": "一句话推荐理由(20-40字)", '
        '"risk": "一句话风险或提醒(20-40字)", "verdict": "四选一：强烈推荐/可投/观望/不建议"}\n'
        "重要：『元/天』『元/时』多为兼职或计件、真实月收入其实很低，务必按真实月薪判断；"
        "薪资明显虚高按销售提成/虚标存疑；HR久未活跃疑似僵尸号要降分并在 risk 点明。"
    )


def build_batch_prompt(profile_brief, jobs):
    lines = [f"【求职者画像】\n{profile_brief}", "", f"【待精排岗位，共 {len(jobs)} 个】"]
    for i, j in enumerate(jobs, 1):
        active = j.get("活跃度", "")
        summary = str(j.get("原始摘要", "")).strip()
        extra = (f"｜HR活跃:{active}" if active else "")
        lines.append(
            f"{i}. [job_id={j.get('job_id')}] {j.get('岗位名称','')}｜{j.get('公司名称','')}"
            f"｜薪资{j.get('薪资范围','')}｜{j.get('工作地点','')}"
            f"｜经验{j.get('经验要求','')}｜学历{j.get('学历要求','')}{extra}"
            + (f"\n   摘要:{summary[:140]}" if summary else ""))
    lines.append("")
    lines.append("请逐个精排，输出一个 JSON 数组，每个元素对应上面一个 job_id：")
    lines.append('[{"job_id":"对应ID","match_score":0到100整数,"reason":"一句话推荐理由",'
                 '"risk":"一句话风险提醒","verdict":"强烈推荐/可投/观望/不建议"}]')
    lines.append("注意：『元/天』『元/时』多为兼职计件、真实月薪很低，按真实月薪判断；"
                 "警惕传销/电销/驻场/纯体力/夜班/擦边岗，给低分；"
                 "薪资明显虚高(普通入门岗开十几K)按销售提成/虚标存疑给低分；"
                 "HR久未活跃(几个月前活跃/已招满未下架)疑似僵尸号要降分并在 risk 点明。"
                 "只输出 JSON 数组，不要多余文字。")
    return "\n".join(lines)


def call_ollama(model, system, user, timeout=120):
    # Qwen3 默认会"思考"(<think>...)，会拖慢并干扰JSON；用 /no_think 关掉
    if "qwen3" in model.lower():
        system = system + "\n/no_think"
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        "stream": False,
        "format": "json",
        "options": {"temperature": 0.2},
        "think": False,
    }
    r = requests.post(f"{OLLAMA_HOST}/api/chat", json=payload, timeout=timeout)
    if r.status_code == 400:
        # 老版本 Ollama 不认 think 字段，去掉重试
        payload.pop("think", None)
        r = requests.post(f"{OLLAMA_HOST}/api/chat", json=payload, timeout=timeout)
    r.raise_for_status()
    content = r.json().get("message", {}).get("content", "")
    # 万一仍带 <think>，剥掉
    if "</think>" in content:
        content = content.split("</think>", 1)[1].strip()
    return content


def find_claude():
    for name in ("claude.cmd", "claude.exe", "claude"):
        p = shutil.which(name)
        if p:
            return p
    fallback = os.path.expandvars(r"%APPDATA%\npm\claude.cmd")
    return fallback if os.path.exists(fallback) else None


def call_claude(prompt, model="sonnet", timeout=300):
    exe = find_claude()
    if not exe:
        raise RuntimeError("未找到 claude 命令行（Claude Code CLI）")
    cmd = [exe, "-p", "--output-format", "json", "--model", model]
    proc = subprocess.run(cmd, input=prompt, capture_output=True,
                          text=True, encoding="utf-8", errors="replace",
                          timeout=timeout)
    if proc.returncode != 0:
        raise RuntimeError(f"claude 调用失败：{(proc.stderr or proc.stdout or '')[:300]}")
    data = json.loads(proc.stdout)
    if data.get("is_error"):
        raise RuntimeError(f"claude 返回错误：{str(data.get('result', ''))[:200]}")
    return data.get("result", "")


def parse_verdict(v):
    v = str(v or "").strip()
    for vv in VALID_VERDICTS:
        if vv in v:
            return vv
    return "观望"


def normalize_one(jid, obj):
    try:
        score = int(float(obj.get("match_score", obj.get("score", 0))))
    except Exception:
        score = 0
    score = max(0, min(100, score))
    reason = (str(obj.get("reason", "")).strip()[:120]) or "（模型未给理由）"
    risk = (str(obj.get("risk", "")).strip()[:120]) or "（模型未给风险提醒）"
    verdict = parse_verdict(obj.get("verdict"))
    return {"job_id": jid, "match_score": score,
            "reason": reason, "risk": risk, "verdict": verdict}


def coerce_result(jid, raw):
    """解析单条模型输出（dict 字符串）。失败返回 None。"""
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
        return None
    return normalize_one(jid, obj)


def coerce_array(raw):
    """解析批量输出（JSON 数组）。返回 list。"""
    obj = None
    try:
        obj = json.loads(raw)
    except Exception:
        m = re.search(r"\[.*\]", raw or "", re.S)
        if m:
            try:
                obj = json.loads(m.group(0))
            except Exception:
                obj = None
    if isinstance(obj, dict):
        obj = obj.get("results") or obj.get("岗位") or obj.get("data") or []
    return obj if isinstance(obj, list) else []


def check_ollama(model):
    try:
        r = requests.get(f"{OLLAMA_HOST}/api/tags", timeout=5)
        r.raise_for_status()
        models = [m.get("name", "") for m in r.json().get("models", [])]
        if model not in models and not any(model.split(":")[0] in m for m in models):
            print(f"  ⚠️ 本地未找到模型 {model}，已装：{models}")
            print(f"     可先运行：ollama pull {model}")
        return True
    except Exception as e:
        print(f"  ✗ 连不上本地 Ollama（{OLLAMA_HOST}）。请确认 Ollama 已启动。详情：{e}")
        return False


def run_ollama_each(profile_brief, jobs, model, retries):
    print(f"  精排大脑：本地 {model}　逐条精排 {len(jobs)} 条")
    print(f"  （首条会稍慢，模型要先加载到内存，请耐心等）\n")
    results = []
    for i, job in enumerate(jobs, 1):
        jid = str(job.get("job_id"))
        name = str(job.get("岗位名称", ""))
        user = build_user_prompt(profile_brief, job)
        res = None
        last_err = ""
        for attempt in range(retries + 1):
            try:
                raw = call_ollama(model, SYSTEM_PROMPT, user)
                res = coerce_result(jid, raw)
                if res:
                    break
            except Exception as e:
                last_err = str(e)
        if not res:
            rs = job.get("rule_score", 0) or 0
            res = {"job_id": jid, "match_score": int(min(100, max(0, rs))),
                   "reason": "本地模型未能给出结果，暂按规则分排序，建议人工复核。",
                   "risk": f"AI精排失败({last_err[:40]})，请人工确认薪资口径与驻场情况。",
                   "verdict": "观望"}
        results.append(res)
        print(f"    [{i:>2}/{len(jobs)}] {res['verdict']:<4} {res['match_score']:>3}分  {name[:24]}")
    return results


def run_claude_batch(profile_brief, jobs, claude_model):
    print(f"  精排大脑：Claude（{claude_model}）　一次性批量 {len(jobs)} 条")
    print("  正在调用 Claude，请稍候（约 20-40 秒）…\n")
    prompt = SYSTEM_PROMPT + "\n\n" + build_batch_prompt(profile_brief, jobs)
    raw = call_claude(prompt, model=claude_model, timeout=300)
    arr = coerce_array(raw)
    by_id = {str(x.get("job_id")): x for x in arr if isinstance(x, dict)}
    results, missing = [], 0
    for job in jobs:
        jid = str(job.get("job_id"))
        a = by_id.get(jid)
        if isinstance(a, dict):
            res = normalize_one(jid, a)
        else:
            missing += 1
            rs = job.get("rule_score", 0) or 0
            res = {"job_id": jid, "match_score": int(min(100, max(0, rs))),
                   "reason": "Claude 本次未覆盖此条，暂按规则分排序。",
                   "risk": "建议人工复核或重跑精排。", "verdict": "观望"}
        results.append(res)
        print(f"    {res['verdict']:<4} {res['match_score']:>3}分  {str(job.get('岗位名称',''))[:24]}")
    if missing:
        print(f"\n  ⚠️ 有 {missing} 条 Claude 未覆盖，已用规则分兜底。")
    return results


def main(argv=None):
    configure_console_encoding()
    ap = argparse.ArgumentParser(description="AI 岗位精排（本地 qwen 或 Claude）")
    ap.add_argument("--backend", default="ollama", choices=["claude", "ollama"],
                    help="精排大脑：claude=最准(批量,走额度)；ollama=本地免费(默认)")
    ap.add_argument("--model", default=DEFAULT_MODEL, help="Ollama 模型名")
    ap.add_argument("--claude-model", default="sonnet", help="Claude 模型(claude 后端)")
    ap.add_argument("--limit", type=int, default=0, help="只精排前 N 条(调试用)，0=全部")
    ap.add_argument("--retries", type=int, default=2, help="单条失败重试次数(ollama)")
    args = ap.parse_args(argv)

    profile = load_json(os.path.join(BASE_DIR, "match_profile.json"), {})
    pending = load_json(os.path.join(INBOX, "pending_jobs.json"), {})
    jobs = pending.get("岗位", []) if isinstance(pending, dict) else []
    if not jobs:
        print("  ✗ 没有 analysis_inbox/pending_jobs.json，请先运行 job_matcher.py。")
        return 1
    if args.limit > 0:
        jobs = jobs[:args.limit]

    profile_brief = build_profile_brief(profile)

    results = None
    backend = args.backend
    if backend == "claude":
        try:
            results = run_claude_batch(profile_brief, jobs, args.claude_model)
        except Exception as e:
            print(f"  ⚠️ Claude 批量精排失败（{e}），自动回退本地 Ollama 逐条。")
            backend = "ollama"
    if backend == "ollama":
        if not check_ollama(args.model):
            return 2
        results = run_ollama_each(profile_brief, jobs, args.model, args.retries)

    os.makedirs(INBOX, exist_ok=True)
    out = os.path.join(INBOX, "analysis_result.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    print(f"\n  ✅ AI 精排完成（{len(results)} 条），已写入：")
    print(f"     {out}")
    print(f"  下一步：运行  python merge_analysis.py  生成最终推荐表。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
