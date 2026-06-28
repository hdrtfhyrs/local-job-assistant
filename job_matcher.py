# -*- coding: utf-8 -*-
"""
岗位匹配引擎（规则粗筛 + 生成『AI精排任务』）

流程：
  1. 读取爬虫产出的岗位 Excel（默认自动找最新的 *岗位采集*.xlsx / *推荐*.xlsx）
  2. 用 match_profile.json + 已有方向/关键词库做规则打分、硬过滤、去重
  3. 把规则 Top-N 写成两份交接文件，放进 analysis_inbox/：
       - pending_jobs.json   机器可读，含规则分和稳定 job_id
       - AI分析任务.md    给AI（Gemini）读的指令 + 岗位清单
  4. 你在AI里打开 .md，让它按指令把结果写成 analysis_inbox/analysis_result.json
  5. 再运行 merge_analysis.py 合并成最终推荐表

只新增、不改动已有爬虫。
"""

import argparse
import datetime as dt
import glob
import hashlib
import json
import os
import re
import sys

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def configure_console_encoding():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def load_json(name, default=None):
    path = os.path.join(BASE_DIR, name)
    if not os.path.exists(path):
        return default if default is not None else {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"  ⚠️ 读取 {name} 失败：{e}")
        return default if default is not None else {}


def newest_job_excel():
    """自动挑选最近修改的岗位结果 Excel。"""
    patterns = [
        "*岗位采集*.xlsx",
        "*IT过渡岗位推荐*.xlsx",
        "*自定义岗位搜集*.xlsx",
        "*本科软件工程入门岗位*.xlsx",
        "*采集*.xlsx",
    ]
    excluded = ("投递", "方向", "精排", "规则打分", "口碑线索")
    files = []
    for p in patterns:
        files.extend(glob.glob(os.path.join(BASE_DIR, p)))
    files = [
        f for f in set(files)
        if not any(word in os.path.basename(f) for word in excluded)
    ]
    if not files:
        return None
    return max(files, key=os.path.getmtime)


def infer_platform_from_filename(path):
    """老结果表可能没有平台列，用文件名做一个温和兜底。"""
    name = os.path.basename(path or "").lower()
    if "boss" in name or "it过渡岗位推荐" in name:
        return "BOSS直聘"
    if "前程" in name or "51job" in name:
        return "前程无忧"
    if "智联" in name or "zhaopin" in name:
        return "智联招聘"
    if "猎聘" in name or "liepin" in name:
        return "猎聘"
    if "拉勾" in name or "lagou" in name:
        return "拉勾招聘"
    return ""


def parse_salary_floor_k(text):
    """把薪资文本解析成『月薪下限(K)』。无法解析或面议返回 None。
    正确处理 K/千/万、元/月、元/天、元/小时、年薪 等不同口径，
    避免把『50-200元/天』『4000元/月』误当成几十上百 K。"""
    s = str(text or "")
    if not s or "面议" in s:
        return None
    # 取第一个数字区间的下限（没有区间就取第一个数字）
    m = re.search(r"(\d+(?:\.\d+)?)\s*[-~至到]\s*\d+", s)
    if m:
        low = float(m.group(1))
    else:
        m1 = re.search(r"(\d+(?:\.\d+)?)", s)
        if not m1:
            return None
        low = float(m1.group(1))

    # 第一步：把数值统一换算成『元』
    if "万" in s:
        low_yuan = low * 10000.0
    elif "千" in s or "k" in s.lower():
        low_yuan = low * 1000.0
    elif "元" in s:
        low_yuan = low                       # 元/月、元/天、元/时 里的『元』
    else:
        low_yuan = low * 1000.0 if low < 1000 else low  # 纯数字：<1000 当 K，否则当元

    # 第二步：按计时/计日/计年口径换算成『月薪元』(21.75 工作日/月，每天 8 小时)
    if re.search(r"小时|时薪|/\s*时|元/?\s*时|/h", s, re.I):
        monthly = low_yuan * 21.75 * 8       # 时薪 -> 月薪
    elif re.search(r"天|日薪|/\s*日|元/?\s*日", s):
        monthly = low_yuan * 21.75           # 日薪 -> 月薪
    elif "年" in s:
        monthly = low_yuan / 12.0            # 年薪 -> 月薪
    else:
        monthly = low_yuan                   # 默认按月

    return round(monthly / 1000.0, 1)        # 元 -> K


def parse_active_recency(text):
    """从岗位卡片文本读 HR『活跃近况』，判断近期活跃 / 僵尸号 / 已结束。
    返回 (label, kind)，kind ∈ {'recent','mid','zombie','closed',None}。"""
    s = str(text or "")
    if not s:
        return "", None
    if re.search(r"已下线|已结束|停止招聘|招聘结束|已招满|职位关闭|暂停招聘", s):
        return "疑似已结束/招满", "closed"
    if re.search(r"刚刚活跃|今日活跃|今天活跃|当前在线|正在招聘|在线", s):
        return "近期活跃", "recent"
    m = re.search(r"(\d+)\s*日内活跃|(\d+)\s*天内活跃", s)
    if m:
        d = int(m.group(1) or m.group(2) or 0)
        return f"{d}日内活跃", ("recent" if d <= 14 else "mid")
    if re.search(r"本周活跃|一周内|7日内|三日内|3日内", s):
        return "本周活跃", "recent"
    if re.search(r"本月活跃|两周内|半月内|14日内", s):
        return "本月活跃", "mid"
    m2 = re.search(r"(\d+)\s*个?月(?:前|内)活跃", s)
    if m2:
        mo = int(m2.group(1) or 0)
        return f"{mo}个月前活跃", ("mid" if mo <= 1 else "zombie")
    if re.search(r"半年前|去年|年前活跃|\d+\s*周前活跃", s):
        return "较久未活跃", "zombie"
    return "", None


def classify_job(row, profile):
    """把岗位分成『成长型』(对口能成长)/『轻松型』(门槛低好上手)/『通用』，供推荐表标注。"""
    blob = f"{row.get('岗位名称','')} {row.get('原始摘要','')}"
    for kw in profile.get("growth_keywords", []):
        if kw in blob:
            return "成长型"
    for kw in profile.get("easy_keywords", []):
        if kw in blob:
            return "轻松型"
    return "通用"


def job_id(row):
    raw = f"{row.get('岗位名称','')}|{row.get('公司名称','')}|{row.get('工作地点','')}|{row.get('岗位链接','')}"
    return hashlib.md5(raw.encode("utf-8")).hexdigest()[:10]


def direction_index(career):
    """职业方向名/别名 -> 该方向综合先验分。"""
    idx = []
    for d in career.get("directions", []):
        prior = (d.get("entry_score", 0) + d.get("outlook_score", 0)
                 + d.get("local_fit_score", 0) + d.get("hours_score", 0)) / 4.0
        names = [d.get("name", "")] + d.get("aliases", [])
        idx.append((d.get("name", ""), [n for n in names if n], round(prior, 1)))
    return idx


def score_job(row, profile, career, beginner):
    """返回 (总分, 命中明细dict)。纯规则、可解释。"""
    name = str(row.get("岗位名称", ""))
    blob = f"{name} {row.get('原始摘要','')} {row.get('公司名称','')}"
    w = profile.get("weights", {})
    hits = {"技能命中": [], "方向": "", "薪资": "", "加减分": []}

    # —— 硬过滤 ——
    for kw in profile.get("hard_block", []):
        if kw in name:
            return None, {"淘汰原因": f"命中硬过滤词「{kw}」"}

    score = 0.0

    # —— 技能匹配（你最看重）——
    have = profile.get("skills", {}).get("have", [])
    grow = profile.get("skills", {}).get("want_to_grow", [])
    for sk in have:
        if sk.lower() in blob.lower():
            score += w.get("skill_have", 4)
            hits["技能命中"].append(sk)
    for sk in grow:
        if sk in blob:
            score += w.get("skill_grow", 6)
            hits["技能命中"].append(sk + "(方向)")

    # —— 职业方向先验 ——
    best_dir, best_prior = "", 0
    for dname, names, prior in direction_index(career):
        if any(n in name for n in names):
            if prior > best_prior:
                best_dir, best_prior = dname, prior
    if best_dir:
        score += best_prior * w.get("direction_prior", 0.4)
        hits["方向"] = f"{best_dir}(先验{best_prior})"

    # —— 薪资门槛（你最看重）——
    floor = parse_salary_floor_k(row.get("薪资范围", ""))
    want = profile.get("salary_floor_k", 0)
    ceiling = profile.get("salary_ceiling_k", 0)
    if floor is None:
        hits["薪资"] = "面议/未知（需面试核实）"
    elif floor < want * 0.8:
        return None, {"淘汰原因": f"薪资下限约{floor}K，明显低于底线{want}K"}
    elif ceiling and floor >= ceiling:
        score += w.get("salary_suspect", -12)
        hits["薪资"] = f"⚠下限约{floor}K 远高于本地同类({ceiling}K线)，谨防销售提成/夸大/挂羊头"
        hits["加减分"].append("高薪存疑-")
    elif floor >= want:
        score += w.get("salary_pass", 15)
        score += min(floor - want, 6) * w.get("salary_bonus_per_k", 1.5)
        hits["薪资"] = f"下限约{floor}K ≥ 底线{want}K ✓"
    else:
        hits["薪资"] = f"下限约{floor}K（略低于{want}K）"

    # —— 本地匹配 ——
    city = profile.get("city", "")
    loc = str(row.get("工作地点", ""))
    if city and (city in loc or city in blob):
        score += w.get("local_fit", 12)
        hits["加减分"].append("本地+")

    # —— 位置远近(离住处太远的往后排，不删)——
    lf = profile.get("location_filter", {})
    if any(a in loc for a in lf.get("veryfar_areas", [])):
        score += w.get("loc_veryfar", -25)
        hits["加减分"].append("⚠地点很远-")
    elif any(a in loc for a in lf.get("far_areas", [])):
        score += w.get("loc_far", -12)
        hits["加减分"].append("地点偏远-")
    elif any(a in loc for a in lf.get("near_areas", [])):
        score += w.get("loc_near", 6)
        hits["加减分"].append("离家近+")

    # —— 经验 ——
    exp = str(row.get("经验要求", "")) + " " + blob
    if any(e in exp for e in profile.get("experience_ok", [])):
        score += w.get("experience_ok", 8)
        hits["加减分"].append("经验友好+")
    if any(e in exp for e in profile.get("experience_penalty", [])):
        score += w.get("experience_penalty", -10)
        hits["加减分"].append("经验偏高-")

    # —— 出差/驻场（不接受时扣分）——
    if not profile.get("accept_relocation", False):
        if re.search(r"出差|驻场|外派", blob):
            score += w.get("avoid_keyword", -8)
            hits["加减分"].append("出差/驻场-")

    # —— 口碑词（复用 career_direction_profile）——
    for kw in career.get("positive_review_keywords", []):
        if kw in blob:
            score += w.get("positive_review", 3)
            hits["加减分"].append(f"{kw}+")
    for kw in career.get("negative_review_keywords", []):
        if kw in blob:
            score += w.get("negative_review", -5)
            hits["加减分"].append(f"{kw}-")

    # —— beginner avoid 关键词（名称里出现，温和扣分）——
    for kw in beginner.get("avoid_keywords", []):
        if kw in name and kw not in profile.get("hard_block", []):
            score += w.get("avoid_keyword", -8)
            hits["加减分"].append(f"{kw}-")

    # —— 重体力『软信号』(常藏在摘要里，名称重体力已被硬过滤) ——
    summary_text = str(row.get("原始摘要", ""))
    for kw in profile.get("physical_labor_keywords", []):
        if kw in summary_text:
            score += w.get("physical_labor", -10)
            hits["加减分"].append(f"重体力({kw})-")
            break

    # —— 能学 AI / 数字化 → 优先(用户明确想上班能学 AI)——
    for kw in profile.get("learn_ai_keywords", []):
        if kw in blob:
            score += w.get("learn_ai", 6)
            hits["加减分"].append(f"可学AI({kw})+")
            break

    # —— HR 活跃度(滤僵尸号 / 已招满未下架)——
    recency_text = str(row.get("活跃度", "")) or str(row.get("原始摘要", ""))
    act_label, act_kind = parse_active_recency(recency_text)
    if act_kind == "closed":
        return None, {"淘汰原因": "岗位疑似已结束/已招满但未下架"}
    if act_kind == "recent":
        score += w.get("active_recent", 5)
        hits["加减分"].append("近期活跃✓")
    elif act_kind == "zombie":
        score += w.get("inactive_zombie", -12)
        hits["加减分"].append(f"⚠HR久未活跃疑僵尸({act_label})")

    return round(score, 1), hits


def build_handoff_markdown(profile, top_rows):
    """生成给AI读的精排指令 + 岗位清单。"""
    pf_have = "、".join(profile.get("skills", {}).get("have", []))
    pf_grow = "、".join(profile.get("skills", {}).get("want_to_grow", []))
    floor = profile.get("salary_floor_k")
    city = profile.get("city")
    relo = "可接受出差/驻场" if profile.get("accept_relocation") else "不接受长期出差/驻场，优先本地"

    lines = []
    lines.append("# AI岗位精排任务")
    lines.append("")
    lines.append("> 给AI（Gemini）的指令。读完本文件后，请按文末「输出要求」把结果写到 "
                 "`analysis_inbox/analysis_result.json`。**不要改动本文件。**")
    lines.append("")
    lines.append("## 一、求职者画像（这就是『我本身』）")
    lines.append(f"- 城市：{city}；{relo}")
    lines.append(f"- 期望月薪下限：{floor}K")
    lines.append(f"- 已具备技能：{pf_have}")
    lines.append(f"- 想发展方向：{pf_grow}")
    lines.append("- 背景：普通本科软件工程，技术基础一般，找能入门、能积累、不重体力、"
                 "工作时长别太长、最好结合景德镇陶瓷/电商/文旅特色的过渡岗。")
    lines.append("")
    lines.append("## 二、你的任务")
    lines.append("下面每个岗位已有 Python 规则初筛分（rule_score，仅供参考）。请你**逐个**结合岗位"
                 "摘要和上面的画像，做语义精排，给出：")
    lines.append("1. `match_score`：0–100 的契合度（综合技能、薪资、本地、成长、踩坑风险）。")
    lines.append("2. `reason`：一句话推荐理由（为什么适合 TA）。")
    lines.append("3. `risk`：一句话风险/提醒（如『虽写实施但可能长期驻场』『薪资面议需面试确认』）。")
    lines.append("4. `verdict`：`强烈推荐` / `可投` / `观望` / `不建议` 之一。")
    lines.append("")
    lines.append("## 三、待精排岗位")
    lines.append("")
    for i, (jid, score, row, hits) in enumerate(top_rows, 1):
        lines.append(f"### {i}. [{jid}] {row.get('岗位名称','')}")
        lines.append(f"- 公司：{row.get('公司名称','') or '—'}　|　平台：{row.get('平台','')}")
        lines.append(f"- 薪资：{row.get('薪资范围','') or '面议'}　|　地点：{row.get('工作地点','') or '—'}"
                     f"　|　经验：{row.get('经验要求','') or '—'}　|　学历：{row.get('学历要求','') or '—'}")
        lines.append(f"- 规则分 rule_score：{score}　命中：{hits}")
        summary = str(row.get("原始摘要", "")).strip()
        if summary:
            lines.append(f"- 摘要：{summary[:300]}")
        lines.append(f"- 链接：{row.get('岗位链接','') or '—'}")
        lines.append("")
    lines.append("## 四、输出要求（写到 analysis_inbox/analysis_result.json）")
    lines.append("严格输出 JSON 数组，每条对应上面一个岗位的 job_id：")
    lines.append("```json")
    lines.append("[")
    lines.append('  {"job_id": "上面方括号里的ID", "match_score": 88, '
                 '"reason": "……", "risk": "……", "verdict": "可投"}')
    lines.append("]")
    lines.append("```")
    lines.append("只输出这一个 JSON 文件即可，其余分析过程可以放在对话里。")
    return "\n".join(lines)


def main(argv=None):
    configure_console_encoding()
    parser = argparse.ArgumentParser(description="岗位规则打分 + 生成AI精排任务")
    parser.add_argument("--input", help="岗位 Excel 路径；不填自动找最新的采集结果")
    parser.add_argument("--top", type=int, help="送给AI精排的条数（默认读 match_profile）")
    parser.add_argument("--all-output", action="store_true",
                        help="额外导出全部规则打分结果到 Excel（含被淘汰原因）")
    args = parser.parse_args(argv)

    profile = load_json("match_profile.json")
    career = load_json("career_direction_profile.json")
    beginner = load_json("beginner_job_profile.json")
    if not profile:
        print("  ✗ 缺少 match_profile.json，先创建画像。")
        return

    in_path = args.input or newest_job_excel()
    if not in_path or not os.path.exists(in_path):
        print("  ✗ 没找到岗位 Excel。先用爬虫采集，或用 --input 指定文件。")
        return
    source_platform = infer_platform_from_filename(in_path)
    print(f"  读取岗位文件：{os.path.basename(in_path)}")

    df = pd.read_excel(in_path).fillna("")
    print(f"  原始岗位 {len(df)} 条，开始规则打分…")

    scored, dropped, seen = [], 0, set()
    for _, r in df.iterrows():
        row = r.to_dict()
        jid = job_id(row)
        if jid in seen:
            continue
        seen.add(jid)
        score, hits = score_job(row, profile, career, beginner)
        if score is None:
            dropped += 1
            continue
        scored.append((jid, score, row, hits))

    scored.sort(key=lambda x: x[1], reverse=True)
    print(f"  规则筛选：保留 {len(scored)} 条，淘汰 {dropped} 条（硬过滤/薪资过低/重复）")

    top_n = args.top or profile.get("handoff", {}).get("top_n_for_llm", 25)
    top_rows = scored[:top_n]

    inbox = os.path.join(BASE_DIR, profile.get("handoff", {}).get("inbox_dir", "analysis_inbox"))
    os.makedirs(inbox, exist_ok=True)

    # 机器可读交接文件
    pending = [{
        "job_id": jid, "rule_score": score, "hits": hits,
        "岗位名称": row.get("岗位名称", ""), "公司名称": row.get("公司名称", ""),
        "平台": row.get("平台", "") or source_platform, "薪资范围": row.get("薪资范围", ""),
        "工作地点": row.get("工作地点", ""), "经验要求": row.get("经验要求", ""),
        "学历要求": row.get("学历要求", ""), "岗位链接": row.get("岗位链接", ""),
        "活跃度": row.get("活跃度", ""),
        "类型": classify_job(row, profile),
        "原始摘要": str(row.get("原始摘要", ""))[:500],
    } for jid, score, row, hits in top_rows]
    with open(os.path.join(inbox, "pending_jobs.json"), "w", encoding="utf-8") as f:
        json.dump({"生成时间": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                   "来源文件": os.path.basename(in_path), "岗位": pending},
                  f, ensure_ascii=False, indent=2)

    # 给AI读的指令文件
    md = build_handoff_markdown(profile, top_rows)
    md_path = os.path.join(inbox, "AI分析任务.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md)

    if args.all_output:
        allrows = []
        for jid, score, row, hits in scored:
            row2 = dict(row)
            row2["规则分"] = score
            row2["命中明细"] = json.dumps(hits, ensure_ascii=False)
            allrows.append(row2)
        out = os.path.join(BASE_DIR, "规则打分_全部.xlsx")
        pd.DataFrame(allrows).to_excel(out, index=False)
        print(f"  全部规则打分已导出：{os.path.basename(out)}")

    print(f"\n  ✅ 已生成待精排岗位（{len(top_rows)} 条）：")
    print(f"     {os.path.join(inbox, 'pending_jobs.json')}")
    print("  下一步：用「本地AI精排」(Claude 或本地 qwen) 自动精排，再生成推荐表。")


if __name__ == "__main__":
    main()
