# -*- coding: utf-8 -*-
"""
====================================================
  个性化引擎 profile_engine v1.0
====================================================

让"不同背景的人"都能用：根据用户自己的背景/目标，自动算出
  1) 该在招聘网站搜的岗位关键词
  2) 该自动跳过的"不对口/需别的专业资质"的岗位关键词（绝不含用户自己的方向）
  3) 一句话招呼语
结果存进 user_profile.json，给爬虫/投递/精排统一读取。

优先用本地大模型(qwen)生成；模型没开就用内置职业模板兜底。

用法：
  生成/更新画像（交互问几句）：  python profile_engine.py --setup
  用指定参数生成（脚本/GUI调用）：python profile_engine.py --city 景德镇 --edu 本科 \
        --background "本科软件工程,会装机网络" --targets "IT实施 技术支持 运维 电商运营 数据助理"
  只看当前画像：                 python profile_engine.py --show
"""

import argparse
import json
import os
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_PATH = os.path.join(BASE_DIR, "user_profile.json")

# 兜底用的职业大类库：方向 -> (搜索关键词, 该方向的岗位名特征词)
# off_fit = 用户没选的所有大类的特征词 之和（即"不对口的别投"）
CAREER_FAMILIES = {
    "IT软件": {
        "search": ["软件实施工程师", "技术支持工程师", "IT运维工程师", "测试工程师", "网络管理员"],
        "titles": ["软件", "实施", "技术支持", "运维", "网络", "测试", "程序员", "开发", "前端", "后端", "系统管理"],
    },
    "电商运营": {
        "search": ["电商运营助理", "跨境电商运营", "新媒体运营", "店铺运营", "内容运营"],
        "titles": ["电商", "运营", "新媒体", "店铺", "直通车", "选品", "内容运营", "社群"],
    },
    "数据": {
        "search": ["数据助理", "数据分析师", "数据标注", "数据专员"],
        "titles": ["数据", "标注", "分析师", "BI", "报表"],
    },
    "财务会计": {
        "search": ["会计", "出纳", "财务专员", "成本会计", "税务专员"],
        "titles": ["财务", "会计", "出纳", "审计", "税务", "成本核算"],
    },
    "设计": {
        "search": ["平面设计师", "UI设计师", "电商美工", "室内设计师"],
        "titles": ["设计师", "美工", "平面", "UI", "UX", "视觉", "修图"],
    },
    "销售": {
        "search": ["销售代表", "销售专员", "客户经理", "业务员"],
        "titles": ["销售", "业务员", "客户经理", "渠道", "招商", "电销"],
    },
    "行政人事": {
        "search": ["行政专员", "人事专员", "前台", "招聘专员"],
        "titles": ["行政", "人事", "HR", "前台", "招聘", "文员", "助理"],
    },
    "客服": {
        "search": ["客服专员", "在线客服", "客户成功专员"],
        "titles": ["客服", "客户成功", "售后", "呼叫"],
    },
    "技工电力": {
        "search": ["电工", "维修电工", "数控操作工", "机修"],
        "titles": ["电工", "焊工", "钳工", "车工", "铣工", "数控", "机修", "电气", "电力", "电站", "变电", "钣金", "锅炉", "水电"],
    },
    "医护": {
        "search": ["护士", "护理", "药剂师"],
        "titles": ["护士", "护理", "医生", "医师", "药剂", "药师", "检验", "麻醉"],
    },
    "教育": {
        "search": ["教师", "助教", "课程顾问"],
        "titles": ["教师", "老师", "幼师", "助教", "讲师", "教练"],
    },
    "餐饮": {
        "search": ["厨师", "面点师", "餐饮服务员"],
        "titles": ["厨师", "面点", "烘焙", "厨房", "服务员", "传菜"],
    },
    "建筑工程": {
        "search": ["施工员", "造价员", "预算员"],
        "titles": ["造价", "预算", "施工员", "测量", "测绘", "钢筋", "架子工", "瓦工", "木工"],
    },
    "司机物流": {
        "search": ["司机", "配送员", "仓管"],
        "titles": ["司机", "驾驶", "配送", "快递", "搬运", "装卸"],
    },
    "其它专业": {
        "search": [],
        "titles": ["律师", "法务", "翻译", "兽医", "美容", "美发", "理疗", "推拿", "按摩",
                   "主播", "直播", "保安", "保洁", "月嫂", "育儿", "中介", "置业顾问", "房产经纪"],
    },
}

DEFAULT_PROFILE = {
    "city": "景德镇",
    "education": "本科",
    "background": "本科软件工程毕业，技术基础一般，会装机/网络/基础编程",
    "targets": "IT实施 技术支持 运维 电商运营 数据助理",
    "avoid": "不接受长期出差/驻场",
    "salary_floor_k": 3.5,
    "generated": {},  # {search_keywords, off_fit_keywords, greeting}
}


def load_profile():
    if os.path.exists(PROFILE_PATH):
        try:
            with open(PROFILE_PATH, "r", encoding="utf-8") as f:
                p = json.load(f)
            merged = dict(DEFAULT_PROFILE)
            merged.update({k: v for k, v in p.items() if v is not None})
            return merged
        except Exception:
            pass
    return dict(DEFAULT_PROFILE)


def save_profile(profile):
    with open(PROFILE_PATH, "w", encoding="utf-8") as f:
        json.dump(profile, f, ensure_ascii=False, indent=2)


# ---------------- 用大模型生成 ----------------

def generate_via_llm(profile, model="qwen3:8b"):
    try:
        from local_ai_matcher import call_ollama, check_ollama
    except Exception as e:
        raise RuntimeError(f"无法加载本地大模型调用：{e}")
    if not check_ollama(model):
        raise RuntimeError("本地大模型(Ollama)没启动或没有该模型")

    system = (
        "你是求职助手的'个性化配置'生成器。根据用户的求职背景与目标，"
        "只输出一个 JSON，字段如下：\n"
        '{"search_keywords": ["在招聘网站要搜的岗位词,8-15个,贴合用户目标"],\n'
        ' "off_fit_keywords": ["要自动跳过的岗位名关键词,即需要别的专业证书/手艺、与用户方向无关的"],\n'
        ' "greeting": "给HR的一句话招呼语,50字内,体现用户背景与求职诚意"}\n'
        "铁律：off_fit_keywords 里【绝对不能】出现用户自己想做的方向相关词"
        "（例如用户做财务，就不能把'财务/会计'放进 off_fit）。"
        "off_fit 只放明显不对口、需要专门资质或手艺的行当。只输出 JSON，不要解释。"
    )
    user = (
        f"城市：{profile.get('city')}\n学历：{profile.get('education')}\n"
        f"背景：{profile.get('background')}\n想做的方向：{profile.get('targets')}\n"
        f"不接受：{profile.get('avoid')}\n薪资底线(K)：{profile.get('salary_floor_k')}"
    )
    raw = call_ollama(model, system, user, timeout=120)
    data = raw if isinstance(raw, dict) else json.loads(raw)
    sk = [str(x).strip() for x in data.get("search_keywords", []) if str(x).strip()]
    off = [str(x).strip() for x in data.get("off_fit_keywords", []) if str(x).strip()]
    greeting = str(data.get("greeting", "")).strip()
    if not sk:
        raise RuntimeError("大模型没给出搜索关键词")
    # 安全网：把用户目标里出现的词从 off_fit 中剔除，绝不自伤
    tgt = (profile.get("targets", "") + profile.get("background", "")).lower()
    off = [w for w in off if w.lower() not in tgt]
    return {"search_keywords": sk[:15], "off_fit_keywords": off, "greeting": greeting}


# ---------------- 规则兜底 ----------------

def _match_families(profile):
    """从用户目标/背景文本里，判断他属于哪些职业大类。"""
    text = (str(profile.get("targets", "")) + " " + str(profile.get("background", ""))).lower()
    chosen = set()
    for fam, info in CAREER_FAMILIES.items():
        for t in info["titles"] + info["search"]:
            if t.lower() in text:
                chosen.add(fam)
                break
    if not chosen:
        chosen.add("电商运营")  # 啥都没匹配上时给个最通用的入门方向兜底
    return chosen


def generate_via_template(profile):
    chosen = _match_families(profile)
    search, mine = [], set()
    for fam in chosen:
        search.extend(CAREER_FAMILIES[fam]["search"])
        mine.update(CAREER_FAMILIES[fam]["titles"])
    off = []
    for fam, info in CAREER_FAMILIES.items():
        if fam in chosen:
            continue
        for t in info["titles"]:
            if t not in mine and t not in off:
                off.append(t)
    edu = profile.get("education", "")
    greeting = (f"您好！我{edu}毕业，"
                f"对这个岗位很感兴趣，踏实肯学、希望长期稳定发展，期待和您进一步沟通，谢谢！")
    return {"search_keywords": search[:15], "off_fit_keywords": off, "greeting": greeting,
            "_families": sorted(chosen)}


def generate(profile, prefer_llm=True, model="qwen3:8b"):
    """生成个性化配置，优先大模型，失败用模板兜底。返回 (gen, 来源)。"""
    if prefer_llm:
        try:
            return generate_via_llm(profile, model), "大模型"
        except Exception as e:
            print(f"  （大模型生成失败，改用职业模板兜底：{e}）", flush=True)
    return generate_via_template(profile), "模板兜底"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--setup", action="store_true", help="交互式问几句生成画像")
    ap.add_argument("--show", action="store_true", help="只看当前画像")
    ap.add_argument("--no-llm", action="store_true", help="强制用模板，不调大模型")
    ap.add_argument("--city")
    ap.add_argument("--edu")
    ap.add_argument("--background")
    ap.add_argument("--targets")
    ap.add_argument("--avoid")
    ap.add_argument("--salary", type=float)
    args = ap.parse_args()

    profile = load_profile()
    if args.show:
        print(json.dumps(profile, ensure_ascii=False, indent=2))
        return

    if args.setup:
        def ask(label, cur):
            v = input(f"  {label}（回车保留：{cur}）：").strip()
            return v or cur
        print("===== 填一下你的求职资料（直接回车=保留默认）=====")
        profile["city"] = ask("你在哪个城市", profile["city"])
        profile["education"] = ask("最高学历", profile["education"])
        profile["background"] = ask("一句话背景(专业/会什么)", profile["background"])
        profile["targets"] = ask("想做哪些方向(空格隔开)", profile["targets"])
        profile["avoid"] = ask("不接受什么(如出差)", profile["avoid"])
        sf = ask("薪资底线(K,如3.5)", str(profile["salary_floor_k"]))
        try:
            profile["salary_floor_k"] = float(sf)
        except ValueError:
            pass
    else:
        for k, a in [("city", args.city), ("education", args.edu),
                     ("background", args.background), ("targets", args.targets),
                     ("avoid", args.avoid)]:
            if a:
                profile[k] = a
        if args.salary is not None:
            profile["salary_floor_k"] = args.salary

    print("  正在生成个性化配置……", flush=True)
    gen, source = generate(profile, prefer_llm=not args.no_llm)
    profile["generated"] = gen
    save_profile(profile)
    print(f"  ✅ 已生成（来源：{source}）并保存到 user_profile.json")
    print(f"  搜索关键词({len(gen['search_keywords'])})：{gen['search_keywords']}")
    print(f"  跳过不对口({len(gen['off_fit_keywords'])})：{gen['off_fit_keywords'][:20]}")
    print(f"  招呼语：{gen['greeting']}")
    if gen.get("_families"):
        print(f"  匹配到的职业大类：{gen['_families']}")


if __name__ == "__main__":
    main()
