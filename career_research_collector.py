import argparse
import json
import os
import sys
from urllib.parse import quote

from job_scraper_core import configure_console_encoding, connect_chrome, safe_filename, save_jobs_excel


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_PATH = os.path.join(BASE_DIR, "career_direction_profile.json")


def load_profile():
    with open(PROFILE_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def total_score(item):
    return (
        item["entry_score"]
        + item["hours_score"]
        + item["physical_score"]
        + item["outlook_score"]
        + item["local_fit_score"]
        - item["review_risk"]
    )


def make_search_links(direction, city):
    name = direction["name"]
    aliases = " OR ".join(direction["aliases"][:3])
    review_query = f'{name} 工作体验 加班 前景 网友评价 避坑'
    local_query = f'{city} {name} 招聘 前景'
    local_feature_query = f'{city} 陶瓷 文旅 电商 {name}'
    return {
        "岗位口碑搜索": f"https://www.bing.com/search?q={quote(review_query)}",
        "本地机会搜索": f"https://www.bing.com/search?q={quote(local_query)}",
        "当地特色结合": f"https://www.bing.com/search?q={quote(local_feature_query)}",
        "知乎/经验搜索": f"https://www.bing.com/search?q={quote(name + ' 知乎 经验 评价')}",
        "招聘关键词": "、".join([name] + direction["aliases"]),
        "组合查询词": f'{city} ({aliases}) 双休 五险一金 不加班',
    }


def build_rows(profile, top=None):
    city = profile.get("city", "景德镇")
    rows = []
    directions = sorted(profile["directions"], key=total_score, reverse=True)
    if top:
        directions = directions[:top]

    for rank, item in enumerate(directions, start=1):
        links = make_search_links(item, city)
        rows.append({
            "平台": "职业方向研究",
            "搜索城市": city,
            "搜索关键词": links["招聘关键词"],
            "岗位名称": f"{rank}. {item['name']}",
            "薪资范围": "",
            "公司名称": item["category"],
            "工作地点": city,
            "经验要求": "应届/入门可尝试",
            "学历要求": "大专/本科常见",
            "岗位链接": links["岗位口碑搜索"],
            "原始摘要": (
                f"推荐分={total_score(item)}；理由：{item['why']}；"
                f"注意：{item['watch_out']}；先学：{'、'.join(item['learn_first'])}；"
                f"本地机会：{links['本地机会搜索']}；特色结合：{links['当地特色结合']}；"
                f"知乎/经验：{links['知乎/经验搜索']}；组合查询词：{links['组合查询词']}"
            ),
            "采集时间": "",
        })
    return rows


def print_recommendations(profile, top=8):
    print("\n综合推荐职业方向")
    print("=" * 60)
    print(profile["target_person"])
    print("评分说明：" + profile["scoring"])
    for rank, item in enumerate(sorted(profile["directions"], key=total_score, reverse=True)[:top], start=1):
        print(f"\n{rank}. {item['name']}｜{item['category']}｜推荐分 {total_score(item)}")
        print(f"   推荐理由：{item['why']}")
        print(f"   注意避坑：{item['watch_out']}")
        print(f"   先补技能：{'、'.join(item['learn_first'])}")


def collect_search_snippets(page, query, limit=8):
    page.get(f"https://www.bing.com/search?q={quote(query)}")
    script = """
    (() => Array.from(document.querySelectorAll('li.b_algo, .b_ans, .b_entityTP, .b_top')).slice(0, 10).map((el) => {
      const title = (el.querySelector('h2')?.innerText || el.querySelector('a')?.innerText || '').trim();
      const link = el.querySelector('a[href]')?.href || '';
      const text = (el.innerText || '').replace(/\\s+/g, ' ').trim();
      return {title, link, text: text.slice(0, 500)};
    }).filter(x => x.title || x.text))();
    """
    try:
        return (page.run_js(script) or [])[:limit]
    except Exception:
        return []


def build_parser():
    parser = argparse.ArgumentParser(description="职业方向联网研究与口碑线索收集")
    parser.add_argument("--top", type=int, default=10, help="导出前 N 个推荐方向")
    parser.add_argument("--output", help="输出 Excel 文件名或路径")
    parser.add_argument("--print", action="store_true", help="在终端打印推荐方向")
    parser.add_argument("--dry-run", action="store_true", help="只打印/生成研究链接，不启动浏览器")
    parser.add_argument("--collect-snippets", action="store_true", help="打开搜索页并收集可见搜索摘要")
    parser.add_argument("--debug-port", type=int, default=9222, help="Chrome 调试端口")
    return parser


def main():
    configure_console_encoding()
    args = build_parser().parse_args()
    profile = load_profile()

    if args.print or args.dry_run:
        print_recommendations(profile, top=args.top)

    rows = build_rows(profile, top=args.top)

    if args.dry_run:
        return

    if args.collect_snippets and not args.dry_run:
        page = connect_chrome(args.debug_port)
        snippet_rows = []
        for row in rows:
            query = f"{profile['city']} {row['岗位名称']} 工作体验 加班 前景 网友评价"
            print(f"  联网查询：{query}")
            for snippet in collect_search_snippets(page, query):
                snippet_rows.append({
                    "平台": "联网口碑摘要",
                    "搜索城市": profile["city"],
                    "搜索关键词": query,
                    "岗位名称": row["岗位名称"],
                    "薪资范围": "",
                    "公司名称": snippet.get("title", ""),
                    "工作地点": profile["city"],
                    "经验要求": "",
                    "学历要求": "",
                    "岗位链接": snippet.get("link", ""),
                    "原始摘要": snippet.get("text", ""),
                    "采集时间": "",
                })
        rows.extend(snippet_rows)

    output = args.output or safe_filename(f"{profile['city']}_职业方向推荐与口碑线索.xlsx")
    path, count = save_jobs_excel(rows, output)
    print("\n已生成研究表：")
    print(path)
    print(f"共 {count} 行")


if __name__ == "__main__":
    main()
