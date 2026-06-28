import argparse
import json
import os

from multi_platform_scraper import PLATFORMS, run


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_PATH = os.path.join(BASE_DIR, "beginner_job_profile.json")


def load_profile():
    with open(PROFILE_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def build_parser():
    parser = argparse.ArgumentParser(description="普通软件工程本科入门岗位采集器")
    parser.add_argument("--platform", choices=sorted(PLATFORMS), default="51job", help="招聘平台，默认前程无忧")
    parser.add_argument("--city", help="搜索城市，不填则使用画像默认城市")
    parser.add_argument("--keywords", help="自定义关键词；不填则使用入门岗位画像关键词")
    parser.add_argument("--pages", type=int, default=1, help="最多采集页数")
    parser.add_argument("--scrolls", type=int, default=4, help="每页滚动采集次数")
    parser.add_argument("--output", help="输出 xlsx 文件名或绝对路径")
    parser.add_argument("--manual", action="store_true", help="人工搜索到结果页后再采集当前页")
    parser.add_argument("--wait-before-collect", type=int, default=0, help="打开页面后等待 N 秒再采集，便于人工登录/验证/搜索")
    parser.add_argument("--dry-run", action="store_true", help="只打印计划，不启动浏览器")
    parser.add_argument("--yes", action="store_true", help="跳过人工确认")
    parser.add_argument("--debug-port", type=int, default=9222, help="Chrome 远程调试端口")
    parser.add_argument("--fast", action="store_true", help="快速模式：缩短防爬等待，更快但更易触发验证")
    parser.add_argument("--show-profile", action="store_true", help="打印当前求职画像和推荐方向")
    return parser


def print_profile(profile):
    print("\n普通软件工程本科入门求职画像")
    print("=" * 55)
    print(profile["profile"])
    print("\n优先方向：")
    for item in profile["best_first_targets"]:
        print(f"- {item['title']}（{item['score']}分）：{item['why']}")
    print("\n投递关键词：")
    print("、".join(profile["search_keywords"]))
    print("\n先少投或慎投：")
    print("、".join(profile["avoid_keywords"]))


def main():
    profile = load_profile()
    args = build_parser().parse_args()

    if args.show_profile:
        print_profile(profile)
        if args.dry_run:
            return

    city = args.city or profile.get("city_default") or "景德镇"
    keywords = args.keywords or ",".join(profile["search_keywords"])
    platform_name = PLATFORMS[args.platform]["name"]
    output = args.output or f"{platform_name}_{city}_本科软件工程入门岗位.xlsx"

    forwarded = [
        "--platform", args.platform,
        "--city", city,
        "--keywords", keywords,
        "--pages", str(max(1, args.pages)),
        "--scrolls", str(max(1, args.scrolls)),
        "--output", output,
        "--debug-port", str(args.debug_port),
        "--wait-before-collect", str(max(0, args.wait_before_collect)),
    ]
    if args.manual:
        forwarded.append("--manual")
    if args.fast:
        forwarded.append("--fast")
    if args.dry_run:
        forwarded.append("--dry-run")
    if args.yes:
        forwarded.append("--yes")

    run(argv=forwarded)


if __name__ == "__main__":
    main()
