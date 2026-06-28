import argparse
import sys

from job_scraper_core import (
    build_search_url,
    collect_jobs,
    collect_jobs_via_api,
    configure_console_encoding,
    connect_chrome,
    safe_filename,
    save_jobs_excel,
    set_pace,
    split_keywords,
)


PLATFORMS = {
    "51job": {
        "name": "前程无忧",
        "home_url": "https://www.51job.com/",
        "search_url_template": "https://we.51job.com/pc/search?keyword={keyword}&searchType=2&sortType=0",
        "link_patterns": ["jobs.51job.com", "we.51job.com/pc/search/job"],
        # 前程无忧搜索接口（不同版本走 search-pc / search-result）
        "api_patterns": ["search-pc", "/job/search", "/open/noauth/search", "searchType"],
        "detail_url_template": "https://jobs.51job.com/all/{id}.html",
        "manual_note": "如页面没有自动进入结果页，请在页面内搜索岗位并切换城市后继续。",
    },
    "zhaopin": {
        "name": "智联招聘",
        "home_url": "https://www.zhaopin.com/",
        "search_url_template": "https://www.zhaopin.com/sou/?kw={keyword}",
        "link_patterns": ["jobs.zhaopin.com", "jobdetail", "sou.zhaopin.com"],
        # 智联搜索接口（fe-api / i.zhaopin）
        "api_patterns": ["c/i/sou", "search/positions", "/c/i/", "fe-api.zhaopin.com", "/api/sou"],
        "detail_url_template": "https://jobs.zhaopin.com/{id}.htm",
        "manual_note": "智联页面结构变化较频繁，建议登录后手动完成搜索再继续采集。",
    },
    "liepin": {
        "name": "猎聘",
        "home_url": "https://www.liepin.com/zhaogongzuo/",
        "search_url_template": "https://www.liepin.com/zhaogongzuo/?key={keyword}",
        "link_patterns": ["job.liepin.com", "liepin.com/job", "liepin.com/a/"],
        # 猎聘搜索接口
        "api_patterns": ["pc-search-job", "searchfront", "/api/com.liepin", "search-job"],
        "detail_url_template": "https://www.liepin.com/job/{id}.shtml",
        "manual_note": "猎聘可能按账号状态展示不同页面，若未出现岗位列表，请手动搜索后继续。",
    },
    "lagou": {
        "name": "拉勾招聘",
        "home_url": "https://www.lagou.com/wn/",
        "search_url_template": "https://www.lagou.com/wn/jobs?kd={keyword}",
        "link_patterns": ["lagou.com/wn/jobs", "lagou.com/jobs"],
        # 拉勾搜索接口
        "api_patterns": ["positionAjax", "/jobs/list", "/jobs/positionAjax.json", "/v1/entry/search"],
        "detail_url_template": "https://www.lagou.com/wn/jobs/{id}.html",
        "manual_note": "拉勾常出现访问验证，请人工完成验证和搜索后继续；脚本不会绕过验证码。",
    },
}


def build_parser(default_platform=None):
    parser = argparse.ArgumentParser(description="多平台招聘岗位可见页采集器")
    if default_platform:
        parser.set_defaults(platform=default_platform)
    else:
        parser.add_argument("--platform", choices=sorted(PLATFORMS), required=True, help="招聘平台")
    parser.add_argument("--city", default="景德镇", help="搜索城市，仅用于记录和辅助构造搜索页")
    parser.add_argument("--keywords", default="软件实施工程师,IT运维工程师", help="岗位关键词，多个用逗号分隔")
    parser.add_argument("--pages", type=int, default=1, help="最多采集页数")
    parser.add_argument("--scrolls", type=int, default=4, help="每页滚动采集次数")
    parser.add_argument("--output", help="输出 xlsx 文件名或绝对路径")
    parser.add_argument("--manual", action="store_true", help="只打开平台首页，人工搜索后再采集当前页")
    parser.add_argument("--wait-before-collect", type=int, default=0, help="打开页面后等待 N 秒再采集，便于人工登录/验证/搜索")
    parser.add_argument("--dry-run", action="store_true", help="只打印计划，不启动浏览器")
    parser.add_argument("--yes", action="store_true", help="跳过人工确认，适合页面已经准备好的调试浏览器")
    parser.add_argument("--debug-port", type=int, default=9222, help="Chrome 远程调试端口")
    parser.add_argument("--fast", action="store_true", help="快速模式：缩短防爬等待，更快但更易触发验证")
    return parser


def run(default_platform=None, argv=None):
    configure_console_encoding()
    parser = build_parser(default_platform)
    args = parser.parse_args(argv)
    config = PLATFORMS[args.platform]
    keywords = split_keywords(args.keywords) or [""]
    interactive = (not args.yes) and sys.stdin.isatty()
    output = args.output or safe_filename(f"{config['name']}_{args.city}_岗位采集.xlsx")
    set_pace(args.fast)

    print("\n" + "=" * 55)
    print(f"  {config['name']} 岗位采集器")
    print("=" * 55)
    print(f"  城市：{args.city}")
    print(f"  关键词：{', '.join(keywords)}")
    print(f"  页数：{max(1, args.pages)}")
    print(f"  每页滚动次数：{max(1, args.scrolls)}")
    print(f"  输出：{output}")
    print(f"  模式：{'人工搜索当前页' if args.manual else '尝试打开搜索页后采集'}")
    print(f"  运行节奏：{'⚡ 快速（防爬等待已缩短）' if args.fast else '常规'}")
    print("  说明：" + config["manual_note"])

    if args.dry_run:
        print("  dry-run 完成：未启动浏览器、未采集。")
        return

    page = connect_chrome(args.debug_port)
    all_rows = []

    for index, keyword in enumerate(keywords, start=1):
        print("\n" + "-" * 55)
        print(f"  [{index}/{len(keywords)}] 关键词：{keyword or '当前页'}")
        url = config["home_url"] if args.manual else build_search_url(config, keyword, args.city)
        print(f"  打开页面：{url}")
        page.get(url)

        if args.wait_before_collect > 0:
            print(f"  等待 {args.wait_before_collect} 秒，请在浏览器中完成登录、验证、城市切换或搜索...")
            import time
            time.sleep(args.wait_before_collect)

        if interactive and not args.yes:
            print("  请在浏览器里完成登录、验证、城市切换或搜索。")
            try:
                input("  >>> 页面出现岗位列表后按回车继续采集：")
            except (EOFError, OSError):
                # 没有可交互的控制台（如从 GUI 后台启动）：不阻塞，直接继续
                print("  （无控制台输入，跳过回车确认，直接开始采集）")

        rows = None
        if config.get("api_patterns"):
            print("  优先尝试拦截平台搜索接口（更准更全）...")
            rows = collect_jobs_via_api(
                page,
                config,
                keyword=keyword,
                city=args.city,
                pages=max(1, args.pages),
                scrolls=max(1, args.scrolls),
                manual=args.manual,
            )
        if rows is None:
            print("  未捕获到接口数据，改用页面可见链接采集...")
            rows = collect_jobs(
                page,
                config,
                keyword=keyword,
                city=args.city,
                pages=max(1, args.pages),
                scrolls=max(1, args.scrolls),
            )
        print(f"  本关键词采集到 {len(rows)} 条")
        all_rows.extend(rows)

    result_path, count = save_jobs_excel(all_rows, output)
    print("\n" + "=" * 55)
    if result_path:
        print(f"  完成：共保存 {count} 条岗位")
        print(f"  文件：{result_path}")
    else:
        print("  未采集到岗位。建议使用 --manual，登录并手动搜索到结果页后再继续。")
    print("=" * 55)


if __name__ == "__main__":
    run()
