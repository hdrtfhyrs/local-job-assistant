# -*- coding: utf-8 -*-
r"""
====================================================
  求职助手 · 一条龙全自动 v1.0
====================================================

无人值守跑完整条流程，不用点任何按钮：
  确保调试Chrome在跑且已登录 → 爬岗位 → 规则筛选 → 本地AI精排
  → 生成推荐表 → 生成网页视图 → 自动投递 → 检查回信

用法：
  双击桌面「求职全自动.bat」，或：
  venv312\Scripts\python.exe auto_run_all.py

注意：
  - 必须已经在调试Chrome里登录过BOSS（登录态会保存在Chrome配置里）。
    若登录过期，脚本会提示扫码登录后停下，不会瞎跑。
  - 本机Ollama(qwen2.5)要常驻，AI精排才用得上。
  - 配置在 自动运行配置.json，可改城市/关键词/每页数/每日投递上限。
"""

import datetime as dt
import json
import os
import subprocess
import sys
import time

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

try:
    from notify import notify
except Exception:
    def notify(title, msg):
        pass
PY = sys.executable  # 由 venv312 的 python 启动本脚本，子进程沿用
CONFIG_PATH = os.path.join(BASE_DIR, "自动运行配置.json")
LOG_PATH = os.path.join(BASE_DIR, "自动运行日志.txt")
DEBUG_PORT = 9222

DEFAULT_CONFIG = {
    "city": "景德镇",
    # 不限行业、有前景就投：覆盖本地量大且有前景的入门方向
    "keywords": [
        "软件实施工程师", "技术支持工程师", "IT运维工程师",
        "陶瓷电商运营助理", "跨境电商运营", "新媒体运营",
        "数据助理", "数据标注", "客户成功专员", "产品助理",
    ],
    "platforms": "boss,51job,zhaopin",  # 抓哪些平台(可加 liepin,lagou)
    "max_pages": 3,         # 每个关键词翻几页（页多=岗位多但慢，也更易触发验证）
    "daily_apply_cap": 200, # 今天最多自动投多少（设高=投到BOSS拦/岗位投完为止）
    "do_apply": True,       # 是否自动投递（False=只爬+出表+出网页，不投）
}


def now():
    return dt.datetime.now().strftime("%H:%M:%S")


def log(msg):
    line = f"[{now()}] {msg}"
    print(line, flush=True)
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def load_config():
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg.update({k: v for k, v in json.load(f).items() if v is not None})
        except Exception as e:
            log(f"（配置读取失败，用默认：{e}）")
    else:
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
            log(f"已生成默认配置：{os.path.basename(CONFIG_PATH)}（以后可改这个文件）")
        except Exception:
            pass
    return cfg


# ---------------- Chrome / 登录 ----------------

def find_chrome():
    cands = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.join(os.environ.get("LocalAppData", ""), r"Google\Chrome\Application\chrome.exe"),
    ]
    for p in cands:
        if p and os.path.exists(p):
            return p
    return None


def connect_chrome():
    from DrissionPage import ChromiumPage, ChromiumOptions
    co = ChromiumOptions()
    co.set_local_port(DEBUG_PORT)
    return ChromiumPage(co)


def ensure_chrome_and_login():
    """确保调试Chrome在跑且已登录BOSS。返回 page 或 None。"""
    # 先试直接连
    page = None
    try:
        page = connect_chrome()
        _ = page.url
        log("已连接到调试 Chrome。")
    except Exception:
        log("没检测到调试 Chrome，正在启动……")
        chrome = find_chrome()
        if not chrome:
            log("❌ 找不到 chrome.exe，请先安装 Chrome。")
            return None
        profile = os.path.join(os.environ.get("TEMP", BASE_DIR), "boss_scraper_debug")
        try:
            subprocess.Popen([
                chrome, f"--remote-debugging-port={DEBUG_PORT}",
                f"--user-data-dir={profile}", "--no-first-run",
                "--no-default-browser-check", "https://www.zhipin.com/",
            ])
        except Exception as e:
            log(f"❌ 启动 Chrome 失败：{e}")
            return None
        time.sleep(8)
        try:
            page = connect_chrome()
        except Exception as e:
            log(f"❌ 连接新 Chrome 失败：{e}")
            return None

    # 校验登录：访问消息页，看是否被踢去登录
    try:
        tab = page.new_tab("https://www.zhipin.com/web/geek/chat")
        time.sleep(5)
        url = tab.url or ""
        try:
            tab.close()
        except Exception:
            pass
        if "/web/geek/chat" in url and "login" not in url:
            log("✅ BOSS 已登录。")
            return page
        log("⚠️ BOSS 未登录或登录已过期。请在弹出的 Chrome 里扫码登录后，重新运行本流程。")
        return None
    except Exception as e:
        log(f"⚠️ 登录校验出错：{e}")
        return None


# ---------------- 步骤执行 ----------------

def run_step(title, args):
    log("-" * 50)
    log(f"▶ {title}")
    try:
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        proc = subprocess.Popen(
            [PY] + args, cwd=BASE_DIR,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace", env=env,
        )
        # 子步骤输出也写进日志文件——后台定时跑时报错才查得到
        for line in proc.stdout:
            line = line.rstrip()
            if line:
                log("    " + line)
        code = proc.wait()
        if code == 0:
            log(f"  {title} 完成 ✅")
        else:
            log(f"  ⚠️ {title} 失败（退出码 {code}），具体报错看上面几行")
        return code == 0
    except Exception as e:
        log(f"  {title} 失败：{e}")
        return False


def main():
    log("=" * 50)
    log("求职助手 · 一条龙全自动 开始")
    log("=" * 50)
    cfg = load_config()
    city = cfg["city"]
    keywords = ",".join(cfg["keywords"]) if isinstance(cfg["keywords"], list) else str(cfg["keywords"])
    pages = str(cfg["max_pages"])

    # 个性化画像优先：用 user_profile.json 的城市 + 自动生成的搜索关键词
    prof_path = os.path.join(BASE_DIR, "user_profile.json")
    if os.path.exists(prof_path):
        try:
            with open(prof_path, "r", encoding="utf-8") as f:
                prof = json.load(f)
            if prof.get("city"):
                city = prof["city"]
            sk = (prof.get("generated") or {}).get("search_keywords")
            if sk:
                keywords = ",".join(sk)
            log(f"已按个性化画像 user_profile.json 设定：城市={city}、关键词{len(sk or [])}个")
        except Exception as e:
            log(f"（读 user_profile.json 失败，用默认配置：{e}）")

    # 0. Chrome + 登录
    page = ensure_chrome_and_login()
    if page is None:
        log("流程中止（Chrome 未就绪 / 未登录）。")
        notify("求职助手：BOSS 登录过期", "今天的自动流程停了。去桌面那个调试 Chrome 扫码登录，再重跑一次「求职全自动」。")
        return

    results = []

    def step(title, args):
        results.append((title, run_step(title, args)))

    # 1. 爬岗位（多平台一起抓，合并成一张全平台表）
    platforms = cfg.get("platforms", "boss,51job,zhaopin")
    if isinstance(platforms, list):
        platforms = ",".join(platforms)
    step("① 全平台爬取岗位", [
        "全平台抓取.py", "--city", city, "--keywords", keywords,
        "--platforms", platforms, "--pages", pages,
    ])

    # 2. 规则筛选 → 生成待精排候选
    step("② 规则筛选", ["job_matcher.py", "--top", "25"])

    # 3. 本地 AI 精排（Ollama）
    step("③ 本地AI精排", ["local_ai_matcher.py", "--backend", "ollama"])

    # 4. 生成最终推荐表
    step("④ 生成推荐表", ["merge_analysis.py"])

    # 5. 生成网页视图
    step("⑤ 生成网页视图", ["build_web.py"])

    # 6. 自动投递
    if cfg.get("do_apply", True):
        step("⑥ 自动投递", ["auto_apply.py", "--max", str(cfg["daily_apply_cap"])])
    else:
        log("（配置 do_apply=false，跳过自动投递）")

    # 7. 检查回信
    step("⑦ 检查回信", ["reply_monitor.py"])

    failed = [t for t, ok in results if not ok]
    log("=" * 50)
    if failed:
        log(f"⚠️ 流程跑完，但有 {len(failed)} 步失败：{'、'.join(failed)}")
        log("   失败原因已记录在本日志里，翻到对应步骤即可看到报错。")
        notify("求职助手：今天有步骤失败", "、".join(failed) + "。详情看 自动运行日志.txt")
    else:
        log("✅ 一条龙全自动 全部跑完！")
        notify("求职助手：今天跑完了 ✅", "推荐网页和回信记录都已更新。")
    log("   推荐网页：桌面\\求职推荐_网页版.html")
    log("   回信记录：回信记录.xlsx")
    log("=" * 50)


if __name__ == "__main__":
    main()
