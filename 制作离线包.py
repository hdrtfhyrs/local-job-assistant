r"""
制作"离线一键装包"（开发者用，VPN开着跑一次）。把 app + Python安装器 + 依赖wheel
 + Ollama安装器 + qwen3:8b模型 全打进一个文件夹，用户拷过去双击 一键安装.bat 即可，
全程不用联网下载大文件。默认输出到 D:\求职助手离线包（C盘空间小）。
用法：python 制作离线包.py  [输出目录]
"""
import json
import os
import shutil
import subprocess
import sys
import urllib.request

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

BASE = os.path.dirname(os.path.abspath(__file__))
OUT_ROOT = sys.argv[1] if len(sys.argv) > 1 else r"D:\求职助手离线包"
PKG = os.path.join(OUT_ROOT, "求职助手")
BUNDLE = os.path.join(PKG, "bundle")
OLLAMA_MODELS = os.path.join(os.path.expanduser("~"), ".ollama", "models")
MODEL_TAG = ("qwen3", "8b")
PY_URL = "https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe"
OLLAMA_URL = "https://ollama.com/download/OllamaSetup.exe"

CODE = [
    "job_assistant_app.py", "boss_scraper.py", "job_scraper_core.py",
    "multi_platform_scraper.py", "scrape_beginner_jobs.py", "job_matcher.py",
    "local_ai_matcher.py", "merge_analysis.py", "build_web.py", "auto_apply.py",
    "reply_monitor.py", "auto_run_all.py", "profile_engine.py", "career_advisor.py",
    "career_research_collector.py", "application_package.py",
]
DATA = ["requirements.txt", "cities.json", "city_dict.json", "recommended_directions.json",
        "local_niche_channels.json", "match_profile.json", "beginner_job_profile.json",
        "career_direction_profile.json"]
INSTALLER = ["install.py", "一键安装.bat"]
NEVER = {"cookie.txt", "user_profile.json", "投递记录.json", "回信快照.json",
         "投递配置.json", "自动运行配置.json"}

LAUNCHER = "@echo off\nchcp 65001 >nul\ncd /d \"%~dp0\"\nif exist \"runtime\\Scripts\\pythonw.exe\" (\n  start \"\" \"runtime\\Scripts\\pythonw.exe\" \"job_assistant_app.py\"\n) else (\n  echo [!] Please run 一键安装.bat first.\n  pause\n)\n"
AUTO = "@echo off\nchcp 65001 >nul\ncd /d \"%~dp0\"\n\"runtime\\Scripts\\python.exe\" auto_run_all.py\npause\n"
README = ("求职助手 · 离线版 使用说明\n========================\n\n"
          "【安装】双击 一键安装.bat —— 自动装好环境、AI模型(已自带,免下载)、检查Chrome、建桌面图标。\n"
          "        (第一次几分钟; 没装Chrome会帮你打开下载页。)\n\n"
          "【使用】双击桌面【求职助手】→ 先点【🧑 我的资料】填情况 → 点【🌐 打开网站·登录BOSS】扫码 → 用。\n"
          "        想全自动: 双击 求职全自动.bat。\n\n"
          "【提示】投递用你自己的BOSS账号，请遵守平台规则、注意节奏。投递时别去操作弹出的那个Chrome。\n")


def dl(url, dst, label):
    if os.path.exists(dst) and os.path.getsize(dst) > 1024:
        print(f"  [跳过] {label} 已存在 ({os.path.getsize(dst)/1048576:.0f}MB)")
        return True
    print(f"  下载 {label} …… {url}")
    try:
        def hook(b, bs, total):
            if total > 0:
                pct = min(100, b * bs * 100 // total)
                if b % 200 == 0:
                    print(f"    {pct}%", end="\r", flush=True)
        urllib.request.urlretrieve(url, dst, hook)
        print(f"  ✅ {label} 完成 ({os.path.getsize(dst)/1048576:.0f}MB)")
        return True
    except Exception as e:
        print(f"  ❌ {label} 下载失败：{e}")
        return False


def copy_model():
    print("== 拷贝 AI 模型 (qwen3:8b, ~5GB, 慢) ==")
    man_src = os.path.join(OLLAMA_MODELS, "manifests", "registry.ollama.ai", "library", *MODEL_TAG)
    if not os.path.exists(man_src):
        print(f"  ❌ 没找到模型清单：{man_src}（先 ollama pull qwen3:8b）")
        return False
    man = json.load(open(man_src, encoding="utf-8"))
    digests = [man["config"]["digest"]] + [l["digest"] for l in man["layers"]]
    # 清单
    man_dst = os.path.join(BUNDLE, "models", "manifests", "registry.ollama.ai", "library", *MODEL_TAG)
    os.makedirs(os.path.dirname(man_dst), exist_ok=True)
    shutil.copy2(man_src, man_dst)
    # 数据块
    bdst = os.path.join(BUNDLE, "models", "blobs")
    os.makedirs(bdst, exist_ok=True)
    for dg in digests:
        fn = dg.replace(":", "-")
        s = os.path.join(OLLAMA_MODELS, "blobs", fn)
        d = os.path.join(bdst, fn)
        if os.path.exists(d) and os.path.getsize(d) == os.path.getsize(s):
            print(f"  [跳过] {fn[:20]} 已拷")
            continue
        print(f"  拷 {fn[:20]}… ({os.path.getsize(s)/1048576:.0f}MB)")
        shutil.copy2(s, d)
    print("  ✅ 模型拷贝完成")
    return True


def main():
    print("=" * 56)
    print(f"  制作离线包 → {PKG}")
    print("=" * 56)
    os.makedirs(BUNDLE, exist_ok=True)

    # 1. app 代码
    print("== 拷贝程序代码 ==")
    n = 0
    for name in CODE + DATA + INSTALLER:
        if name in NEVER:
            continue
        s = os.path.join(BASE, name)
        if os.path.exists(s):
            shutil.copy2(s, os.path.join(PKG, name)); n += 1
    for fn, txt in (("启动求职助手.bat", LAUNCHER), ("求职全自动.bat", AUTO),
                    ("README_使用说明.txt", README)):
        with open(os.path.join(PKG, fn), "w", encoding="utf-8") as f:
            f.write(txt)
    print(f"  ✅ {n} 个代码/配置文件")

    # 2. Python 安装器
    print("== Python 安装器 ==")
    dl(PY_URL, os.path.join(BUNDLE, "python-3.12.7-amd64.exe"), "Python")

    # 3. 依赖 wheels（离线装用）
    print("== 下载依赖 wheels ==")
    wheels = os.path.join(BUNDLE, "wheels")
    os.makedirs(wheels, exist_ok=True)
    subprocess.run([sys.executable, "-m", "pip", "download", "-r",
                    os.path.join(BASE, "requirements.txt"), "-d", wheels], check=False)

    # 4. Ollama 安装器
    print("== Ollama 安装器 (~700MB) ==")
    dl(OLLAMA_URL, os.path.join(BUNDLE, "OllamaSetup.exe"), "Ollama")

    # 5. 模型
    copy_model()

    # 隐私检查
    leaked = [fn for r, _, fs in os.walk(PKG) for fn in fs
              if fn in NEVER or fn == "cookie.txt"]
    print("=" * 56)
    if leaked:
        print(f"  ❌ 警告：混入隐私文件 {leaked}")
    else:
        print("  ✅ 隐私检查通过（无 cookie/投递记录/个人画像）")
    # 总大小
    total = sum(os.path.getsize(os.path.join(r, f)) for r, _, fs in os.walk(PKG) for f in fs)
    print(f"  📦 离线包总大小：{total/1073741824:.2f} GB")
    print(f"  📁 位置：{PKG}")
    print("  把这个文件夹整个传给别人(网盘/U盘)，对方双击里面的 一键安装.bat 即可。")
    print("=" * 56)


if __name__ == "__main__":
    main()
