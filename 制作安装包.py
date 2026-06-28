# -*- coding: utf-8 -*-
r"""
制作"一键配装包"（开发者用，把工具打包成发给别人的安装包）。
用白名单只打包【代码+配置模板+安装器】，绝不打包 cookie/投递记录等个人隐私。
产出：dist\求职助手\  和  dist\求职助手安装包.zip
用法：python 制作安装包.py
"""
import os
import shutil
import sys
import zipfile

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

BASE = os.path.dirname(os.path.abspath(__file__))
DIST = os.path.join(BASE, "dist")
PKG = os.path.join(DIST, "求职助手")

# 程序代码（运行必需）
CODE = [
    "job_assistant_app.py", "boss_scraper.py", "job_scraper_core.py",
    "multi_platform_scraper.py", "scrape_beginner_jobs.py",
    "job_matcher.py", "local_ai_matcher.py", "merge_analysis.py",
    "build_web.py", "auto_apply.py", "reply_monitor.py", "auto_run_all.py",
    "profile_engine.py", "career_advisor.py", "career_research_collector.py",
    "application_package.py",
]
# 配置/字典模板（结构性数据，非个人隐私）
DATA = [
    "requirements.txt", "cities.json", "city_dict.json",
    "recommended_directions.json", "local_niche_channels.json",
    # 下面这几个是"默认画像模板"，用户用【我的资料】会覆盖成自己的
    "match_profile.json", "beginner_job_profile.json", "career_direction_profile.json",
]
# 安装器
INSTALLER = ["install.py", "一键安装.bat"]

# 绝不打包（个人隐私/本机产物）——白名单之外的本来就不会进，这里只是双保险检查
NEVER = {"cookie.txt", "user_profile.json", "投递记录.json", "回信快照.json",
         "投递配置.json", "自动运行配置.json", "自动运行日志.txt"}

LAUNCHER = """@echo off
chcp 65001 >nul
cd /d "%~dp0"
if exist "runtime\\Scripts\\pythonw.exe" (
  start "" "runtime\\Scripts\\pythonw.exe" "job_assistant_app.py"
) else (
  echo [!] runtime not found. Please run 一键安装.bat first.
  pause
)
"""

AUTO_LAUNCHER = """@echo off
chcp 65001 >nul
cd /d "%~dp0"
"runtime\\Scripts\\python.exe" auto_run_all.py
pause
"""

README = """求职助手 · 使用说明
============================

【第一次用，三步】
1) 双击  一键安装.bat   —— 自动装好运行环境、AI模型(qwen3:8b)、检查Chrome、建桌面图标。
   （第一次要下载AI模型约5GB，耐心等；没装Chrome会帮你打开下载页。）
2) 双击桌面【求职助手】图标打开程序。先点【🧑 我的资料】填你的城市/学历/想做的方向。
3) 点【②打开网站并登录】，在弹出的Chrome里扫码登录BOSS直聘。

【日常用】
- 想全自动跑：双击  求职全自动.bat   （爬岗位→筛选→AI精排→出网页→自动投递→查回信）
- 或打开程序，用左边按钮一步步来。

【说明】
- AI精排需要本地Ollama+模型，安装器会自动装好。
- 自动投递用的是你自己登录的BOSS账号，请遵守平台规则，注意投递节奏。
"""


def main():
    if os.path.exists(PKG):
        shutil.rmtree(PKG)
    os.makedirs(PKG)

    copied, missing = [], []
    for name in CODE + DATA + INSTALLER:
        src = os.path.join(BASE, name)
        if name in NEVER:
            continue  # 双保险
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(PKG, name))
            copied.append(name)
        else:
            missing.append(name)

    # 生成启动器 + 说明
    with open(os.path.join(PKG, "启动求职助手.bat"), "w", encoding="utf-8") as f:
        f.write(LAUNCHER)
    with open(os.path.join(PKG, "求职全自动.bat"), "w", encoding="utf-8") as f:
        f.write(AUTO_LAUNCHER)
    with open(os.path.join(PKG, "README_使用说明.txt"), "w", encoding="utf-8") as f:
        f.write(README)

    # 隐私检查：确保包里没有任何敏感文件
    leaked = []
    for root, _, files in os.walk(PKG):
        for fn in files:
            if fn in NEVER or fn.endswith(".xlsx") or fn == "cookie.txt":
                leaked.append(fn)

    # 打包 zip
    zip_path = os.path.join(DIST, "求职助手安装包.zip")
    if os.path.exists(zip_path):
        os.remove(zip_path)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(PKG):
            for fn in files:
                fp = os.path.join(root, fn)
                z.write(fp, os.path.relpath(fp, DIST))

    print("=" * 50)
    print(f"  打包完成：{zip_path}")
    print(f"  含文件 {len(copied)} 个 + 启动器/说明")
    if missing:
        print(f"  ⚠️ 缺失(跳过)：{missing}")
    if leaked:
        print(f"  ❌ 警告！包里混入了隐私文件：{leaked}")
    else:
        print("  ✅ 隐私检查通过：无 cookie/投递记录/个人画像 等敏感文件")
    size = os.path.getsize(zip_path) / 1024
    print(f"  压缩包大小：{size:.0f} KB（不含运行环境和模型，由安装器现装）")
    print("=" * 50)


if __name__ == "__main__":
    main()
