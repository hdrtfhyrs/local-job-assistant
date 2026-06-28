r"""
求职助手 · 离线一键安装（在用户电脑上跑，全程不用联网下载大文件）
被 一键安装.bat 调用（bat 已确保有 Python）。用包里自带的东西装好：
  1) 建运行环境 runtime\ 并用 bundle\wheels 离线装依赖
  2) 装 Ollama（用 bundle\OllamaSetup.exe）
  3) 把自带的 AI 模型 bundle\models 拷到 Ollama 目录（免下载！）
  4) 检查 Chrome
  5) 桌面建快捷方式
单步失败不致命，会提示怎么手动补。
"""
import os
import shutil
import subprocess
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# 打包成 exe 后 __file__ 在临时目录，要用 exe 所在目录
if getattr(sys, "frozen", False):
    BASE = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE = os.path.dirname(os.path.abspath(__file__))
BUNDLE = os.path.join(BASE, "bundle")
RUNTIME = os.path.join(BASE, "runtime")
PY_RUNTIME = os.path.join(RUNTIME, "Scripts", "python.exe")
PYW_RUNTIME = os.path.join(RUNTIME, "Scripts", "pythonw.exe")
MODEL = "qwen3:8b"
TOTAL = 5


def step(n, msg):
    print(f"\n[{n}/{TOTAL}] {msg}", flush=True)


def run(cmd, **kw):
    show = cmd if isinstance(cmd, str) else " ".join(str(c) for c in cmd)
    print("    > " + show[:120], flush=True)
    return subprocess.run(cmd, **kw)


def real_python():
    """找一个真实的 Python 解释器(用来建 venv)。打包成 exe 后 sys.executable 是 exe 本身，不能用。"""
    if not getattr(sys, "frozen", False):
        return sys.executable
    for name in ("py", "python"):
        p = shutil.which(name)
        if p:
            return p
    return None


def ensure_python():
    """没装 Python 就用包里自带的安装器装。返回可用的 python 路径。"""
    py = real_python()
    if py:
        return py
    setup = None
    import glob
    cand = glob.glob(os.path.join(BUNDLE, "python-*.exe"))
    if cand:
        setup = cand[0]
    if setup:
        print("    没检测到 Python，正在用自带安装包静默安装……")
        run([setup, "/quiet", "InstallAllUsers=0", "PrependPath=1", "Include_pip=1"], check=False)
    else:
        print("    ⚠️ 没找到 Python 也没有自带安装包，请先装 Python(python.org)。")
    return real_python()


def ensure_runtime():
    step(1, "创建运行环境并离线安装依赖……")
    py = ensure_python()
    if not py:
        print("    ❌ 没有可用的 Python，无法建运行环境。")
        return
    if not os.path.exists(PY_RUNTIME):
        run([py, "-m", "venv", RUNTIME], check=False)
    pip = [PY_RUNTIME, "-m", "pip"]
    wheels = os.path.join(BUNDLE, "wheels")
    req = os.path.join(BASE, "requirements.txt")
    if os.path.isdir(wheels):
        r = run(pip + ["install", "--no-index", "--find-links", wheels, "-r", req], check=False)
    else:
        r = run(pip + ["install", "-r", req], check=False)  # 兜底联网
    print("    ✅ 依赖装好" if r.returncode == 0 else "    ⚠️ 依赖可能没装全，可重跑安装器")


def find_ollama():
    p = shutil.which("ollama")
    if p:
        return p
    cand = os.path.expandvars(r"%LOCALAPPDATA%\Programs\Ollama\ollama.exe")
    return cand if os.path.exists(cand) else None


def ensure_ollama():
    step(2, "安装本地 AI 引擎 Ollama……")
    ol = find_ollama()
    if not ol:
        setup = os.path.join(BUNDLE, "OllamaSetup.exe")
        if not os.path.exists(setup):
            # 在线版（GitHub代码下载者）：自己下 Ollama 安装包
            print("    没有自带安装包，正在从官网下载 Ollama……")
            try:
                import urllib.request
                urllib.request.urlretrieve("https://ollama.com/download/OllamaSetup.exe", setup)
            except Exception as e:
                print(f"    ⚠️ 下载 Ollama 失败：{e}，请到 https://ollama.com/download 手动装。")
        if os.path.exists(setup):
            print("    静默安装 Ollama……")
            run([setup, "/VERYSILENT", "/NORESTART"], check=False)
        ol = find_ollama()
    print("    ✅ Ollama 就绪" if ol else "    ⚠️ 没找到 Ollama，可能要重启电脑后再跑一次本安装器。")
    return ol


def install_model():
    step(3, "导入自带 AI 模型（免下载）……")
    src = os.path.join(BUNDLE, "models")
    if not os.path.isdir(src):
        # 在线版：没有自带模型就联网拉（国内可能较慢/需代理）
        print("    没有自带模型，改为联网拉取 qwen3:8b（国内可能慢，失败可手动 ollama pull qwen3:8b）……")
        ol = find_ollama()
        if ol:
            run([ol, "pull", MODEL], check=False)
        return
    dst = os.path.join(os.path.expanduser("~"), ".ollama", "models")
    os.makedirs(os.path.join(dst, "blobs"), exist_ok=True)
    # 拷数据块（已存在的跳过）
    sb = os.path.join(src, "blobs")
    n = 0
    if os.path.isdir(sb):
        for fn in os.listdir(sb):
            d = os.path.join(dst, "blobs", fn)
            if not os.path.exists(d):
                shutil.copy2(os.path.join(sb, fn), d)
                n += 1
    # 拷清单
    sm = os.path.join(src, "manifests")
    if os.path.isdir(sm):
        for root, _, files in os.walk(sm):
            rel = os.path.relpath(root, sm)
            outd = os.path.join(dst, "manifests", rel)
            os.makedirs(outd, exist_ok=True)
            for fn in files:
                shutil.copy2(os.path.join(root, fn), os.path.join(outd, fn))
    print(f"    ✅ 模型已导入（新增 {n} 个数据块）。可用 'ollama list' 确认有 {MODEL}")


def check_chrome():
    cands = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ]
    if any(os.path.exists(c) for c in cands):
        print("    ✅ 已检测到 Chrome。")
        return
    print("    ⚠️ 没装 Chrome，本工具需要它，正在打开下载页……")
    try:
        import webbrowser
        webbrowser.open("https://www.google.cn/chrome/")
    except Exception:
        print("    请手动装 Chrome：https://www.google.cn/chrome/")


def make_shortcut():
    launcher = os.path.join(BASE, "启动求职助手.bat")
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    lnk = os.path.join(desktop, "求职助手.lnk")
    icon = PYW_RUNTIME if os.path.exists(PYW_RUNTIME) else launcher
    ps = ("$s=(New-Object -ComObject WScript.Shell).CreateShortcut('{lnk}');"
          "$s.TargetPath='{tgt}';$s.WorkingDirectory='{wd}';$s.IconLocation='{icon}';$s.Save()").format(
        lnk=lnk.replace("'", "''"), tgt=launcher.replace("'", "''"),
        wd=BASE.replace("'", "''"), icon=icon.replace("'", "''"))
    try:
        run(["powershell", "-NoProfile", "-Command", ps], check=False)
        print("    ✅ 桌面已建快捷方式【求职助手】")
    except Exception as e:
        print(f"    ⚠️ 建快捷方式失败：{e}（可手动双击 启动求职助手.bat）")


def self_check():
    print("=" * 56)
    print("  安装器自检（不实际安装）")
    print("=" * 56)
    print(f"  运行模式: {'打包exe' if getattr(sys,'frozen',False) else '脚本'}")
    print(f"  安装目录 BASE: {BASE}")
    print(f"  bundle 存在: {os.path.isdir(BUNDLE)}")
    if os.path.isdir(BUNDLE):
        for sub in ("python-3.12.7-amd64.exe", "OllamaSetup.exe", "wheels", "models"):
            print(f"    - {sub}: {'有' if os.path.exists(os.path.join(BUNDLE, sub)) else '无'}")
    print(f"  找到真实 Python: {real_python() or '无(将用自带安装器)'}")
    print(f"  requirements.txt: {os.path.exists(os.path.join(BASE, 'requirements.txt'))}")
    print("  ✅ 自检完成（路径解析正常即说明 exe 能找对位置）")


def main():
    if "--check" in sys.argv:
        self_check()
        return
    print("=" * 56)
    print("  求职助手 · 离线一键安装")
    print("=" * 56)
    ensure_runtime()
    ensure_ollama()
    install_model()
    step(4, "检查 Chrome……")
    check_chrome()
    step(5, "创建桌面快捷方式……")
    make_shortcut()
    print("\n" + "=" * 56)
    print("  ✅ 安装完成！")
    print("  双击桌面【求职助手】打开 → 先点【🧑 我的资料】填情况 →")
    print("  点【🌐 打开网站·登录BOSS】扫码登录 → 就能用了。")
    print("=" * 56)
    try:
        input("\n按回车关闭……")
    except Exception:
        pass


if __name__ == "__main__":
    main()
