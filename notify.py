# -*- coding: utf-8 -*-
"""
====================================================
  Windows 桌面通知（右下角弹窗）
====================================================

给无人值守的流程用：登录过期、有新回信、步骤失败时，
在屏幕右下角弹一条系统通知，不用盯日志。

不依赖任何第三方库——用 Windows 自带的 PowerShell 调系统
toast 接口（Win10/11 都支持），后台静默执行、不闪黑窗。

用法：
  from notify import notify
  notify("标题", "内容")

命令行自测：
  python notify.py            # 弹一条测试通知
"""

import base64
import os
import subprocess
import sys
from xml.sax.saxutils import escape

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# 借 PowerShell 的注册身份发通知（自己不是打包应用，没有 AppId）
_APP_ID = r"{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe"

_PS_TEMPLATE = """
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
$xml = @'
<toast duration="long"><visual><binding template="ToastGeneric"><text>__TITLE__</text><text>__BODY__</text></binding></visual></toast>
'@
$doc = New-Object Windows.Data.Xml.Dom.XmlDocument
$doc.LoadXml($xml)
$toast = [Windows.UI.Notifications.ToastNotification]::new($doc)
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('__APPID__').Show($toast)
"""

CREATE_NO_WINDOW = 0x08000000


def notify(title, msg):
    """弹一条桌面通知。失败不抛异常（打日志代替），返回 True/False。"""
    title = str(title or "求职助手")
    msg = str(msg or "")
    script = (
        _PS_TEMPLATE
        .replace("__TITLE__", escape(title))
        .replace("__BODY__", escape(msg))
        .replace("__APPID__", _APP_ID)
    )
    try:
        encoded = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded],
            capture_output=True, timeout=20, creationflags=CREATE_NO_WINDOW,
        )
        if r.returncode == 0:
            return True
        print(f"(桌面通知发送失败：{(r.stderr or b'').decode('gbk', 'replace').strip()[:200]})", flush=True)
        return False
    except Exception as e:
        print(f"(桌面通知发送失败：{e})", flush=True)
        return False


if __name__ == "__main__":
    t = sys.argv[1] if len(sys.argv) > 1 else "求职助手 · 通知测试"
    m = sys.argv[2] if len(sys.argv) > 2 else "看到这条弹窗说明通知功能正常。"
    ok = notify(t, m)
    print("通知已发出 ✅" if ok else "通知失败 ❌")
    sys.exit(0 if ok else 1)
