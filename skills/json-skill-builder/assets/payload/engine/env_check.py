# -*- coding: utf-8 -*-
"""环境自检：换到新电脑后先跑这个，确认能不能用。

    python scripts/env_check.py
"""

from __future__ import annotations

import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

problems = []
notes = []

print("skill 目录：%s" % ROOT)
print("系统：%s %s" % (sys.platform, "（Windows）" if sys.platform == "win32" else "（macOS / Linux）"))
print("Python：%s" % sys.version.split()[0])
if sys.version_info < (3, 10):
    problems.append("Python 版本过低，需要 3.10 以上")

try:
    import docx  # noqa: F401

    print("python-docx：已安装")
except Exception:  # noqa: BLE001
    problems.append("缺少 python-docx，请执行：pip install python-docx")

required = [
    "scripts/wf.py",
    "scripts/wordfill/engine.py",
    "references/json-rules.md",
    "references/how-to-fill.md",
]
assets = os.path.join(ROOT, "assets")
if os.path.isdir(assets):
    for name in sorted(os.listdir(assets)):
        required.append(os.path.join("assets", name))
for item in required:
    if not os.path.exists(os.path.join(ROOT, item)):
        problems.append("缺少文件：%s" % item)

try:
    from wordfill import engine  # noqa: F401

    print("填充引擎：可用")
except Exception as exc:  # noqa: BLE001
    problems.append("引擎加载失败：%s" % exc)

def find_libreoffice():
    for name in ("soffice", "soffice.com", "soffice.exe"):
        found = shutil.which(name)
        if found:
            return found
    for path in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
        "/opt/homebrew/bin/soffice",
    ):
        if os.path.exists(path):
            return path
    return None


if find_libreoffice():
    notes.append("检测到 LibreOffice，可以用 --pdf 导出 PDF")
else:
    notes.append(
        "没有 LibreOffice：不能用 --pdf 导出，docx 不受影响"
        "（可以用 Word / Pages 打开后另存为 PDF）"
    )

for item in notes:
    print("[提示] %s" % item)
for item in problems:
    print("[问题] %s" % item)
if problems:
    print("环境检查未通过：%d 个问题" % len(problems))
    sys.exit(1)
print("环境检查通过，可以用了：python scripts/wf.py check ... / fill ...")
