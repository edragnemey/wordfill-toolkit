#!/bin/bash
# macOS / Linux：双击本文件（或终端执行 bash 启动填写界面.command）打开填写界面。
HERE="$(cd "$(dirname "$0")" && pwd)"
GUI="$HERE/skills/json-skill-builder/scripts/skill_builder_gui.py"
[ -f "$GUI" ] || GUI="$HOME/.codex/skills/json-skill-builder/scripts/skill_builder_gui.py"

if [ ! -f "$GUI" ]; then
  echo "找不到 skill_builder_gui.py"
  echo "期望位置：$HERE/skills/json-skill-builder/scripts/skill_builder_gui.py"
  read -r -p "按回车关闭…" _
  exit 1
fi

export PYTHONUTF8=1
PY=""
for candidate in python3 python; do
  if command -v "$candidate" >/dev/null 2>&1; then PY="$candidate"; break; fi
done
if [ -z "$PY" ]; then
  echo "没有找到 Python。请先安装 Python 3.10+，然后执行：  pip3 install python-docx"
  read -r -p "按回车关闭…" _
  exit 1
fi

"$PY" "$GUI"
