#!/bin/bash
# macOS / Linux：双击本文件（或在终端执行 bash 启动工具.command）打开 Word 模板编译器界面。
cd "$(dirname "$0")" || exit 1

export PYTHONUTF8=1
PY=""
for candidate in python3 python; do
  if command -v "$candidate" >/dev/null 2>&1; then PY="$candidate"; break; fi
done
if [ -z "$PY" ]; then
  echo "没有找到 Python。请先安装 Python 3.10 或更高版本，然后执行："
  echo "    pip3 install python-docx"
  read -r -p "按回车关闭…" _
  exit 1
fi

"$PY" "start_tool.py"
