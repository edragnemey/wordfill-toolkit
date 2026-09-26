# -*- coding: utf-8 -*-
"""启动 Word 模板编译器。

用法：
    双击 一键启动.bat
    或把 Word 文件直接拖到 一键启动.bat 上（会用这个文件当模板打开）
    或命令行执行  python start_tool.py "D:\\合同.docx"
"""

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)


def _show_error(message: str) -> None:
    # pythonw 启动时没有控制台，用弹窗把错误显示出来
    try:
        print(message)
    except Exception:  # noqa: BLE001
        pass
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(0, message, "Word 模板编译器", 0x10)
    except Exception:  # noqa: BLE001
        pass


def main() -> int:
    template = None
    arguments = [item for item in sys.argv[1:] if item and not item.startswith("-")]
    if arguments:
        candidate = os.path.abspath(arguments[0].strip('"'))
        if os.path.exists(candidate):
            template = candidate

    try:
        import docx  # noqa: F401
    except Exception:  # noqa: BLE001
        _show_error(
            "缺少组件 python-docx，无法启动。\n\n"
            "请在命令行里执行下面这行命令安装后再打开：\n\n"
            "    pip install python-docx"
        )
        return 1

    try:
        from wordfill.gui import main as gui_main
    except Exception as exc:  # noqa: BLE001
        _show_error(
            "启动失败：%s\n\n"
            "请确认已经安装 Python 3.10 以上版本，并且执行过：\n\n"
            "    pip install python-docx" % exc
        )
        return 1

    try:
        return gui_main(template)
    except Exception as exc:  # noqa: BLE001
        import traceback

        _show_error("程序运行出错：\n\n%s\n\n%s" % (exc, traceback.format_exc(limit=3)))
        return 1


if __name__ == "__main__":
    sys.exit(main())
