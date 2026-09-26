# -*- coding: utf-8 -*-
"""Word 模板编译器 命令行入口。"""

from __future__ import annotations

import argparse
import json
import os
import sys

from .engine import (
    MissingFieldsError,
    build_skeleton,
    export_pdf,
    fill_template,
    load_json,
    scan_template,
)


def _print_fields(info) -> None:
    if info.is_empty:
        print("这个文档里没有发现特殊字段。")
        print("在 Word 里写上 {{字段名}} 这样的标记后重新扫描即可。")
        return
    print("=" * 60)
    print("普通字段（%d 个）" % len(info.fields))
    for field in info.fields:
        extra = ""
        if field.modifiers:
            extra = "  格式:%s" % "、".join(field.modifiers)
        print("  · %-16s 位置:%s%s" % (field.name, field.location or "-", extra))
    if info.loops:
        print("-" * 60)
        print("循环（%d 个）" % len(info.loops))
        for loop in info.loops:
            print("  · %-16s 次数由数据决定，项目字段:%s" % (loop.name, "、".join(loop.fields) or "-"))
    if info.conditions:
        print("-" * 60)
        print("条件（%d 个）" % len(info.conditions))
        for expression, location in info.conditions:
            print("  · %-16s 位置:%s" % (expression, location or "-"))
    if info.images:
        print("-" * 60)
        print("图片字段：%s" % "、".join(info.images))
    if info.tables:
        print("-" * 60)
        print("表格字段（整张表由数据生成）：%s" % "、".join("{{表格:%s}}" % name for name in info.tables))
    if info.builtins:
        print("内置字段（自动填充，无需提供）：%s" % "、".join("$" + name for name in info.builtins))
    print("=" * 60)
    print("需要提供的 JSON 键：%s" % "、".join(info.data_keys))


def _resolve_output(template: str, output: str | None, suffix: str = "_已填充") -> str:
    if output:
        return output
    stem, ext = os.path.splitext(os.path.abspath(template))
    return "%s%s%s" % (stem, suffix, ext or ".docx")


def _cmd_scan(args) -> int:
    info = scan_template(args.template)
    if args.as_json:
        payload = {
            "字段": [
                {"名称": f.name, "类型": f.kind, "位置": f.location, "格式": list(f.modifiers)}
                for f in info.fields
            ],
            "循环": [{"名称": l.name, "项目字段": l.fields, "位置": l.location} for l in info.loops],
            "条件": [{"表达式": e, "位置": loc} for e, loc in info.conditions],
            "图片": info.images,
            "内置": info.builtins,
        }
        text = json.dumps(payload, ensure_ascii=False, indent=2)
        if args.as_json == "-":
            print(text)
        else:
            with open(args.as_json, "w", encoding="utf-8") as fh:
                fh.write(text)
            print("字段清单已写入 %s" % args.as_json)
    else:
        _print_fields(info)
    return 0


def _cmd_skeleton(args) -> int:
    info = scan_template(args.template)
    skeleton = build_skeleton(info)
    target = args.output or os.path.splitext(os.path.abspath(args.template))[0] + "_数据模板.json"
    with open(target, "w", encoding="utf-8") as fh:
        json.dump(skeleton, fh, ensure_ascii=False, indent=2)
    print("已根据模板生成 JSON 数据骨架：%s" % target)
    print("把里面的空值改成真实内容，再用「填充」命令生成文档。")
    return 0


def _cmd_fill(args) -> int:
    data = load_json(args.data)
    if not isinstance(data, dict):
        print("JSON 最外层必须是对象（键值对），例如 {\"姓名\": \"张三\"}")
        return 2
    output = _resolve_output(args.template, args.output)
    try:
        result = fill_template(
            args.template,
            data,
            output,
            blank_missing=args.blank,
            strict=args.strict,
            fill_controls=not args.no_controls,
        )
    except MissingFieldsError as exc:
        print("生成失败：%s" % exc)
        print("补全数据后重试，或去掉 --严格 参数（缺数据的字段会原样保留）。")
        return 3
    print("已生成：%s" % result.output_path)
    print(result.summary())
    if args.pdf:
        try:
            pdf = export_pdf(result.output_path)
            print("已导出 PDF：%s" % pdf)
        except Exception as exc:  # noqa: BLE001
            print("PDF 导出跳过：%s" % exc)
    return 0


def _cmd_check(args) -> int:
    info = scan_template(args.template)
    data = load_json(args.data)
    keys = set(info.data_keys)
    missing = [key for key in info.data_keys if _lookup(data, key) is None and not _has_key(data, key)]
    extra = [key for key in data.keys() if key not in keys]
    print("模板需要 %d 个顶层键，JSON 提供了 %d 个。" % (len(keys), len(data.keys())))
    if missing:
        print("缺少：" + "、".join(missing))
    else:
        print("数据齐全，没有缺少的顶层键。")
    if extra:
        print("JSON 里多出来的键（不影响生成，可能用不上）：" + "、".join(extra))
    return 0 if not missing else 1


def _has_key(data, key: str) -> bool:
    parts = key.split(".")
    node = data
    for part in parts:
        if isinstance(node, dict) and part in node:
            node = node[part]
        else:
            return False
    return True


def _lookup(data, key: str):
    node = data
    for part in key.split("."):
        if isinstance(node, dict):
            node = node.get(part)
        else:
            return None
    return node


def _cmd_pdf(args) -> int:
    pdf = export_pdf(args.document, args.output_dir)
    print("已导出 PDF：%s" % pdf)
    return 0


def _cmd_gui(_args) -> int:
    from .gui import main as gui_main

    return gui_main()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="wordfill",
        description="Word 模板编译器：把 Word 文档里的 {{特殊字段}} 用 JSON 数据替换掉。",
    )
    sub = parser.add_subparsers(dest="command")

    scan = sub.add_parser("扫描", aliases=["scan", "fields"], help="列出文档里的所有特殊字段")
    scan.add_argument("template", help="Word 模板文件（.docx/.dotx）")
    scan.add_argument("--json", dest="as_json", nargs="?", const="-", help="输出字段清单 JSON")
    scan.set_defaults(func=_cmd_scan)

    skeleton = sub.add_parser(
        "生成数据", aliases=["skeleton", "json"], help="根据模板生成一份 JSON 数据骨架"
    )
    skeleton.add_argument("template")
    skeleton.add_argument("-o", "--output", help="保存的 JSON 路径")
    skeleton.set_defaults(func=_cmd_skeleton)

    fill = sub.add_parser("填充", aliases=["fill"], help="用 JSON 数据生成文档")
    fill.add_argument("template", help="Word 模板文件")
    fill.add_argument("data", help="JSON 数据文件")
    fill.add_argument("-o", "--output", help="输出文件（默认在模板旁生成 *_已填充.docx）")
    fill.add_argument("--blank", action="store_true", help="没有数据的字段替换为空（默认原样保留）")
    fill.add_argument("--strict", action="store_true", help="缺少数据时报错，不生成文档")
    fill.add_argument("--no-controls", action="store_true", help="不填充 Word 表单域（内容控件）")
    fill.add_argument("--pdf", action="store_true", help="同时导出 PDF（需要本机安装 LibreOffice）")
    fill.set_defaults(func=_cmd_fill)

    check = sub.add_parser("检查", aliases=["check"], help="检查 JSON 是否覆盖了模板需要的字段")
    check.add_argument("template")
    check.add_argument("data")
    check.set_defaults(func=_cmd_check)

    pdf = sub.add_parser("转PDF", aliases=["pdf"], help="把 Word 文档转成 PDF")
    pdf.add_argument("document")
    pdf.add_argument("-d", "--output-dir", dest="output_dir")
    pdf.set_defaults(func=_cmd_pdf)

    gui = sub.add_parser("界面", aliases=["gui"], help="打开图形界面")
    gui.set_defaults(func=_cmd_gui)

    return parser


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if not getattr(args, "command", None):
        parser.print_help()
        return 0
    try:
        return args.func(args)
    except FileNotFoundError as exc:
        print("错误：%s" % exc)
        return 2
    except Exception as exc:  # noqa: BLE001
        print("错误：%s" % exc)
        return 2


if __name__ == "__main__":
    sys.exit(main())
