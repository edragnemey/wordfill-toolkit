# -*- coding: utf-8 -*-
"""wordfill skill 的统一命令行入口。

    python wf.py scan     模板.docx
    python wf.py skeleton 模板.docx -o 数据.json
    python wf.py check    模板.docx 数据.json [--config 字段配置.json]
    python wf.py check    --keys-from 骨架.json 数据.json     # 没有模板时
    python wf.py check    --config 字段配置.json 数据.json     # 只按配置校验
    python wf.py fill     模板.docx 数据.json -o 结果.docx [--pdf] [--blank]
    python wf.py pdf      结果.docx

所有子命令都返回退出码：0 = 没问题，1 = 有问题（check 失败或生成失败）。
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from wordfill import engine  # noqa: E402


# --------------------------------------------------------------------------
# 工具函数
# --------------------------------------------------------------------------


def load_json(path):
    with open(path, "rb") as fh:
        raw = fh.read()
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return json.loads(raw.decode(encoding))
        except (UnicodeDecodeError, json.JSONDecodeError):
            continue
    raise ValueError("无法解析 JSON：%s" % path)


def dump_json(data, path):
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)


def lookup(data, path):
    node = data
    for part in str(path).split("."):
        if isinstance(node, dict) and part in node:
            node = node[part]
        else:
            return False, None
    return True, node


def strip_private(data):
    """去掉以下划线开头的说明键（_字段说明 / _注释 等），生成文档时不用它们。"""
    if isinstance(data, dict):
        return {key: strip_private(value) for key, value in data.items() if not str(key).startswith("_")}
    if isinstance(data, list):
        return [strip_private(item) for item in data]
    return data


def is_empty(value) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    if isinstance(value, (list, tuple, dict)):
        return len(value) == 0
    return False


# --------------------------------------------------------------------------
# check：生成前自检
# --------------------------------------------------------------------------


def collect_config_fields(config, prefix=""):
    """把字段配置摊平成 {字段路径: 说明文字}。"""
    found = {}
    if isinstance(config, dict):
        for key, value in config.items():
            if str(key).startswith("_"):
                continue
            path = "%s.%s" % (prefix, key) if prefix else str(key)
            if isinstance(value, dict) and any(
                str(k).startswith("_") for k in value.keys()
            ):
                for meta_key, meta_value in value.items():
                    if str(meta_key).startswith("_"):
                        found[path] = meta_value
            elif isinstance(value, str):
                found[path] = value
            elif isinstance(value, dict):
                found.update(collect_config_fields(value, path))
    return found


def keys_from_skeleton(skeleton, prefix=""):
    """从 JSON 骨架里推断 需要的键 与 表格键。"""
    fields = []
    tables = []
    if isinstance(skeleton, dict):
        for key, value in skeleton.items():
            if str(key).startswith("_"):
                continue
            path = "%s.%s" % (prefix, key) if prefix else str(key)
            if isinstance(value, dict) and any(
                item in value for item in ("表头", "行", "rows", "数据", "明细")
            ) or (
                isinstance(value, dict)
                and any(str(item).startswith("_行") for item in value.keys())
            ):
                tables.append(path)
            elif isinstance(value, dict):
                sub_fields, sub_tables = keys_from_skeleton(value, path)
                fields.extend(sub_fields)
                tables.extend(sub_tables)
            elif isinstance(value, list) and value and isinstance(value[0], dict):
                tables.append(path)
            else:
                fields.append(path)
    return fields, tables


def config_table_paths(config, prefix=""):
    """字段配置里带 _行说明 的节点就是表格字段。"""
    paths = []
    if isinstance(config, dict):
        for key, value in config.items():
            if str(key).startswith("_"):
                continue
            path = "%s.%s" % (prefix, key) if prefix else str(key)
            if isinstance(value, dict):
                if any(str(item).startswith("_行") for item in value.keys()):
                    paths.append(path)
                paths.extend(config_table_paths(value, path))
    return paths


def table_problems(name, value):
    """检查表格字段的取值结构。"""
    problems = []
    if isinstance(value, list):
        rows = value
        header = None
    elif isinstance(value, dict):
        rows = value.get("行") or value.get("rows") or value.get("数据") or []
        header = value.get("表头") or value.get("header") or value.get("columns")
        if not isinstance(rows, list):
            problems.append("表格 %s 的「行」必须是数组" % name)
            return problems
        if header is not None and isinstance(header, dict):
            header = list(header.keys())
    else:
        return ["表格 %s 的值必须是对象（表头/行）或数组，现在给的是 %s" % (name, type(value).__name__)]

    if not rows:
        problems.append("表格 %s 没有任何数据行" % name)
        return problems
    if not header and isinstance(rows[0], dict):
        header = []
        for row in rows:
            if isinstance(row, dict):
                for key in row.keys():
                    if key not in header:
                        header.append(key)
    if header:
        columns = [str(item) for item in header]
        for index, row in enumerate(rows, start=1):
            if isinstance(row, dict):
                missing = [column for column in columns if column not in row]
                extra = [key for key in row.keys() if key not in columns]
                if missing:
                    problems.append("表格 %s 第 %d 行缺少列：%s" % (name, index, "、".join(missing)))
                if extra:
                    problems.append("表格 %s 第 %d 行有多余的键：%s" % (name, index, "、".join(extra)))
            elif isinstance(row, (list, tuple)):
                if len(row) != len(columns):
                    problems.append(
                        "表格 %s 第 %d 行有 %d 个值，表头是 %d 列"
                        % (name, index, len(row), len(columns))
                    )
    return problems


def cmd_check(args) -> int:
    data = load_json(args.data)
    if not isinstance(data, dict):
        print("[错误] 数据 JSON 的最外层必须是对象 { }")
        return 1
    data = strip_private(data)

    config = load_json(args.config) if args.config else None
    errors = []
    warnings = []
    conditions_fields = {}

    if args.template:
        info = engine.scan_template(args.template)
        needed = list(info.data_keys)
        tables = list(info.tables)
        conditions_fields = dict(info.conditions_fields or {})
        source = "模板 %s" % os.path.basename(args.template)
    elif args.keys_from:
        fields, tables = keys_from_skeleton(load_json(args.keys_from))
        needed = fields + [name for name in tables if name not in fields]
        source = "JSON 骨架 %s" % os.path.basename(args.keys_from)
    elif config is not None:
        fields, tables = keys_from_skeleton(config)
        needed = [key for key in config.keys() if not str(key).startswith("_")]
        tables = [path for path in config_table_paths(config) if "." not in path]
        source = "字段配置 %s" % os.path.basename(args.config)
    else:
        print("[错误] 至少给出模板，或用 --keys-from / --config 指定字段清单来源")
        return 1

    print("字段清单来源：%s" % source)

    # 条件为否的分支里的字段，空值是正常的
    conditional_ok = {}
    for expression, names in conditions_fields.items():
        key = engine._condition_key(expression)
        if not key:
            continue
        found, value = lookup(data, key)
        if found and not engine.to_bool(value):
            for name in names:
                conditional_ok[name] = key

    for key in needed:
        found, value = lookup(data, key)
        if not found:
            errors.append("缺少键：%s" % key)
            continue
        if is_empty(value):
            if key in conditional_ok:
                warnings.append(
                    "字段 %s 为空，但它位于条件「%s」为否的分支里，本次不会出现在文档中"
                    % (key, conditional_ok[key])
                )
            elif not args.template:
                warnings.append(
                    "字段 %s 的值为空；没有模板信息时无法判断它是否必填，请确认本次是否确实不需要"
                    % key
                )
            else:
                errors.append("键 %s 的值是空的" % key)

    for name in tables:
        found, value = lookup(data, name)
        if found and not is_empty(value):
            errors.extend(table_problems(name, value))

    leftover = [
        key
        for key in data.keys()
        if key not in needed
        and not any(key == item.split(".")[0] for item in needed)
    ]
    if leftover:
        warnings.append("数据里多了模板用不到的键：%s" % "、".join(leftover))

    config_fields = {}
    if config is not None:
        config_fields = collect_config_fields(config)
        for path, note in config_fields.items():
            found, value = lookup(data, path)
            text = str(note)
            required = not any(word in text for word in ("可空", "选填", "非必填", "没有就"))
            if not found:
                target = errors if required else warnings
                target.append("字段配置里要求的 %s 没有出现在数据里（说明：%s）" % (path, text[:40]))
            elif is_empty(value) and required:
                errors.append("字段 %s 是空的（配置说明：%s）" % (path, text[:40]))

    placeholders = [
        key for key in data.keys()
        if isinstance(data.get(key), str) and re.search(r"【(待确认|待补充|TODO)", data[key])
    ]
    if placeholders:
        warnings.append("以下字段还带着占位符，交付前请确认：" + "、".join(placeholders))

    print("需要的键（%d）：%s" % (len(needed), "、".join(needed) or "-"))
    if config_fields:
        print("字段配置覆盖（%d）：%s" % (len(config_fields), "、".join(sorted(config_fields))))
    for item in warnings:
        print("[提示] %s" % item)
    for item in errors:
        print("[问题] %s" % item)
    if errors or (args.strict and warnings):
        print(
            "检查未通过：%d 个问题%s"
            % (len(errors), "（--strict：提示也算问题）" if args.strict and warnings and not errors else "")
        )
        return 1
    print("检查通过：数据齐全，可以生成文档")
    return 0


# --------------------------------------------------------------------------
# 其它子命令
# --------------------------------------------------------------------------


def cmd_scan(args) -> int:
    info = engine.scan_template(args.template)
    print("普通字段：%s" % "、".join(field.name for field in info.fields) or "-")
    if info.tables:
        print("表格字段：%s" % "、".join("{{表格:%s}}" % name for name in info.tables))
    if info.loops:
        print("循环：%s" % "、".join(loop.name for loop in info.loops))
    if info.conditions:
        print("条件：%s" % "、".join(item[0] for item in info.conditions))
    if info.images:
        print("图片：%s" % "、".join(info.images))
    if info.builtins:
        print("内置（自动填）：%s" % "、".join("$" + name for name in info.builtins))
    print("数据里需要的顶层键：%s" % "、".join(info.data_keys))
    return 0


def cmd_skeleton(args) -> int:
    info = engine.scan_template(args.template)
    skeleton = engine.build_skeleton(info)
    if not skeleton:
        print("[错误] 模板里没有发现任何字段，先在 Word 里写 {{字段名}} 标记")
        return 1
    target = args.output or os.path.splitext(os.path.abspath(args.template))[0] + "_数据.json"
    dump_json(skeleton, target)
    print("已生成数据骨架：%s" % target)
    print("接下来按字段填写说明把每个空值改成真实内容，再用 wf.py check / fill。")
    return 0


def cmd_fill(args) -> int:
    data = strip_private(load_json(args.data))
    if not isinstance(data, dict):
        print("[错误] 数据 JSON 的最外层必须是对象 { }")
        return 1
    output = args.output or os.path.splitext(os.path.abspath(args.template))[0] + "_已填充.docx"
    result = engine.fill_template(
        args.template,
        data,
        output,
        blank_missing=args.blank,
        strict=False,
    )
    print("已生成：%s" % result.output_path)
    head = "已替换 %d 处字段" % result.filled
    extras = []
    if result.images:
        extras.append("%d 张图片" % result.images)
    if result.tables:
        extras.append("%d 张自动生成的表格" % result.tables)
    if extras:
        head += "（含 %s）" % "、".join(extras)
    print(head)
    if result.missing:
        print("[注意] 缺少数据：%s" % "、".join(result.missing))
    if result.leftover:
        print("[注意] 文档里仍保留的标记：%s" % "、".join(result.leftover))
    for warning in result.warnings:
        print("[提示] %s" % warning)
    if result.unused:
        print("[提示] 以下键本次没有用到（通常是条件分支未命中，属正常）：%s" % "、".join(result.unused))
    if args.pdf:
        try:
            print("已导出 PDF：%s" % engine.export_pdf(result.output_path))
        except Exception as exc:  # noqa: BLE001
            print("[提示] PDF 导出跳过：%s" % exc)
    if result.missing or result.leftover:
        print("[注意] 文档里仍有未替换的标记，交付前请补齐数据或改用 --blank")
        return 1
    return 0


def cmd_pdf(args) -> int:
    print(engine.export_pdf(args.document))
    return 0


def build_parser():
    parser = argparse.ArgumentParser(prog="wf", description="Word 模板编译器 skill 命令行")
    sub = parser.add_subparsers(dest="command")

    scan = sub.add_parser("scan", help="列出模板里的字段")
    scan.add_argument("template")
    scan.set_defaults(func=cmd_scan)

    skeleton = sub.add_parser("skeleton", help="从模板生成数据骨架")
    skeleton.add_argument("template")
    skeleton.add_argument("-o", "--output")
    skeleton.set_defaults(func=cmd_skeleton)

    check = sub.add_parser("check", help="生成前检查数据是否齐全、表格结构是否正确")
    check.add_argument("template", nargs="?", help="模板 docx；用 --keys-from / --config 时可以省略")
    check.add_argument("data")
    check.add_argument("--config", help="字段填写配置（同构说明文件）")
    check.add_argument("--keys-from", dest="keys_from", help="从 JSON 骨架取字段清单（没有模板时用）")
    check.add_argument("--strict", action="store_true", help="把提示也当成问题")
    check.set_defaults(func=cmd_check)

    fill = sub.add_parser("fill", help="用数据生成文档")
    fill.add_argument("template")
    fill.add_argument("data")
    fill.add_argument("-o", "--output")
    fill.add_argument("--pdf", action="store_true", help="同时导出 PDF（需要 LibreOffice）")
    fill.add_argument("--blank", action="store_true", help="没有数据的字段替换为空")
    fill.set_defaults(func=cmd_fill)

    pdf = sub.add_parser("pdf", help="把 docx 导出成 PDF")
    pdf.add_argument("document")
    pdf.set_defaults(func=cmd_pdf)

    return parser


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if not getattr(args, "command", None):
        parser.print_help()
        return 0
    try:
        return args.func(args)
    except Exception as exc:  # noqa: BLE001
        print("[错误] %s" % exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
