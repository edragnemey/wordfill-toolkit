# -*- coding: utf-8 -*-
"""JSON 字段 Skill 生成器。

把「一份 JSON 骨架 + 每个字段的填写描述」变成一个专用 skill：AI 之后按描述和对话上下文填这份 JSON。

    python gen_skill.py init    <骨架.json> [-o 字段说明.json]
    python gen_skill.py check   <骨架.json> <字段说明.json>
    python gen_skill.py build   <骨架.json> <字段说明.json> --name <skill名> [--display-name 名称] [--template 模板.docx] [--force]
    python gen_skill.py list    <骨架.json>
    python gen_skill.py validate <skill目录>
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)
PAYLOAD = os.path.join(SKILL_DIR, "assets", "payload")
DEFAULT_SKILLS_DIR = os.path.join(os.path.expanduser("~"), ".codex", "skills")

TABLE_KEYS = ("表头", "行", "rows", "数据", "明细")


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


def set_path(data: dict, path: str, value) -> None:
    """按「a.b.c」把值写进嵌套字典。"""
    parts = [part for part in str(path).split(".") if part]
    if not parts:
        return
    node = data
    for part in parts[:-1]:
        child = node.get(part)
        if not isinstance(child, dict):
            child = {}
            node[part] = child
        node = child
    node[parts[-1]] = value


# --------------------------------------------------------------------------
# 结构分析
# --------------------------------------------------------------------------


def is_table(node) -> bool:
    if isinstance(node, dict):
        if any(key in node for key in TABLE_KEYS):
            return True
        if str(node.get("_类型", "")).startswith("表格"):
            return True
    if isinstance(node, list) and node and isinstance(node[0], dict):
        return True
    return False


def table_columns(node):
    columns = []
    if isinstance(node, dict):
        header = node.get("表头") or node.get("header") or node.get("columns")
        if isinstance(header, dict):
            columns = list(header.keys())
        elif isinstance(header, list):
            columns = list(header)
        rows = node.get("行") or node.get("rows") or node.get("数据") or node.get("明细")
        if not columns and isinstance(rows, list) and rows and isinstance(rows[0], dict):
            columns = list(rows[0].keys())
        if not columns:
            columns = [key for key in node.keys() if not str(key).startswith("_")]
    elif isinstance(node, list) and node and isinstance(node[0], dict):
        columns = list(node[0].keys())
    return [str(item) for item in columns if not str(item).startswith("_")]


def is_placeholder_columns(columns) -> bool:
    """骨架里由模板自动生成的占位列名，如 列1 / 列2 / col1。"""
    if not columns:
        return False
    return all(re.fullmatch(r"(列|col)\s*\d+", str(item), re.I) for item in columns)


def rebuild_tables(node, description):
    """用字段说明里的列名替换骨架里的占位列名，返回 (新节点, 提示列表)。"""
    notes = []
    if isinstance(node, dict):
        if is_table(node):
            columns = [key for key in (description or {}).keys() if not str(key).startswith("_")]
            old_columns = table_columns(node)
            if columns and old_columns != columns:
                if is_placeholder_columns(old_columns) or not old_columns:
                    header = node.get("表头") if isinstance(node, dict) else None
                    new_header = (
                        {column: (header.get(column, "") if isinstance(header, dict) else "") for column in columns}
                        if isinstance(header, dict)
                        else columns
                    )
                    rebuilt = {"表头": new_header, "行": [{column: "" for column in columns}]}
                    for key in ("合并列", "列格式", "字号", "无表头"):
                        if isinstance(node, dict) and key in node:
                            rebuilt[key] = node[key]
                    notes.append("表格列名按字段说明重建：%s" % "、".join(columns))
                    return rebuilt, notes
                notes.append(
                    "字段说明里的列（%s）与骨架里的列（%s）不一致，已保留骨架列名，请确认"
                    % ("、".join(columns), "、".join(old_columns))
                )
            return node, notes
        result = {}
        for key, value in node.items():
            if str(key).startswith("_"):
                result[key] = value
                continue
            child_desc = description.get(key) if isinstance(description, dict) else None
            result[key], child_notes = rebuild_tables(value, child_desc)
            notes.extend(child_notes)
        return result, notes
    if isinstance(node, list) and node and isinstance(node[0], dict):
        columns = [key for key in (description or {}).keys() if not str(key).startswith("_")]
        old_columns = table_columns(node)
        row_template = {
            key: value for key, value in node[0].items() if not str(key).startswith("_")
        }
        if columns and (is_placeholder_columns(old_columns) or old_columns != columns):
            rows = [{column: "" for column in columns}]
            notes.append("数组字段按字段说明重建列：%s" % "、".join(columns))
            return rows, notes
        if len(row_template) != len(node[0]):
            # 去掉 _index 这类由程序自动提供的键
            return [row_template], notes
    return node, notes


def field_type(node) -> str:
    if is_table(node):
        return "表格"
    if isinstance(node, dict):
        return "对象"
    if isinstance(node, list):
        return "列表"
    if isinstance(node, bool):
        return "是/否"
    if isinstance(node, (int, float)):
        return "数字"
    text = str(node or "").strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}", text):
        return "日期"
    if text.lower().endswith((".png", ".jpg", ".jpeg", ".gif", ".bmp")) or "图片" in text:
        return "图片"
    return "文本"


def iter_fields(node, prefix=""):
    """摊平字段：返回 [(路径, 节点, 类型)]，表格字段只返回自身，不下钻。"""
    items = []
    if isinstance(node, dict):
        if is_table(node) or not node:
            if prefix:
                items.append((prefix, node, field_type(node)))
            return items
        for key, value in node.items():
            if str(key).startswith("_"):
                continue
            path = "%s.%s" % (prefix, key) if prefix else str(key)
            if isinstance(value, dict) and not is_table(value) and value:
                items.extend(iter_fields(value, path))
            else:
                items.append((path, value, field_type(value)))
    elif prefix:
        items.append((prefix, node, field_type(node)))
    return items


# --------------------------------------------------------------------------
# init：生成待填写的字段说明
# --------------------------------------------------------------------------


def describe_skeleton(node):
    if isinstance(node, dict):
        if is_table(node):
            result = {"_类型": "表格", "_行说明": ""}
            for column in table_columns(node):
                result[column] = ""
            return result
        result = {}
        for key, value in node.items():
            if str(key).startswith("_"):
                continue
            result[key] = describe_skeleton(value)
        return result
    if isinstance(node, list) and node and isinstance(node[0], dict):
        result = {"_类型": "表格", "_行说明": ""}
        for column in table_columns(node):
            result[column] = ""
        return result
    return ""


def cmd_init(args) -> int:
    skeleton = load_json(args.skeleton)
    description = describe_skeleton(skeleton)
    target = args.output or os.path.splitext(os.path.abspath(args.skeleton))[0] + "_字段说明.json"
    dump_json(description, target)
    print("已生成字段说明骨架：%s" % target)
    print("请把每个字段的值改成「这个字段要填什么」的描述（表格字段写在列名上，_行说明写整表的填写要求）。")
    print("改完执行：python gen_skill.py check \"%s\" \"%s\"" % (args.skeleton, target))
    return 0


# --------------------------------------------------------------------------
# check：说明是否覆盖了所有字段
# --------------------------------------------------------------------------


def cmd_check(args) -> int:
    skeleton = load_json(args.skeleton)
    description = load_json(args.description)
    problems, filled, hints = check_descriptions(skeleton, description)
    for item in hints:
        print("[提示] %s" % item)
    for item in problems:
        print("[问题] %s" % item)
    if problems:
        print("字段说明还不完整：%d 个问题" % len(problems))
        return 1
    print("字段说明覆盖完整（%d 个字段），可以生成 skill：python gen_skill.py build ..." % filled)
    return 0


def check_descriptions(skeleton, description):
    """检查字段说明是否覆盖所有字段，返回 (问题列表, 已填数量, 提示列表)。"""
    problems = []
    hints = []
    filled = 0

    for path, node, kind in iter_fields(skeleton):
        parts = path.split(".")
        node_desc = description
        for part in parts:
            if isinstance(node_desc, dict) and part in node_desc:
                node_desc = node_desc[part]
            else:
                node_desc = None
                break
        if node_desc is None:
            problems.append("字段说明里缺少：%s（%s）" % (path, kind))
            continue
        if kind == "表格":
            if not isinstance(node_desc, dict):
                problems.append(
                    "表格 %s 的说明要写成对象（_行说明 + 每列说明），现在写的是：%s"
                    % (path, str(node_desc)[:30])
                )
                continue
            columns = table_columns(node)
            row_note = str(node_desc.get("_行说明", "")).strip()
            empty_columns = [column for column in columns if not str(node_desc.get(column, "")).strip()]
            if not row_note and not any(str(node_desc.get(column, "")).strip() for column in columns):
                problems.append("表格 %s 没有任何填写说明（_行说明 或 列说明）" % path)
            elif empty_columns and not row_note:
                problems.append("表格 %s 这些列还没写说明：%s" % (path, "、".join(empty_columns)))
            else:
                filled += 1
        else:
            if not str(node_desc).strip():
                problems.append("字段 %s 的说明还是空的" % path)
            else:
                filled += 1

    for path, node, kind in iter_fields(skeleton):
        if kind != "表格":
            continue
        desc = describe_for(path, description)
        if not isinstance(desc, dict):
            continue
        wanted = [key for key in desc.keys() if not str(key).startswith("_")]
        current = table_columns(node)
        if wanted and current and wanted != current:
            if is_placeholder_columns(current):
                hints.append(
                    "表格 %s 用的是占位列名（%s），生成 skill 时会按你的说明重建为：%s"
                    % (path, "、".join(current), "、".join(wanted))
                )
            else:
                hints.append(
                    "表格 %s 的列名与说明不一致：骨架=%s，说明=%s（生成时保留骨架列名）"
                    % (path, "、".join(current), "、".join(wanted))
                )

    return problems, filled, hints


# --------------------------------------------------------------------------
# build：生成 skill
# --------------------------------------------------------------------------


def slugify(name: str) -> str:
    text = str(name).strip().lower()
    text = re.sub(r"[\s_]+", "-", text)
    text = re.sub(r"[^a-z0-9\-]", "", text)
    text = re.sub(r"-{2,}", "-", text).strip("-")
    return text[:63]


def describe_for(path: str, description):
    node = description
    for part in path.split("."):
        if isinstance(node, dict) and part in node:
            node = node[part]
        else:
            return None
    return node


def render_field_table(skeleton, description):
    lines = ["| 字段 | 类型 | 填写要求 |", "| --- | --- | --- |"]
    highlights = {"required": [], "tables": [], "optional": []}
    for path, node, kind in iter_fields(skeleton):
        desc = describe_for(path, description)
        if kind == "表格":
            note = ""
            columns = []
            if isinstance(desc, dict):
                note = str(desc.get("_行说明", "")).strip()
                columns = [
                    "%s：%s" % (column, str(desc.get(column, "")).strip())
                    for column in table_columns(node)
                    if str(desc.get(column, "")).strip()
                ]
            requirement = note or "按行填写"
            if columns:
                requirement += "；列说明：" + "；".join(columns)
            highlights["tables"].append(path)
        else:
            requirement = str(desc).strip() if desc is not None else ""
            if isinstance(desc, str) and any(word in requirement for word in ("可空", "选填", "非必填")):
                highlights["optional"].append(path)
            else:
                highlights["required"].append(path)
        requirement = requirement.replace("|", "／").replace("\n", "<br>")
        lines.append("| `%s` | %s | %s |" % (path, kind, requirement or "—"))
    return "\n".join(lines), highlights


def render_highlights(highlights):
    rows = []
    if highlights["tables"]:
        rows.append(
            "- **表格字段（按行增加，列名不要改）**：%s"
            % "、".join("`%s`" % name for name in highlights["tables"])
        )
    if highlights["required"]:
        rows.append("- **必须填**：%s" % "、".join("`%s`" % name for name in highlights["required"]))
    if highlights["optional"]:
        rows.append("- **可空/选填**：%s" % "、".join("`%s`" % name for name in highlights["optional"]))
    return "\n".join(rows) if rows else ""


FIELD_RULES = """1. 契约里写明"取什么、从哪取"的，严格照办，不要自由发挥。
2. 对话、附件、代码里已经出现的技术名词（表名、字段名、接口名、类名、SQL）原样搬进 JSON，不改写。
3. 背景、目标、影响、测试要点这类描述性内容，用上下文里的事实组织成工程文档口吻；不要口号、套话，不要编数字和结论。
4. 没有依据的事实类内容（人名、日期、版本号、金额、数量、SQL）写 `【待确认：xxx】`，并在回复里说明还差什么。
5. 缺信息时一次最多问 3 个关键问题，其余用【待确认】占位，先把可用初稿交出去。"""


def copy_payload(skill_root: str) -> None:
    engine_src = os.path.join(PAYLOAD, "engine")
    engine_dst = os.path.join(skill_root, "scripts")
    if os.path.exists(engine_dst):
        shutil.rmtree(engine_dst)
    shutil.copytree(engine_src, engine_dst)
    refs_src = os.path.join(PAYLOAD, "references")
    refs_dst = os.path.join(skill_root, "references")
    if os.path.exists(refs_dst):
        shutil.rmtree(refs_dst)
    shutil.copytree(refs_src, refs_dst)


def clean_pycache(root: str) -> None:
    for current, dirs, _files in os.walk(root):
        for name in list(dirs):
            if name == "__pycache__":
                shutil.rmtree(os.path.join(current, name), ignore_errors=True)
                dirs.remove(name)


def make_zip(skill_root: str, zip_path: str) -> str:
    """把整个 skill 目录打成 zip，方便拷到别的电脑。"""
    zip_path = os.path.abspath(zip_path)
    if not zip_path.lower().endswith(".zip"):
        zip_path += ".zip"
    os.makedirs(os.path.dirname(zip_path), exist_ok=True)
    base = os.path.splitext(zip_path)[0]
    archive = shutil.make_archive(
        base, "zip", root_dir=os.path.dirname(skill_root), base_dir=os.path.basename(skill_root)
    )
    return archive


def cmd_build(args) -> int:
    skeleton = load_json(args.skeleton)
    description = load_json(args.description)
    skeleton, notes = rebuild_tables(skeleton, description)

    name = slugify(args.name)
    if not name:
        print("[错误] skill 名只能用英文小写字母、数字和连字符，例如 tech-solution-json")
        return 1
    skills_dir = os.path.abspath(args.skills_dir or DEFAULT_SKILLS_DIR)
    skill_root = os.path.join(skills_dir, name)
    if os.path.exists(skill_root) and not args.force:
        print("[错误] 已存在同名 skill：%s（要覆盖请加 --force）" % skill_root)
        return 1
    if os.path.exists(skill_root):
        # --force：先清空，避免上一次生成留下的旧资产混在里面
        shutil.rmtree(skill_root)

    if args.template and not os.path.exists(args.template):
        print("[错误] 找不到 Word 模板：%s" % args.template)
        return 1

    skeleton_name = os.path.basename(args.skeleton)
    skeleton_base = os.path.splitext(skeleton_name)[0]
    config_name = skeleton_base + "_字段说明.json"
    template_asset = os.path.basename(args.template) if args.template else None

    json_name = skeleton_base
    if json_name.lower() in ("骨架", "数据", "模板", "skeleton", "data", "template", "json"):
        json_name = "本 JSON"
    label = re.sub(r"(数据)?(骨架|模板|skeleton|template)$", "", json_name, flags=re.I).strip() or json_name
    display_name = args.display_name or ("%s 数据填写" % label)

    field_table, highlights = render_field_table(skeleton, description)
    description_text = args.description_text or (
        "按字段说明和对话上下文生成「%s」的 JSON：逐字段自动填充内容，表格字段按行增加%s。"
        % (label, "，并用自带的 Word 模板出文档" if template_asset else "")
    )

    bundle = [
        "- `assets/%s` — JSON 结构样板（键名、层级、表格列）" % skeleton_name,
        "- `assets/%s` — 每个字段填什么的说明（字段契约）" % config_name,
    ]
    if template_asset:
        bundle.append("- `assets/%s` — 配套 Word 模板（已打包，生成文档直接用）" % template_asset)
    bundle.append("- `scripts/wf.py` + `scripts/wordfill/` — 校验与填充引擎（自带，不用另装工具）")
    bundle.append("- `references/*.md` — JSON 写法与填写规则")
    bundle_list = "\n".join(bundle)

    if template_asset:
        template_section = (
            "\n## Word 模板\n\n模板已经打包在 `assets/%s`，生成数据后用上面第二条命令出文档；"
            "模板本身不需要分析。用户另外指定模板时，把路径换成那个文件即可。\n" % template_asset
        )
        template_ref = "- `assets/%s` — 配套 Word 模板（生成文档用）\n" % template_asset
        fill_command = (
            "# 用自带的 Word 模板生成文档\n"
            "python scripts/wf.py fill \"assets/%s\" \"数据.json\" -o \"结果.docx\" [--pdf]\n"
            % template_asset
        )
    else:
        template_section = (
            "\n> 本 skill 只负责产出数据 JSON。用户提供 Word 模板时，用\n"
            "> `python scripts/wf.py fill \"<模板.docx>\" \"数据.json\" -o \"结果.docx\"` 生成文档。\n"
        )
        template_ref = ""
        fill_command = ""

    os.makedirs(skill_root, exist_ok=True)
    os.makedirs(os.path.join(skill_root, "agents"), exist_ok=True)
    os.makedirs(os.path.join(skill_root, "assets"), exist_ok=True)
    copy_payload(skill_root)
    if template_asset:
        shutil.copy2(args.template, os.path.join(skill_root, "assets", template_asset))

    with open(os.path.join(PAYLOAD, "SKILL.md.tpl"), encoding="utf-8") as fh:
        skill_template = fh.read()
    for placeholder, value in (
        ("{{SKILL_NAME}}", name),
        ("{{DESCRIPTION}}", description_text.replace("\n", " ")),
        ("{{DISPLAY_NAME}}", display_name),
        ("{{SKELETON_NAME}}", skeleton_name),
        ("{{CONFIG_NAME}}", config_name),
        ("{{BUNDLE_LIST}}", bundle_list),
        ("{{FIELD_TABLE}}", field_table),
        ("{{HIGHLIGHTS}}", render_highlights(highlights)),
        ("{{FIELD_RULES}}", FIELD_RULES),
        ("{{FILL_COMMAND}}", fill_command),
        ("{{TEMPLATE_SECTION}}", template_section),
        ("{{TEMPLATE_REF}}", template_ref),
    ):
        skill_template = skill_template.replace(placeholder, value)
    with open(os.path.join(skill_root, "SKILL.md"), "w", encoding="utf-8") as fh:
        fh.write(skill_template)

    with open(os.path.join(PAYLOAD, "openai.yaml.tpl"), encoding="utf-8") as fh:
        yaml_template = fh.read()
    yaml_template = (
        yaml_template.replace("{{SKILL_NAME}}", name)
        .replace("{{DISPLAY_NAME}}", display_name)
        .replace("{{SHORT_DESCRIPTION}}", "按字段说明和上下文填充 %s" % json_name)
    )
    with open(os.path.join(skill_root, "agents", "openai.yaml"), "w", encoding="utf-8") as fh:
        fh.write(yaml_template)

    dump_json(skeleton, os.path.join(skill_root, "assets", skeleton_name))
    dump_json(description, os.path.join(skill_root, "assets", config_name))
    clean_pycache(skill_root)

    zip_path = None
    if args.zip is not None:
        zip_path = args.zip or os.path.join(skills_dir, name + ".zip")
        zip_path = make_zip(skill_root, zip_path)

    print("已生成 skill：%s" % skill_root)
    print("  字段 %d 个（表格 %d 个）" % (len(iter_fields(skeleton)), len(highlights["tables"])))
    print("  打包内容：assets/%s、assets/%s%s" % (skeleton_name, config_name,
          "、assets/%s（Word 模板）" % template_asset if template_asset else ""))
    print("  引擎与规则：scripts/wf.py + scripts/wordfill/、references/*.md（自带，换台电脑也能跑）")
    if zip_path:
        print("  便携压缩包：%s" % zip_path)
    for note in notes:
        print("  [提示] %s" % note)
    print("校验：python \"%s\" \"%s\"" % (os.path.join(HERE, "gen_skill.py"), skill_root))
    print("用法（新开对话）：$%s 按我的描述和刚才讨论的内容填这份 JSON" % name)
    if template_asset:
        print("  出文档：python scripts/wf.py fill \"assets/%s\" \"数据.json\" -o \"结果.docx\"" % template_asset)
    return 0


# --------------------------------------------------------------------------
# list / validate
# --------------------------------------------------------------------------


def cmd_list(args) -> int:
    skeleton = load_json(args.skeleton)
    for path, _node, kind in iter_fields(skeleton):
        print("%-8s %s" % (kind, path))
    return 0


def build_skill(
    skeleton_path,
    description_path,
    name,
    display_name=None,
    template=None,
    skills_dir=None,
    force=False,
    description_text=None,
    zip_path=None,
):
    """给界面/脚本调用的编程接口，返回 {skill_dir, name, output}。"""
    import contextlib
    import io

    class _Args:
        pass

    args = _Args()
    args.skeleton = skeleton_path
    args.description = description_path
    args.name = name
    args.display_name = display_name
    args.template = template
    args.skills_dir = skills_dir
    args.force = force
    args.description_text = description_text
    args.zip = zip_path

    buffer = io.StringIO()
    with contextlib.redirect_stdout(buffer):
        code = cmd_build(args)
    output = buffer.getvalue()
    if code != 0:
        raise ValueError(output.strip() or "生成失败")
    slug = slugify(name)
    root = os.path.join(os.path.abspath(skills_dir or DEFAULT_SKILLS_DIR), slug)
    archive = None
    if zip_path is not None:
        archive = make_zip(root, zip_path or os.path.join(os.path.dirname(root), slug + ".zip"))
    return {"skill_dir": root, "name": slug, "output": output, "zip": archive}


def cmd_validate(args) -> int:
    root = os.path.abspath(args.skill_dir)
    problems = []
    skill_md = os.path.join(root, "SKILL.md")
    if not os.path.exists(skill_md):
        print("[问题] 缺少 SKILL.md")
        return 1
    with open(skill_md, encoding="utf-8") as fh:
        content = fh.read()
    if not content.startswith("---"):
        problems.append("SKILL.md 缺少 frontmatter")
    for token in ("{{", "}}"):
        if token in content:
            problems.append("SKILL.md 里还有未替换的占位符 %s" % token)
    for required in (
        "SKILL.md",
        "agents/openai.yaml",
        "references/json-rules.md",
        "references/how-to-fill.md",
        "references/troubleshooting.md",
        "scripts/wf.py",
        "scripts/env_check.py",
        "scripts/wordfill/engine.py",
    ):
        if not os.path.exists(os.path.join(root, required)):
            problems.append("缺少文件：%s" % required)
    assets_dir = os.path.join(root, "assets")
    json_assets = []
    templates = []
    if os.path.isdir(assets_dir):
        for item in sorted(os.listdir(assets_dir)):
            if item.lower().endswith(".json"):
                json_assets.append(item)
            elif item.lower().endswith((".docx", ".dotx")):
                templates.append(item)
    else:
        problems.append("缺少 assets 目录")
    if len(json_assets) < 2:
        problems.append("assets 里应该有 JSON 骨架和字段说明两份 JSON，现在只有：%s" % (json_assets or "无"))
    for item in templates:
        if item not in content:
            problems.append("assets 里有 Word 模板 %s，但 SKILL.md 没有引用它" % item)
    if re.search(r"[A-Za-z]:[\\/]", content):
        problems.append("SKILL.md 里还有本机绝对路径（如 C:\\...），换台电脑会失效")
    name = re.search(r"^name:\s*(\S+)", content, re.M)
    if not name or not re.fullmatch(r"[a-z0-9][a-z0-9\-]*", name.group(1)):
        problems.append("frontmatter 里的 name 不合法（只能小写字母、数字、连字符）")
    for item in problems:
        print("[问题] %s" % item)
    if problems:
        return 1
    print("skill 结构正常：%s" % root)
    print(
        "  自带资源：%s%s"
        % ("、".join(json_assets), ("、" + "、".join(templates)) if templates else "")
    )
    return 0


def build_parser():
    parser = argparse.ArgumentParser(prog="gen_skill", description="JSON 字段 Skill 生成器")
    sub = parser.add_subparsers(dest="command")

    init = sub.add_parser("init", help="从 JSON 骨架生成待填写的字段说明")
    init.add_argument("skeleton")
    init.add_argument("-o", "--output")
    init.set_defaults(func=cmd_init)

    check = sub.add_parser("check", help="检查字段说明是否覆盖所有字段")
    check.add_argument("skeleton")
    check.add_argument("description")
    check.set_defaults(func=cmd_check)

    build = sub.add_parser("build", help="生成 skill")
    build.add_argument("skeleton")
    build.add_argument("description")
    build.add_argument("--name", required=True, help="skill 名（英文小写+连字符）")
    build.add_argument("--display-name", help="界面上显示的名称")
    build.add_argument("--skills-dir", help="skill 安装目录（默认 ~/.codex/skills）")
    build.add_argument("--template", help="配套的 Word 模板路径（可选）")
    build.add_argument("--description-text", help="skill 的 description（可选）")
    build.add_argument(
        "--zip",
        nargs="?",
        const="",
        help="同时打成 zip（方便拷到别的电脑）；不给路径就放在 skill 目录旁边",
    )
    build.add_argument("--force", action="store_true", help="覆盖同名 skill")
    build.set_defaults(func=cmd_build)

    listing = sub.add_parser("list", help="列出骨架里的字段")
    listing.add_argument("skeleton")
    listing.set_defaults(func=cmd_list)

    validate = sub.add_parser("validate", help="检查生成的 skill 结构")
    validate.add_argument("skill_dir")
    validate.set_defaults(func=cmd_validate)

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
