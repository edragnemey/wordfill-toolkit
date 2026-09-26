---
name: {{SKILL_NAME}}
description: {{DESCRIPTION}}
---

# {{DISPLAY_NAME}}

你只做一件事：**产出一份数据 JSON**，结构以 `assets/{{SKELETON_NAME}}` 为准，每个字段的内容按下面的字段契约和对话上下文来写。

## 这个 skill 是自包含的

所有要用的文件都在本 skill 目录里，整个文件夹拷到别的电脑也能直接用：

{{BUNDLE_LIST}}

运行环境：Python 3.10 以上 + `python-docx`（`pip install python-docx`）。
填充引擎已经打包在 `scripts/` 里，**不需要**另外安装工具或依赖别的 skill。
换机器后先跑一次环境自检：`python scripts/env_check.py`。

## 字段契约（唯一权威）

{{FIELD_TABLE}}

完整说明见 `assets/{{CONFIG_NAME}}`。**不要**为了搞清楚字段含义去分析 Word 模板或猜结构——契约里写的就是全部要求。

{{HIGHLIGHTS}}

## 怎么填

{{FIELD_RULES}}

通用填写规则见 [references/how-to-fill.md](references/how-to-fill.md)，JSON 写法（格式修饰符、表格、循环、条件、图片）见 [references/json-rules.md](references/json-rules.md)。

## 表格 / 数组字段

- 表格字段按行增加：每个对象一行，列的键名与契约一致，不要自己改列名。
- 行数由真实内容决定，不凑数、不留空行；一行内容都没有时给空数组 `[]`，并在回复里说明。
- 表头顺序、列宽、列格式、合并列按契约里的写法给，例如
  `"表头": {"表名": "4cm", "字段": "3.1cm"}, "合并列": ["表名"]`。
- 多行文本（SQL、代码、多条说明）用 `\n` 换行。

## 校验与交付

```
# 结构校验（不需要 Word 模板）
python scripts/wf.py check --keys-from "assets/{{SKELETON_NAME}}" "数据.json" --config "assets/{{CONFIG_NAME}}"
{{FILL_COMMAND}}```

`check` 退出码为 0 才算完成：缺键、空值、表格结构不对都会被报出来。
交付时用一两句话说明每个字段的来源（上下文 / 附件 / 代码 / 推断 / 待确认），不要把全文贴进回复。
{{TEMPLATE_SECTION}}
## 参考文件

- `assets/{{SKELETON_NAME}}` — 结构样板（键名、层级、表格列）
- `assets/{{CONFIG_NAME}}` — 每个字段填什么的完整说明
{{TEMPLATE_REF}}
- [references/how-to-fill.md](references/how-to-fill.md) — 内容生成规则：什么能推断、什么必须问、不许编造
- [references/json-rules.md](references/json-rules.md) — JSON 写法完整规则
- [references/troubleshooting.md](references/troubleshooting.md) — 字段为空、表格没生成、换行丢失等排查
