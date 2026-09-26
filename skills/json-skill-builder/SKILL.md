---
name: json-skill-builder
description: 由一份 JSON 骨架和逐字段的填写描述，生成一个专用填充 skill —— 生成的 skill 里已经内置 JSON 编写规则、引擎、骨架和字段说明，AI 之后按描述和对话上下文自动填字段值、按行增加表格数据。用户说"把我的 JSON 做成 skill""按我的字段描述生成一个 skill""让它以后自动帮我填这份 JSON"时使用。
---

# JSON 字段 Skill 生成器

把「一份 JSON 骨架 + 你写的字段描述」变成一个专用 skill。
生成的 skill 里已经**预制好 JSON 编写规则**（格式修饰符、表格、循环、条件、图片、合并单元格）、
内置填充引擎、并带上你的骨架和字段说明；以后 AI 只要拿到需求上下文，就能按描述把这份 JSON 填出来。

## 先看这里：有没有图形界面

**有。** 用户不喜欢写 JSON 的时候，让他双击 skill 目录下的 `一键启动.bat`，
或者你直接运行：

```
python "<本 skill 目录>/scripts/skill_builder_gui.py"
```

macOS / Linux：双击 `启动填写界面.command`（首次可能要 `chmod +x` 一次），
或终端执行 `python3 "<本 skill 目录>/scripts/skill_builder_gui.py"`。
`.bat` 只能在 Windows 用。

界面里做的事：选 JSON 骨架（或直接从 Word 模板生成骨架）→ **一个字段一个输入框地写"这个字段填什么"**，
表格字段还能改列名、加列 → 检查漏填 → 填 skill 名 → 生成 skill。
描述自动存到骨架同目录的 `xxx_字段说明.json`，下次打开继续编辑。

给用户的操作说明就一句话：双击 `一键启动.bat`，从上往下填，最后点「生成 skill」。

## 三个输入

1. **JSON 骨架**：要长期复用的那份 JSON（键名、层级、表格结构）。用户可能直接给文件，也可能贴内容 —— 贴内容就落成一个 `.json` 文件再用。
2. **字段描述**：每个字段该填什么。用户可能写成文件，也可能只是在对话里说 ——
   例如"prd_name 填需求名称""改造清单这张表按工程合并，列是工程/类/说明/类型"。
   你要么用 `init` 生成待填骨架让他补，要么直接把他的话整理成 `字段说明.json`。
3. **skill 名**：英文小写 + 连字符（如 `tech-solution-json`、`prd-filler`）。用户没给就按 JSON 用途起一个并告诉他。

## 工作流

有界面时：让用户用 `一键启动.bat` 自己填（推荐，尤其是"一个个字段说描述"的场景）。
没有界面或要批量处理时，走下面的命令行：

```bash
# 1) 生成待填写的字段说明（也可以直接手写/由你生成这份文件）
python scripts/gen_skill.py init "骨架.json" -o "字段说明.json"

# 2) 逐字段补描述：表格字段写 _行说明（整表要求）+ 每列的说明

# 3) 检查说明是否覆盖了所有字段
python scripts/gen_skill.py check "骨架.json" "字段说明.json"

# 4) 生成 skill（默认装到 ~/.codex/skills/<name>）
python scripts/gen_skill.py build "骨架.json" "字段说明.json" --name "tech-solution-json" \
    [--display-name "技术方案数据填充"] [--template "技术方案模板.docx"] \
    [--zip "输出目录/tech-solution-json.zip"] [--force]

# 5) 结构自检
python scripts/gen_skill.py validate "%USERPROFILE%\.codex\skills\tech-solution-json"
```

生成后告诉用户：**新开一个对话**（skill 列表在会话开始时加载），用
`$<skill名> 按我的描述和刚才讨论的内容填这份 JSON` 即可。

## 生成的 skill 是自包含的（可搬走）

用户给了 Word 模板，就必须**把模板文件复制进 skill 里**，不能只在文档里记一个本机路径——
否则换台电脑就失效。`build --template` 会自动做这件事：模板拷进 `assets/`，
`SKILL.md` 里写成相对路径（如 `assets/技术方案模板.docx`），并附上出文档的命令。

```
<skill名>/
├─ SKILL.md                  字段契约 + 自带文件清单 + 运行环境 + 命令
├─ assets/xxx_数据骨架.json    JSON 结构样板          ← 你的 JSON 模板
├─ assets/xxx_字段说明.json    每个字段填什么          ← 你的字段描述
├─ assets/xxx.docx            Word 模板              ← 你的模板（--template 自动打包）
├─ references/*.md            JSON 写法与填写规则（预制）
└─ scripts/wf.py + wordfill/ + env_check.py   校验/填充引擎（自带）
```

- 模板、骨架、字段说明三份文件都在 skill 里，**整个文件夹拷到别的电脑就能用**；
- 引擎和规则一起打包，不依赖 `wordfill` 或任何外部工具；
- 新机器只需要 Python 3.10+ 和 `python-docx`，先跑 `python scripts/env_check.py` 自检；
- 要发给同事时加 `--zip`，生成 `<skill名>.zip`，解压到 `~/.codex/skills/` 即用；
- `validate` 会检查模板是否真的在 assets 里、SKILL.md 里是否还残留 `C:\...` 本机路径。

## 字段说明的格式

与 JSON 骨架同构，叶子节点写"这里填什么"；表格字段是一个对象：

```json
{
  "prd_name": "需求名称，取用户这次需求的标题，20 字以内",
  "编制日期": "编写日期，格式 YYYY-MM-DD；不知道就写【待确认：编制日期】",
  "改造清单": {
    "_类型": "表格",
    "_行说明": "每个改造点一行；没有的不要凑数",
    "工程": "工程名，如 dceppay",
    "模块或类": "类名，取真实代码里的名字",
    "改造说明": "本处改了什么，接口名原样保留",
    "类型": "新增 / 修改 / 删除"
  }
}
```

可选的下划线键（不会被写进文档，只作说明）：

- `_类型`：`表格` / `列表`，用于骨架里看不出结构时；
- `_行说明`：整个表格怎么填（有它就够了，列说明可省略）；
- `_必填`、`_来源`、`_示例`：补充要求（"来源"写 对话/附件/代码/用户确认）。

字段描述里写清楚"取什么、从哪取、格式或长度要求"最有用；写"按要求填写"这类空话没有价值。

## 用户只在对话里描述时

不用强求他填文件。把他说的整理成 `字段说明.json`，逐条对照骨架看有没有漏，
再把整理结果回给他确认（尤其是：表格的列、哪些字段可空、哪些必须来自代码或附件）。
确认后按上面第 3 步继续。

## 生成的 skill 里有什么

```
<skill名>/
├─ SKILL.md                  字段契约表（从你的描述渲染）+ 填写规则 + 命令
├─ agents/openai.yaml        界面显示名
├─ assets/数据骨架.json       结构样板
├─ assets/字段说明.json       你的逐字段描述（唯一权威）
├─ references/
│   ├─ json-rules.md         JSON 编写规则（预制）
│   ├─ how-to-fill.md        内容生成规则：什么能推断、什么必须问、不许编造
│   └─ troubleshooting.md    排查
└─ scripts/wf.py + wordfill/ 校验 / 填充引擎
```

生成的 skill 只关心 JSON；有 Word 模板时用 `--template` 记下路径，它就能顺带出 docx。

## 注意

- skill 名必须是英文小写字母、数字、连字符；重复生成同一个名字要加 `--force`。
- 不要为了写得更好而擅自改字段名或表格列名 —— 架构、下游脚本可能依赖它们。
- 生成的 skill 属于用户的产出，不要顺手改动别的已有 skill。
