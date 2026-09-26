# -*- coding: utf-8 -*-
"""Word 模板编译器 图形界面。"""

from __future__ import annotations

import json
import os
import sys
import traceback
from tkinter import BOTH, END, LEFT, RIGHT, X, BooleanVar, StringVar, Text, Tk, filedialog, messagebox
from tkinter import ttk

from .engine import (
    MissingFieldsError,
    build_skeleton,
    export_pdf,
    fill_template,
    load_json,
    scan_template,
)

APP_TITLE = "Word 模板编译器"

HELP_TEXT = """三步生成文档

1. 在 Word 里写模板
   把需要变化的地方写成 {{字段名}} 这样的标记，例如 甲方：{{甲方名称}}。
   标记可以随便改字号、加粗，替换时会保留原来的格式。

2. 准备一份 JSON 数据
   点「生成数据模板」会自动扫描模板里的字段，生成一份 JSON 骨架，
   用记事本把里面的空值改成真实内容即可。

3. 点「开始生成文档」
   选好模板和 JSON，点生成，就会得到填好内容的 Word 文档。

常用写法速查

  {{姓名}}                     普通字段
  {{客户.电话}}                多级字段（JSON 里嵌套一层）
  {{金额|千分位}}              1,234,567.89
  {{金额|0.00}}               保留两位小数
  {{金额|人民币大写}}           壹佰贰拾叁万肆仟伍佰陆拾柒元捌角玖分
  {{日期|yyyy年M月d日}}        2026年9月17日
  {{日期|yyyy-MM-dd}}         2026-09-17
  {{姓名|默认:张三}}            数据里没有时用默认值
  {{@公章}}                    插入图片，JSON 里填图片路径
  {{@公章|宽=4cm}}             指定图片宽度
  {{表格:明细}}                在这个位置插入一整张表格，表头/列数/行数都由 JSON 决定
  {{表格:明细|无表头}}          同上，但不显示表头
  {{$today}} / {{$now}}        自动填当天日期 / 当前时间
  {{#if 加急}}…{{else}}…{{/if}}       条件内容，整段或表格行都可以
  {{#each 员工}}…{{/each}}            重复内容，整段或表格行都可以

小技巧

  · 表格里做循环：把 {{#each 员工}} 放在第一行的某个单元格、{{/each}} 放在同一行
    最后一个单元格（或者单独占一行），中间写 {{姓名}} 这类字段，生成时会自动复制行。
  · 表格行数（甚至列数、表头）都要跟着数据变：模板里只写一行 {{表格:员工}}，
    表格数据写在 JSON 里，例如
      "员工": { "表头": {"姓名": "3cm", "月薪": "2.5cm"},
                "列格式": {"月薪": "千分位"},
                "合并列": ["姓名"],
                "行": [ {"姓名": "李四", "月薪": 18600} ] }
    也可以直接写成 [ {"姓名": "李四", "月薪": 18600} ]，表头会自动取数据的键名。
  · 循环里可以用 {{_index}} 拿到序号（从 1 开始）。
  · 条件判断支持 {{#if 金额 > 1000}}、{{#if 状态 == 已签}}、{{#if !备注}} 这类写法。
  · 数据里没提供的字段，默认原样保留 {{字段名}}，方便检查缺了什么；
    勾选「缺失字段留空」就替换成空白。
  · Word 表单域（内容控件）也能填：给控件设置标题或标记等于 JSON 的键名即可。
"""


def _apply_windows_dpi() -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:  # noqa: BLE001
        pass


def _pick_font() -> str:
    try:
        from tkinter import font as tkfont

        families = set(tkfont.families())
    except Exception:  # noqa: BLE001
        return "TkDefaultFont"
    for name in ("Microsoft YaHei UI", "微软雅黑", "Microsoft YaHei", "PingFang SC", "SimHei"):
        if name in families:
            return name
    return "TkDefaultFont"


def _preview(value) -> str:
    if isinstance(value, dict):
        return "包含 %d 项" % len(value)
    if isinstance(value, list):
        if not value:
            return "空列表"
        first = value[0]
        if isinstance(first, dict):
            return "%d 条记录，字段：%s" % (len(value), "、".join(list(first.keys())[:4]))
        return "%d 项：%s" % (len(value), "、".join(str(item) for item in value[:3]))
    text = str(value)
    return text if len(text) <= 40 else text[:40] + "…"


USAGE_KINDS = [
    "{{字段}}　普通字段",
    "{{字段|默认:默认值}}　带默认值",
    "{{字段|0.00}}　保留两位小数",
    "{{字段|千分位}}　千分位",
    "{{字段|人民币大写}}　人民币大写",
    "{{字段|yyyy年M月d日}}　中文日期",
    "{{字段|yyyy-MM-dd}}　日期",
    "{{@字段}}　图片",
    "{{@字段|宽=4cm}}　图片（指定宽度）",
    "{{表格:字段}}　整张表由数据生成",
    "{{表格:字段|无表头}}　整张表（不带表头）",
    "{{表格:字段|合并首列}}　整张表（首列相同内容合并）",
    "{{#if 字段}}内容{{/if}}　条件",
    "{{#if 字段}}内容{{else}}否则{{/if}}　条件（带否则）",
    "{{#each 字段}}内容{{/each}}　循环",
    "{{_index}}　循环序号",
    "{{$today}}　今天的日期",
]


class App(Tk):
    """主窗口。"""

    def __init__(self, initial_template=None):
        super().__init__()
        _apply_windows_dpi()
        self.title(APP_TITLE)
        self.geometry("1040x740")
        self.minsize(900, 620)

        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except Exception:  # noqa: BLE001
            pass
        font_name = _pick_font()
        for style_name in ("TLabel", "TButton", "TEntry", "TCheckbutton", "TLabelframe.Label", "Treeview", "TNotebook.Tab"):
            try:
                style.configure(style_name, font=(font_name, 10))
            except Exception:  # noqa: BLE001
                pass
        style.configure("Title.TLabel", font=(font_name, 16, "bold"))
        style.configure("Hint.TLabel", font=(font_name, 9), foreground="#555555")
        style.configure("Go.TButton", font=(font_name, 11, "bold"))
        self.option_add("*Font", (font_name, 10))

        self.template_path = StringVar()
        self.data_path = StringVar()
        self.output_path = StringVar()
        self.field_name = StringVar()
        self.usage_kind = StringVar(value=USAGE_KINDS[0])
        self.status_text = StringVar(value="请选择模板文件和数据文件")
        self.option_blank = BooleanVar(value=False)
        self.option_pdf = BooleanVar(value=False)
        self.option_open = BooleanVar(value=True)
        self.option_controls = BooleanVar(value=True)
        self._template_info = None
        self._data = None

        self._build_header()
        notebook = ttk.Notebook(self)
        notebook.pack(fill=BOTH, expand=True, padx=12, pady=(0, 6))
        self._build_fill_tab(notebook)
        self._build_helper_tab(notebook)
        self._build_help_tab(notebook)
        self._build_menu()

        ttk.Label(self, textvariable=self.status_text, anchor="w", padding=(14, 6)).pack(fill=X, side="bottom")
        if initial_template:
            self.use_template(initial_template)

    # -- 界面骨架 ----------------------------------------------------------

    def _build_header(self) -> None:
        header = ttk.Frame(self, padding=(16, 12, 16, 6))
        header.pack(fill=X)
        ttk.Label(header, text=APP_TITLE, style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="在 Word 里写 {{字段}}，用一份 JSON 数据一次性替换全部内容，版式和格式保持不变。",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(4, 0))

    def _build_menu(self) -> None:
        from tkinter import Menu

        menubar = Menu(self)
        file_menu = Menu(menubar, tearoff=0)
        file_menu.add_command(label="选择模板文件…", command=self.pick_template)
        file_menu.add_command(label="选择数据文件…", command=self.pick_data)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.destroy)
        menubar.add_cascade(label="文件", menu=file_menu)
        help_menu = Menu(menubar, tearoff=0)
        help_menu.add_command(label="关于", command=self.show_about)
        menubar.add_cascade(label="帮助", menu=help_menu)
        self.config(menu=menubar)

    def _build_fill_tab(self, notebook: ttk.Notebook) -> None:
        frame = ttk.Frame(notebook, padding=14)
        notebook.add(frame, text="  ① 填充文档  ")
        frame.columnconfigure(1, weight=1)

        ttk.Label(
            frame,
            text="① 选要填充的 Word 文件（模板）　→　② 选数据 JSON　→　③ 点最下面的「开始生成文档」",
            style="Hint.TLabel",
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 8))

        ttk.Label(frame, text="Word 模板：").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.template_path).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Button(frame, text="浏览…", command=self.pick_template).grid(row=1, column=2, padx=(8, 0))

        ttk.Label(frame, text="数据文件：").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.data_path).grid(row=2, column=1, sticky="ew", pady=4)
        ttk.Button(frame, text="浏览…", command=self.pick_data).grid(row=2, column=2, padx=(8, 0))
        ttk.Button(frame, text="查看/编辑", command=self.edit_data).grid(row=2, column=3, padx=(8, 0))

        actions = ttk.Frame(frame)
        actions.grid(row=3, column=0, columnspan=4, sticky="ew", pady=(8, 10))
        ttk.Button(actions, text="读取字段并检查", command=self.read_fields).pack(side=LEFT)
        ttk.Button(actions, text="生成数据模板", command=self.save_skeleton).pack(side=LEFT, padx=8)
        ttk.Button(actions, text="只看模板字段", command=self.scan_only).pack(side=LEFT)

        tree = ttk.Treeview(frame, columns=("field", "state", "value"), show="headings", height=11)
        tree.heading("field", text="字段 / 循环")
        tree.heading("state", text="状态")
        tree.heading("value", text="数据内容")
        tree.column("field", width=230, anchor="w")
        tree.column("state", width=90, anchor="center")
        tree.column("value", width=560, anchor="w")
        tree.tag_configure("ok", foreground="#137333")
        tree.tag_configure("missing", foreground="#c5221f")
        tree.tag_configure("group", foreground="#1a4b8c")
        tree.grid(row=4, column=0, columnspan=4, sticky="nsew", pady=(0, 8))
        frame.rowconfigure(4, weight=1)
        scroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        scroll.grid(row=4, column=4, sticky="ns")
        tree.configure(yscrollcommand=scroll.set)
        self.tree = tree

        ttk.Label(frame, text="输出文件：").grid(row=5, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.output_path).grid(row=5, column=1, sticky="ew", pady=4)
        ttk.Button(frame, text="另存为…", command=self.pick_output).grid(row=5, column=2, padx=(8, 0))

        options = ttk.Frame(frame)
        options.grid(row=6, column=0, columnspan=4, sticky="w", pady=(6, 0))
        ttk.Checkbutton(options, text="生成后打开文档", variable=self.option_open).pack(side=LEFT)
        ttk.Checkbutton(options, text="同时导出 PDF", variable=self.option_pdf).pack(side=LEFT, padx=12)
        ttk.Checkbutton(options, text="缺失字段留空", variable=self.option_blank).pack(side=LEFT)
        ttk.Checkbutton(options, text="填充 Word 表单域", variable=self.option_controls).pack(side=LEFT, padx=12)

        ttk.Button(frame, text="开始生成文档", style="Go.TButton", command=self.generate).grid(
            row=7, column=0, columnspan=4, sticky="ew", pady=(14, 0), ipady=6
        )

    def _build_helper_tab(self, notebook: ttk.Notebook) -> None:
        frame = ttk.Frame(notebook, padding=14)
        notebook.add(frame, text="  ② 字段助手  ")
        frame.columnconfigure(0, weight=1)

        ttk.Label(
            frame,
            text="做模板时用这里生成字段标记：填好字段名、选好用法，点复制，回到 Word 里 Ctrl+V 粘贴即可。",
            style="Hint.TLabel",
        ).grid(row=0, column=0, sticky="w")

        row = ttk.Frame(frame)
        row.grid(row=1, column=0, sticky="ew", pady=12)
        ttk.Label(row, text="字段名：").pack(side=LEFT)
        ttk.Entry(row, textvariable=self.field_name, width=22).pack(side=LEFT)
        ttk.Combobox(
            row, textvariable=self.usage_kind, values=USAGE_KINDS, width=36, state="readonly"
        ).pack(side=LEFT, padx=8)
        ttk.Button(row, text="复制标记", command=self.copy_marker).pack(side=LEFT)

        ttk.Separator(frame, orient="horizontal").grid(row=2, column=0, sticky="ew", pady=6)

        box = ttk.Frame(frame)
        box.grid(row=3, column=0, sticky="nsew")
        box.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)
        ttk.Label(box, text="当前模板里已有的字段（双击即可复制标记）：").grid(row=0, column=0, sticky="w")
        self.field_list = ttk.Treeview(box, columns=("name", "kind", "where"), show="headings", height=11)
        self.field_list.heading("name", text="标记写法")
        self.field_list.heading("kind", text="类型")
        self.field_list.heading("where", text="位置")
        self.field_list.column("name", width=280, anchor="w")
        self.field_list.column("kind", width=110, anchor="center")
        self.field_list.column("where", width=380, anchor="w")
        self.field_list.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
        self.field_list.bind("<Double-1>", self.copy_selected_field)
        box.rowconfigure(1, weight=1)

        bar = ttk.Frame(box)
        bar.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(bar, text="扫描当前模板", command=self.scan_only).pack(side=LEFT)
        ttk.Button(bar, text="生成数据模板", command=self.save_skeleton).pack(side=LEFT, padx=8)

    def _build_help_tab(self, notebook: ttk.Notebook) -> None:
        frame = ttk.Frame(notebook, padding=14)
        notebook.add(frame, text="  ③ 使用说明  ")
        text = Text(frame, wrap="word", relief="flat", background="#fbfbfb", padx=12, pady=10)
        text.pack(fill=BOTH, expand=True)
        text.insert("1.0", HELP_TEXT)
        text.configure(state="disabled")

    # -- 交互 --------------------------------------------------------------

    def set_status(self, message: str) -> None:
        self.status_text.set(message)
        self.update_idletasks()

    def show_about(self) -> None:
        messagebox.showinfo(
            APP_TITLE,
            "Word 模板编译器\n\n把 Word 文档里的 {{特殊字段}} 用 JSON 数据替换掉，\n版式保持不变，支持整段循环、条件内容和图片。",
        )

    def pick_template(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Word 模板",
            filetypes=[("Word 文档", "*.docx *.dotx *.docm *.dotm"), ("所有文件", "*.*")],
            initialdir=os.path.dirname(self.template_path.get()) or None,
        )
        if path:
            self.use_template(path)

    def use_template(self, path: str) -> None:
        """选定模板文件，并自动推算输出文件名。"""
        if not path or not os.path.exists(path):
            return
        self.template_path.set(path)
        stem, ext = os.path.splitext(path)
        self.output_path.set("%s_已填充%s" % (stem, ext or ".docx"))
        self.title("%s - %s" % (APP_TITLE, os.path.basename(path)))
        self.set_status("当前模板：%s　请再选择 JSON 数据文件" % path)
        self.read_fields()

    def pick_data(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 JSON 数据文件",
            filetypes=[("JSON 数据", "*.json"), ("所有文件", "*.*")],
            initialdir=os.path.dirname(self.data_path.get()) or None,
        )
        if path:
            self.data_path.set(path)
            self.read_fields()

    def pick_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="保存为",
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")],
            initialfile=os.path.basename(self.output_path.get() or "生成结果.docx"),
        )
        if path:
            self.output_path.set(path)

    def load_data(self):
        path = self.data_path.get().strip()
        if not path:
            messagebox.showwarning(APP_TITLE, "请先选择 JSON 数据文件。")
            return None
        if not os.path.exists(path):
            messagebox.showwarning(APP_TITLE, "找不到数据文件：\n%s" % path)
            return None
        try:
            data = load_json(path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_TITLE, "JSON 读取失败：\n%s" % exc)
            return None
        if not isinstance(data, dict):
            messagebox.showerror(APP_TITLE, "JSON 最外层必须是对象，例如：\n{\n    \"姓名\": \"张三\"\n}")
            return None
        self._data = data
        return data

    def scan_only(self) -> None:
        path = self.template_path.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning(APP_TITLE, "请先选择 Word 模板文件。")
            return
        self.set_status("正在扫描模板字段…")
        try:
            self._template_info = scan_template(path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_TITLE, "扫描失败：\n%s" % exc)
            return
        self.fill_field_list()
        info = self._template_info
        if info.is_empty:
            messagebox.showinfo(
                APP_TITLE,
                "这个文档里还没有特殊字段。\n\n在 Word 里写上 {{字段名}} 这样的标记后重新扫描即可。",
            )
        self.set_status(
            "扫描完成：%d 个字段、%d 个循环、%d 个条件"
            % (len(info.fields), len(info.loops), len(info.conditions))
        )

    def fill_field_list(self) -> None:
        for item in self.field_list.get_children():
            self.field_list.delete(item)
        info = self._template_info
        if info is None:
            return
        for field in info.fields:
            marker = "{{@%s}}" % field.name if field.kind == "image" else "{{%s}}" % field.name
            kind = "图片" if field.kind == "image" else ("列表" if field.kind == "list" else "字段")
            self.field_list.insert("", END, values=(marker, kind, field.location))
        for loop in info.loops:
            self.field_list.insert("", END, values=("{{#each %s}}" % loop.name, "循环", loop.location))
        for name in info.images:
            self.field_list.insert("", END, values=("{{@%s}}" % name, "图片", "自动插入"))
        for name in info.tables:
            self.field_list.insert("", END, values=("{{表格:%s}}" % name, "表格", "整张表自动生成"))
        for name in info.builtins:
            self.field_list.insert("", END, values=("{{$%s}}" % name, "内置", "自动填充"))

    @staticmethod
    def _lookup(data, path: str):
        node = data
        for part in (path or "").split("."):
            if isinstance(node, dict) and part in node:
                node = node[part]
            else:
                return False, None
        return True, node

    def read_fields(self) -> None:
        path = self.template_path.get().strip()
        data_path = self.data_path.get().strip()
        data = None
        if data_path:
            if not os.path.exists(data_path):
                messagebox.showwarning(APP_TITLE, "找不到数据文件：\n%s" % data_path)
            else:
                try:
                    data = load_json(data_path)
                except Exception as exc:  # noqa: BLE001
                    messagebox.showerror(APP_TITLE, "JSON 读取失败：\n%s" % exc)
        if not path or not os.path.exists(path):
            if data is not None:
                self.set_status("已读取数据文件")
            return
        try:
            self._template_info = scan_template(path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_TITLE, "扫描模板失败：\n%s" % exc)
            return
        self.fill_field_list()
        for item in self.tree.get_children():
            self.tree.delete(item)
        info = self._template_info
        total = 0
        missing = 0
        if info.is_empty:
            self.tree.insert("", END, values=("（没有发现特殊字段）", "", "在 Word 里写 {{字段名}} 即可"), tags=("group",))
        entries = [(field.name, "字段", field.name) for field in info.fields]
        entries += [("循环 %s" % loop.name, "循环", loop.name) for loop in info.loops]
        entries += [("图片 %s" % name, "图片", name) for name in info.images]
        entries += [("表格 %s" % name, "表格", name) for name in info.tables]
        for label, kind, key in entries:
            total += 1
            present = False
            value = None
            if data is not None:
                present, value = self._lookup(data, key)
            if not present:
                missing += 1
            self.tree.insert(
                "",
                END,
                values=(label, "已填" if present else "缺数据", _preview(value) if present else ""),
                tags=("ok" if present else "missing",),
            )
        if data is None:
            self.set_status("已列出模板字段；选择 JSON 数据文件后可以看到填写情况")
        else:
            self.set_status("共 %d 个字段，其中 %d 个缺数据" % (total, missing))

    def copy_marker(self) -> None:
        name = self.field_name.get().strip() or "字段名"
        template = self.usage_kind.get().split("　")[0]
        marker = template.replace("字段", name)
        self.clipboard_clear()
        self.clipboard_append(marker)
        self.set_status("已复制：%s　回到 Word 里 Ctrl+V 粘贴即可" % marker)

    def copy_selected_field(self, _event=None) -> None:
        selection = self.field_list.selection()
        if not selection:
            return
        marker = self.field_list.item(selection[0], "values")[0]
        self.clipboard_clear()
        self.clipboard_append(marker)
        self.set_status("已复制：%s" % marker)

    def save_skeleton(self) -> None:
        path = self.template_path.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning(APP_TITLE, "请先选择 Word 模板文件。")
            return
        try:
            info = scan_template(path)
            skeleton = build_skeleton(info)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_TITLE, "生成数据模板失败：\n%s" % exc)
            return
        if not skeleton:
            messagebox.showinfo(APP_TITLE, "模板里没有发现字段，先在 Word 里写 {{字段名}} 吧。")
            return
        target = filedialog.asksaveasfilename(
            title="保存数据模板",
            defaultextension=".json",
            filetypes=[("JSON 数据", "*.json")],
            initialfile=os.path.splitext(os.path.basename(path))[0] + "_数据.json",
        )
        if not target:
            return
        with open(target, "w", encoding="utf-8") as fh:
            json.dump(skeleton, fh, ensure_ascii=False, indent=2)
        self.data_path.set(target)
        self.set_status("数据模板已生成：%s" % target)
        if messagebox.askyesno(APP_TITLE, "数据模板已生成：\n%s\n\n现在打开编辑吗？" % target):
            self.edit_data()

    def edit_data(self) -> None:
        from tkinter import Toplevel

        path = self.data_path.get().strip()
        if not path:
            path = filedialog.asksaveasfilename(
                title="新建数据文件",
                defaultextension=".json",
                filetypes=[("JSON 数据", "*.json")],
                initialfile="数据.json",
            )
            if not path:
                return
            self.data_path.set(path)

        window = Toplevel(self)
        window.title("编辑数据文件 - %s" % os.path.basename(path))
        window.geometry("760x640")
        window.transient(self)
        editor = Text(window, wrap="none", undo=True)
        editor.pack(fill=BOTH, expand=True, padx=10, pady=(10, 0))

        content = ""
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8-sig") as fh:
                    content = fh.read()
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror(APP_TITLE, "读取失败：%s" % exc)
                window.destroy()
                return
        elif self.template_path.get().strip():
            try:
                content = json.dumps(
                    build_skeleton(scan_template(self.template_path.get().strip())),
                    ensure_ascii=False,
                    indent=2,
                )
            except Exception:  # noqa: BLE001
                content = "{\n  \n}"
        editor.insert("1.0", content)

        bar = ttk.Frame(window)
        bar.pack(fill=X, padx=10, pady=10)

        def format_json() -> None:
            try:
                payload = json.loads(editor.get("1.0", END))
                editor.delete("1.0", END)
                editor.insert("1.0", json.dumps(payload, ensure_ascii=False, indent=2))
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror(APP_TITLE, "JSON 格式有误：\n%s" % exc, parent=window)

        def save() -> None:
            body = editor.get("1.0", END).strip()
            if not body:
                messagebox.showwarning(APP_TITLE, "内容是空的。", parent=window)
                return
            try:
                payload = json.loads(body)
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror(APP_TITLE, "JSON 格式有误，先修好再保存：\n%s" % exc, parent=window)
                return
            if not isinstance(payload, dict):
                messagebox.showerror(APP_TITLE, "JSON 最外层必须是对象 { }。", parent=window)
                return
            try:
                with open(path, "w", encoding="utf-8") as fh:
                    json.dump(payload, fh, ensure_ascii=False, indent=2)
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror(APP_TITLE, "保存失败：%s" % exc, parent=window)
                return
            self.set_status("已保存数据文件：%s" % path)
            window.destroy()
            self.read_fields()

        ttk.Button(bar, text="保存", command=save).pack(side=LEFT)
        ttk.Button(bar, text="格式化", command=format_json).pack(side=LEFT, padx=8)
        ttk.Button(bar, text="取消", command=window.destroy).pack(side=RIGHT)
        ttk.Label(bar, text="Ctrl+S 保存", style="Hint.TLabel").pack(side=RIGHT, padx=12)
        window.bind("<Control-s>", lambda _event: save())
        editor.focus_set()

    def generate(self) -> None:
        template = self.template_path.get().strip()
        if not template or not os.path.exists(template):
            messagebox.showwarning(APP_TITLE, "请先选择 Word 模板文件。")
            return
        data = self.load_data()
        if data is None:
            return
        output = self.output_path.get().strip()
        if not output:
            stem, ext = os.path.splitext(template)
            output = "%s_已填充%s" % (stem, ext or ".docx")
            self.output_path.set(output)
        self.set_status("正在生成文档…")
        try:
            result = fill_template(
                template,
                data,
                output,
                blank_missing=self.option_blank.get(),
                fill_controls=self.option_controls.get(),
            )
        except MissingFieldsError as exc:
            messagebox.showerror(APP_TITLE, "生成失败：%s" % exc)
            return
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_TITLE, "生成失败：\n%s" % exc)
            return

        self.set_status("已生成：%s" % os.path.basename(result.output_path))
        messagebox.showinfo(
            APP_TITLE, result.summary() + "\n\n输出文件：\n%s" % result.output_path
        )

        if self.option_pdf.get():
            try:
                self.set_status("正在导出 PDF…")
                pdf = export_pdf(result.output_path)
                self.set_status("已生成 PDF：%s" % os.path.basename(pdf))
            except Exception as exc:  # noqa: BLE001
                messagebox.showwarning(
                    APP_TITLE,
                    "文档已生成，但 PDF 导出没有成功：\n%s\n\n（PDF 导出需要本机安装 LibreOffice）" % exc,
                )
        if self.option_open.get():
            self.open_path(result.output_path)

    @staticmethod
    def open_path(path: str) -> None:
        try:
            if sys.platform == "win32":
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system('open "%s"' % path)
            else:
                os.system('xdg-open "%s"' % path)
        except Exception:  # noqa: BLE001
            pass


def main(initial_template: str | None = None) -> int:
    app = App(initial_template=initial_template)
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
