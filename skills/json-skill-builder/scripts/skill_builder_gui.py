# -*- coding: utf-8 -*-
"""JSON 字段说明填写器：可视化地逐条写"这个字段填什么"，然后生成 skill。

双击 skill 目录下的 一键启动.bat 即可打开；也可以直接运行本文件。
"""

from __future__ import annotations

import json
import os
import sys
from tkinter import (
    BOTH,
    END,
    LEFT,
    RIGHT,
    VERTICAL,
    X,
    Y,
    BooleanVar,
    StringVar,
    Text,
    Tk,
    filedialog,
    messagebox,
)
from tkinter import ttk

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(SKILL_DIR, "assets", "payload", "engine"))

import gen_skill  # noqa: E402

TITLE = "JSON 字段说明填写器"


def apply_dpi() -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:  # noqa: BLE001
        pass


def pick_font() -> str:
    try:
        from tkinter import font as tkfont

        families = set(tkfont.families())
    except Exception:  # noqa: BLE001
        return "TkDefaultFont"
    for name in ("Microsoft YaHei UI", "微软雅黑", "Microsoft YaHei", "SimHei"):
        if name in families:
            return name
    return "TkDefaultFont"


class SimpleEditor:
    """普通字段：一行标题 + 一个描述框。"""

    def __init__(self, parent, path: str, kind: str, value: str, on_change):
        self.path = path
        self.kind = kind
        self.frame = ttk.Frame(parent)
        self.frame.pack(fill=X, pady=(6, 2), padx=2)
        ttk.Label(self.frame, text=path, font=("", 10, "bold")).pack(anchor="w")
        ttk.Label(self.frame, text="类型：%s" % kind, foreground="#666666").pack(anchor="w")
        self.text = Text(self.frame, height=2, wrap="word", undo=True)
        self.text.pack(fill=X, pady=(2, 0))
        self.text.insert("1.0", value or "")
        self.text.bind("<FocusOut>", lambda _e: on_change())
        self.text.bind("<KeyRelease>", lambda _e: on_change())

    def value(self):
        return self.text.get("1.0", END).strip()


class TableEditor:
    """表格字段：整表说明 + 每列一行（列名可改、可加列）。"""

    def __init__(self, parent, path: str, description: dict, columns, on_change):
        self.path = path
        self.kind = "表格"
        self.on_change = on_change
        self.column_rows = []
        self.frame = ttk.LabelFrame(parent, text="表格：%s" % path, padding=8)
        self.frame.pack(fill=X, pady=(8, 4), padx=2)

        ttk.Label(self.frame, text="整表怎么填（_行说明）：").pack(anchor="w")
        self.note = Text(self.frame, height=2, wrap="word", undo=True)
        self.note.pack(fill=X, pady=(2, 6))
        self.note.insert("1.0", str((description or {}).get("_行说明", "")))
        self.note.bind("<KeyRelease>", lambda _e: self.on_change())
        self.note.bind("<FocusOut>", lambda _e: self.on_change())

        ttk.Label(self.frame, text="每一列填什么（列名可以改）：").pack(anchor="w")
        self.rows_box = ttk.Frame(self.frame)
        self.rows_box.pack(fill=X)
        names = [str(item) for item in (columns or [])]
        if not names:
            names = [key for key in (description or {}).keys() if not str(key).startswith("_")]
        for name in names:
            self.add_column(name, str((description or {}).get(name, "")))
        bar = ttk.Frame(self.frame)
        bar.pack(fill=X, pady=(6, 0))
        ttk.Button(bar, text="＋ 添加一列", command=lambda: self.add_column("", "")).pack(side=LEFT)
        ttk.Label(bar, text="  行数由数据决定，这里只定义列名和每列填什么", foreground="#666666").pack(
            side=LEFT
        )

    def add_column(self, name: str, note: str) -> None:
        row = ttk.Frame(self.rows_box)
        row.pack(fill=X, pady=2)
        name_var = StringVar(value=name)
        name_entry = ttk.Entry(row, textvariable=name_var, width=18)
        name_entry.pack(side=LEFT)
        note_text = Text(row, height=2, wrap="word", undo=True, width=70)
        note_text.pack(side=LEFT, fill=X, expand=True, padx=6)
        note_text.insert("1.0", note or "")
        note_text.bind("<KeyRelease>", lambda _e: self.on_change())
        note_text.bind("<FocusOut>", lambda _e: self.on_change())
        name_entry.bind("<KeyRelease>", lambda _e: self.on_change())

        holder = {"frame": row, "name": name_var, "note": note_text}
        self.column_rows.append(holder)

        def remove() -> None:
            row.destroy()
            self.column_rows.remove(holder)
            self.on_change()

        ttk.Button(row, text="✕", width=3, command=remove).pack(side=LEFT)

    def value(self):
        result = {"_类型": "表格", "_行说明": self.note.get("1.0", END).strip()}
        for holder in self.column_rows:
            name = holder["name"].get().strip()
            if not name:
                continue
            result[name] = holder["note"].get("1.0", END).strip()
        return result


class App(Tk):
    def __init__(self):
        super().__init__()
        apply_dpi()
        self.title(TITLE)
        self.geometry("1060x780")
        self.minsize(900, 620)

        font_name = pick_font()
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except Exception:  # noqa: BLE001
            pass
        for name in ("TLabel", "TButton", "TEntry", "TCheckbutton", "TLabelframe.Label", "TNotebook.Tab"):
            try:
                style.configure(name, font=(font_name, 10))
            except Exception:  # noqa: BLE001
                pass
        style.configure("Title.TLabel", font=(font_name, 15, "bold"))
        style.configure("Hint.TLabel", font=(font_name, 9), foreground="#555555")
        style.configure("Go.TButton", font=(font_name, 11, "bold"))
        self.option_add("*Font", (font_name, 10))

        self.skeleton_path = StringVar()
        self.skill_name = StringVar()
        self.display_name = StringVar()
        self.template_path = StringVar()
        self.status_text = StringVar(value="请先选择 JSON 骨架文件，或从 Word 模板生成一个")
        self.option_force = BooleanVar(value=True)
        self.option_zip = BooleanVar(value=True)
        self.zip_path = StringVar()

        self.skeleton = None
        self.editors = []

        self._build_header()
        self._build_source_row()
        self._build_fields_area()
        self._build_bottom()
        ttk.Label(self, textvariable=self.status_text, anchor="w", padding=(14, 6)).pack(
            fill=X, side="bottom"
        )

    # -- 界面 --------------------------------------------------------------

    def _build_header(self) -> None:
        header = ttk.Frame(self, padding=(16, 12, 16, 4))
        header.pack(fill=X)
        ttk.Label(header, text=TITLE, style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="① 选 JSON 骨架　② 一个字段一个框地写「这个字段填什么」　③ 点生成 skill —— 之后 AI 就按你的描述填这份 JSON",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(4, 0))

    def _build_source_row(self) -> None:
        row = ttk.Frame(self, padding=(16, 6, 16, 6))
        row.pack(fill=X)
        ttk.Label(row, text="JSON 骨架：").pack(side=LEFT)
        ttk.Entry(row, textvariable=self.skeleton_path).pack(side=LEFT, fill=X, expand=True)
        ttk.Button(row, text="浏览…", command=self.pick_skeleton).pack(side=LEFT, padx=(8, 0))
        ttk.Button(row, text="从 Word 模板生成骨架…", command=self.skeleton_from_word).pack(
            side=LEFT, padx=(8, 0)
        )
        ttk.Button(row, text="重新载入", command=self.load_skeleton).pack(side=LEFT, padx=(8, 0))

    def _build_fields_area(self) -> None:
        wrapper = ttk.Frame(self)
        wrapper.pack(fill=BOTH, expand=True, padx=16, pady=(0, 6))
        self.canvas = __import__("tkinter").Canvas(wrapper, highlightthickness=0)
        self.canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar = ttk.Scrollbar(wrapper, orient=VERTICAL, command=self.canvas.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.inner = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind(
            "<Configure>", lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.bind(
            "<Configure>",
            lambda event: self.canvas.itemconfigure(self.canvas.find_all()[0], width=event.width),
        )
        self.canvas.bind_all("<MouseWheel>", self._on_wheel)

    def _on_wheel(self, event) -> None:
        try:
            self.canvas.yview_scroll(int(-event.delta / 120), "units")
        except Exception:  # noqa: BLE001
            pass

    def _build_bottom(self) -> None:
        box = ttk.LabelFrame(self, text="生成 skill（第 3 步）", padding=10)
        box.pack(fill=X, padx=16, pady=(0, 6))
        line1 = ttk.Frame(box)
        line1.pack(fill=X)
        ttk.Label(line1, text="skill 名（英文小写+连字符）：").pack(side=LEFT)
        ttk.Entry(line1, textvariable=self.skill_name, width=26).pack(side=LEFT)
        ttk.Label(line1, text="   界面显示名：").pack(side=LEFT)
        ttk.Entry(line1, textvariable=self.display_name, width=26).pack(side=LEFT)
        ttk.Checkbutton(line1, text="同名时覆盖", variable=self.option_force).pack(side=LEFT, padx=10)

        line2 = ttk.Frame(box)
        line2.pack(fill=X, pady=(8, 0))
        ttk.Label(line2, text="配套 Word 模板（会一起打包进 skill）：").pack(side=LEFT)
        ttk.Entry(line2, textvariable=self.template_path).pack(side=LEFT, fill=X, expand=True)
        ttk.Button(line2, text="浏览…", command=self.pick_template).pack(side=LEFT, padx=(8, 0))

        line2b = ttk.Frame(box)
        line2b.pack(fill=X, pady=(6, 0))
        ttk.Checkbutton(line2b, text="同时打包成 zip（拷到别的电脑用）→", variable=self.option_zip).pack(
            side=LEFT
        )
        ttk.Entry(line2b, textvariable=self.zip_path).pack(side=LEFT, fill=X, expand=True, padx=6)
        ttk.Button(line2b, text="另存为…", command=self.pick_zip).pack(side=LEFT)

        line3 = ttk.Frame(box)
        line3.pack(fill=X, pady=(10, 0))
        ttk.Button(line3, text="保存描述", command=self.save_description).pack(side=LEFT)
        ttk.Button(line3, text="检查有没有漏填", command=self.check).pack(side=LEFT, padx=8)
        ttk.Button(line3, text="生成 skill", style="Go.TButton", command=self.build).pack(
            side=LEFT, padx=8
        )
        ttk.Button(line3, text="打开 skill 目录", command=self.open_skills_dir).pack(side=LEFT)

    # -- 数据 --------------------------------------------------------------

    def _on_change(self) -> None:
        self.update_status()

    def update_status(self) -> None:
        if not self.editors:
            return
        description = self.collect()
        problems, filled, _hints = gen_skill.check_descriptions(self.skeleton, description)
        total = len(self.editors)
        self.status_text.set(
            "共 %d 个字段，已写 %d 个，还差 %d 个没写%s"
            % (total, filled, max(total - filled, 0), "（描述保存在模型外的空值也算缺）" if problems else "，可以生成 skill 了")
        )

    def pick_skeleton(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 JSON 骨架",
            filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")],
            initialdir=os.path.dirname(self.skeleton_path.get()) or None,
        )
        if path:
            self.skeleton_path.set(path)
            self.load_skeleton()

    def skeleton_from_word(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Word 模板",
            filetypes=[("Word 文档", "*.docx *.dotx"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            from wordfill import engine
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(TITLE, "内置引擎加载失败：%s" % exc)
            return
        try:
            info = engine.scan_template(path)
            skeleton = engine.build_skeleton(info)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(TITLE, "扫描模板失败：%s" % exc)
            return
        if not skeleton:
            messagebox.showinfo(TITLE, "这个文档里没有找到 {{字段}} 标记，先在 Word 里写标记吧。")
            return
        target = os.path.splitext(path)[0] + "_数据骨架.json"
        with open(target, "w", encoding="utf-8") as fh:
            json.dump(skeleton, fh, ensure_ascii=False, indent=2)
        self.skeleton_path.set(target)
        self.template_path.set(path)
        if not self.display_name.get():
            self.display_name.set("%s 数据填写" % os.path.splitext(os.path.basename(path))[0])
        self.load_skeleton()
        messagebox.showinfo(TITLE, "已从模板生成骨架：\n%s\n\n共 %d 个字段，接下来逐个写描述。"
                            % (target, len(self.editors)))

    def load_skeleton(self) -> None:
        path = self.skeleton_path.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning(TITLE, "请先选择 JSON 骨架文件。")
            return
        try:
            with open(path, "r", encoding="utf-8-sig") as fh:
                self.skeleton = json.load(fh)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(TITLE, "JSON 读取失败：%s" % exc)
            return
        if not isinstance(self.skeleton, dict):
            messagebox.showerror(TITLE, "JSON 最外层必须是对象 { }。")
            return

        base = os.path.splitext(path)[0]
        description = {}
        for candidate in (base + "_字段说明.json", os.path.join(os.path.dirname(path), "字段说明.json")):
            if os.path.exists(candidate):
                try:
                    with open(candidate, "r", encoding="utf-8-sig") as fh:
                        description = json.load(fh)
                    self.status_text.set("已载入之前的描述：%s" % candidate)
                    break
                except Exception:  # noqa: BLE001
                    continue
        if not description:
            description = gen_skill.describe_skeleton(self.skeleton)

        for editor in self.editors:
            editor.frame.destroy()
        self.editors = []
        for path_name, node, kind in gen_skill.iter_fields(self.skeleton):
            desc = gen_skill.describe_for(path_name, description)
            if kind == "表格":
                editor = TableEditor(
                    self.inner,
                    path_name,
                    desc if isinstance(desc, dict) else {},
                    gen_skill.table_columns(node),
                    self._on_change,
                )
            else:
                editor = SimpleEditor(
                    self.inner, path_name, kind, desc if isinstance(desc, str) else "", self._on_change
                )
            self.editors.append(editor)
        if not self.editors:
            self.status_text.set("这份 JSON 里没有可填的字段。")
        else:
            self.update_status()

    def collect(self) -> dict:
        result = {}
        for editor in self.editors:
            value = editor.value()
            if isinstance(value, dict) and not any(
                str(key).strip() for key in value.keys() if not str(key).startswith("_")
            ) and not value.get("_行说明"):
                value = {}
            gen_skill.set_path(result, editor.path, value)
        return result

    def description_target(self) -> str:
        path = self.skeleton_path.get().strip()
        base = os.path.splitext(path)[0]
        return base + "_字段说明.json"

    def save_description(self) -> bool:
        if not self.editors:
            return False
        target = self.description_target()
        try:
            with open(target, "w", encoding="utf-8") as fh:
                json.dump(self.collect(), fh, ensure_ascii=False, indent=2)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(TITLE, "保存失败：%s" % exc)
            return False
        self.status_text.set("描述已保存：%s" % target)
        return True

    def check(self) -> bool:
        if not self.editors:
            return False
        self.save_description()
        problems, filled, hints = gen_skill.check_descriptions(self.skeleton, self.collect())
        message = []
        if hints:
            message.append("提示：")
            message += ["  · %s" % item for item in hints]
        if problems:
            message.append("还差这些没写：")
            message += ["  · %s" % item for item in problems]
            messagebox.showwarning(TITLE, "\n".join(message))
            self.status_text.set("还差 %d 个字段没写" % len(problems))
            return False
        messagebox.showinfo(TITLE, "字段说明都写好了，可以生成 skill。")
        self.status_text.set("字段说明检查通过（%d 个字段）" % filled)
        return True

    def pick_template(self) -> None:
        path = filedialog.askopenfilename(
            title="选择配套的 Word 模板（可选）",
            filetypes=[("Word 文档", "*.docx *.dotx"), ("所有文件", "*.*")],
        )
        if path:
            self.template_path.set(path)

    def pick_zip(self) -> None:
        path = filedialog.asksaveasfilename(
            title="skill 打包保存为",
            defaultextension=".zip",
            filetypes=[("ZIP 压缩包", "*.zip")],
            initialfile=(gen_skill.slugify(self.skill_name.get()) or "my-skill") + ".zip",
        )
        if path:
            self.zip_path.set(path)
            self.option_zip.set(True)

    def build(self) -> None:
        if not self.editors:
            messagebox.showwarning(TITLE, "请先载入 JSON 骨架。")
            return
        name = self.skill_name.get().strip()
        if not name:
            messagebox.showwarning(TITLE, "请填 skill 名，例如 tech-solution-json（英文小写+连字符）。")
            return
        if not self.check():
            return
        if not self.save_description():
            return
        zip_path = None
        if self.option_zip.get():
            zip_path = self.zip_path.get().strip()
            if not zip_path:
                zip_path = os.path.join(
                    os.path.expanduser("~"),
                    ".codex",
                    "skills",
                    (gen_skill.slugify(name) or name) + ".zip",
                )
                self.zip_path.set(zip_path)
        try:
            result = gen_skill.build_skill(
                self.skeleton_path.get().strip(),
                self.description_target(),
                name,
                display_name=self.display_name.get().strip() or None,
                template=self.template_path.get().strip() or None,
                force=bool(self.option_force.get()),
                zip_path=zip_path,
            )
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(TITLE, "生成失败：\n%s" % exc)
            return
        self.status_text.set("已生成 skill：%s" % result["skill_dir"])
        lines = ["生成成功：", result["skill_dir"], ""]
        if self.template_path.get().strip():
            lines.append("Word 模板已一起打包进 assets/，换台电脑也能用。")
        if result.get("zip"):
            lines.append("便携压缩包：%s" % result["zip"])
        lines += [
            "",
            "新开一个对话就能用它：",
            "$%s 按我的描述和刚才讨论的内容填这份 JSON" % result["name"],
        ]
        messagebox.showinfo(TITLE, "\n".join(lines))

    def open_skills_dir(self) -> None:
        target = os.path.join(os.path.expanduser("~"), ".codex", "skills")
        try:
            if sys.platform == "win32":
                os.startfile(target)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system('open "%s"' % target)
            else:
                os.system('xdg-open "%s"' % target)
        except Exception:  # noqa: BLE001
            messagebox.showinfo(TITLE, target)


def main() -> int:
    app = App()
    app.mainloop()
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:  # noqa: BLE001
        import traceback

        message = "界面启动失败：%s\n\n%s" % (exc, traceback.format_exc(limit=3))
        print(message)
        try:
            import ctypes

            ctypes.windll.user32.MessageBoxW(0, message, TITLE, 0x10)
        except Exception:  # noqa: BLE001
            pass
        sys.exit(1)
