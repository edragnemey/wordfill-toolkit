# -*- coding: utf-8 -*-
"""Word 模板编译器（核心引擎）。

在 Word 文档里写好"特殊字段"标记，再用一份 JSON 数据把它们替换掉。

字段标记语法
------------
    {{姓名}}                     普通字段
    {{客户.名称}}                 多级字段（对应嵌套 JSON）
    {{金额|千分位}}               加千分位，如 1,234,567.89
    {{金额|0.00}}                数字格式
    {{金额|人民币大写}}            壹佰贰拾叁万肆仟伍佰陆拾柒元捌角玖分
    {{日期|yyyy年M月d日}}         日期格式
    {{姓名|默认:张三}}             数据里没有时使用的默认值
    {{@公章}}                    图片字段（JSON 里填图片路径）
    {{@公章|宽=4cm}}             指定图片宽度，高度按比例自动
    {{$today}} / {{$now}}        内置：当前日期 / 当前日期时间
    {{#if 是否加急}} ... {{/if}}   条件段落（也支持 {{else}}）
    {{#unless 是否加急}} ... {{/unless}} 反向条件
    {{#each 员工}} ... {{/each}}   重复内容（整段重复，或表格行重复）

未提供数据时，字段标记会原样保留，方便一眼看出缺什么；加 --blank 则替换为空。
"""

from __future__ import annotations

import copy
import datetime as _dt
import io
import json
import os
import re
import zipfile
from dataclasses import dataclass, field as dc_field
from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from docx import Document
from docx.image.image import Image as DocxImage
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Inches, Length, Pt
from docx.text.paragraph import Paragraph
from docx.text.run import Run

__all__ = [
    "Context",
    "FieldInfo",
    "LoopInfo",
    "TemplateInfo",
    "FillResult",
    "scan_template",
    "build_skeleton",
    "fill_template",
    "export_pdf",
    "load_json",
    "value_to_text",
    "rmb_upper",
]

# --------------------------------------------------------------------------
# 基础工具
# --------------------------------------------------------------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _q(tag: str) -> str:
    """返回 w: 命名空间下的完整标签名。"""
    return qn("w:" + tag)


def load_json(path: str) -> Any:
    """读取 JSON 数据文件（兼容带 BOM 的 UTF-8 与 GBK）。"""
    with open(path, "rb") as fh:
        raw = fh.read()
    for enc in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return json.loads(raw.decode(enc))
        except (UnicodeDecodeError, json.JSONDecodeError):
            continue
    raise ValueError("无法解析 JSON 文件：%s" % path)


# --------------------------------------------------------------------------
# 字段标记解析
# --------------------------------------------------------------------------

TAG_RE = re.compile(r"\{\{(.*?)\}\}", re.S)
IMG_MARK = "\ue000"
IMG_MARK_END = "\ue001"
IMG_MARK_RE = re.compile("\ue000(\\d+)\ue001")
TABLE_MARK = "\ue002"
TABLE_MARK_END = "\ue003"
TABLE_MARK_RE = re.compile("\ue002(\\d+)\ue003")


@dataclass
class Tag:
    """文档里出现的一个 {{...}} 标记。"""

    start: int
    end: int
    raw: str
    body: str
    kind: str          # field / image / builtin / open / close / else
    name: str          # 控件名(if/each/else) 或字段名
    expr: str = ""     # 控件的判断表达式 / 循环数据名
    modifiers: Tuple[str, ...] = ()
    default: Optional[str] = None
    para: Any = None   # 该标记所在的段落元素

    @property
    def is_control(self) -> bool:
        return self.kind in ("open", "close", "else")

    @property
    def label(self) -> str:
        if self.kind in ("field", "image", "builtin"):
            return self.name
        return self.expr or self.name


def _split_body(body: str) -> List[str]:
    return [part.strip() for part in body.split("|")]


def parse_tag(start: int, end: int, raw: str, body: str) -> Tag:
    """把一个标记文本解析成 Tag 对象。"""
    text = body.strip()
    low = text.lower()

    if text.startswith("/"):
        word = text[1:].strip().lower()
        return Tag(start, end, raw, body, "close", word)

    if text.startswith("^"):
        parts = text[1:].strip().split(None, 1)
        word = parts[0].lower() if parts else ""
        expr = parts[1].strip() if len(parts) > 1 else ""
        if word == "if":
            return Tag(start, end, raw, body, "open", "unless", expr)

    if text.startswith("#"):
        parts = text[1:].strip().split(None, 1)
        word = parts[0].lower() if parts else ""
        expr = parts[1].strip() if len(parts) > 1 else ""
        return Tag(start, end, raw, body, "open", word, expr)

    if low in ("else", "否则", "other"):
        return Tag(start, end, raw, body, "else", "else")

    if text.startswith("@"):
        segments = _split_body(text[1:].strip())
        return Tag(start, end, raw, body, "image", segments[0], "", tuple(segments[1:]))

    if text.startswith("$"):
        segments = _split_body(text[1:].strip())
        return Tag(start, end, raw, body, "builtin", segments[0], "", tuple(segments[1:]))

    for prefix in ("表格:", "表格：", "table:", "table：", "表:"):
        if text.startswith(prefix):
            segments = _split_body(text[len(prefix):].strip())
            name = segments[0] if segments else ""
            return Tag(start, end, raw, body, "table", name, "", tuple(segments[1:]))

    segments = _split_body(text)
    name = segments[0]
    modifiers: List[str] = []
    default: Optional[str] = None
    for seg in segments[1:]:
        lower = seg.lower()
        if lower.startswith("默认:") or lower.startswith("默认："):
            default = seg[3:]
        elif lower.startswith("default:"):
            default = seg[len("default:"):]
        elif seg:
            modifiers.append(seg)
    return Tag(start, end, raw, body, "field", name, "", tuple(modifiers), default)


def find_tags(text: str) -> List[Tag]:
    """找出文本里的全部字段标记。"""
    tags: List[Tag] = []
    for match in TAG_RE.finditer(text):
        body = match.group(1)
        if "{{" in body or "}}" in body:
            continue
        tags.append(parse_tag(match.start(), match.end(), match.group(0), body))
    return tags


# --------------------------------------------------------------------------
# 段落文字定位（支持把标记写在多个 run 里）
# --------------------------------------------------------------------------

_DESCEND_TAGS = {
    "hyperlink",
    "ins",
    "del",
    "moveFrom",
    "moveTo",
    "smartTag",
    "sdt",
    "fldSimple",
    "customXml",
    "dir",
    "bdo",
}


def _iter_run_elements(root) -> Iterable[Any]:
    """按文档顺序取出 root 下的 w:r（不进入图形内部）。"""
    for child in root:
        tag = child.tag
        if not isinstance(tag, str):
            continue
        name = tag.rsplit("}", 1)[-1]
        if name == "r":
            yield child
        elif name in _DESCEND_TAGS:
            for item in _iter_run_elements(child):
                yield item


def _iter_text_elements(root) -> Iterable[Any]:
    """按文档顺序取出 root 下的 w:t。"""
    for run in _iter_run_elements(root):
        for child in run:
            if child.tag == _q("t"):
                yield child


def write_text(t_el, text: str) -> None:
    """把文本写进 w:t；含换行时自动拆成 w:t + w:br，Word 里就是真的换行。"""
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    if "\n" not in text:
        t_el.text = text
        if text and text != text.strip():
            t_el.set(qn("xml:space"), "preserve")
        return
    run = t_el.getparent()
    if run is None:
        t_el.text = text.replace("\n", " ")
        return
    position = list(run).index(t_el)
    run.remove(t_el)
    for offset, piece in enumerate(text.split("\n")):
        if offset:
            run.insert(position, OxmlElement("w:br"))
            position += 1
        element = OxmlElement("w:t")
        element.set(qn("xml:space"), "preserve")
        element.text = piece
        run.insert(position, element)
        position += 1


class ParaText:
    """段落文本 + 每个字符落在哪个 w:t 上，用于跨 run 精确替换。"""

    def __init__(self, p_el):
        self.p_el = p_el
        self.atoms: List[Tuple[Any, int]] = []
        chars: List[str] = []
        for run in _iter_run_elements(p_el):
            for t in run:
                if t.tag != _q("t"):
                    continue
                value = t.text or ""
                for idx in range(len(value)):
                    self.atoms.append((t, idx))
                    chars.append(value[idx])
        self.text = "".join(chars)

    def tags(self) -> List[Tag]:
        tags = find_tags(self.text)
        for tag in tags:
            tag.para = self.p_el
        return tags

    def apply(self, ops: Sequence[Tuple[int, int, str]]) -> None:
        """执行替换。ops = [(起点, 终点, 替换文本)]，区间互不重叠。"""
        if not ops:
            return
        deleted: Dict[Any, set] = {}
        inserted: Dict[Tuple[Any, int], str] = {}
        for start, end, payload in ops:
            if end > start:
                for t, offset in self.atoms[start:end]:
                    deleted.setdefault(t, set()).add(offset)
            if payload:
                t, offset = self.atoms[start]
                inserted[(t, offset)] = payload + inserted.get((t, offset), "")

        seen: List[Any] = []
        for t, _ in self.atoms:
            if all(t is not other for other in seen):
                seen.append(t)

        for t in seen:
            original = t.text or ""
            drop = deleted.get(t, set())
            has_insert = any(key[0] is t for key in inserted)
            if not drop and not has_insert:
                continue
            out: List[str] = []
            for idx, ch in enumerate(original):
                payload = inserted.get((t, idx))
                if payload:
                    out.append(payload)
                if idx in drop:
                    continue
                out.append(ch)
            new_text = "".join(out)
            write_text(t, new_text)


# --------------------------------------------------------------------------
# 取值与格式化
# --------------------------------------------------------------------------

_DIGITS = "零壹贰叁肆伍陆柒捌玖"
_UNITS4 = ("仟", "佰", "拾", "")
_BIG_UNITS = ("", "万", "亿", "万亿")


def _section_to_chinese(num: int) -> str:
    out = ""
    started = False
    zero_pending = False
    for pos, unit in enumerate(_UNITS4):
        digit = (num // (10 ** (3 - pos))) % 10
        if digit == 0:
            if started:
                zero_pending = True
        else:
            if zero_pending:
                out += "零"
                zero_pending = False
            out += _DIGITS[digit] + unit
            started = True
    return out


def _int_to_chinese(num: int) -> str:
    if num == 0:
        return "零"
    sections: List[int] = []
    while num > 0:
        sections.append(num % 10000)
        num //= 10000
    out = ""
    for idx in range(len(sections) - 1, -1, -1):
        sec = sections[idx]
        if sec == 0:
            if out and not out.endswith("零"):
                out += "零"
            continue
        if out and sec < 1000 and not out.endswith("零"):
            out += "零"
        piece = _section_to_chinese(sec)
        unit = _BIG_UNITS[idx] if idx < len(_BIG_UNITS) else ""
        out += piece + unit
    return out


def rmb_upper(value: Any) -> str:
    """把数字转成人民币大写，例如 1234.56 -> 壹仟贰佰叁拾肆元伍角陆分。"""
    try:
        amount = Decimal(str(value).replace(",", "").replace("￥", "").replace("¥", "").strip())
    except Exception:
        return str(value)
    sign = "负" if amount < 0 else ""
    amount = abs(amount)
    cents = int((amount * 100).to_integral_value(rounding=ROUND_HALF_UP))
    if cents == 0:
        return "零元整"
    yuan, rest = divmod(cents, 100)
    jiao, fen = divmod(rest, 10)
    out = ""
    if yuan:
        out += _int_to_chinese(yuan) + "元"
    if jiao == 0 and fen == 0:
        out += "整"
    else:
        if jiao:
            out += _DIGITS[jiao] + "角"
        elif yuan:
            out += "零"
        if fen:
            out += _DIGITS[fen] + "分"
    return sign + out


def coerce_datetime(value: Any) -> Optional[_dt.datetime]:
    """尽量把各种写法的时间值转成 datetime。"""
    if isinstance(value, _dt.datetime):
        return value
    if isinstance(value, _dt.date):
        return _dt.datetime(value.year, value.month, value.day)
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        # Excel 日期序列号
        if 20000 < float(value) < 60000:
            return _dt.datetime(1899, 12, 30) + _dt.timedelta(days=float(value))
        return None
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        for fmt in (
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d %H:%M",
            "%Y-%m-%d",
            "%Y/%m/%d %H:%M:%S",
            "%Y/%m/%d",
            "%Y.%m.%d",
            "%Y年%m月%d日",
            "%Y%m%d",
        ):
            try:
                return _dt.datetime.strptime(text, fmt)
            except ValueError:
                continue
        try:
            return _dt.datetime.fromisoformat(text)
        except ValueError:
            return None
    return None


def format_datetime(value: Any, pattern: str) -> Optional[str]:
    """按 Word 风格的模式格式化日期，如 yyyy年M月d日、yyyy-MM-dd HH:mm:ss。"""
    moment = coerce_datetime(value)
    if moment is None:
        return None
    out: List[str] = []
    idx = 0
    length = len(pattern)
    while idx < length:
        chunk = pattern[idx:]
        if chunk.startswith("yyyy"):
            out.append("%04d" % moment.year)
            idx += 4
        elif chunk.startswith("yy"):
            out.append("%02d" % (moment.year % 100))
            idx += 2
        elif chunk.startswith("MM"):
            out.append("%02d" % moment.month)
            idx += 2
        elif pattern[idx] == "M":
            out.append(str(moment.month))
            idx += 1
        elif chunk.startswith("dd"):
            out.append("%02d" % moment.day)
            idx += 2
        elif pattern[idx] == "d":
            out.append(str(moment.day))
            idx += 1
        elif chunk.startswith("HH"):
            out.append("%02d" % moment.hour)
            idx += 2
        elif pattern[idx] == "H":
            out.append(str(moment.hour))
            idx += 1
        elif chunk.startswith("mm"):
            out.append("%02d" % moment.minute)
            idx += 2
        elif pattern[idx] == "m":
            out.append(str(moment.minute))
            idx += 1
        elif chunk.startswith("ss"):
            out.append("%02d" % moment.second)
            idx += 2
        elif pattern[idx] == "s":
            out.append(str(moment.second))
            idx += 1
        else:
            out.append(pattern[idx])
            idx += 1
    return "".join(out)


def format_number(value: Any, pattern: str) -> Optional[str]:
    """按 0.00 / #,##0.00 之类的模式格式化数字。"""
    try:
        number = float(str(value).replace(",", "").strip())
    except Exception:
        return None
    decimals = 0
    if "." in pattern:
        decimals = len(re.findall(r"[0#]", pattern.split(".", 1)[1]))
    if "," in pattern:
        return format(number, ",.%df" % decimals)
    return format(number, ".%df" % decimals)


_NUMBER_PATTERN_RE = re.compile(r"^[#,0.\s]+$")
_DATE_PATTERN_RE = re.compile(r"^[\s年月日yMdHhms:./\-]+$")
_TRUTHY_TEXT = {"1", "true", "yes", "y", "是", "对", "有", "开启"}
_FALSY_TEXT = {"", "0", "false", "no", "n", "否", "无", "不", "关闭", "null", "none"}


def _plain_number_text(value: Any) -> str:
    if isinstance(value, bool):
        return "是" if value else "否"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value == int(value) and abs(value) < 1e16:
            return str(int(value))
        return ("%f" % value).rstrip("0").rstrip(".")
    return str(value)


def to_bool(value: Any) -> bool:
    """把各种写法转成真/假。"""
    if value is None:
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    if isinstance(value, (list, tuple, dict, set)):
        return len(value) > 0
    text = str(value).strip()
    low = text.lower()
    if low in _FALSY_TEXT:
        return False
    if low in _TRUTHY_TEXT:
        return True
    return bool(text)


def value_to_text(value: Any, modifiers: Sequence[str] = ()) -> str:
    """把数据值转成写入文档的文本，并按修饰符格式化。"""
    if value is None:
        text = ""
    elif isinstance(value, (_dt.datetime, _dt.date)):
        text = value.strftime("%Y-%m-%d")
    elif isinstance(value, (dict, list)):
        text = json.dumps(value, ensure_ascii=False)
    else:
        text = _plain_number_text(value)

    for modifier in modifiers or ():
        mod = (modifier or "").strip()
        if not mod:
            continue
        low = mod.lower()
        if low in ("大写", "upper"):
            text = text.upper()
        elif low in ("小写", "lower"):
            text = text.lower()
        elif low in ("去空格", "trim"):
            text = text.strip()
        elif mod in ("人民币大写", "大写金额", "金额大写") or low in ("rmb", "rmb_upper"):
            text = rmb_upper(value)
        elif mod in ("千分位", "货币"):
            text = format_number(value, ",0.00" if mod == "货币" else ",0") or text
        elif mod == "中文日期":
            text = format_datetime(value, "yyyy年M月d日") or text
        elif mod == "日期":
            text = format_datetime(value, "yyyy-MM-dd") or text
        elif _NUMBER_PATTERN_RE.match(mod) and any(c in mod for c in "0#"):
            text = format_number(value, mod) or text
        elif _DATE_PATTERN_RE.match(mod) and re.search(r"[yMdHhms]", mod):
            text = format_datetime(value, mod) or text
        elif low in ("原文", "raw"):
            text = str(value)
    return text


def _split_top(expr: str, sep: str) -> List[str]:
    """按顶层分隔符切分表达式（忽略引号与括号内部）。"""
    parts: List[str] = []
    depth = 0
    quote = ""
    current: List[str] = []
    idx = 0
    while idx < len(expr):
        ch = expr[idx]
        if quote:
            current.append(ch)
            if ch == quote:
                quote = ""
            idx += 1
            continue
        if ch in ('"', "'"):
            quote = ch
            current.append(ch)
            idx += 1
            continue
        if ch in "([{":
            depth += 1
        elif ch in ")]}":
            depth -= 1
        if depth == 0 and expr.startswith(sep, idx):
            parts.append("".join(current))
            current = []
            idx += len(sep)
            continue
        current.append(ch)
        idx += 1
    parts.append("".join(current))
    return parts


# --------------------------------------------------------------------------
# 数据上下文
# --------------------------------------------------------------------------


class Context:
    """带作用域的 JSON 数据，支持 {{a.b}}、循环内的上级字段查找。"""

    def __init__(self, data: Any = None, parent: Optional["Context"] = None, source_name: str = ""):
        self.data = data if data is not None else {}
        self.parent = parent
        self.source_name = source_name or (parent.source_name if parent else "")
        self.used: set = set()
        self.missing: List[str] = []
        self.warnings: List[str] = []

    # -- 作用域 ------------------------------------------------------------

    def root(self) -> "Context":
        node = self
        while node.parent is not None:
            node = node.parent
        return node

    def push(self, item: Any, index: Optional[int] = None) -> "Context":
        frame = item if isinstance(item, dict) else {"_值": item}
        payload = dict(frame)
        if index is not None:
            payload.setdefault("_index", index)
            payload.setdefault("_序号", index)
            payload.setdefault("序号", index)
        return Context(payload, parent=self)

    # -- 取值 --------------------------------------------------------------

    def _lookup_here(self, path: str) -> Tuple[bool, Any]:
        data = self.data
        if isinstance(data, dict) and path in data:
            return True, data[path]
        parts = [p for p in path.split(".") if p != ""]
        if not parts:
            return False, None
        current = data
        for part in parts:
            if isinstance(current, dict):
                if part in current:
                    current = current[part]
                else:
                    return False, None
            elif isinstance(current, (list, tuple)):
                if part.lstrip("-").isdigit() and -len(current) <= int(part) < len(current):
                    current = current[int(part)]
                else:
                    return False, None
            else:
                return False, None
        return True, current

    def _builtin(self, name: str) -> Tuple[bool, Any]:
        key = name.strip().lower()
        now = _dt.datetime.now()
        if key in ("today", "date", "今天", "日期"):
            return True, now.date()
        if key in ("now", "datetime", "现在", "当前时间"):
            return True, now
        if key in ("time", "时间"):
            return True, now.strftime("%H:%M:%S")
        if key in ("year", "年"):
            return True, now.year
        if key in ("month", "月"):
            return True, now.month
        if key in ("day", "日"):
            return True, now.day
        if key in ("filename", "文件名"):
            return True, os.path.basename(self.root().source_name)
        return False, None

    def get(self, path: str) -> Tuple[bool, Any]:
        path = (path or "").strip()
        if not path:
            return False, None
        if path.startswith("$"):
            return self._builtin(path[1:])
        node: Optional[Context] = self
        while node is not None:
            found, value = node._lookup_here(path)
            if found:
                self.root().used.add(path.split(".")[0])
                return True, value
            node = node.parent
        return False, None

    def get_list(self, expr: str) -> List[Any]:
        found, value = self.get(expr)
        if not found or value is None:
            self.root().warnings.append("循环数据 %s 没有提供，该部分内容已跳过" % expr)
            return []
        if isinstance(value, (list, tuple)):
            return list(value)
        if isinstance(value, dict):
            return [value]
        return [value]

    def missing_key(self, name: str) -> None:
        root = self.root()
        if name not in root.missing:
            root.missing.append(name)


def _literal(text: str) -> Any:
    value = text.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in ('"', "'"):
        return value[1:-1]
    low = value.lower()
    if low in ("true", "yes", "是", "对"):
        return True
    if low in ("false", "no", "否", "无"):
        return False
    if low in ("null", "none", ""):
        return None
    try:
        number = float(value)
        return int(number) if number == int(number) else number
    except ValueError:
        return value


def _loose_equal(left: Any, right: Any) -> bool:
    if isinstance(left, bool) or isinstance(right, bool):
        return to_bool(left) == to_bool(right)
    try:
        return float(str(left).replace(",", "").strip()) == float(str(right).replace(",", "").strip())
    except (ValueError, TypeError):
        pass
    return str(left).strip() == str(right).strip()


def _compare(ctx: Context, left_text: str, op: str, right_text: str) -> bool:
    left_found, left_value = ctx.get(left_text)
    if not left_found:
        left_value = _literal(left_text)
    right_found, right_value = ctx.get(right_text)
    if not right_found:
        right_value = _literal(right_text)

    if op == "==":
        return _loose_equal(left_value, right_value)
    if op == "!=":
        return not _loose_equal(left_value, right_value)

    left_number: Optional[float]
    right_number: Optional[float]
    try:
        left_number = float(str(left_value).replace(",", "").strip())
    except (ValueError, TypeError):
        left_number = None
    try:
        right_number = float(str(right_value).replace(",", "").strip())
    except (ValueError, TypeError):
        right_number = None

    if left_number is not None and right_number is not None:
        pairs: Tuple[Any, Any] = (left_number, right_number)
    else:
        pairs = (str(left_value if left_value is not None else ""), str(right_value if right_value is not None else ""))
    first, second = pairs
    if op == ">":
        return first > second
    if op == "<":
        return first < second
    if op == ">=":
        return first >= second
    if op == "<=":
        return first <= second
    return False


def eval_condition(expr: str, ctx: Context) -> bool:
    """求值 {{#if ...}} 里的判断表达式。"""
    expr = (expr or "").strip()
    if not expr:
        return False
    parts = _split_top(expr, "||")
    if len(parts) > 1:
        return any(eval_condition(part, ctx) for part in parts)
    parts = _split_top(expr, "&&")
    if len(parts) > 1:
        return all(eval_condition(part, ctx) for part in parts)
    if expr.startswith("!"):
        return not eval_condition(expr[1:], ctx)
    if expr.lower().startswith("not "):
        return not eval_condition(expr[4:], ctx)
    if expr.startswith("(") and expr.endswith(")"):
        return eval_condition(expr[1:-1], ctx)
    for op in ("==", "!=", ">=", "<=", ">", "<"):
        parts = _split_top(expr, op)
        if len(parts) > 1:
            return _compare(ctx, parts[0].strip(), op, op.join(parts[1:]).strip())
    found, value = ctx.get(expr)
    return to_bool(value) if found else False


# --------------------------------------------------------------------------
# 图片与文档工具
# --------------------------------------------------------------------------


class _PartShim:
    """给 Run 提供一个能取到 part 的父对象，便于插入图片。"""

    def __init__(self, part):
        self.part = part


def _tag_name(element) -> str:
    tag = getattr(element, "tag", None)
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


_LENGTH_RE = re.compile(r"^\s*([\d.]+)\s*(cm|厘米|mm|毫米|in|inch|英寸|pt|磅|px|像素)?\s*$", re.I)


def parse_length(text: Any) -> Optional[Length]:
    """把 '4cm' / '3.5厘米' / 4 转成 Word 长度。"""
    if text is None or text == "":
        return None
    if isinstance(text, Length):
        return text
    if isinstance(text, (int, float)):
        return Cm(float(text))
    match = _LENGTH_RE.match(str(text))
    if not match:
        return None
    number = float(match.group(1))
    unit = (match.group(2) or "cm").lower()
    if unit in ("cm", "厘米"):
        return Cm(number)
    if unit in ("mm", "毫米"):
        return Emu(int(Cm(number).emu / 10))
    if unit in ("in", "inch", "英寸"):
        return Inches(number)
    if unit in ("pt", "磅"):
        return Pt(number)
    if unit in ("px", "像素"):
        return Emu(int(number * 9525))
    return Cm(number)


def _open_docx(path: str):
    """打开 Word 文档，顺带兼容 .dotx/.dotm 模板。"""
    ext = os.path.splitext(path)[1].lower()
    if not os.path.exists(path):
        raise FileNotFoundError("找不到文件：%s" % path)
    if ext == ".doc":
        raise ValueError("不支持 .doc 旧格式，请先在 Word 里另存为 .docx 再使用")
    if not zipfile.is_zipfile(path):
        raise ValueError("这不是有效的 Word 文档（.docx/.dotx）：%s" % path)
    try:
        return Document(path)
    except Exception:
        pass
    with zipfile.ZipFile(path) as source:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
            for item in source.infolist():
                payload = source.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    text = payload.decode("utf-8", "ignore")
                    text = text.replace(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                    ).replace(
                        "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                    )
                    payload = text.encode("utf-8")
                target.writestr(item, payload)
        buffer.seek(0)
    try:
        return Document(buffer)
    except Exception as exc:  # noqa: BLE001
        raise ValueError("无法打开这个 Word 文档，请确认文件没有损坏：%s" % exc)


# --------------------------------------------------------------------------
# 渲染器
# --------------------------------------------------------------------------


class _Renderer:
    def __init__(self, ctx: Context, blank_missing: bool = False, base_dir: str = "", doc=None):
        self.ctx = ctx
        self.blank_missing = blank_missing
        self.base_dir = base_dir
        self.doc = doc
        self.image_specs: List[Dict[str, Any]] = []
        self.table_specs: List[Dict[str, Any]] = []
        self.image_count = 0
        self.table_count = 0
        self.filled = 0
        self._seen_boxes: set = set()

    # -- 通用 --------------------------------------------------------------

    @staticmethod
    def _block_children(el) -> List[Any]:
        return [child for child in el if _tag_name(child) in ("p", "tbl", "sdt")]

    def _marker_tags(self, p_el) -> Optional[List[Tag]]:
        pt = ParaText(p_el)
        tags = pt.tags()
        if not tags:
            return None
        if not all(tag.is_control for tag in tags):
            return None
        if TAG_RE.sub("", pt.text).strip():
            return None
        return tags

    def _row_control_tags(self, row) -> List[Tag]:
        """表格行里的控制标记。

        {{#each}} / {{/each}} 允许和普通内容写在同一个单元格里（表格行循环最常用的写法）；
        {{#if}} 只有在"整段只有标记"时才按整行条件处理，避免和单元格内的行内条件冲突。
        """
        out: List[Tag] = []
        for cell in row.findall(_q("tc")):
            for child in cell:
                if child.tag != _q("p"):
                    continue
                pt = ParaText(child)
                if "{{" not in pt.text:
                    continue
                tags = pt.tags()
                # 同一段落里成对的标记（例如 {{#if}}…{{/if}}、{{#each}}…{{/each}}）
                # 属于单元格内联写法，不算表格行标记。
                is_marker = not TAG_RE.sub("", pt.text).strip()
                paired: set = set()
                stack: List[int] = []
                for pos, tag in enumerate(tags):
                    if tag.kind == "open":
                        stack.append(pos)
                    elif tag.kind == "close" and stack:
                        paired.add(stack.pop())
                        paired.add(pos)
                for pos, tag in enumerate(tags):
                    if pos in paired:
                        continue
                    if tag.kind == "close":
                        out.append(tag)
                    elif tag.kind == "open" and (tag.name == "each" or is_marker):
                        out.append(tag)
                    elif tag.kind == "else" and is_marker:
                        out.append(tag)
        return out

    @staticmethod
    def _kinds_match(open_name: str, close_name: str) -> bool:
        if open_name == close_name:
            return True
        return {open_name, close_name} <= {"if", "unless"}

    def _eval_control(self, open_tag: Tag, ctx: Context) -> bool:
        value = eval_condition(open_tag.expr, ctx)
        return value if open_tag.name == "if" else (not value)

    # -- 段落清理 ----------------------------------------------------------

    def _cleanup_empty_paragraph(self, p_el) -> None:
        if ParaText(p_el).text.strip():
            return
        for name in ("drawing", "pict", "object", "fldChar", "sym", "br"):
            if p_el.findall(".//" + _q(name)):
                return
        parent = p_el.getparent()
        if parent is None:
            return
        if _tag_name(parent) == "tc":
            paragraphs = [c for c in parent if c.tag == _q("p")]
            if len(paragraphs) <= 1:
                return
        parent.remove(p_el)

    def _strip_paragraph_tags(self, p_el, tags: Sequence[Tag]) -> None:
        if p_el is None or not tags:
            return
        pt = ParaText(p_el)
        ops = sorted(((tag.start, tag.end, "") for tag in tags), key=lambda item: -item[0])
        pt.apply(ops)
        self._cleanup_empty_paragraph(p_el)

    def _remove_tag_text(self, tag: Tag) -> None:
        self._strip_paragraph_tags(tag.para, [tag])

    def _strip_row_tags(self, row, tags: Sequence[Tag]) -> None:
        grouped: Dict[Any, List[Tag]] = {}
        for tag in tags:
            if tag.para is None:
                continue
            grouped.setdefault(tag.para, []).append(tag)
        for para, group in grouped.items():
            self._strip_paragraph_tags(para, group)

    def _strip_row_tag(self, row, tag: Tag) -> None:
        candidates = self._row_control_tags(row)
        match = None
        for item in candidates:
            if item.kind == tag.kind and item.expr == tag.expr and item.name == tag.name:
                match = item
                break
        if match is None and candidates:
            match = candidates[0]
        if match is not None:
            self._remove_tag_text(match)

    def _strip_block_tags(self, block) -> None:
        if _tag_name(block) != "p":
            return
        tags = self._marker_tags(block)
        if tags:
            self._strip_paragraph_tags(block, tags)

    def _strip_block_tag(self, block, tag: Tag) -> None:
        if _tag_name(block) != "p":
            return
        for item in self._marker_tags(block) or []:
            if item.kind == tag.kind and item.expr == tag.expr and item.name == tag.name:
                self._remove_tag_text(item)
                return

    # -- 正文 / 单元格 ------------------------------------------------------

    def render_story(self, root_el, part) -> None:
        self.render_container(root_el, self.ctx, part)
        for box in list(root_el.iter(_q("txbxContent"))):
            if box in self._seen_boxes:
                continue
            self._seen_boxes.add(box)
            self.render_container(box, self.ctx, part)

    def render_container(self, el, ctx: Context, part) -> None:
        children = self._block_children(el)
        index = 0
        guard = 0
        while index < len(children):
            guard += 1
            if guard > 500000:
                self.ctx.warnings.append("文档结构过于复杂，处理已停止")
                break
            block = children[index]
            name = _tag_name(block)
            if name == "p":
                marker = self._marker_tags(block)
                if marker:
                    match = self._match_block_region(children, index)
                    if match is not None:
                        resume = self._handle_block_region(children, index, part, match, ctx)
                        children = self._block_children(el)
                        index = self._index_of(children, resume)
                        continue
                    self._strip_paragraph_tags(block, marker)
                    if block.getparent() is None:
                        # 标记段落已被删除，后面的内容前移
                        children = self._block_children(el)
                        continue
                    index += 1
                    continue
                self.render_paragraph(block, ctx, part)
                index += 1
            elif name == "tbl":
                self.render_table(block, ctx, part)
                index += 1
            elif name == "sdt":
                content = block.find(_q("sdtContent"))
                if content is not None:
                    self.render_container(content, ctx, part)
                index += 1
            else:
                index += 1

    def _match_block_region(self, children: Sequence[Any], index: int):
        first = self._marker_tags(children[index])
        if not first or first[0].kind != "open":
            return None
        stack: List[str] = []
        else_index = None
        else_tag = None
        for pos in range(index, len(children)):
            if _tag_name(children[pos]) != "p":
                continue
            for tag in self._marker_tags(children[pos]) or []:
                if tag.kind == "open":
                    stack.append(tag.name)
                elif tag.kind == "close":
                    if not stack or not self._kinds_match(stack[-1], tag.name):
                        return None
                    stack.pop()
                    if not stack:
                        return pos, tag, else_index, else_tag
                elif tag.kind == "else" and len(stack) == 1:
                    else_index, else_tag = pos, tag
        return None

    @staticmethod
    def _index_of(children: Sequence[Any], element) -> int:
        if element is None:
            return len(children)
        for pos, child in enumerate(children):
            if child is element:
                return pos
        return len(children)

    def _handle_block_region(self, children: List[Any], index: int, part, match, ctx: Context):
        """处理整段循环/条件，返回继续扫描的位置（区域后面的那个元素）。"""
        close_index, close_tag, else_index, else_tag = match
        marker = self._marker_tags(children[index])
        if not marker:
            return children[index].getnext()
        open_tag = marker[0]
        region = children[index:close_index + 1]
        resume = region[-1].getnext()

        if open_tag.name == "each":
            items = ctx.get_list(open_tag.expr)
            anchor = region[-1]
            copies: List[Tuple[int, int, Any]] = []
            contexts: List[Context] = []
            for item_index, item in enumerate(items):
                contexts.append(ctx.push(item, index=item_index + 1))
                for pos, block in enumerate(region):
                    clone = copy.deepcopy(block)
                    anchor.addnext(clone)
                    anchor = clone
                    copies.append((pos, item_index, clone))
            last = len(region) - 1
            for pos, item_index, clone in copies:
                if pos == 0:
                    self._strip_block_tag(clone, open_tag)
                if pos == last:
                    self._strip_block_tag(clone, close_tag)
                self._render_block(clone, contexts[item_index], part)
            for block in region:
                parent = block.getparent()
                if parent is not None:
                    parent.remove(block)
            return resume

        keep = self._eval_control(open_tag, ctx)
        if keep:
            stop = else_index if else_index is not None else close_index + 1
            keep_blocks = children[index:stop]
            drop_blocks = children[stop:close_index + 1]
        else:
            start = else_index if else_index is not None else close_index + 1
            keep_blocks = children[start:close_index + 1] if else_index is not None else []
            drop_blocks = children[index:start] if else_index is not None else children[index:close_index + 1]
        for block in drop_blocks:
            if block not in keep_blocks:
                parent = block.getparent()
                if parent is not None:
                    parent.remove(block)
        for block in keep_blocks:
            self._strip_block_tags(block)
            self._render_block(block, ctx, part)
        return resume

    def _render_block(self, block, ctx: Context, part) -> None:
        name = _tag_name(block)
        if name == "p":
            self.render_paragraph(block, ctx, part)
        elif name == "tbl":
            self.render_table(block, ctx, part)
        elif name == "sdt":
            content = block.find(_q("sdtContent"))
            if content is not None:
                self.render_container(content, ctx, part)

    # -- 表格 --------------------------------------------------------------

    def render_table(self, tbl, ctx: Context, part) -> None:
        rows: List[Tuple[Any, Context]] = [(row, ctx) for row in tbl.findall(_q("tr"))]
        index = 0
        guard = 0
        while index < len(rows):
            guard += 1
            if guard > 500000:
                self.ctx.warnings.append("表格结构过于复杂，处理已停止")
                break
            row, row_ctx = rows[index]
            if row.getparent() is None:
                index += 1
                continue
            tags = self._row_control_tags(row)
            if not tags:
                self._render_row_cells(row, row_ctx, part)
                index += 1
                continue
            match = self._match_row_region(rows, index)
            if match is None:
                self._strip_row_tags(row, tags)
                self._render_row_cells(row, row_ctx, part)
                index += 1
                continue
            open_tag, close_index, close_tag, else_index, else_tag = match
            region = rows[index:close_index + 1]
            template_rows = [entry[0] for entry in region]

            if open_tag.name == "each":
                items = row_ctx.get_list(open_tag.expr)
                anchor = template_rows[-1]
                new_entries: List[Tuple[Any, Context]] = []
                last = len(template_rows) - 1
                for item_index, item in enumerate(items):
                    item_ctx = row_ctx.push(item, index=item_index + 1)
                    for pos, template in enumerate(template_rows):
                        clone = copy.deepcopy(template)
                        anchor.addnext(clone)
                        marker_only = not TAG_RE.sub("", self._row_text(clone)).strip()
                        if pos == 0:
                            self._strip_row_tag(clone, open_tag)
                        if pos == last:
                            self._strip_row_tag(clone, close_tag)
                        if marker_only and self._drop_blank_row(clone):
                            continue
                        anchor = clone
                        new_entries.append((clone, item_ctx))
                for template in template_rows:
                    parent = template.getparent()
                    if parent is not None:
                        parent.remove(template)
                rows[index:index + len(region)] = new_entries
                continue

            keep = self._eval_control(open_tag, row_ctx)
            if keep:
                stop = else_index if else_index is not None else close_index + 1
                keep_entries = region[:stop - index]
                drop_entries = region[stop - index:]
            else:
                start = else_index if else_index is not None else close_index + 1
                keep_entries = region[start - index:] if else_index is not None else []
                drop_entries = region[:start - index] if else_index is not None else region
            for entry, _ in drop_entries:
                parent = entry.getparent()
                if parent is not None:
                    parent.remove(entry)
            kept_entries: List[Tuple[Any, Context]] = []
            for entry, entry_ctx in keep_entries:
                marker_only = not TAG_RE.sub("", self._row_text(entry)).strip()
                self._strip_row_tags(entry, self._row_control_tags(entry))
                if marker_only and self._drop_blank_row(entry):
                    continue
                kept_entries.append((entry, entry_ctx))
            rows[index:index + len(region)] = kept_entries
            continue

    def _match_row_region(self, rows: Sequence[Tuple[Any, Context]], index: int):
        first = self._row_control_tags(rows[index][0])
        if not first or first[0].kind != "open":
            return None
        open_tag = first[0]
        stack: List[str] = []
        else_index = None
        else_tag = None
        for pos in range(index, len(rows)):
            for tag in self._row_control_tags(rows[pos][0]):
                if tag.kind == "open":
                    stack.append(tag.name)
                elif tag.kind == "close":
                    if stack and self._kinds_match(stack[-1], tag.name):
                        stack.pop()
                        if not stack:
                            return open_tag, pos, tag, else_index, else_tag
                elif tag.kind == "else" and len(stack) == 1:
                    else_index, else_tag = pos, tag
        return None

    def _render_row_cells(self, row, ctx: Context, part) -> None:
        for cell in row.findall(_q("tc")):
            self.render_container(cell, ctx, part)

    @staticmethod
    def _row_text(row) -> str:
        return "".join(
            "".join(node.text or "" for node in paragraph.iter(_q("t")))
            for paragraph in row.iter(_q("p"))
        )

    def _drop_blank_row(self, row) -> bool:
        """整行只剩空格时删掉这一行（循环标记单独占一行时用它收尾）。"""
        if self._row_text(row).strip():
            return False
        for name in ("drawing", "pict", "object"):
            if row.findall(".//" + _q(name)):
                return False
        parent = row.getparent()
        if parent is None:
            return False
        parent.remove(row)
        return True

    # -- 段落 --------------------------------------------------------------

    def render_paragraph(self, p_el, ctx: Context, part) -> None:
        pt = ParaText(p_el)
        if "{{" not in pt.text:
            return
        tags = pt.tags()
        if not tags:
            return
        nodes, _, _ = self._build_nodes(tags)
        ops: List[Tuple[int, int, str]] = []
        self._walk(nodes, ctx, ops, pt.text)
        if ops:
            pt.apply(ops)
        if self.image_specs:
            self._materialize_images(p_el, part)
        if self.table_specs:
            self._materialize_tables(p_el, part)

    def _build_nodes(self, tags: Sequence[Tag], index: int = 0, stop=None):
        nodes: List[Any] = []
        while index < len(tags):
            tag = tags[index]
            if tag.kind == "open" and tag.name in ("if", "unless"):
                then_nodes, term, next_index = self._build_nodes(tags, index + 1, {"else", "close"})
                else_nodes: List[Any] = []
                else_tag = None
                if term is not None and term.kind == "else":
                    else_tag = term
                    else_nodes, term, next_index = self._build_nodes(tags, next_index + 1, {"close"})
                close_tag = None
                if term is not None:
                    close_tag = term
                    next_index += 1
                nodes.append(("if", tag, then_nodes, else_nodes, close_tag, else_tag))
                index = next_index
            elif tag.kind == "open" and tag.name == "each":
                body, term, next_index = self._build_nodes(tags, index + 1, {"close"})
                if term is not None:
                    next_index += 1
                nodes.append(("each", tag, body, term))
                index = next_index
            elif stop and tag.kind in stop:
                return nodes, tag, index
            else:
                nodes.append(("tag", tag))
                index += 1
        return nodes, None, index

    def _walk(
        self,
        nodes: Sequence[Any],
        ctx: Context,
        ops: List[Tuple[int, int, str]],
        text: str = "",
    ) -> None:
        for node in nodes:
            kind = node[0]
            if kind == "tag":
                value = self._tag_text(node[1], ctx)
                if value is not None:
                    ops.append((node[1].start, node[1].end, value))
            elif kind == "if":
                _, open_tag, then_nodes, else_nodes, close_tag, else_tag = node
                keep = self._eval_control(open_tag, ctx)
                if keep:
                    ops.append((open_tag.start, open_tag.end, ""))
                    if else_tag is not None and close_tag is not None:
                        ops.append((else_tag.start, close_tag.end, ""))
                    elif close_tag is not None:
                        ops.append((close_tag.start, close_tag.end, ""))
                    self._walk(then_nodes, ctx, ops, text)
                else:
                    if else_tag is not None:
                        ops.append((open_tag.start, else_tag.end, ""))
                        self._walk(else_nodes, ctx, ops, text)
                        if close_tag is not None:
                            ops.append((close_tag.start, close_tag.end, ""))
                    elif close_tag is not None:
                        ops.append((open_tag.start, close_tag.end, ""))
                    else:
                        ops.append((open_tag.start, open_tag.end, ""))
            elif kind == "each":
                _, open_tag, body, close_tag = node
                items = ctx.get_list(open_tag.expr)
                inner_start = open_tag.end
                inner_end = close_tag.start if close_tag is not None else open_tag.end
                pieces: List[str] = []
                for item_index, item in enumerate(items):
                    item_ctx = ctx.push(item, index=item_index + 1)
                    pieces.append(
                        self._render_nodes_text(body, item_ctx, text, inner_start, inner_end)
                    )
                end = close_tag.end if close_tag is not None else open_tag.end
                ops.append((open_tag.start, end, "".join(pieces)))

    def _render_nodes_text(
        self, nodes: Sequence[Any], ctx: Context, text: str, start: int, end: int
    ) -> str:
        """把一组节点渲染成纯文本（行内循环使用），保留标记之间的字面文字。"""
        out: List[str] = []
        cursor = start
        for node in nodes:
            kind = node[0]
            if kind == "tag":
                tag = node[1]
                if tag.start > cursor:
                    out.append(text[cursor:tag.start])
                value = self._tag_text(tag, ctx)
                out.append(tag.raw if value is None else value)
                cursor = tag.end
            elif kind == "if":
                _, open_tag, then_nodes, else_nodes, close_tag, else_tag = node
                if open_tag.start > cursor:
                    out.append(text[cursor:open_tag.start])
                keep = self._eval_control(open_tag, ctx)
                if keep:
                    inner_start = open_tag.end
                    inner_end = (
                        else_tag.start
                        if else_tag is not None
                        else (close_tag.start if close_tag is not None else open_tag.end)
                    )
                    out.append(
                        self._render_nodes_text(then_nodes, ctx, text, inner_start, inner_end)
                    )
                elif else_tag is not None:
                    inner_start = else_tag.end
                    inner_end = close_tag.start if close_tag is not None else else_tag.end
                    out.append(
                        self._render_nodes_text(else_nodes, ctx, text, inner_start, inner_end)
                    )
                cursor = close_tag.end if close_tag is not None else open_tag.end
            elif kind == "each":
                _, open_tag, body, close_tag = node
                if open_tag.start > cursor:
                    out.append(text[cursor:open_tag.start])
                inner_start = open_tag.end
                inner_end = close_tag.start if close_tag is not None else open_tag.end
                for item_index, item in enumerate(ctx.get_list(open_tag.expr)):
                    out.append(
                        self._render_nodes_text(
                            body,
                            ctx.push(item, index=item_index + 1),
                            text,
                            inner_start,
                            inner_end,
                        )
                    )
                cursor = close_tag.end if close_tag is not None else open_tag.end
        if end > cursor:
            out.append(text[cursor:end])
        return "".join(out)

    def _tag_text(self, tag: Tag, ctx: Context) -> Optional[str]:
        """返回替换文本；None 表示保持原样。"""
        if tag.kind in ("close", "else"):
            return ""
        if tag.kind == "builtin":
            found, value = ctx.get("$" + tag.name)
            if not found:
                return None
            self.filled += 1
            return value_to_text(value, tag.modifiers)
        if tag.kind == "image":
            found, value = ctx.get(tag.name)
            if not found or value in (None, ""):
                if tag.default:
                    value, found = tag.default, True
                elif self.blank_missing:
                    return ""
                else:
                    ctx.missing_key(tag.name)
                    return None
            spec = self._image_spec(value, tag)
            if spec is None:
                self.ctx.warnings.append("图片字段 %s 的取值不是有效路径：%r" % (tag.name, value))
                return ""
            position = len(self.image_specs)
            self.image_specs.append(spec)
            self.filled += 1
            return "%s%d%s" % (IMG_MARK, position, IMG_MARK_END)

        if tag.kind == "table":
            found, value = ctx.get(tag.name)
            if not found or value in (None, ""):
                if tag.default:
                    value, found = tag.default, True
                elif self.blank_missing:
                    return ""
                else:
                    ctx.missing_key(tag.name)
                    return None
            position = len(self.table_specs)
            self.table_specs.append(
                {"value": value, "modifiers": tag.modifiers, "field": tag.name}
            )
            self.filled += 1
            return "%s%d%s" % (TABLE_MARK, position, TABLE_MARK_END)

        found, value = ctx.get(tag.name)
        if not found or value is None:
            if tag.default is not None:
                self.filled += 1
                return tag.default
            if self.blank_missing:
                return ""
            ctx.missing_key(tag.name)
            return None
        self.filled += 1
        return value_to_text(value, tag.modifiers)

    @staticmethod
    def _image_spec(value: Any, tag: Tag) -> Optional[Dict[str, Any]]:
        path = value
        width = height = None
        if isinstance(value, dict):
            for key in ("路径", "文件", "path", "src", "file", "值"):
                if value.get(key):
                    path = value[key]
                    break
            else:
                path = ""
            width = value.get("宽") or value.get("宽度") or value.get("width")
            height = value.get("高") or value.get("高度") or value.get("height")
        if not path or not isinstance(path, str):
            return None
        for modifier in tag.modifiers:
            match = re.match(r"^(宽|宽度|width)\s*[=:：]\s*(.+)$", modifier, re.I)
            if match:
                width = match.group(2).strip()
                continue
            match = re.match(r"^(高|高度|height)\s*[=:：]\s*(.+)$", modifier, re.I)
            if match:
                height = match.group(2).strip()
        return {"path": path, "width": width, "height": height, "field": tag.name}

    def _materialize_images(self, p_el, part) -> None:
        while True:
            target = None
            for t in _iter_text_elements(p_el):
                match = IMG_MARK_RE.search(t.text or "")
                if match:
                    target = (t, match)
                    break
            if target is None:
                return
            t, match = target
            spec = self.image_specs[int(match.group(1))]
            original = t.text or ""
            t.text = original[:match.start()]
            tail = original[match.end():]
            run_el = t.getparent()
            if run_el is None:
                continue

            path = spec["path"]
            if not os.path.isabs(path):
                path = os.path.join(self.base_dir or os.getcwd(), path)
            if not os.path.exists(path):
                self.ctx.warnings.append("找不到图片文件：%s" % spec["path"])
                continue

            picture_run = OxmlElement("w:r")
            run_el.addnext(picture_run)
            if tail:
                tail_run = OxmlElement("w:r")
                properties = run_el.find(_q("rPr"))
                if properties is not None:
                    tail_run.append(copy.deepcopy(properties))
                text_el = OxmlElement("w:t")
                text_el.set(qn("xml:space"), "preserve")
                text_el.text = tail
                tail_run.append(text_el)
                picture_run.addnext(tail_run)

            width = parse_length(spec.get("width"))
            height = parse_length(spec.get("height"))
            if width is None and height is None:
                try:
                    image = DocxImage.from_file(path)
                    limit = Cm(15.5)
                    if image.width and image.width > limit:
                        width = limit
                except Exception:  # noqa: BLE001
                    width = None
            try:
                runner = Run(picture_run, _PartShim(part))
                runner.add_picture(path, width=width, height=height)
                self.image_count += 1
            except Exception as exc:  # noqa: BLE001
                self.ctx.warnings.append("插入图片失败：%s（%s）" % (spec["path"], exc))

    # -- 由数据生成整张表格 ------------------------------------------------

    @staticmethod
    def _column_modifier(formats, column, index) -> str:
        if isinstance(formats, dict):
            for key in (column, str(index), str(index + 1)):
                if key is not None and key in formats:
                    return str(formats[key] or "")
        elif isinstance(formats, (list, tuple)) and index < len(formats):
            return str(formats[index] or "")
        return ""

    def _table_parts(self, spec):
        """把 JSON 里的表格数据整理成 表头 / 数据行 / 列宽。"""
        value = spec.get("value")
        modifiers = [item or "" for item in spec.get("modifiers", ())]
        header = None
        raw_rows = None
        widths = None
        formats = None
        merge = None
        if isinstance(value, dict):
            for key in ("表头", "标题行", "列名", "header", "columns"):
                if key in value:
                    header = value[key]
                    break
            for key in ("行", "数据行", "数据", "明细", "rows", "data"):
                if key in value:
                    raw_rows = value[key]
                    break
            widths = value.get("列宽") or value.get("宽度") or value.get("widths")
            formats = value.get("列格式") or value.get("格式") or value.get("formats")
            merge = value.get("合并列") or value.get("合并单元格") or value.get("merge")
        else:
            raw_rows = value

        if isinstance(raw_rows, dict):
            raw_rows = [raw_rows]
        elif not isinstance(raw_rows, (list, tuple)):
            raw_rows = [] if raw_rows is None else [raw_rows]
        raw_rows = list(raw_rows)

        if isinstance(header, dict):
            if widths is None:
                widths = dict(header)
            header = list(header.keys())

        if header is None:
            header = []
            for row in raw_rows:
                if isinstance(row, dict):
                    for key in row.keys():
                        if key not in header:
                            header.append(key)
        header = [str(item) for item in header]

        if isinstance(widths, dict):
            widths = [widths.get(name) for name in header]
        elif widths is not None:
            widths = list(widths)

        rows: List[List[str]] = []
        for row in raw_rows:
            if isinstance(row, dict):
                cells = []
                for index, name in enumerate(header):
                    modifier = self._column_modifier(formats, name, index)
                    cells.append(value_to_text(row.get(name), [modifier]))
                rows.append(cells)
            elif isinstance(row, (list, tuple)):
                cells = []
                for index, item in enumerate(row):
                    name = header[index] if index < len(header) else None
                    modifier = self._column_modifier(formats, name, index)
                    cells.append(value_to_text(item, [modifier]))
                rows.append(cells)
            elif row is not None:
                modifier = self._column_modifier(formats, header[0] if header else None, 0)
                rows.append([value_to_text(row, [modifier])])

        hide_header = any(
            item in ("无表头", "不要表头", "没有表头", "noheader", "no-header")
            for item in modifiers
        )
        for item in modifiers:
            if item in ("合并首列", "合并第一列", "合并", "merge"):
                merge = True
            elif item.startswith("合并=") or item.startswith("合并:"):
                text = item.split("=", 1)[-1] if "=" in item else item.split(":", 1)[-1]
                merge = [part.strip() for part in re.split(r"[、,，]", text) if part.strip()]
        columns = max([len(header)] + [len(row) for row in rows] + [0])
        for row in rows:
            while len(row) < columns:
                row.append("")
        if not hide_header and columns == 0:
            return header, rows, widths, False, merge
        show_header = (not hide_header) and bool(header)
        return header, rows, widths, show_header, merge

    def _build_table_element(self, rows: int, columns: int):
        if self.doc is None:
            raise RuntimeError("生成整张表格需要文档对象")
        table = self.doc.add_table(rows=rows, cols=columns)
        tbl_el = table._tbl
        parent = tbl_el.getparent()
        if parent is not None:
            parent.remove(tbl_el)
        tbl_pr = tbl_el.tblPr
        for child in list(tbl_pr):
            tbl_pr.remove(child)
        width_el = OxmlElement("w:tblW")
        width_el.set(qn("w:type"), "pct")
        width_el.set(qn("w:w"), "5000")
        tbl_pr.append(width_el)
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            element = OxmlElement("w:" + edge)
            element.set(qn("w:val"), "single")
            element.set(qn("w:sz"), "6")
            element.set(qn("w:space"), "0")
            element.set(qn("w:color"), "D9D9D9")
            borders.append(element)
        tbl_pr.append(borders)
        layout = OxmlElement("w:tblLayout")
        layout.set(qn("w:type"), "autofit")
        tbl_pr.append(layout)
        margins = OxmlElement("w:tblCellMar")
        for edge, value in (("top", 60), ("left", 108), ("bottom", 60), ("right", 108)):
            element = OxmlElement("w:" + edge)
            element.set(qn("w:w"), str(value))
            element.set(qn("w:type"), "dxa")
            margins.append(element)
        tbl_pr.append(margins)
        return tbl_el

    @staticmethod
    def _number_like(text: str) -> bool:
        value = str(text).strip().replace(",", "").replace("￥", "").replace("¥", "").replace("%", "")
        if not value:
            return False
        try:
            float(value)
            return True
        except ValueError:
            return False

    def _column_alignments(self, header, rows, show_header) -> List[str]:
        total = max(len(header), max([len(row) for row in rows] + [0]))
        money_words = ("金额", "小计", "单价", "合计", "工资", "薪", "费用", "税额", "余额", "总价")
        center_words = (
            "序号",
            "编号",
            "编码",
            "日期",
            "时间",
            "状态",
            "是否",
            "数量",
            "个数",
            "单位",
            "类型",
            "默认",
            "是否为空",
        )
        aligns: List[str] = []
        for index in range(total):
            name = header[index] if index < len(header) else ""
            values = [row[index] for row in rows if index < len(row) and str(row[index]).strip() != ""]
            numeric = bool(values) and all(self._number_like(item) for item in values)
            if numeric and any(word in name for word in money_words):
                aligns.append("right")
            elif any(word in name for word in center_words):
                aligns.append("center")
            elif numeric:
                aligns.append("right")
            elif values and all(len(str(item)) <= 6 for item in values):
                aligns.append("center")
            else:
                aligns.append("left")
        return aligns

    def _fill_table_cell(self, tc, text, *, bold, align, size_pt, shade_color=None) -> None:
        for child in list(tc):
            if child.tag != _q("tcPr"):
                tc.remove(child)
        tc_pr = tc.find(_q("tcPr"))
        if tc_pr is None:
            tc_pr = OxmlElement("w:tcPr")
            tc.insert(0, tc_pr)
        for name in ("shd", "vAlign"):
            existing = tc_pr.find(_q(name))
            if existing is not None:
                tc_pr.remove(existing)
        if shade_color:
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), shade_color)
            tc_pr.append(shd)
        v_align = OxmlElement("w:vAlign")
        v_align.set(qn("w:val"), "center")
        tc_pr.append(v_align)

        paragraph = OxmlElement("w:p")
        properties = OxmlElement("w:pPr")
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:before"), "40")
        spacing.set(qn("w:after"), "40")
        spacing.set(qn("w:line"), "300")
        spacing.set(qn("w:lineRule"), "auto")
        properties.append(spacing)
        justify = OxmlElement("w:jc")
        justify.set(qn("w:val"), {"right": "right", "center": "center"}.get(align, "left"))
        properties.append(justify)
        paragraph.append(properties)
        if text:
            run = OxmlElement("w:r")
            rpr = OxmlElement("w:rPr")
            size = OxmlElement("w:sz")
            size.set(qn("w:val"), str(int(size_pt * 2)))
            rpr.append(size)
            size_cs = OxmlElement("w:szCs")
            size_cs.set(qn("w:val"), str(int(size_pt * 2)))
            rpr.append(size_cs)
            if bold:
                rpr.append(OxmlElement("w:b"))
            run.append(rpr)
            for offset, piece in enumerate(
                str(text).replace("\r\n", "\n").replace("\r", "\n").split("\n")
            ):
                if offset:
                    run.append(OxmlElement("w:br"))
                text_el = OxmlElement("w:t")
                text_el.set(qn("xml:space"), "preserve")
                text_el.text = piece
                run.append(text_el)
            paragraph.append(run)
        tc.append(paragraph)

    @staticmethod
    def _insert_tc_property(tc_pr, element) -> None:
        """按 Word 的字段顺序把属性插到 tcPr 里。"""
        tail = ("tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark")
        position = len(tc_pr)
        for index, child in enumerate(tc_pr):
            if _tag_name(child) in tail:
                position = index
                break
        tc_pr.insert(position, element)

    def _merge_column(self, trs, start_row: int, column: int) -> None:
        """把同一列里连续相同的内容纵向合并。"""
        groups = []
        current = None
        for index in range(start_row, len(trs)):
            cells = trs[index].findall(_q("tc"))
            if column >= len(cells):
                break
            text = "".join(node.text or "" for node in cells[column].iter(_q("t")))
            if current is None or text != current[0]:
                current = [text, []]
                groups.append(current)
            current[1].append(index)
        for text, indexes in groups:
            if len(indexes) < 2 or not text.strip():
                continue
            for position, row_index in enumerate(indexes):
                tc = trs[row_index].findall(_q("tc"))[column]
                tc_pr = tc.find(_q("tcPr"))
                if tc_pr is None:
                    tc_pr = OxmlElement("w:tcPr")
                    tc.insert(0, tc_pr)
                v_merge = OxmlElement("w:vMerge")
                if position == 0:
                    v_merge.set(qn("w:val"), "restart")
                    self._insert_tc_property(tc_pr, v_merge)
                else:
                    v_merge.set(qn("w:val"), "continue")
                    self._insert_tc_property(tc_pr, v_merge)
                    self._fill_table_cell(tc, "", bold=False, align="center", size_pt=10.5)

    def _insert_table_after(self, p_el, spec) -> None:
        header, rows, widths, show_header, merge = self._table_parts(spec)
        columns = max([len(header)] + [len(row) for row in rows] + [0])
        if columns == 0:
            self.ctx.warnings.append("表格字段 %s 没有数据行，已跳过" % spec.get("field", ""))
            return
        total_rows = len(rows) + (1 if show_header else 0)
        tbl_el = self._build_table_element(total_rows, columns)
        aligns = self._column_alignments(header, rows, show_header)
        override = None
        if isinstance(spec.get("value"), dict):
            override = spec["value"].get("字号") or spec["value"].get("字体大小")
        if override:
            try:
                size_pt = float(str(override).replace("pt", "").replace("磅", "").strip())
            except ValueError:
                size_pt = 10.5
        elif columns <= 4:
            size_pt = 10.5
        elif columns <= 7:
            size_pt = 9.5
        else:
            size_pt = 8.5

        trs = tbl_el.findall(_q("tr"))
        cursor = 0
        if show_header:
            tr = trs[cursor]
            tr_pr = OxmlElement("w:trPr")
            cant_split = OxmlElement("w:cantSplit")
            tr_pr.append(cant_split)
            header_flag = OxmlElement("w:tblHeader")
            tr_pr.append(header_flag)
            tr.insert(0, tr_pr)
            cells = tr.findall(_q("tc"))
            for index, tc in enumerate(cells):
                self._fill_table_cell(
                    tc,
                    header[index] if index < len(header) else "",
                    bold=True,
                    align="center",
                    size_pt=size_pt,
                    shade_color="D9D9D9",
                )
            cursor += 1
        for row in rows:
            tr = trs[cursor]
            tr_pr = OxmlElement("w:trPr")
            cant_split = OxmlElement("w:cantSplit")
            tr_pr.append(cant_split)
            tr.insert(0, tr_pr)
            cells = tr.findall(_q("tc"))
            for index, tc in enumerate(cells):
                self._fill_table_cell(
                    tc,
                    row[index] if index < len(row) else "",
                    bold=False,
                    align=aligns[index] if index < len(aligns) else "left",
                    size_pt=size_pt,
                )
            cursor += 1

        if widths:
            grid = tbl_el.find(_q("tblGrid"))
            grid_cols = grid.findall(_q("gridCol")) if grid is not None else []
            for index, col in enumerate(grid_cols):
                length = parse_length(widths[index]) if index < len(widths) else None
                if length is not None:
                    col.set(qn("w:w"), str(int(length.twips)))
            for tr in trs:
                for index, tc in enumerate(tr.findall(_q("tc"))):
                    length = parse_length(widths[index]) if index < len(widths) else None
                    if length is None:
                        continue
                    tc_pr = tc.find(_q("tcPr"))
                    if tc_pr is None:
                        tc_pr = OxmlElement("w:tcPr")
                        tc.insert(0, tc_pr)
                    tc_w = tc_pr.find(_q("tcW"))
                    if tc_w is None:
                        tc_w = OxmlElement("w:tcW")
                        tc_pr.insert(0, tc_w)
                    tc_w.set(qn("w:type"), "dxa")
                    tc_w.set(qn("w:w"), str(int(length.twips)))
            layout = tbl_el.tblPr.find(_q("tblLayout"))
            if layout is not None:
                layout.set(qn("w:type"), "fixed")

        if merge:
            data_start = 1 if show_header else 0
            targets: List[int] = []
            if merge is True:
                targets = [0]
            elif isinstance(merge, (list, tuple)):
                for item in merge:
                    if isinstance(item, bool) or item is None:
                        continue
                    if isinstance(item, int):
                        targets.append(item - 1 if item > 0 else 0)
                        continue
                    text = str(item).strip()
                    if not text:
                        continue
                    if text.isdigit():
                        targets.append(int(text) - 1)
                    elif text in header:
                        targets.append(header.index(text))
            for column in targets:
                if 0 <= column < columns:
                    self._merge_column(trs, data_start, column)

        p_el.addnext(tbl_el)
        self.table_count += 1

    def _materialize_tables(self, p_el, part) -> None:
        inserted = False
        while True:
            target = None
            for t in _iter_text_elements(p_el):
                match = TABLE_MARK_RE.search(t.text or "")
                if match:
                    target = (t, match)
                    break
            if target is None:
                break
            t, match = target
            spec = self.table_specs[int(match.group(1))]
            original = t.text or ""
            t.text = original[:match.start()] + original[match.end():]
            t.set(qn("xml:space"), "preserve")
            self._insert_table_after(p_el, spec)
            inserted = True
        if inserted:
            self._cleanup_empty_paragraph(p_el)

    # -- 内容控件（Word 表单域） --------------------------------------------

    def fill_content_controls(self, story_roots: Sequence[Tuple[Any, Any]]) -> None:
        for _part, root in story_roots:
            for sdt in root.iter(_q("sdt")):
                self._fill_one_control(sdt)

    def _fill_one_control(self, sdt) -> None:
        properties = sdt.find(_q("sdtPr"))
        if properties is None:
            return
        keys: List[str] = []
        for name in ("tag", "alias"):
            element = properties.find(_q(name))
            if element is not None:
                value = element.get(_q("val"))
                if value:
                    keys.append(value)
        if not keys:
            return
        for key in keys:
            found, value = self.ctx.get(key)
            if not found or value is None:
                continue
            content = sdt.find(_q("sdtContent"))
            if content is None:
                continue
            texts = list(content.iter(_q("t")))
            if not texts:
                continue
            text = value_to_text(value)
            texts[0].text = text
            texts[0].set(qn("xml:space"), "preserve")
            for other in texts[1:]:
                other.text = ""
            self.filled += 1
            return

    # -- 收尾 --------------------------------------------------------------

    def blank_leftovers(self, root_el) -> None:
        for p_el in root_el.iter(_q("p")):
            pt = ParaText(p_el)
            if "{{" not in pt.text:
                continue
            ops = [(tag.start, tag.end, "") for tag in find_tags(pt.text)]
            if ops:
                pt.apply(ops)


# --------------------------------------------------------------------------
# 对外接口
# --------------------------------------------------------------------------


@dataclass
class FieldInfo:
    """模板里出现的一个普通字段。"""

    name: str
    kind: str = "text"      # text / image / builtin
    modifiers: Tuple[str, ...] = ()
    location: str = ""
    loop: str = ""


@dataclass
class LoopInfo:
    """模板里的一段循环（{{#each 员工}} ... {{/each}}）。"""

    name: str
    fields: List[str] = dc_field(default_factory=list)
    location: str = ""


@dataclass
class TemplateInfo:
    """模板扫描结果。"""

    fields: List[FieldInfo] = dc_field(default_factory=list)
    loops: List[LoopInfo] = dc_field(default_factory=list)
    conditions: List[Tuple[str, str]] = dc_field(default_factory=list)
    images: List[str] = dc_field(default_factory=list)
    tables: List[str] = dc_field(default_factory=list)
    builtins: List[str] = dc_field(default_factory=list)
    conditions_fields: Dict[str, List[str]] = dc_field(default_factory=dict)

    @property
    def data_keys(self) -> List[str]:
        keys: List[str] = []
        for name in self.images:
            if name not in keys:
                keys.append(name)
        for name in self.tables:
            if name not in keys:
                keys.append(name)
        for expression, _location in self.conditions:
            key = _condition_key(expression)
            if key and key not in keys:
                keys.append(key)
        for field in self.fields:
            key = field.name.split(".")[0]
            if key not in keys:
                keys.append(key)
        for loop in self.loops:
            key = loop.name.split(".")[0]
            if key not in keys:
                keys.append(key)
        return keys

    @property
    def is_empty(self) -> bool:
        return not (self.fields or self.loops or self.conditions or self.images or self.tables)


@dataclass
class FillResult:
    """一次填充的结果。"""

    output_path: str = ""
    filled: int = 0
    images: int = 0
    tables: int = 0
    missing: List[str] = dc_field(default_factory=list)
    unused: List[str] = dc_field(default_factory=list)
    warnings: List[str] = dc_field(default_factory=list)
    leftover: List[str] = dc_field(default_factory=list)

    def summary(self) -> str:
        extra = []
        if self.images:
            extra.append("%d 张图片" % self.images)
        if self.tables:
            extra.append("%d 张自动生成的表格" % self.tables)
        head = "已替换 %d 处字段" % self.filled
        if extra:
            head += "（含 %s）" % "、".join(extra)
        lines = [head]
        if self.missing:
            lines.append("缺少数据：%s" % "、".join(self.missing))
        if self.unused:
            lines.append("未用到：%s" % "、".join(self.unused))
        if self.leftover:
            lines.append("仍保留的字段标记：%s" % "、".join(self.leftover))
        for warning in self.warnings:
            lines.append("提示：%s" % warning)
        return "\n".join(lines)


class MissingFieldsError(ValueError):
    """严格模式下缺少数据时抛出。"""

    def __init__(self, fields: Sequence[str]):
        self.fields = list(fields)
        super().__init__("缺少以下数据：%s" % "、".join(self.fields))


def _condition_key(expression: str) -> str:
    text = (expression or "").strip()
    text = text.lstrip("!").strip()
    if text.lower().startswith("not "):
        text = text[4:].strip()
    if not text:
        return ""
    if re.search(r"[=<>]|\|\||&&", text):
        return ""
    if text.startswith("$"):
        return ""
    return text


def _inside(element, name: str) -> bool:
    node = element.getparent()
    while node is not None:
        if _tag_name(node) == name:
            return True
        node = node.getparent()
    return False


def _location_label(p_el, base: str) -> str:
    if _inside(p_el, "txbxContent"):
        return base + "文本框"
    if _inside(p_el, "tbl"):
        return base + "表格"
    return base


def _story_roots(doc) -> List[Tuple[Any, Any, str]]:
    """返回 (part, 根元素, 位置名称)。"""
    roots: List[Tuple[Any, Any, str]] = [(doc.part, doc.element.body, "正文")]
    seen: set = set()
    for section in doc.sections:
        holders = []
        for label, attribute in (
            ("页眉", "header"),
            ("页脚", "footer"),
            ("首页页眉", "first_page_header"),
            ("首页页脚", "first_page_footer"),
            ("偶数页页眉", "even_page_header"),
            ("偶数页页脚", "even_page_footer"),
        ):
            try:
                holders.append((label, getattr(section, attribute)))
            except Exception:  # noqa: BLE001
                continue
        for label, holder in holders:
            try:
                if holder.is_linked_to_previous:
                    continue
                element = holder._element
            except Exception:  # noqa: BLE001
                continue
            if element is None or element in seen:
                continue
            seen.add(element)
            roots.append((holder.part, element, label))
    return roots


def scan_template(template_path: str) -> TemplateInfo:
    """扫描模板，列出所有特殊字段。"""
    doc = _open_docx(template_path)
    info = TemplateInfo()
    seen_fields: set = set()
    seen_images: set = set()
    seen_tables: set = set()
    seen_builtins: set = set()
    for _part, root, label in _story_roots(doc):
        flat: List[Tuple[Tag, str]] = []
        for p_el in root.iter(_q("p")):
            pt = ParaText(p_el)
            if "{{" not in pt.text:
                continue
            location = _location_label(p_el, label)
            for tag in pt.tags():
                flat.append((tag, location))

        stack: List[Any] = []
        active_conditions: List[str] = []
        for tag, location in flat:
            if tag.kind == "open":
                if tag.name == "each":
                    stack.append(LoopInfo(name=tag.expr, location=location))
                else:
                    stack.append(("cond", tag.expr, location))
                    info.conditions.append((tag.expr, location))
                    active_conditions.append(tag.expr)
            elif tag.kind == "close":
                if stack:
                    top = stack.pop()
                    if isinstance(top, LoopInfo):
                        info.loops.append(top)
                        if stack and isinstance(stack[-1], LoopInfo):
                            stack[-1].fields.append(top.name)
                        elif top.name not in seen_fields:
                            seen_fields.add(top.name)
                            info.fields.append(FieldInfo(name=top.name, kind="list", location=top.location))
                    elif active_conditions:
                        active_conditions.pop()
            elif tag.kind == "else":
                continue
            elif tag.kind == "image":
                if tag.name not in seen_images:
                    seen_images.add(tag.name)
                    info.images.append(tag.name)
                for expression in active_conditions:
                    info.conditions_fields.setdefault(expression, [])
                    if tag.name not in info.conditions_fields[expression]:
                        info.conditions_fields[expression].append(tag.name)
            elif tag.kind == "table":
                owner = stack[-1] if stack else None
                if isinstance(owner, LoopInfo):
                    if tag.name not in owner.fields:
                        owner.fields.append(tag.name)
                elif tag.name not in seen_tables:
                    seen_tables.add(tag.name)
                    info.tables.append(tag.name)
                for expression in active_conditions:
                    info.conditions_fields.setdefault(expression, [])
                    if tag.name not in info.conditions_fields[expression]:
                        info.conditions_fields[expression].append(tag.name)
            elif tag.kind == "builtin":
                if tag.name not in seen_builtins:
                    seen_builtins.add(tag.name)
                    info.builtins.append(tag.name)
            else:
                owner = stack[-1] if stack else None
                if isinstance(owner, LoopInfo):
                    if tag.name not in owner.fields:
                        owner.fields.append(tag.name)
                elif tag.name not in seen_fields:
                    seen_fields.add(tag.name)
                    info.fields.append(
                        FieldInfo(
                            name=tag.name,
                            kind="text",
                            modifiers=tag.modifiers,
                            location=location,
                        )
                    )
                for expression in active_conditions:
                    info.conditions_fields.setdefault(expression, [])
                    if tag.name not in info.conditions_fields[expression]:
                        info.conditions_fields[expression].append(tag.name)
    return info


def _sample_value(field: FieldInfo) -> Any:
    joined = " ".join(field.modifiers or ())
    if any(word in joined for word in ("人民币大写", "大写金额", "金额大写")):
        return 0
    if joined in ("日期", "中文日期") or (
        _DATE_PATTERN_RE.match(joined) and re.search(r"[yMdHhms]", joined)
    ):
        return "2026-01-01"
    if any(word in joined for word in ("千分位", "货币")) or (_NUMBER_PATTERN_RE.match(joined) and re.search(r"[0#]", joined)):
        return 0
    if field.kind == "list":
        return []
    return ""


def _set_path(data: Dict[str, Any], path: str, value: Any) -> None:
    parts = [part for part in path.split(".") if part]
    if not parts:
        return
    node = data
    for part in parts[:-1]:
        nxt = node.get(part)
        if not isinstance(nxt, dict):
            nxt = {}
            node[part] = nxt
        node = nxt
    node.setdefault(parts[-1], value)


def build_skeleton(info: TemplateInfo) -> Dict[str, Any]:
    """根据扫描结果生成一份可编辑的 JSON 数据骨架。"""
    data: Dict[str, Any] = {}
    for name in info.images:
        _set_path(data, name, "图片文件路径.png")
    for name in info.tables:
        _set_path(
            data,
            name,
            {"表头": ["列1", "列2"], "行": [{"列1": "示例内容", "列2": "示例内容"}]},
        )
    seen_conditions: set = set()
    for expression, _location in info.conditions:
        key = _condition_key(expression)
        if key and key not in seen_conditions:
            seen_conditions.add(key)
            _set_path(data, key, False)
    field_index = {field.name: field for field in info.fields}
    for field in info.fields:
        if field.kind == "list":
            continue
        _set_path(data, field.name, _sample_value(field))
    for loop in info.loops:
        item: Dict[str, Any] = {}
        for name in loop.fields:
            if str(name).startswith("_"):
                # _index、_序号 这类由循环自动提供，不需要用户填
                continue
            inner = field_index.get(name)
            _set_path(item, name, _sample_value(inner) if inner is not None else "")
        payload = [item] if item else []
        _set_path(data, loop.name, payload)
    return data


def fill_template(
    template_path: str,
    data: Dict[str, Any],
    output_path: Optional[str] = None,
    *,
    blank_missing: bool = False,
    strict: bool = False,
    fill_controls: bool = True,
    base_dir: Optional[str] = None,
) -> FillResult:
    """用 data 填充模板并保存为 output_path。"""
    if not isinstance(data, dict):
        raise ValueError("JSON 数据的最外层必须是对象（键值对）")
    # 以 _ 开头的键是给数据留的注释/说明，不参与文档生成
    data = {key: value for key, value in data.items() if not str(key).startswith("_")}
    doc = _open_docx(template_path)
    ctx = Context(data, source_name=os.path.basename(template_path))
    renderer = _Renderer(
        ctx,
        blank_missing=blank_missing,
        base_dir=base_dir or os.path.dirname(os.path.abspath(template_path)),
        doc=doc,
    )
    roots = _story_roots(doc)
    if fill_controls:
        renderer.fill_content_controls([(part, element) for part, element, _label in roots])
    for part, element, _label in roots:
        renderer.render_story(element, part)

    leftover: List[str] = []
    for _part, element, _label in roots:
        for p_el in element.iter(_q("p")):
            text = ParaText(p_el).text
            if "{{" not in text:
                continue
            for tag in find_tags(text):
                if tag.kind in ("field", "image", "builtin") and tag.name not in leftover:
                    leftover.append(tag.name)

    if blank_missing:
        for _part, element, _label in roots:
            renderer.blank_leftovers(element)

    if output_path is None:
        stem, ext = os.path.splitext(os.path.abspath(template_path))
        output_path = "%s_已填充%s" % (stem, ext or ".docx")
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)

    unused: List[str] = []
    for key in data.keys():
        if key not in ctx.used and not isinstance(data.get(key), dict):
            unused.append(key)
        elif key not in ctx.used and isinstance(data.get(key), dict):
            # 嵌套对象：只要其子键都没被用到，就算未使用
            sub_used = any(name.startswith(key + ".") for name in ctx.used)
            if not sub_used:
                unused.append(key)

    result = FillResult(
        output_path=output_path,
        filled=renderer.filled,
        images=renderer.image_count,
        tables=renderer.table_count,
        missing=list(ctx.missing),
        unused=unused,
        warnings=list(ctx.warnings),
        leftover=leftover,
    )
    if strict and result.missing:
        raise MissingFieldsError(result.missing)
    return result


def _find_soffice() -> Optional[str]:
    import shutil

    for name in ("soffice.com", "soffice.exe", "soffice"):
        found = shutil.which(name)
        if found:
            return found
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.com",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.com",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def export_pdf(docx_path: str, out_dir: Optional[str] = None, soffice: Optional[str] = None) -> str:
    """用 LibreOffice 把 Word 文档转成 PDF（需要本机装 LibreOffice）。"""
    import subprocess
    import tempfile

    executable = soffice or _find_soffice()
    if not executable:
        raise RuntimeError("没有找到 LibreOffice，无法导出 PDF。可以先用 Word 打开文档后另存为 PDF。")
    docx_path = os.path.abspath(docx_path)
    if not os.path.exists(docx_path):
        raise FileNotFoundError("找不到要转换的文档：%s" % docx_path)
    out_dir = os.path.abspath(out_dir or os.path.dirname(docx_path))
    os.makedirs(out_dir, exist_ok=True)
    with tempfile.TemporaryDirectory() as profile:
        command = [
            executable,
            "-env:UserInstallation=file:///%s" % profile.replace("\\", "/").lstrip("/"),
            "--headless",
            "--norestore",
            "--convert-to",
            "pdf",
            "--outdir",
            out_dir,
            docx_path,
        ]
        try:
            completed = subprocess.run(command, capture_output=True, text=True, timeout=240)
        except subprocess.TimeoutExpired:
            completed = None
    target = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    # Windows 上 soffice.exe 会立刻返回，等文件真正写出来
    import time

    deadline = time.time() + 45
    while not os.path.exists(target) and time.time() < deadline:
        time.sleep(0.4)
    if not os.path.exists(target):
        message = ""
        if completed is not None:
            message = (completed.stderr or completed.stdout or "").strip()
        raise RuntimeError("PDF 导出失败：%s" % (message or "未知原因"))
    return target
