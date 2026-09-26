# -*- coding: utf-8 -*-
"""Word 模板编译器：把 Word 文档里的特殊字段用 JSON 数据替换掉。"""

from .engine import (  # noqa: F401
    Context,
    FieldInfo,
    FillResult,
    LoopInfo,
    MissingFieldsError,
    TemplateInfo,
    build_skeleton,
    export_pdf,
    fill_template,
    load_json,
    rmb_upper,
    scan_template,
    value_to_text,
)

__version__ = "1.0.0"

__all__ = [
    "Context",
    "FieldInfo",
    "FillResult",
    "LoopInfo",
    "MissingFieldsError",
    "TemplateInfo",
    "build_skeleton",
    "export_pdf",
    "fill_template",
    "load_json",
    "rmb_upper",
    "scan_template",
    "value_to_text",
]
