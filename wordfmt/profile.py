# -*- coding: utf-8 -*-
"""格式方案（Profile）：数据模型 + YAML 读写。

方案 = 页面设置 + 若干「部分（role）」的格式定义。
每个 role 有三件事：
  1) 怎么认出这一部分：match_styles（样式名，优先）+ match_regex（内容正则，兜底）
  2) 套哪个样式：style（目标文档里要用的段落样式名，不存在会自动创建）
  3) 这个样式长什么样：font / para
"""
from __future__ import annotations

import dataclasses
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import yaml

ALIGN_MAP = {
    "left": "left",
    "center": "center",
    "right": "right",
    "justify": "justify",
    "两端对齐": "justify",
    "居中": "center",
    "左对齐": "left",
    "右对齐": "right",
}


@dataclass
class Font:
    east_asia: Optional[str] = None      # 中文字体，如 仿宋 / 黑体 / 宋体
    ascii: Optional[str] = None          # 西文字体，如 Times New Roman
    size: Optional[float] = None         # 字号（磅），14 = 四号，10.5 = 五号
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    color: Optional[str] = None          # "auto" 或 "000000"
    superscript: Optional[bool] = None
    char_spacing_pt: Optional[float] = None   # 字间距（磅），正=加宽


@dataclass
class Para:
    align: Optional[str] = None          # left/center/right/justify
    line_spacing: Optional[float] = None       # 倍数行距，1.5 = 1.5 倍
    line_exact_pt: Optional[float] = None     # 固定行距（磅），与 line_spacing 二选一
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None
    auto_space_before: Optional[bool] = None   # Word「自动调整段落间距」
    auto_space_after: Optional[bool] = None
    first_line_chars: Optional[float] = None  # 首行缩进（字符），2 = 空两格
    left_chars: Optional[float] = None        # 左缩进（字符）
    right_chars: Optional[float] = None
    outline_level: Optional[int] = None       # 大纲级别 0-8，0=一级标题
    keep_with_next: Optional[bool] = None
    page_break_before: Optional[bool] = None
    snap_to_grid: Optional[bool] = None


@dataclass
class Role:
    name: str = ""                       # 中文名，如「正文」「一级标题」
    style: str = ""                      # 目标样式名（不存在则创建）
    match_styles: List[str] = field(default_factory=list)
    match_regex: List[str] = field(default_factory=list)
    regex_mode: str = "search"           # search=包含即命中；fullmatch=整段匹配
    font: Font = field(default_factory=Font)
    para: Para = field(default_factory=Para)
    strip_direct: Optional[str] = None   # font（默认）/ all / none
    enabled: bool = True
    note: str = ""                       # 备注：这一部分通常长什么样

    def describe(self) -> str:
        bits = []
        if self.font.east_asia:
            bits.append(self.font.east_asia)
        if self.font.ascii:
            bits.append(self.font.ascii)
        if self.font.size:
            bits.append("%gpt" % self.font.size)
        if self.font.bold:
            bits.append("加粗")
        if self.para.align:
            bits.append(ALIGN_MAP.get(self.para.align, self.para.align))
        if self.para.line_spacing:
            bits.append("%g倍行距" % self.para.line_spacing)
        if self.para.line_exact_pt:
            bits.append("固定%gpt" % self.para.line_exact_pt)
        if self.para.first_line_chars:
            bits.append("首行缩进%g字符" % self.para.first_line_chars)
        return " ".join(bits) or "（未设置）"


@dataclass
class Page:
    width_cm: float = 21.0
    height_cm: float = 29.7
    margin_top_cm: float = 2.54
    margin_bottom_cm: float = 2.54
    margin_left_cm: float = 3.17
    margin_right_cm: float = 3.17
    gutter_cm: float = 0.0
    header_distance_cm: float = 1.7
    footer_distance_cm: float = 1.8

    header_text: str = ""
    header_style: str = "header"
    header_align: str = "center"
    header_size: Optional[float] = None
    header_font_east_asia: Optional[str] = None
    header_font_ascii: Optional[str] = None
    footer_text: str = ""
    footer_style: str = "footer"
    footer_align: str = "center"
    footer_size: Optional[float] = None
    footer_font_east_asia: Optional[str] = None
    footer_font_ascii: Optional[str] = None

    page_number: bool = True
    page_number_position: str = "footer_center"   # footer_center / footer_right / header_right
    page_number_format: str = "decimal"           # decimal / upperRoman / lowerRoman / chineseCounting
    page_number_start: Optional[int] = 1
    page_number_size: float = 9.0
    page_number_font_east_asia: str = "宋体"
    page_number_font_ascii: str = "Times New Roman"

    heading1_new_page: bool = False


@dataclass
class Profile:
    name: str = "未命名格式方案"
    source_template: str = ""
    note: str = ""
    page: Page = field(default_factory=Page)
    roles: List[Role] = field(default_factory=list)
    style_overrides: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    table: Dict[str, Any] = field(default_factory=dict)
    options: Dict[str, Any] = field(default_factory=dict)

    # ---------- 便捷访问 ----------
    def role(self, name: str) -> Optional[Role]:
        for r in self.roles:
            if r.name == name:
                return r
        return None

    def enabled_roles(self) -> List[Role]:
        return [r for r in self.roles if r.enabled]

    def option(self, key: str, default=None):
        return self.options.get(key, default)

    # ---------- 读写 ----------
    @classmethod
    def load(cls, path: str) -> "Profile":
        with open(path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
        return cls.from_dict(data)

    def save(self, path: str) -> None:
        with open(path, "w", encoding="utf-8") as f:
            yaml.safe_dump(self.to_dict(), f, allow_unicode=True,
                           sort_keys=False, default_flow_style=False, width=200)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "Profile":
        prof = cls()
        for key, val in (data or {}).items():
            if not hasattr(prof, key):
                raise ValueError("格式方案里有未知字段：%s" % key)
            cur = getattr(prof, key)
            if key == "roles":
                prof.roles = [_build(Role, r, "roles") for r in (val or [])]
            elif key == "page":
                prof.page = _build(Page, val, "page")
            else:
                if dataclasses.is_dataclass(cur) and isinstance(val, dict):
                    val = _build(type(cur), val, key)
                setattr(prof, key, val)
        return prof

    def to_dict(self) -> Dict[str, Any]:
        return _dump(self)


def _build(cls, data: Optional[Dict[str, Any]], where: str = ""):
    obj = cls()
    for key, val in (data or {}).items():
        if not hasattr(obj, key):
            raise ValueError("格式方案 %s 里有未知字段：%s" % (where or cls.__name__, key))
        cur = getattr(obj, key)
        if dataclasses.is_dataclass(cur) and isinstance(val, dict):
            setattr(obj, key, _build(type(cur), val, where))
        else:
            setattr(obj, key, val)
    return obj


def _dump(obj):
    if dataclasses.is_dataclass(obj):
        out = {}
        for f in dataclasses.fields(obj):
            v = getattr(obj, f.name)
            if v is None or v == [] or v == {}:
                continue
            out[f.name] = _dump(v)
        return out
    if isinstance(obj, list):
        return [_dump(x) for x in obj]
    if isinstance(obj, dict):
        return {k: _dump(v) for k, v in obj.items() if v is not None}
    return obj
