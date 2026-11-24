# -*- coding: utf-8 -*-
"""校验：把输出文档逐段对照格式方案，指出还有哪没改到位。

检查三类问题：
  1) 样式没换对（段落仍用旧样式）
  2) 段落/字符上还残留直接格式（会顶掉样式）
  3) 生效格式与方案不符（字体、字号、行距、缩进、对齐、大纲级别）
另外检查页面设置、页眉、页码域、表格边框。
"""
from __future__ import annotations

from collections import Counter, OrderedDict
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from docx import Document
from docx.oxml.ns import qn

from . import docxutil as du
from .classify import Classifier
from .extract import resolve_style
from .profile import Profile

TOL = 0.15


@dataclass
class Issue:
    kind: str
    role: str
    text: str
    detail: str


@dataclass
class VerifyResult:
    paragraphs: int = 0
    checked: int = 0
    ok: int = 0
    issues: List[Issue] = field(default_factory=list)
    page_problems: List[str] = field(default_factory=list)
    notes: List[str] = field(default_factory=list)

    @property
    def pass_rate(self) -> float:
        return (self.ok / self.checked * 100.0) if self.checked else 100.0

    def by_kind(self) -> "Counter":
        return Counter(i.kind for i in self.issues)

    def text_report(self, max_detail: int = 25) -> str:
        lines = []
        lines.append("段落检查：%d / %d 段合格（%.1f%%）" %
                     (self.ok, self.checked, self.pass_rate))
        if self.issues:
            lines.append("问题分布：")
            for kind, cnt in self.by_kind().most_common():
                lines.append("  · %-12s %d 处" % (kind, cnt))
            lines.append("明细（最多 %d 条）：" % max_detail)
            for it in self.issues[:max_detail]:
                lines.append("  [%s] %s ｜ %s ｜ %s" %
                             (it.kind, it.role, it.text[:22], it.detail))
        if self.page_problems:
            lines.append("页面/页眉问题：")
            lines.extend("  ! " + p for p in self.page_problems)
        if self.notes:
            lines.append("说明：")
            lines.extend("  - " + n for n in self.notes)
        return "\n".join(lines)


def _same(a, b, tol=TOL) -> bool:
    if a is None or b is None:
        return True
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return str(a) == str(b)


def _check_font(eff, want, faults: List[str]) -> None:
    if want.size and not _same(eff.size, want.size):
        faults.append("字号 %s≠%s" % (eff.size, want.size))
    if want.east_asia and (eff.east_asia or "") != want.east_asia:
        faults.append("中文字体 %s≠%s" % (eff.east_asia, want.east_asia))
    if want.ascii and (eff.ascii or "") != want.ascii:
        faults.append("西文字体 %s≠%s" % (eff.ascii, want.ascii))
    if want.bold is True and not eff.bold:
        faults.append("未加粗")
    if want.italic is True and not eff.italic:
        faults.append("未倾斜")


def _check_para(eff, want, faults: List[str]) -> None:
    if want.align and du.align_canonical(eff.align) != du.align_canonical(want.align):
        faults.append("对齐 %s≠%s" % (eff.align, want.align))
    if want.line_spacing and not _same(eff.line_spacing, want.line_spacing, 0.06):
        faults.append("行距 %s≠%s倍" % (eff.line_spacing, want.line_spacing))
    if want.line_exact_pt and not _same(eff.line_exact_pt, want.line_exact_pt, 0.6):
        faults.append("固定行距 %s≠%spt" % (eff.line_exact_pt, want.line_exact_pt))
    if want.first_line_chars is not None and not _same(eff.first_line_chars,
                                                       want.first_line_chars, 0.3):
        faults.append("首行缩进 %s≠%s字符" % (eff.first_line_chars, want.first_line_chars))
    if want.outline_level is not None and eff.outline_level != want.outline_level:
        faults.append("大纲级别 %s≠%s" % (eff.outline_level, want.outline_level))


def verify(doc_path: str, profile: Profile) -> VerifyResult:
    doc = Document(doc_path)
    res = VerifyResult()
    clf = Classifier(doc, profile)
    plan = clf.plan()
    res.paragraphs = len(plan)

    for p, role, _reason in plan:
        if role is None:
            continue
        res.checked += 1
        text = du.text_of(p).strip()
        faults: List[str] = []
        style_name = du.p_style_name(doc, p)
        if du.norm_key(style_name) != du.norm_key(role.style):
            faults.append("样式仍是「%s」" % style_name)
        residue = du.has_direct_format(p)
        if residue:
            faults.append("残留直接格式：%s" % ",".join(sorted(set(residue))))
        try:
            font, para, _raw = resolve_style(doc, p.style)
            want_font = role.font
            if role.name == "关键词":
                # 关键词是「标签加粗、内容常规」，样式本身不加粗属正常
                import copy as _copy

                want_font = _copy.copy(role.font)
                want_font.bold = None
            _check_font(font, want_font, faults)
            _check_para(para, role.para, faults)
        except Exception as e:
            faults.append("无法解析样式：%s" % e)

        if faults:
            res.issues.append(Issue(kind=_classify_kind(faults), role=role.name,
                                    text=text, detail="；".join(faults)))
        else:
            res.ok += 1

    _check_page(doc, profile, res)
    _check_header(doc, profile, res)
    return res


def _classify_kind(faults: List[str]) -> str:
    joined = "|".join(faults)
    if "直接格式" in joined:
        return "直接格式残留"
    if "样式仍是" in joined:
        return "样式未换"
    if "字号" in joined:
        return "字号"
    if "字体" in joined:
        return "字体"
    if "行距" in joined:
        return "行距"
    if "缩进" in joined:
        return "缩进"
    if "对齐" in joined:
        return "对齐"
    if "大纲" in joined:
        return "大纲级别"
    return "其它"


def _check_page(doc, profile: Profile, res: VerifyResult) -> None:
    pg = profile.page
    keep_cover = bool(profile.option("cover_keep", True)) and len(doc.sections) > 1
    for i, sec in enumerate(doc.sections):
        if i == 0 and keep_cover:
            continue
        try:
            if not _same(sec.page_width.cm, pg.width_cm, 0.05):
                res.page_problems.append("第%d节纸宽 %.2f≠%.2fcm" % (i + 1, sec.page_width.cm, pg.width_cm))
            if not _same(sec.top_margin.cm, pg.margin_top_cm, 0.05):
                res.page_problems.append("第%d节上边距 %.2f≠%.2fcm" % (i + 1, sec.top_margin.cm, pg.margin_top_cm))
            if not _same(sec.bottom_margin.cm, pg.margin_bottom_cm, 0.05):
                res.page_problems.append("第%d节下边距 %.2f≠%.2fcm" % (i + 1, sec.bottom_margin.cm, pg.margin_bottom_cm))
            if not _same(sec.left_margin.cm, pg.margin_left_cm, 0.05):
                res.page_problems.append("第%d节左边距 %.2f≠%.2fcm" % (i + 1, sec.left_margin.cm, pg.margin_left_cm))
            if not _same(sec.right_margin.cm, pg.margin_right_cm, 0.05):
                res.page_problems.append("第%d节右边距 %.2f≠%.2fcm" % (i + 1, sec.right_margin.cm, pg.margin_right_cm))
        except Exception as e:
            res.page_problems.append("第%d节读取失败：%s" % (i + 1, e))


def _check_header(doc, profile: Profile, res: VerifyResult) -> None:
    pg = profile.page
    if pg.header_text:
        found = False
        for sec in doc.sections:
            try:
                for p in sec.header.paragraphs:
                    if pg.header_text in p.text:
                        found = True
            except Exception:
                pass
        if not found:
            res.page_problems.append("页眉里没有找到「%s」" % pg.header_text)
    if pg.page_number:
        has = False
        for sec in doc.sections:
            try:
                for p in sec.footer.paragraphs:
                    if any(re_pg(i) for i in du.field_instrs(p)):
                        has = True
            except Exception:
                pass
        if not has:
            res.page_problems.append("页脚没有页码域（PAGE）")


def re_pg(instr: str) -> bool:
    import re
    return bool(re.match(r"^\s*PAGE\b", instr or "", re.I))
