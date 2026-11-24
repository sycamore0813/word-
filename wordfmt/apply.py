# -*- coding: utf-8 -*-
"""格式化引擎：把格式方案落到文档上。

思路是「样式层重置」而不是逐段刷格式：
  1) 先按方案把目标样式的定义改成标准格式（字体/字号/行距/缩进/对齐/大纲级别）；
  2) 再给每个段落换上对应样式；
  3) 最后删掉段落和字符上的直接格式，让它老老实实回到样式。
这样改完的文档是「干净」的：以后在 Word 里改样式，全文跟着变。
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm
from docx.text.paragraph import Paragraph

from . import docxutil as du
from .classify import Classifier
from .profile import Profile, Role


@dataclass
class ApplyReport:
    total_paragraphs: int = 0
    styled: int = 0
    skipped_cover: int = 0
    no_role: int = 0
    role_counts: Counter = field(default_factory=Counter)
    warnings: List[str] = field(default_factory=list)
    changed_sections: int = 0
    cover: dict = field(default_factory=dict)

    def summary(self) -> str:
        lines = [
            "段落总数：%d" % self.total_paragraphs,
            "按方案改写：%d 段" % self.styled,
            "封面保留未动：%d 段" % self.skipped_cover,
            "未匹配到部分：%d 段" % self.no_role,
        ]
        for name, cnt in self.role_counts.most_common():
            lines.append("  · %-10s %d 段" % (name, cnt))
        if self.cover:
            cv = self.cover
            line = "封面：%s，%d 行" % (cv.get("mode", "format"), cv.get("lines", 0))
            if cv.get("filled"):
                line += "，已填：" + "、".join(cv["filled"])
            if cv.get("blank"):
                line += "，留空：" + "、".join(cv["blank"])
            lines.append(line)
            if cv.get("warning"):
                lines.append("  ! " + cv["warning"])
        if self.warnings:
            lines.append("提示：")
            lines.extend("  ! " + w for w in self.warnings)
        return "\n".join(lines)


def _progress(cb: Optional[Callable], msg: str) -> None:
    if cb:
        try:
            cb(msg)
        except Exception:
            pass


# ---------------------------------------------------------------- 样式层
def prepare_styles(doc, profile: Profile, report: ApplyReport) -> Dict[str, object]:
    """把每个角色的目标样式改成方案要求的格式。"""
    mapping: Dict[str, object] = {}
    base_name = profile.option("base_style", "Normal")
    for role in profile.enabled_roles():
        if not role.style:
            continue
        st = du.ensure_style(doc, role.style, base_name=base_name)
        if st is None:
            report.warnings.append("样式「%s」创建失败，跳过该部分" % role.style)
            continue
        try:
            if not role.font.ascii:
                base = du.find_style(doc, base_name)
                base_rPr = base.element.find(qn("w:rPr")) if base is not None else None
                base_fonts = base_rPr.find(qn("w:rFonts")) if base_rPr is not None else None
                fallback = base_fonts.get(qn("w:ascii")) if base_fonts is not None else None
                role.font.ascii = fallback or "Times New Roman"
            du.set_font(st, role.font)
            if role.name == "关键词":
                # 关键词只让「关键词：」这几个字加粗，样式本身保持常规
                _rPr = st.element.find(qn("w:rPr"))
                if _rPr is not None:
                    for _tag in ("b", "bCs"):
                        for _el in _rPr.findall(qn("w:" + _tag)):
                            _rPr.remove(_el)
            # 内置标题样式自带蓝色主题色，方案没指定颜色就强制黑色
            if role.font.color is None:
                du.force_font_color(st, "auto")
            # 让样式出现在 Word 的样式库里
            du.mark_style_qformat(st, priority=profile.option("style_priority", 10))
            du.set_para(st, role.para, font_size_hint=role.font.size)
        except Exception as e:  # pragma: no cover
            report.warnings.append("样式「%s」设置失败：%s" % (role.style, e))
            continue
        mapping[role.name] = st
    return mapping


# ---------------------------------------------------------------- 段落
def _strip_emphasis(p, keep_bold, keep_italic) -> None:
    """方案没要求加粗/斜体时，把源文档里残留的加粗斜体去掉。

    不然「摘要正文」这类段落会沿用原文的加粗，看起来像黑体一样又粗又重。
    """
    for run in p.runs:
        if du.is_field_run(run):
            continue
        rPr = run._r.find(qn("w:rPr"))
        if rPr is None:
            continue
        for tag, keep in (("b", keep_bold), ("bCs", keep_bold),
                          ("i", keep_italic), ("iCs", keep_italic)):
            for el in rPr.findall(qn("w:" + tag)):
                if keep is True:
                    # 方案要求加粗：只清掉「显式取消加粗」的 b val="0"，
                    # 否则源文档残留的 bold=False 会压住样式里的加粗（关键词就是这种）
                    val = el.get(qn("w:val"))
                    if val is not None and str(val).lower() in ("0", "false", "off"):
                        rPr.remove(el)
                    continue
                rPr.remove(el)


def _split_keyword_bold(p, role) -> None:
    """「关键词：」标签加粗，冒号后面的内容常规粗细。

    注意：不能整段删了重建——段落里可能藏着换行/分页符/制表符（模板常用
    「关键词」段末尾的分页符把英文摘要推到下一页），重建会把它们弄丢。
    所以只改带文字的 run，其余（空 run、<w:br>、图片）原样保留。
    """
    import copy as _copy
    import re as _re

    from docx.text.run import Run

    done = False
    for run in list(p.runs):
        text = run.text or ""
        if not text:
            continue
        if not done:
            m = _re.search(r"[:：]", text)
            if m:
                label, rest = text[:m.end()], text[m.end():]
                run.text = label
                run.font.bold = True
                if rest:
                    el = _copy.deepcopy(run._r)
                    run._r.addnext(el)
                    new_run = Run(el, p)
                    new_run.text = rest
                    new_run.font.bold = False
                done = True
                continue
            run.font.bold = True
            continue
        if run.font.bold is not False:
            run.font.bold = False


# ---------------------------------------------------------------- 结构完整性
STRUCT_TAGS = {
    "分页符/换行": "w:br",
    "图片/图形": "w:drawing",
    "域": "w:fldChar",
    "制表符": "w:tab",
    "分节符": "w:sectPr",
}


def _count_struct(doc):
    """数一数文档里的“结构件”：换行/分页符、图片、域、制表符、分节符。"""
    counts = {}
    for label, tag in STRUCT_TAGS.items():
        try:
            counts[label] = len(doc.element.body.findall(".//" + qn(tag)))
        except Exception:
            counts[label] = 0
    # 页眉页脚里的域也算（页码）
    for sec in doc.sections:
        for part in (sec.header, sec.footer):
            try:
                counts["域"] += len(part._element.findall(".//" + qn("w:fldChar")))
            except Exception:
                pass
    return counts


def check_structure(before, after, report) -> None:
    """排版前后对比结构件数量；少了就报警，避免“把分页符/图片弄丢”这类事故。"""
    b, a = _count_struct(before), _count_struct(after)
    lost = []
    for label in STRUCT_TAGS:
        if a.get(label, 0) < b.get(label, 0):
            lost.append("%s %d→%d" % (label, b.get(label, 0), a.get(label, 0)))
    if lost:
        report.warnings.append("结构件减少（请检查）：" + "，".join(lost))


def apply_paragraph(p, role: Role, style, report: ApplyReport, default_strip: str) -> None:
    level = role.strip_direct or default_strip
    if role.style:
        try:
            p.style = style
        except Exception as e:
            report.warnings.append("段落样式设置失败：%s" % e)
            return
    if level != "none":
        du.clear_para_direct(p, level)
    _strip_emphasis(p, role.font.bold, role.font.italic)
    if role.name == "关键词":
        try:
            _split_keyword_bold(p, role)
        except Exception as e:
            report.warnings.append("关键词加粗处理失败：%s" % e)
    # 段落内显式字体（样式已定，但域结果里的 run 可能需要显式化以免 Word 沿用旧值）
    report.styled += 1
    report.role_counts[role.name] += 1


# ---------------------------------------------------------------- 页面
def apply_page(doc, profile: Profile, report: ApplyReport) -> None:
    """页面尺寸/边距对所有节生效；页码格式默认保持原样（避免把罗马数字改成阿拉伯数字）。"""
    pg = profile.page
    keep_cover = bool(profile.option("cover_keep", True)) and len(doc.sections) > 1
    for i, sec in enumerate(doc.sections):
        if i == 0 and keep_cover:
            # 封面单独一节、边距常常和正文不同；既然「封面保留不动」，页面设置也不碰它
            continue
        try:
            sec.page_width = Cm(pg.width_cm)
            sec.page_height = Cm(pg.height_cm)
            sec.top_margin = Cm(pg.margin_top_cm)
            sec.bottom_margin = Cm(pg.margin_bottom_cm)
            sec.left_margin = Cm(pg.margin_left_cm)
            sec.right_margin = Cm(pg.margin_right_cm)
            if pg.gutter_cm:
                sec.gutter = Cm(pg.gutter_cm)
            sec.header_distance = Cm(pg.header_distance_cm)
            sec.footer_distance = Cm(pg.footer_distance_cm)
            if pg.page_number:
                # 默认不动原有页码格式（模板前面用罗马数字、正文用阿拉伯数字是常见做法），
                # 只有显式打开 force_page_numbering 才统一改写。
                if profile.option("force_page_numbering", False):
                    du.set_page_numbering(sec, pg.page_number_format,
                                          pg.page_number_start if i == 0 else None)
            report.changed_sections += 1
        except Exception as e:  # pragma: no cover
            report.warnings.append("第 %d 节页面设置失败：%s" % (i + 1, e))


# ---------------------------------------------------------------- 页眉页脚
def _hf_parts(doc, sec, kind: str):
    """取出本节已存在的页眉/页脚部件（default/even/first 全算）。"""
    out = []
    for ref in sec._sectPr.findall(qn("w:%sReference" % kind)):
        rid = ref.get(qn("r:id"))
        if not rid:
            continue
        part = doc.part.related_parts.get(rid)
        if part is not None:
            out.append(part)
    return out


def _rewrite_hf(doc, part, text: str, align: str, size, ea, ascii_,
                page_field: bool, page_size=None, page_ea=None, page_ascii=None) -> None:
    """重写页眉/页脚部件：保留第一段的段落属性（样式、缩进、边框），只换掉里面的内容。"""
    root = part.element
    paras = root.findall(qn("w:p"))
    keep = paras[0] if paras else None
    for el in list(root):
        if el is not keep and el.tag != qn("w:sectPr"):
            root.remove(el)
    if keep is None:
        keep = OxmlElement("w:p")
        root.append(keep)
    for child in list(keep):
        if child.tag != qn("w:pPr"):
            keep.remove(child)
    para = Paragraph(keep, doc)
    if align:
        para.alignment = du.ALIGN.get(align.lower(), WD_ALIGN_PARAGRAPH.CENTER)
    if text:
        run = para.add_run(text)
        _hf_font(run, size, ea, ascii_)
    if page_field:
        du.add_field(para, "PAGE", "1")
        for run in para.runs:
            _hf_font(run, page_size or size, page_ea or ea, page_ascii or ascii_)


def _hf_font(run, size, ea, ascii_) -> None:
    from .profile import Font
    f = Font(size=size, east_asia=ea, ascii=ascii_)
    try:
        du.set_font(run, f)
    except Exception:
        pass


def _part_text(part) -> str:
    return "".join(t.text or "" for t in part.element.iter(qn("w:t")))


def _part_has_field(part) -> bool:
    for el in part.element.iter(qn("w:instrText")):
        if "PAGE" in (el.text or "").upper():
            return True
    return False


def apply_header_footer(doc, profile: Profile, report: ApplyReport) -> None:
    """preserve 模式：只改「本来就有内容」的页眉页脚。

    · 页眉：这一段本来有文字才更新；空的页眉保持空（封面、摘要页通常就是空的）。
    · 页脚：本来有 PAGE 页码域或文字才更新，顺手把旧的缓存值（如 257）清掉。
    · header_footer_mode: all 时，才强制给每一节补上页眉页码。
    """
    pg = profile.page
    mode = profile.option("header_footer_mode", "preserve")
    done = set()
    for i, sec in enumerate(doc.sections):
        try:
            if pg.header_text:
                for part in _hf_parts(doc, sec, "header"):
                    if id(part.element) in done:
                        continue
                    if mode != "all" and not _part_text(part).strip():
                        continue
                    done.add(id(part.element))
                    _rewrite_hf(doc, part, pg.header_text, pg.header_align,
                                pg.header_size, pg.header_font_east_asia,
                                pg.header_font_ascii, False)
            if pg.footer_text or pg.page_number:
                for part in _hf_parts(doc, sec, "footer"):
                    if id(part.element) in done:
                        continue
                    if mode != "all" and not (_part_text(part).strip() or _part_has_field(part)):
                        continue
                    done.add(id(part.element))
                    _rewrite_hf(doc, part, pg.footer_text, pg.footer_align,
                                pg.footer_size, pg.footer_font_east_asia,
                                pg.footer_font_ascii, pg.page_number,
                                page_size=pg.page_number_size,
                                page_ea=pg.page_number_font_east_asia,
                                page_ascii=pg.page_number_font_ascii)
            if mode == "all":
                if pg.header_text and not _hf_parts(doc, sec, "header"):
                    sec.header.is_linked_to_previous = False
                    for part in _hf_parts(doc, sec, "header"):
                        if id(part.element) in done:
                            continue
                        done.add(id(part.element))
                        _rewrite_hf(doc, part, pg.header_text, pg.header_align,
                                    pg.header_size, pg.header_font_east_asia,
                                    pg.header_font_ascii, False)
                if (pg.footer_text or pg.page_number) and not _hf_parts(doc, sec, "footer"):
                    sec.footer.is_linked_to_previous = False
                    for part in _hf_parts(doc, sec, "footer"):
                        if id(part.element) in done:
                            continue
                        done.add(id(part.element))
                        _rewrite_hf(doc, part, pg.footer_text, pg.footer_align,
                                    pg.footer_size, pg.footer_font_east_asia,
                                    pg.footer_font_ascii, pg.page_number,
                                    page_size=pg.page_number_size,
                                    page_ea=pg.page_number_font_east_asia,
                                    page_ascii=pg.page_number_font_ascii)
        except Exception as e:
            report.warnings.append("第 %d 节页眉页脚设置失败：%s" % (i + 1, e))


# ---------------------------------------------------------------- 表格
def style_tables(doc, profile: Profile, report: ApplyReport) -> None:
    if not profile.option("table_borders", True):
        return
    center = bool(profile.option("table_center", True))
    for tbl in doc.element.body.iter(qn("w:tbl")):
        try:
            tblPr = tbl.find(qn("w:tblPr"))
            if tblPr is None:
                tblPr = OxmlElement("w:tblPr")
                tbl.insert(0, tblPr)
            borders = tblPr.find(qn("w:tblBorders"))
            if borders is None:
                borders = OxmlElement("w:tblBorders")
                tblPr.append(borders)
            for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
                el = borders.find(qn("w:" + edge))
                if el is None:
                    el = OxmlElement("w:" + edge)
                    borders.append(el)
                el.set(qn("w:val"), "single")
                el.set(qn("w:sz"), "4")
                el.set(qn("w:space"), "0")
                el.set(qn("w:color"), "000000")
            if center:
                jc = tblPr.find(qn("w:jc"))
                if jc is None:
                    jc = OxmlElement("w:jc")
                    tblPr.append(jc)
                jc.set(qn("w:val"), "center")
            du.reorder_tblpr(tblPr)
        except Exception as e:
            report.warnings.append("表格格式设置失败：%s" % e)


# ---------------------------------------------------------------- 摘要拆行
def split_abstract_paragraphs(doc) -> int:
    """「摘  要：正文正文…」被塞在一段里时，拆成「摘  要」+ 正文两段。

    很多课程论文的摘要标题和正文是同一段，不拆的话标题只能按正文排，
    看起来就是「摘要和正文一个格式」。
    """
    import copy as _copy
    import re as _re

    head_re = _re.compile(r"^(\s*(?:摘\s*要|英文摘要|abstract)\s*[：:]\s*)", _re.I)
    done = 0
    for p in list(du.iter_paragraphs(doc)):
        if du.in_table(p):
            continue
        text = du.text_of(p)
        m = head_re.match(text)
        if not m:
            continue
        rest = text[m.end():]
        if len(rest) < 30:
            continue
        head = m.group(1).strip().rstrip("：:").strip()
        first_rPr = None
        for r in p.runs:
            rp = r._r.find(qn("w:rPr"))
            if rp is not None:
                first_rPr = rp
                break
        new_p = _copy.deepcopy(p._p)
        for el in list(p._p):
            if el.tag != qn("w:pPr"):
                p._p.remove(el)
        for el in list(new_p):
            if el.tag != qn("w:pPr"):
                new_p.remove(el)
        run = p.add_run(head)
        if first_rPr is not None:
            run._r.insert(0, _copy.deepcopy(first_rPr))
        body_para = Paragraph(new_p, doc)
        run2 = body_para.add_run(rest)
        if first_rPr is not None:
            run2._r.insert(0, _copy.deepcopy(first_rPr))
        p._p.addnext(new_p)
        done += 1
    return done


# ---------------------------------------------------------------- 主流程
def apply_profile(src: str, dst: str, profile: Profile,
                  progress: Optional[Callable[[str], None]] = None,
                  cover=None, cover_mode: str = "format") -> ApplyReport:
    report = ApplyReport()
    _progress(progress, "打开文档…")
    doc = Document(src)

    # 每章另起页
    if profile.page.heading1_new_page:
        for r in profile.roles:
            if r.name == "一级标题":
                r.para.page_break_before = True

    try:
        _struct_before = _count_struct(doc)
    except Exception:
        _struct_before = None

    if profile.option("clean_chars", True):
        try:
            from . import sanitize

            st = sanitize.clean_document(
                doc, drop_nbsp=profile.option("nbsp_to_space", True),
                wide_space_as_spaces=profile.option("wide_space_as_spaces", True))
            if st["chars"]:
                report.warnings.append("清掉隐形字符 %d 个（零宽空格/软连字符/不间断空格等）" % st["chars"])
                _progress(progress, "清掉隐形字符 %d 个" % st["chars"])
        except Exception as e:
            report.warnings.append("清理隐形字符失败：%s" % e)

    _progress(progress, "准备样式定义…")
    styles = prepare_styles(doc, profile, report)

    if profile.option("split_abstract", True):
        n = split_abstract_paragraphs(doc)
        if n:
            _progress(progress, "拆出摘要标题 %d 处" % n)

    _progress(progress, "识别文档各部分…")
    clf = Classifier(doc, profile)
    plan = clf.plan()
    report.total_paragraphs = len(plan)

    default_strip = profile.option("default_strip", "font")
    _progress(progress, "套用格式（共 %d 段）…" % len(plan))
    for p, role, reason in plan:
        if role is None:
            report.skipped_cover += 1
            continue
        st = styles.get(role.name)
        if st is None:
            report.no_role += 1
            continue
        apply_paragraph(p, role, st, report, default_strip)

    _progress(progress, "设置页面…")
    apply_page(doc, profile, report)

    _progress(progress, "设置页眉页脚与页码…")
    apply_header_footer(doc, profile, report)

    _progress(progress, "规范表格…")
    style_tables(doc, profile, report)

    if cover is not None:
        _progress(progress, "套用封面（%s）…" % cover_mode)
        try:
            from .cover import apply_cover
            report.cover = apply_cover(doc, cover, mode=cover_mode)
            if report.cover.get("warning"):
                report.warnings.append(report.cover["warning"])
        except Exception as e:
            report.warnings.append("封面套用失败：%s" % e)

    if profile.option("update_fields", True):
        du.ensure_update_fields(doc)

    if profile.option("ensure_sections", False):
        try:
            from . import skeleton

            added = skeleton.ensure_sections(doc, profile)
            if added:
                report.warnings.append("按标准大纲补齐了章节：" + "、".join(added))
                _progress(progress, "补齐章节：%s" % "、".join(added))
        except Exception as e:
            report.warnings.append("补齐章节失败：%s" % e)

    try:
        if _struct_before is not None:
            after = _count_struct(doc)
            lost = []
            for label in STRUCT_TAGS:
                if after.get(label, 0) < _struct_before.get(label, 0):
                    lost.append("%s %d→%d" % (label, _struct_before.get(label, 0),
                                               after.get(label, 0)))
            if lost:
                report.warnings.append("结构件减少（请检查）：" + "，".join(lost))
    except Exception:
        pass

    _progress(progress, "保存…")
    doc.save(dst)
    return report
