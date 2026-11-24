# -*- coding: utf-8 -*-
"""docx 底层工具：样式查找/创建、直接格式清理、域代码、图片遍历等。"""
from __future__ import annotations

import re
from typing import Iterator, List, Optional

from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph
from docx.text.run import Run

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# 会被「直接格式清理」删掉的字符属性
FONT_ATTRS = ["rFonts", "sz", "szCs", "color", "kern"]
EMPH_ATTRS = ["b", "bCs", "i", "iCs", "u", "strike", "dstrike", "caps", "smallCaps",
              "highlight", "shd", "spacing", "w", "em"]
PARA_ATTRS = ["spacing", "ind", "jc", "outlineLvl", "keepNext", "keepLines",
              "pageBreakBefore", "snapToGrid", "contextualSpacing"]

ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

# XML 里 w:jc 的取值：两端对齐叫 both，不是 justify
ALIGN_XML = {
    "left": "left",
    "center": "center",
    "right": "right",
    "justify": "both",
    "两端对齐": "both",
    "居中": "center",
    "左对齐": "left",
    "右对齐": "right",
    "distribute": "distribute",
}


def align_canonical(val: Optional[str]) -> Optional[str]:
    """把 Word 的各种对齐写法归一到 left/center/right/justify。"""
    if val is None:
        return None
    v = str(val).strip().lower()
    if v in ("both", "justify", "distribute"):
        return "justify"
    if v.startswith("center"):
        return "center"
    if v.startswith("left"):
        return "left"
    if v.startswith("right"):
        return "right"
    return v


# ---------------------------------------------------------------- 基础
def text_of(p) -> str:
    if isinstance(p, Paragraph):
        return p.text
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def norm_key(name: str) -> str:
    return re.sub(r"[\s_]+", "", (name or "")).casefold()


def find_style(doc, name: str):
    """按样式名 / 中文名 / styleId 容错查找样式。"""
    if not name:
        return None
    key = norm_key(name)
    styles = doc.styles
    for st in styles:
        if norm_key(st.name) == key:
            return st
    for st in styles:
        try:
            sid = st.element.get(qn("w:styleId"))
        except Exception:
            sid = None
        if sid and norm_key(sid) == key:
            return st
    return None


def _raw_add_style(doc, name: str, style_id: str = None):
    styles_el = doc.styles.element
    existing = {s.get(qn("w:styleId")) for s in styles_el.findall(qn("w:style"))}
    sid = style_id or re.sub(r"[^\w\u4e00-\u9fff-]", "", name)
    if not sid or sid in existing:
        i = 1
        while ("wf%d" % i) in existing:
            i += 1
        sid = "wf%d" % i
    el = OxmlElement("w:style")
    el.set(qn("w:type"), "paragraph")
    el.set(qn("w:styleId"), sid)
    n = OxmlElement("w:name")
    n.set(qn("w:val"), name)
    el.append(n)
    styles_el.append(el)
    from docx.text.style import ParagraphStyle
    return ParagraphStyle(el, doc.styles)


def ensure_style(doc, name: str, base_name: Optional[str] = None):
    st = find_style(doc, name)
    if st is not None:
        return st
    try:
        st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    except Exception:
        st = _raw_add_style(doc, name)
    if base_name:
        base = find_style(doc, base_name)
        if base is not None:
            try:
                st.base_style = base
            except Exception:
                pass
    return st


def p_style_name(doc, p) -> str:
    try:
        if p.style is not None:
            return p.style.name or ""
    except Exception:
        pass
    return ""


# ---------------------------------------------------------------- 直接格式清理
def is_field_run(run) -> bool:
    r = run._r
    return (r.find(qn("w:fldChar")) is not None) or (r.find(qn("w:instrText")) is not None)


def para_has_field(p) -> bool:
    return (p._p.find(".//" + qn("w:fldChar")) is not None) or \
           (p._p.find(".//" + qn("w:instrText")) is not None)


def _drop(parent, tags):
    for tag in tags:
        for el in parent.findall(qn("w:" + tag)):
            parent.remove(el)


def clear_run_direct(run, level: str = "font") -> None:
    """删掉 run 上的直接格式，让它回到样式的定义。level=font/all。"""
    if is_field_run(run):
        return
    rPr = run._r.find(qn("w:rPr"))
    if rPr is None:
        return
    _drop(rPr, FONT_ATTRS)
    if level == "all":
        _drop(rPr, EMPH_ATTRS)
        for el in rPr.findall(qn("w:rStyle")):
            rPr.remove(el)
    # 空的 rPr 直接删掉
    if len(rPr) == 0 and not rPr.attrib:
        run._r.remove(rPr)


def clear_para_direct(p, level: str = "font") -> None:
    pPr = p._p.find(qn("w:pPr"))
    if pPr is not None:
        _drop(pPr, PARA_ATTRS)
        inner = pPr.find(qn("w:rPr"))
        if inner is not None:
            _drop(inner, FONT_ATTRS)
            if level == "all":
                _drop(inner, EMPH_ATTRS)
            if len(inner) == 0:
                pPr.remove(inner)
    for run in p.runs:
        clear_run_direct(run, level)


def has_direct_format(p) -> List[str]:
    """返回段落上残留的直接格式项（用于校验/报告）。"""
    found = []
    pPr = p._p.find(qn("w:pPr"))
    if pPr is not None:
        for tag in PARA_ATTRS:
            if pPr.find(qn("w:" + tag)) is not None:
                found.append(tag)
    for run in p.runs:
        if is_field_run(run):
            continue
        rPr = run._r.find(qn("w:rPr"))
        if rPr is None:
            continue
        for tag in FONT_ATTRS:
            if rPr.find(qn("w:" + tag)) is not None:
                found.append(tag)
    return found


# ---------------------------------------------------------------- 字体/段落属性写入
def set_font(obj, font) -> None:
    """obj 可以是 style 或 run。按 Font 定义写入，None 表示不改。"""
    el_rPr = obj.element.get_or_add_rPr() if hasattr(obj, "element") else obj._r.get_or_add_rPr()
    rFonts = el_rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        el_rPr.insert(0, rFonts)
    if font.ascii:
        rFonts.set(qn("w:ascii"), font.ascii)
        rFonts.set(qn("w:hAnsi"), font.ascii)
    if font.east_asia:
        rFonts.set(qn("w:eastAsia"), font.east_asia)
    if font.east_asia or font.ascii:
        # 清掉主题字体引用，否则 Word 可能按主题覆盖
        for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            if rFonts.get(qn("w:" + attr)) is not None:
                del rFonts.attrib[qn("w:" + attr)]
        # 一律优先用中文字体渲染「中西文共用」的字符（全角空格 U+3000、中文标点等）。
        # 不设的话 Word 会用西文字体去排这些字符，而 Times New Roman 没有 U+3000，
        # 于是标题里的「1　引言」「摘　要」就变成一个个小方框。
        if font.east_asia:
            rFonts.set(qn("w:hint"), "eastAsia")
    if font.size:
        for tag in ("sz", "szCs"):
            el = el_rPr.find(qn("w:" + tag))
            if el is None:
                el = OxmlElement("w:" + tag)
                el_rPr.append(el)
            el.set(qn("w:val"), str(int(round(font.size * 2))))
    for tag, val in (("b", font.bold), ("i", font.italic)):
        if val is None:
            continue
        _set_toggle(el_rPr, tag, val)
    if font.underline is not None:
        el = el_rPr.find(qn("w:u"))
        if font.underline:
            if el is None:
                el = OxmlElement("w:u")
                el_rPr.append(el)
            # 显式写 single：只写空 <w:u/> 时，Word 不会给「行尾的空格」画线
            el.set(qn("w:val"), "single")
        elif el is not None:
            el_rPr.remove(el)
    if font.color:
        el = el_rPr.find(qn("w:color"))
        if el is None:
            el = OxmlElement("w:color")
            el_rPr.append(el)
        el.set(qn("w:val"), "auto" if font.color == "auto" else font.color.replace("#", ""))
    if font.superscript is not None:
        el = el_rPr.find(qn("w:vertAlign"))
        if el is None:
            el = OxmlElement("w:vertAlign")
            el_rPr.append(el)
        el.set(qn("w:val"), "superscript" if font.superscript else "baseline")
    if font.char_spacing_pt is not None:
        el = el_rPr.find(qn("w:spacing"))
        if el is None:
            el = OxmlElement("w:spacing")
            el_rPr.append(el)
        el.set(qn("w:val"), str(int(round(font.char_spacing_pt * 20))))


def _set_toggle(rPr, tag: str, on: bool) -> None:
    el = rPr.find(qn("w:" + tag))
    if on:
        if el is None:
            el = OxmlElement("w:" + tag)
            rPr.append(el)
        el.attrib.pop(qn("w:val"), None)
    elif el is not None:
        rPr.remove(el)


def set_para(obj, para, font_size_hint: Optional[float] = None) -> None:
    """obj 可以是 style 或 paragraph。按 Para 定义写入。"""
    if hasattr(obj, "element") and not isinstance(obj, Paragraph):
        pPr = obj.element.get_or_add_pPr()
    else:
        pPr = obj._p.get_or_add_pPr()

    if para.align:
        _set_val(pPr, "jc", ALIGN_XML.get(para.align.lower(), para.align.lower()))

    sp = pPr.find(qn("w:spacing"))
    if para.line_spacing or para.line_exact_pt or para.space_before_pt is not None \
            or para.space_after_pt is not None or para.auto_space_before is not None \
            or para.auto_space_after is not None:
        if sp is None:
            sp = OxmlElement("w:spacing")
            pPr.append(sp)
            _reorder(pPr)
    if para.line_spacing:
        sp.set(qn("w:line"), str(int(round(para.line_spacing * 240))))
        sp.set(qn("w:lineRule"), "auto")
    if para.line_exact_pt:
        sp.set(qn("w:line"), str(int(round(para.line_exact_pt * 20))))
        sp.set(qn("w:lineRule"), "exact")
    if para.space_before_pt is not None:
        sp.set(qn("w:before"), str(int(round(para.space_before_pt * 20))))
        sp.attrib.pop(qn("w:beforeLines"), None)
    if para.space_after_pt is not None:
        sp.set(qn("w:after"), str(int(round(para.space_after_pt * 20))))
        sp.attrib.pop(qn("w:afterLines"), None)
    if para.auto_space_before is not None:
        if para.auto_space_before:
            sp.set(qn("w:beforeAutospacing"), "1")
        else:
            sp.attrib.pop(qn("w:beforeAutospacing"), None)
    if para.auto_space_after is not None:
        if para.auto_space_after:
            sp.set(qn("w:afterAutospacing"), "1")
        else:
            sp.attrib.pop(qn("w:afterAutospacing"), None)

    if para.first_line_chars is not None or para.left_chars is not None \
            or para.right_chars is not None:
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            pPr.append(ind)
            _reorder(pPr)
        size = font_size_hint or 14.0
        if para.first_line_chars is not None:
            chars = float(para.first_line_chars)
            ind.set(qn("w:firstLineChars"), str(int(round(chars * 100))))
            ind.set(qn("w:firstLine"), str(int(round(chars * size * 20))))
        if para.left_chars is not None:
            chars = float(para.left_chars)
            ind.set(qn("w:leftChars"), str(int(round(chars * 100))))
            ind.set(qn("w:left"), str(int(round(chars * size * 20))))
        if para.right_chars is not None:
            chars = float(para.right_chars)
            ind.set(qn("w:rightChars"), str(int(round(chars * 100))))
            ind.set(qn("w:right"), str(int(round(chars * size * 20))))

    if para.outline_level is not None:
        _set_val(pPr, "outlineLvl", str(int(para.outline_level)))
    if para.keep_with_next is not None:
        _set_toggle(pPr, "keepNext", para.keep_with_next)
    if para.page_break_before is not None:
        _set_toggle(pPr, "pageBreakBefore", para.page_break_before)
    if para.snap_to_grid is not None:
        _set_val(pPr, "snapToGrid", "1" if para.snap_to_grid else "0")
    _reorder(pPr)


def _set_val(parent, tag: str, val: str):
    el = parent.find(qn("w:" + tag))
    if val is None:
        if el is not None:
            parent.remove(el)
        return
    if el is None:
        el = OxmlElement("w:" + tag)
        parent.append(el)
    el.set(qn("w:val"), str(val))


# pPr 子元素有固定顺序，Word 对顺序敏感
_PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
              "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd",
              "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
              "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
              "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
              "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
              "textDirection", "textAlignment", "textboxTightWrap",
              "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]


def _reorder(pPr) -> None:
    children = list(pPr)
    def key(el):
        tag = el.tag.split("}")[-1]
        return _PPR_ORDER.index(tag) if tag in _PPR_ORDER else 99
    ordered = sorted(children, key=key)
    if ordered != children:
        for el in children:
            pPr.remove(el)
        for el in ordered:
            pPr.append(el)


# ---------------------------------------------------------------- 遍历
def iter_paragraphs(container) -> Iterator[Paragraph]:
    """按文档顺序遍历正文与表格单元格里的段落（含嵌套表格）。"""
    def walk(parent):
        body = parent.element.body if hasattr(parent, "element") else parent._element
        for child in body.iterchildren():
            if child.tag == qn("w:p"):
                yield Paragraph(child, parent)
            elif child.tag == qn("w:tbl"):
                for row in child.findall(qn("w:tr")):
                    for cell in row.findall(qn("w:tc")):
                        for p in cell.findall(qn("w:p")):
                            yield Paragraph(p, parent)
                        for tbl in cell.findall(qn("w:tbl")):
                            for row2 in tbl.findall(qn("w:tr")):
                                for cell2 in row2.findall(qn("w:tc")):
                                    for p2 in cell2.findall(qn("w:p")):
                                        yield Paragraph(p2, parent)
    return walk(container)


def in_table(p) -> bool:
    el = p._p.getparent()
    while el is not None:
        if el.tag == qn("w:tc"):
            return True
        if el.tag == qn("w:body"):
            return False
        el = el.getparent()
    return False


def has_picture(p) -> bool:
    return p._p.find(".//" + qn("w:drawing")) is not None or \
           p._p.find(".//" + qn("w:pict")) is not None


def outline_level(p) -> Optional[int]:
    """段落的大纲级别：先看直接格式，再看样式继承链；正文返回 None。"""
    val = None
    pPr = p._p.find(qn("w:pPr"))
    if pPr is not None:
        el = pPr.find(qn("w:outlineLvl"))
        if el is not None:
            val = el.get(qn("w:val"))
    if val is None:
        st = p.style
        seen = set()
        while st is not None and st.style_id not in seen:
            seen.add(st.style_id)
            sppr = st.element.find(qn("w:pPr"))
            if sppr is not None:
                el = sppr.find(qn("w:outlineLvl"))
                if el is not None:
                    val = el.get(qn("w:val"))
                    break
            st = st.base_style
    try:
        n = int(val)
    except (TypeError, ValueError):
        return None
    return None if n >= 9 else n


def mark_style_qformat(style, priority: int = 10) -> None:
    """让自定义样式出现在 Word 的样式库（样式选项卡）里。

    Word 只显示带 qFormat 标记的样式；顺带给个 uiPriority，
    这样它会排在「推荐」区而不是被折叠起来。
    """
    el = style.element
    pri = el.find(qn("w:uiPriority"))
    if pri is None:
        pri = OxmlElement("w:uiPriority")
        el.append(pri)
    pri.set(qn("w:val"), str(int(priority)))
    if el.find(qn("w:qFormat")) is None:
        el.append(OxmlElement("w:qFormat"))


def force_font_color(obj, color: str = "auto") -> None:
    """把字体颜色强制设成黑（清掉主题色），用于盖掉内置标题样式自带的蓝色。"""
    rPr = obj.element.get_or_add_rPr() if hasattr(obj, "element") else obj._r.get_or_add_rPr()
    el = rPr.find(qn("w:color"))
    if el is None:
        el = OxmlElement("w:color")
        rPr.append(el)
    el.set(qn("w:val"), color)
    for attr in ("themeColor", "themeShade", "themeTint"):
        if el.get(qn("w:" + attr)) is not None:
            del el.attrib[qn("w:" + attr)]


# ---------------------------------------------------------------- 域代码
def add_field(paragraph, instr: str, placeholder: str = "1"):
    """在段落末尾插入一个复杂域（begin/instr/separate/结果/end）。"""
    def _run_with(el):
        r = OxmlElement("w:r")
        r.append(el)
        paragraph._p.append(r)
        return r

    b = OxmlElement("w:fldChar"); b.set(qn("w:fldCharType"), "begin")
    _run_with(b)
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve")
    it.text = " %s " % instr
    _run_with(it)
    s = OxmlElement("w:fldChar"); s.set(qn("w:fldCharType"), "separate")
    _run_with(s)
    t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = placeholder
    _run_with(t)
    e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), "end")
    _run_with(e)


def field_instrs(p) -> List[str]:
    return [(el.text or "").strip() for el in p._p.iter(qn("w:instrText"))]


def ensure_update_fields(doc) -> None:
    """在 settings.xml 里置 <w:updateFields/>，打开文档时 Word 自动更新域。"""
    settings = doc.settings.element
    el = settings.find(qn("w:updateFields"))
    if el is None:
        el = OxmlElement("w:updateFields")
        settings.append(el)
    el.set(qn("w:val"), "true")


def set_page_numbering(section, fmt: str = "decimal", start: Optional[int] = 1) -> None:
    sectPr = section._sectPr
    pg = sectPr.find(qn("w:pgNumType"))
    if pg is None:
        pg = OxmlElement("w:pgNumType")
        # pgNumType 必须放在 cols 之前
        cols = sectPr.find(qn("w:cols"))
        if cols is not None:
            cols.addprevious(pg)
        else:
            sectPr.append(pg)
    pg.set(qn("w:fmt"), fmt)
    if start is None:
        pg.attrib.pop(qn("w:start"), None)
    else:
        pg.set(qn("w:start"), str(int(start)))


# ---------------------------------------------------------------- 主题字体
def theme_fonts(doc) -> dict:
    out = {}
    try:
        from docx.oxml import parse_xml
        part = None
        for p in doc.part.package.parts:
            if str(p.partname).endswith("theme/theme1.xml"):
                part = p
                break
        if part is None:
            return out
        root = parse_xml(part.blob)
        A = "http://schemas.openxmlformats.org/drawingml/2006/main"
        for kind, tag in (("major", "majorFont"), ("minor", "minorFont")):
            fs = root.find(".//{%s}%s" % (A, tag))
            if fs is None:
                continue
            latin = fs.find("{%s}latin" % A)
            ea = fs.find("{%s}ea" % A)
            if latin is not None:
                out[kind + "_latin"] = latin.get("typeface")
            if ea is not None:
                out[kind + "_ea"] = ea.get("typeface")
    except Exception:
        pass
    return out


THEME_EA = {"minorEastAsia": "minor_ea", "majorEastAsia": "major_ea"}
THEME_ASCII = {"minorHAnsi": "minor_latin", "majorHAnsi": "major_latin",
               "minorAscii": "minor_latin", "majorAscii": "major_latin"}

_TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
                "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
                "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout",
                "tblCellMar", "tblLook", "tblCaption", "tblDescription"]


def reorder_tblpr(tblPr) -> None:
    children = list(tblPr)
    def key(el):
        tag = el.tag.split("}")[-1]
        return _TBLPR_ORDER.index(tag) if tag in _TBLPR_ORDER else 99
    ordered = sorted(children, key=key)
    if ordered != children:
        for el in children:
            tblPr.remove(el)
        for el in ordered:
            tblPr.append(el)
