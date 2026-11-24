# -*- coding: utf-8 -*-
"""从任意一份「模板」docx 里反推格式方案（Profile）。

自动识别：页面设置、页眉页脚与其对齐方式、每个段落在用的样式、样式的实际
字体/字号/行距/缩进/对齐/大纲级别，并按名称与大纲级别猜出「这一部分是什么」。
模板里定义了但没用到的样式（比如空的「3-图题」）也会拿来补全方案；
摘要前面那行手工排版的题目，会直接量它的实际字号。
识别结果是一份可以直接编辑的 YAML；猜错的地方改 YAML 即可。
"""
from __future__ import annotations

import copy
import re
from collections import Counter, OrderedDict
from typing import Dict, List, Optional, Tuple

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn

from . import docxutil as du
from .profile import Font, Page, Para, Profile, Role

# 每种「部分」的默认内容识别规则（样式名认不出来时的兜底）
DEFAULT_REGEX: Dict[str, List[str]] = {
    "一级标题": [r"^\s*\d{1,2}\s+\S.{0,40}$",
               r"^\s*第\s*[一二三四五六七八九十百]+\s*[章篇]\s*\S.{0,30}$"],
    "二级标题": [r"^\s*\d{1,2}\.\d{1,2}\s+\S.{0,40}$"],
    "三级标题": [r"^\s*\d{1,2}\.\d{1,2}\.\d{1,2}\s+\S.{0,40}$"],
    "四级标题": [r"^\s*\d{1,2}\.\d{1,2}\.\d{1,2}\.\d{1,2}\s+\S.{0,40}$"],
    "摘要标题": [r"^\s*摘\s*要\s*$"],
    "英文摘要标题": [r"^\s*abstract\s*$"],
    "关键词": [r"^\s*关键\s*[词字]\s*[：:]", r"^\s*key\s*words?\s*[：:]"],
    "目录标题": [r"^\s*目\s*录\s*$"],
    "参考文献标题": [r"^\s*参考文献\s*$"],
    "参考文献条目": [r"^\s*\[\d{1,3}\]\s*\S"],
    "致谢标题": [r"^\s*致\s*谢\s*$"],
    "图题": [r"^\s*图\s*\d+([.\-－—]\d+)*[\s:：]+\S.{0,40}$"],
    "表题": [r"^\s*表\s*\d+([.\-－—]\d+)*[\s:：]+\S.{0,40}$"],
    "附录标题": [r"^\s*附\s*录\s*[A-Z一二三四五六七八九十]?\s*$"],
}

# 角色排序：越靠前越优先。目录条目只有样式匹配，放最前面才不会把目录行判成标题。
ROLE_ORDER = ["封面", "论文题目", "英文论文题目",
              "目录条目1", "目录条目2", "目录条目3", "目录条目4", "目录条目5",
              "目录条目",
              "摘要标题", "英文摘要标题", "关键词", "目录标题",
              "一级标题", "二级标题", "三级标题", "四级标题", "五级标题",
              "图题", "表题", "图片",
              "参考文献标题", "参考文献条目", "致谢标题", "附录标题",
              "正文", "表内文字",
              "页眉", "页脚", "其它"]

STYLE_HINTS = [
    (("图题", "图名", "图注", "caption", "figure"), "图题"),
    (("表题", "表名", "table title"), "表题"),
    (("表内", "表格文字", "table text"), "表内文字"),
    (("图片", "插图", "picture"), "图片"),
    (("参考文献", "reference", "bibliography"), "参考文献条目"),
    (("目录", "toc"), "目录条目"),
    (("页眉", "header"), "页眉"),
    (("页脚", "footer"), "页脚"),
    (("摘要", "abstract"), "摘要标题"),
    (("关键词", "keyword"), "关键词"),
    (("正文", "body", "article", "content", "normal", "文章"), "正文"),
]

HEADING_NAMES = {0: "一级标题", 1: "二级标题", 2: "三级标题", 3: "四级标题", 4: "五级标题"}

# 模板里没有的部分，按下面的规则派生（src=以哪个部分为蓝本）
DERIVE = {
    "论文题目":    dict(src="摘要标题", font={"bold": True},
                     para={"align": "center", "outline_level": None,
                           "first_line_chars": 0, "space_before_pt": 6, "space_after_pt": 6}),
    "英文论文题目":  dict(src="论文题目", font={"bold": True},
                     para={"align": "center", "first_line_chars": 0}),
    "摘要标题":    dict(src="一级标题", font={"bold": True},
                     para={"align": "center", "outline_level": 0, "first_line_chars": 0}),
    "英文摘要标题": dict(src="摘要标题", font={"bold": True},
                     para={"align": "center", "first_line_chars": 0}),
    "目录标题":    dict(src="一级标题", font={"bold": True},
                     para={"align": "center", "outline_level": None, "first_line_chars": 0}),
    "参考文献标题": dict(src="一级标题", font={"bold": True},
                     para={"align": "center", "outline_level": 0, "first_line_chars": 0}),
    "致谢标题":    dict(src="一级标题", font={"bold": True},
                     para={"align": "center", "outline_level": 0, "first_line_chars": 0}),
    "附录标题":    dict(src="一级标题", font={"bold": True},
                     para={"align": "center", "outline_level": 0, "first_line_chars": 0}),
    "关键词":      dict(src="正文"),
    "图题":        dict(src="正文", size_delta=-2,
                     para={"first_line_chars": 0, "align": "center",
                           "space_before_pt": 3, "space_after_pt": 6}),
    "表题":        dict(src="正文", size_delta=-2,
                     para={"first_line_chars": 0, "align": "center",
                           "space_before_pt": 6, "space_after_pt": 3}),
    "表内文字":    dict(src="正文", size_delta=-2, para={"first_line_chars": 0}),
    "图片":        dict(src="正文", para={"first_line_chars": 0, "align": "center"}),
    "参考文献条目":  dict(src="正文", para={"first_line_chars": 0, "align": "justify"}),
    "一级标题":    dict(src=None, font={"east_asia": "黑体", "size": 14, "bold": True},
                     para={"align": "center", "outline_level": 0, "line_spacing": 1.5,
                           "first_line_chars": 0, "space_before_pt": 12, "space_after_pt": 6}),
    "二级标题":    dict(src="一级标题", font={"bold": True},
                     para={"align": "left", "outline_level": 1, "first_line_chars": 0}),
    "三级标题":    dict(src="二级标题", para={"align": "left", "outline_level": 2}),
    "四级标题":    dict(src="三级标题", para={"align": "left", "outline_level": 3}),
    "正文":        dict(src=None, font={"east_asia": "宋体", "ascii": "Times New Roman",
                                    "size": 12},
                     para={"align": "justify", "line_spacing": 1.5,
                           "first_line_chars": 2, "space_before_pt": 0, "space_after_pt": 0}),
}


# ---------------------------------------------------------------- 小工具
def _toggle(el) -> bool:
    if el is None:
        return False
    v = el.get(qn("w:val"))
    return str(v) not in ("0", "false", "off")


def _attr(el, name) -> Optional[str]:
    if el is None:
        return None
    return el.get(qn("w:" + name))


def _fnum(v) -> Optional[float]:
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------- 样式解析
def resolve_style(doc, style) -> Tuple[Font, Para, Dict]:
    """沿 basedOn 链 + docDefaults 求样式的「实际生效」格式。"""
    theme = du.theme_fonts(doc)
    chain, seen = [], set()
    cur, depth = style, 0
    while cur is not None and id(cur) not in seen and depth < 12:
        depth += 1
        seen.add(id(cur))
        chain.append(cur)
        try:
            cur = cur.base_style
        except Exception:
            cur = None

    font, para = Font(), Para()
    raw: Dict = {"chain": [s.name for s in chain], "outline": None, "line_rule": None}

    def take_rpr(rPr):
        if rPr is None:
            return
        rf = rPr.find(qn("w:rFonts"))
        if rf is not None:
            if font.ascii is None:
                a = _attr(rf, "ascii") or _attr(rf, "hAnsi")
                if a is None:
                    t = _attr(rf, "asciiTheme") or _attr(rf, "hAnsiTheme")
                    if t:
                        a = theme.get(du.THEME_ASCII.get(t, ""), None)
                if a:
                    font.ascii = a
            if font.east_asia is None:
                e = _attr(rf, "eastAsia")
                if e is None:
                    t = _attr(rf, "eastAsiaTheme")
                    if t:
                        e = theme.get(du.THEME_EA.get(t, ""), None)
                if e:
                    font.east_asia = e
        if font.size is None:
            v = _fnum(_attr(rPr.find(qn("w:sz")), "val"))
            if v:
                font.size = v / 2.0
        for tag, field in (("b", "bold"), ("i", "italic"), ("u", "underline")):
            el = rPr.find(qn("w:" + tag))
            if el is not None and getattr(font, field) is None:
                setattr(font, field, _toggle(el))
        if font.color is None:
            font.color = _attr(rPr.find(qn("w:color")), "val")
        if font.superscript is None:
            va = rPr.find(qn("w:vertAlign"))
            if va is not None:
                font.superscript = _attr(va, "val") == "superscript"
        if font.char_spacing_pt is None:
            v = _fnum(_attr(rPr.find(qn("w:spacing")), "val"))
            if v:
                font.char_spacing_pt = v / 20.0

    def take_ppr(pPr):
        if pPr is None:
            return
        if para.align is None:
            para.align = _attr(pPr.find(qn("w:jc")), "val")
        sp = pPr.find(qn("w:spacing"))
        if sp is not None:
            line = _fnum(_attr(sp, "line"))
            if line and para.line_spacing is None and para.line_exact_pt is None:
                rule = _attr(sp, "lineRule") or "auto"
                raw["line_rule"] = rule
                if rule == "auto":
                    para.line_spacing = line / 240.0
                else:
                    para.line_exact_pt = line / 20.0
            size = font.size or 14.0
            for attr, field in (("beforeLines", "space_before_pt"),
                                ("afterLines", "space_after_pt")):
                v = _fnum(_attr(sp, attr))
                if v is not None and getattr(para, field) is None:
                    setattr(para, field, round(v / 100.0 * size, 2))
            for attr, field in (("before", "space_before_pt"),
                                ("after", "space_after_pt")):
                v = _fnum(_attr(sp, attr))
                if v is not None and getattr(para, field) is None:
                    setattr(para, field, round(v / 20.0, 2))
            for attr, field in (("beforeAutospacing", "auto_space_before"),
                                ("afterAutospacing", "auto_space_after")):
                raw_v = _attr(sp, attr)
                if raw_v is not None and getattr(para, field) is None:
                    setattr(para, field, str(raw_v) in ("1", "true"))
        ind = pPr.find(qn("w:ind"))
        if ind is not None:
            size = font.size or 14.0
            flc = _fnum(_attr(ind, "firstLineChars"))
            if flc is not None and para.first_line_chars is None:
                para.first_line_chars = flc / 100.0
            elif para.first_line_chars is None:
                v = _fnum(_attr(ind, "firstLine"))
                if v:
                    para.first_line_chars = round(v / 20.0 / size, 2)
            lc = _fnum(_attr(ind, "leftChars"))
            if lc is not None and para.left_chars is None:
                para.left_chars = lc / 100.0
            elif para.left_chars is None:
                v = _fnum(_attr(ind, "left"))
                if v:
                    para.left_chars = round(v / 20.0 / size, 2)
        if raw["outline"] is None:
            v = _fnum(_attr(pPr.find(qn("w:outlineLvl")), "val"))
            if v is not None:
                raw["outline"] = int(v)
        for tag, field in (("keepNext", "keep_with_next"),
                           ("pageBreakBefore", "page_break_before")):
            el = pPr.find(qn("w:" + tag))
            if el is not None and getattr(para, field) is None:
                setattr(para, field, _toggle(el))

    for st in chain:
        take_rpr(st.element.find(qn("w:rPr")))
        take_ppr(st.element.find(qn("w:pPr")))

    try:
        dd = doc.styles.element.find(qn("w:docDefaults"))
        if dd is not None:
            take_rpr(dd.find(qn("w:rPrDefault") + "/" + qn("w:rPr")))
            take_ppr(dd.find(qn("w:pPrDefault") + "/" + qn("w:pPr")))
    except Exception:
        pass

    if raw["outline"] is not None:
        para.outline_level = raw["outline"]
    return font, para, raw


def guess_role(name: str, para: Para) -> str:
    n = (name or "").casefold()
    if para.outline_level is not None and 0 <= para.outline_level <= 4:
        return HEADING_NAMES[para.outline_level]
    for keys, role in STYLE_HINTS:
        if any(k in n for k in keys):
            return role
    return "其它"


# ---------------------------------------------------------------- 体检
def _para_align(p) -> Optional[str]:
    jc = p._p.find(qn("w:pPr") + "/" + qn("w:jc"))
    if jc is not None:
        return du.align_canonical(jc.get(qn("w:val")))
    try:
        pPr = p.style.element.find(qn("w:pPr"))
        if pPr is not None:
            jc = pPr.find(qn("w:jc"))
            if jc is not None:
                return du.align_canonical(jc.get(qn("w:val")))
    except Exception:
        pass
    return None


def analyze(path: str) -> Dict:
    doc = Document(path)
    used = Counter()
    for p in du.iter_paragraphs(doc):
        if du.text_of(p).strip() or du.has_picture(p):
            used[du.p_style_name(doc, p) or "(无样式)"] += 1

    used_styles = []
    for name, cnt in used.most_common():
        st = du.find_style(doc, name)
        if st is None:
            continue
        font, para, _raw = resolve_style(doc, st)
        used_styles.append({"style": name, "count": cnt, "font": font, "para": para,
                            "role": guess_role(name, para), "used": True})

    all_styles = []
    try:
        for st in doc.styles:
            if st.type != WD_STYLE_TYPE.PARAGRAPH:
                continue
            font, para, _raw = resolve_style(doc, st)
            all_styles.append({"style": st.name, "count": used.get(st.name, 0),
                               "font": font, "para": para,
                               "role": guess_role(st.name, para),
                               "used": used.get(st.name, 0) > 0})
    except Exception:
        all_styles = list(used_styles)

    header_text, footer_text, page_fields = "", "", []
    header_align, footer_align = None, None
    for sec in doc.sections:
        try:
            for p in sec.header.paragraphs:
                if p.text.strip() and not du.field_instrs(p):
                    if header_align is None:
                        header_align = _para_align(p)
                    if not header_text:
                        header_text = p.text.strip()
            for p in sec.footer.paragraphs:
                if du.field_instrs(p):
                    for i in du.field_instrs(p):
                        if "PAGE" in i.upper():
                            page_fields.append(i)
                    if footer_align is None:
                        footer_align = _para_align(p)
                    continue
                if p.text.strip():
                    if not footer_text:
                        footer_text = p.text.strip()
                    if footer_align is None:
                        footer_align = _para_align(p)
        except Exception:
            pass

    return {"used": used_styles, "all": all_styles, "sections": len(doc.sections),
            "header_text": header_text, "footer_text": footer_text,
            "header_align": header_align, "footer_align": footer_align,
            "footer_fields": page_fields,
            "has_toc": any("TOC" in i.upper() for p in du.iter_paragraphs(doc)
                           for i in du.field_instrs(p))}


# ---------------------------------------------------------------- 组角色
def _role_from(style_name: str, font: Font, para: Para) -> Role:
    return Role(name="", style=style_name, match_styles=[style_name],
                match_regex=[], font=copy.deepcopy(font), para=copy.deepcopy(para))


def _derive(name: str, existing: Dict[str, Role]) -> Role:
    spec = DERIVE.get(name, dict(src="正文"))
    src = existing.get(spec.get("src") or "", None)
    if src is not None:
        font = copy.deepcopy(src.font)
        para = copy.deepcopy(src.para)
    else:
        font = Font(east_asia="宋体", ascii="Times New Roman", size=12)
        para = Para(align="justify", line_spacing=1.5, first_line_chars=2)
    for k, v in (spec.get("font") or {}).items():
        setattr(font, k, v)
    if spec.get("size_delta") and font.size:
        font.size = max(8.0, round(font.size + spec["size_delta"], 1))
    for k, v in (spec.get("para") or {}).items():
        setattr(para, k, v)
    r = Role(name=name, style=name, match_styles=[name],
             match_regex=list(DEFAULT_REGEX.get(name, [])), font=font, para=para)
    r.note = "模板里没找到这一部分，按%s派生，可按需修改" % (spec.get("src") or "通用默认值")
    return r


def build_profile(path: str, name: Optional[str] = None) -> Profile:
    """从模板 docx 生成格式方案。"""
    doc = Document(path)
    info = analyze(path)
    used = info["used"]
    all_styles = info["all"]

    # 正文 = 没被判成标题/图表、且用量最大的那个样式
    body_candidates = [s for s in used if s["role"] in ("正文", "其它")
                       and not re.search(r"toc|页眉|页脚", s["style"], re.I)]
    if body_candidates:
        max(body_candidates, key=lambda s: s["count"])["role"] = "正文"

    grouped: "OrderedDict[str, List[dict]]" = OrderedDict()
    for s in used:
        if s["role"] in ("页眉", "页脚"):
            continue
        if s["role"].startswith("目录条目"):
            continue
        grouped.setdefault(s["role"], []).append(s)

    roles: Dict[str, Role] = {}
    for role_name, items in grouped.items():
        primary = max(items, key=lambda s: s["count"])
        r = _role_from(primary["style"], primary["font"], primary["para"])
        r.name = role_name
        r.match_styles = [s["style"] for s in items]
        r.match_regex = list(DEFAULT_REGEX.get(role_name, []))
        if r.para.line_spacing and r.para.line_exact_pt:
            r.para.line_exact_pt = None
        roles[role_name] = r

    # 用模板里「定义了但没用」的样式补全图题/表题这类空位
    unused_by_role: Dict[str, List[dict]] = {}
    for s in all_styles:
        if s["used"]:
            continue
        unused_by_role.setdefault(s["role"], []).append(s)

    for want in ("图题", "表题", "表内文字", "图片", "参考文献条目"):
        if want in roles:
            continue
        cands = unused_by_role.get(want) or []
        if cands:
            best = min(cands, key=lambda s: (abs(len(s["style"]) - len(want)), s["style"]))
            r = _role_from(best["style"], best["font"], best["para"])
            r.name = want
            r.match_regex = list(DEFAULT_REGEX.get(want, []))
            r.note = "取自模板中未使用但已定义的样式「%s」" % best["style"]
            roles[want] = r

    # 补齐标准部分，按依赖顺序多走几轮
    pending = [k for k in DERIVE.keys() if k not in roles]
    for _ in range(4):
        if not pending:
            break
        for want in list(pending):
            src = DERIVE[want].get("src")
            if src and src not in roles:
                continue
            roles[want] = _derive(want, roles)
            pending.remove(want)
    for want in pending:
        roles[want] = _derive(want, roles)

    # 论文题目：模板里多半是手工排版，直接量「摘要」前面那一段的实际字号
    base_ascii = None
    base = du.find_style(doc, "Normal")
    if base is not None:
        bf, _bp, _br = resolve_style(doc, base)
        base_ascii = bf.ascii
    paras_all = list(du.iter_paragraphs(doc))
    zh_title = _measure_title_before(doc, paras_all, DEFAULT_REGEX["摘要标题"], base_ascii)
    en_title = _measure_title_before(doc, paras_all, DEFAULT_REGEX["英文摘要标题"], base_ascii)
    if zh_title is not None:
        roles["论文题目"].font, roles["论文题目"].para = zh_title
        roles["论文题目"].note = "按模板中「摘要」前一段的实际格式量取"
    if en_title is not None:
        if "英文论文题目" not in roles:
            roles["英文论文题目"] = _derive("英文论文题目", roles)
        roles["英文论文题目"].font, roles["英文论文题目"].para = en_title
        roles["英文论文题目"].note = "按模板中「Abstract」前一段的实际格式量取"

    # 其它派生出来的部分：如果模板里有现成例子，就照着例子量一遍
    _refine_derived_roles(doc, roles)

    # 大纲级别必须保住：目录是按大纲级别收的，量取实际格式时容易把它丢掉
    outline = {"摘要标题": 0, "英文摘要标题": 0, "参考文献标题": 0,
               "致谢标题": 0, "附录标题": 0, "一级标题": 0, "二级标题": 1,
               "三级标题": 2, "四级标题": 3, "五级标题": 4, "目录标题": None}
    for rname, lvl in outline.items():
        r = roles.get(rname)
        if r is not None:
            r.para.outline_level = lvl

    # 目录条目按级别各建一个
    for s in all_styles:
        m = re.match(r"^toc\s*([1-9])$", s["style"], re.I)
        if m and s["used"]:
            lvl = m.group(1)
            roles["目录条目" + lvl] = Role(
                name="目录条目" + lvl, style=s["style"], match_styles=[s["style"]],
                match_regex=[], font=copy.deepcopy(s["font"]), para=copy.deepcopy(s["para"]))

    ordered = sorted(roles.values(),
                     key=lambda r: ROLE_ORDER.index(r.name) if r.name in ROLE_ORDER
                     else len(ROLE_ORDER))

    # 页面设置取「正文所在节」（一般是最后一节）：封面节的页眉距常和正文不同
    sec = doc.sections[-1]
    page = Page()
    try:
        page.width_cm = round(sec.page_width.cm, 2) if sec.page_width else 21.0
        page.height_cm = round(sec.page_height.cm, 2) if sec.page_height else 29.7
        page.margin_top_cm = round(sec.top_margin.cm, 2)
        page.margin_bottom_cm = round(sec.bottom_margin.cm, 2)
        page.margin_left_cm = round(sec.left_margin.cm, 2)
        page.margin_right_cm = round(sec.right_margin.cm, 2)
        page.gutter_cm = round(sec.gutter.cm, 2) if sec.gutter else 0.0
        page.header_distance_cm = round(sec.header_distance.cm, 2)
        page.footer_distance_cm = round(sec.footer_distance.cm, 2)
    except Exception:
        pass
    page.header_text = info["header_text"]
    page.footer_text = info["footer_text"]
    if info.get("header_align"):
        page.header_align = info["header_align"]
    if info.get("footer_align"):
        page.footer_align = info["footer_align"]
    page.page_number = any(re.match(r"^\s*PAGE\s*$", f, re.I)
                           for f in info["footer_fields"])
    hfont = _hf_font_of(doc, "header")
    ffont = _hf_font_of(doc, "footer")
    if hfont:
        page.header_size, page.header_font_east_asia, page.header_font_ascii = hfont
    if ffont:
        page.footer_size, page.footer_font_east_asia, page.footer_font_ascii = ffont

    prof = Profile(
        name=name or "%s 格式" % path.replace("\\", "/").split("/")[-1].rsplit(".", 1)[0],
        source_template=path, roles=ordered, page=page)
    prof.options = {
        "fallback_role": "正文",
        "cover_keep": True,
        "cover_end_regex": r"^\s*摘\s*要|^\s*abstract\b|^\s*目\s*录",
        "default_strip": "font",
        "update_fields": True,
        "backup": True,
        "table_borders": True,
        "table_center": True,
        "base_style": "Normal",
        "header_footer_mode": "preserve",
        "force_page_numbering": False,
    }
    prof.note = ("由「%s」自动提取：%d 个部分、%d 个样式。识别有误可直接改 "
                 "match_styles / match_regex / font / para。" %
                 (path, len(ordered), len(all_styles)))
    return prof


def _hf_font_of(doc, which: str):
    """页眉/页脚文字用的字号与字体（取第一个有文字的段落）。"""
    for sec in doc.sections:
        part = sec.header if which == "header" else sec.footer
        try:
            for p in part.paragraphs:
                if not p.text.strip() or du.field_instrs(p):
                    continue
                st = du.find_style(doc, du.p_style_name(doc, p))
                if st is None:
                    return None
                font, _para, _raw = resolve_style(doc, st)
                return (font.size, font.east_asia, font.ascii)
        except Exception:
            continue
    return None


# ---------------------------------------------------------------- 量取手工排版
def measure_paragraph(doc, p) -> Tuple[Font, Para]:
    """量一个段落的实际格式：样式打底，再用段落里 run 的直接格式覆盖。"""
    font, para = Font(), Para()
    st = du.find_style(doc, du.p_style_name(doc, p))
    if st is not None:
        f0, p0, _ = resolve_style(doc, st)
        font, para = copy.deepcopy(f0), copy.deepcopy(p0)

    pPr = p._p.find(qn("w:pPr"))
    if pPr is not None:
        jc = pPr.find(qn("w:jc"))
        if jc is not None:
            para.align = _attr(jc, "val")
        sp = pPr.find(qn("w:spacing"))
        if sp is not None:
            size = font.size or 14.0
            v = _fnum(_attr(sp, "beforeLines"))
            if v is not None:
                para.space_before_pt = round(v / 100.0 * size, 2)
            v = _fnum(_attr(sp, "afterLines"))
            if v is not None:
                para.space_after_pt = round(v / 100.0 * size, 2)
            v = _fnum(_attr(sp, "before"))
            if v is not None and para.space_before_pt is None:
                para.space_before_pt = round(v / 20.0, 2)
            v = _fnum(_attr(sp, "after"))
            if v is not None and para.space_after_pt is None:
                para.space_after_pt = round(v / 20.0, 2)
            for attr, field in (("beforeAutospacing", "auto_space_before"),
                                ("afterAutospacing", "auto_space_after")):
                raw_v = _attr(sp, attr)
                if raw_v is not None:
                    setattr(para, field, str(raw_v) in ("1", "true"))
            line = _fnum(_attr(sp, "line"))
            if line:
                rule = _attr(sp, "lineRule") or "auto"
                if rule == "auto":
                    para.line_spacing = line / 240.0
                    para.line_exact_pt = None
                else:
                    para.line_exact_pt = line / 20.0
                    para.line_spacing = None
        ind = pPr.find(qn("w:ind"))
        if ind is not None:
            v = _fnum(_attr(ind, "firstLineChars"))
            if v is not None:
                para.first_line_chars = v / 100.0
            elif _fnum(_attr(ind, "firstLine")):
                para.first_line_chars = 0.0

    sizes, bolds, eas, asis = [], [], [], []
    for run in p.runs:
        rPr = run._r.find(qn("w:rPr"))
        if rPr is None:
            continue
        v = _fnum(_attr(rPr.find(qn("w:sz")), "val"))
        if v:
            sizes.append(v / 2.0)
        el = rPr.find(qn("w:b"))
        if el is not None:
            bolds.append(_toggle(el))
        rf = rPr.find(qn("w:rFonts"))
        if rf is not None:
            if _attr(rf, "eastAsia"):
                eas.append(_attr(rf, "eastAsia"))
            if _attr(rf, "ascii"):
                asis.append(_attr(rf, "ascii"))

    # 空段落的字号藏在「段落标记」的 rPr 里，它决定了这个空行有多高
    if pPr is not None:
        mark = pPr.find(qn("w:rPr"))
        if mark is not None:
            v = _fnum(_attr(mark.find(qn("w:sz")), "val"))
            if v:
                sizes.append(v / 2.0)
            rf = mark.find(qn("w:rFonts"))
            if rf is not None:
                if _attr(rf, "eastAsia"):
                    eas.append(_attr(rf, "eastAsia"))
                if _attr(rf, "ascii"):
                    asis.append(_attr(rf, "ascii"))

    if sizes:
        font.size = max(set(sizes), key=sizes.count)
    if bolds:
        font.bold = max(set(bolds), key=bolds.count)
    if eas:
        font.east_asia = max(set(eas), key=eas.count)
    if asis:
        font.ascii = max(set(asis), key=asis.count)
    return font, para


def _refine_derived_roles(doc, roles: Dict[str, Role]) -> None:
    """派生出来的部分若在模板里有现成段落，就量它，别用猜的默认值。"""
    paras = [p for p in du.iter_paragraphs(doc) if du.text_of(p).strip()]
    for name, role in roles.items():
        if not (role.note or "").startswith("模板里没找到"):
            continue
        pats = [re.compile(r, re.I) for r in role.match_regex or []]
        if not pats:
            continue
        for p in paras:
            text = du.text_of(p).strip()
            if not any(pat.fullmatch(text) or pat.search(text) for pat in pats):
                continue
            font, para = measure_paragraph(doc, p)
            if not font.size:
                break
            role.font = font
            role.para = para
            role.note = "按模板里「%s」这一段的实际格式量取" % text[:18]
            break


def _is_cjk_font(name: Optional[str]) -> bool:
    if not name:
        return False
    return any("\u2e80" <= ch <= "\u9fff" or "\uff00" <= ch <= "\uffef" for ch in name)


def _measure_title_before(doc, paras, patterns, base_ascii=None):
    """量「某个标题」前面最近的短行实际格式（页面题目常是手工排版）。

    优先量英文摘要前面那段（通常是英文题目，字号大、能代表题目格式），
    没有再退回中文摘要前面那段。
    """
    pats = [re.compile(r, re.I) for r in patterns]
    anchor = None
    for i, p in enumerate(paras):
        text = du.text_of(p).strip()
        if text and any(pat.search(text) for pat in pats):
            anchor = i
            break
    if anchor is None:
        return None
    for j in range(anchor - 1, max(anchor - 5, -1), -1):
        p = paras[j]
        text = du.text_of(p).strip()
        if not text:
            continue
        if len(text) <= 120 and not text.endswith(("。", "；")):
            font, para = measure_paragraph(doc, p)
            if font.size and font.size >= 15:
                latin = sum(ch.isascii() and ch.isalpha() for ch in text)
                if _is_cjk_font(font.ascii) and latin > len(text) * 0.5:
                    font.ascii = base_ascii or "Times New Roman"
                para.align = "center"
                para.first_line_chars = 0
                return font, para
        break
    return None
