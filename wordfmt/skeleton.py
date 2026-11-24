# -*- coding: utf-8 -*-
"""从模板生成「空白结构稿」：只保留文章的骨架（封面可选、摘要、目录、各级标题、
参考文献、致谢…），正文内容全部删掉，留给使用者往里填。

顺带做两件让结构更规整的事：
  1. 按标题编号（1 / 1.1 / 1.1.1）归一化大纲级别，避免模板里层级混乱；
  2. 目录只留 TOC 域本身，旧的目录条目删掉（打开后更新域会自动生成）。
"""
import copy
import os
import re

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from . import docxutil as du
from .classify import Classifier
from .cover import cover_span

# 这些角色的段落属于“结构”，要保留
STRUCT_ROLES = {
    "摘要标题", "英文摘要标题", "目录标题",
    "一级标题", "二级标题", "三级标题", "四级标题", "五级标题",
    "参考文献标题", "致谢标题", "附录标题",
}

LEVEL_ROLE = {1: "一级标题", 2: "二级标题", 3: "三级标题", 4: "四级标题"}


def _is_toc_style(p):
    name = (p.style.name if p.style is not None else "") or ""
    return name.lower().startswith("toc")


def _has_toc_field(p):
    xml = p._p.xml
    return "TOC \\" in xml or " TOC " in xml


def _has_sectpr(p):
    pPr = p._p.find(qn("w:pPr"))
    return pPr is not None and pPr.find(qn("w:sectPr")) is not None


DATE_LINE = re.compile(r"^[二○〇零0-9一二三四五六七八九十\s年]{2,}[月日]?\s*$")


def _looks_like_cover_date(text):
    t = re.sub(r"\s+", "", text or "")
    return bool(t) and bool(re.match(r"^[二○〇零0-9一二三四五六七八九十]{2,}年[0-9一二三四五六七八九十]{0,2}月$", t))


def _heading_level(text):
    """从「1 引言」「2.1 试验材料」这类文字里看层级。"""
    m = re.match(r"^\s*(\d+(?:\.\d+)*)[\.、]?\s+\S", text or "")
    if not m:
        return None
    depth = m.group(1).count(".") + 1
    return min(depth, 4)


def build_skeleton(src, dst=None, profile=None, keep_cover=False, pad=True,
                   normalize_levels=True, drop_toc_entries=True):
    """生成空白结构稿，返回 (输出路径, 统计信息)。"""
    doc = Document(src)
    if profile is None:
        from .extract import build_profile

        profile = build_profile(src)
    clf = Classifier(doc, profile)

    paras = list(du.iter_paragraphs(doc))
    span = cover_span(doc) if keep_cover else None
    cover_range = set(range(span[0], span[1] + 1)) if span else set()

    keep, stats = [], {"headings": 0, "cover": 0, "removed": 0, "toc": 0, "relabeled": 0}
    toc_seen = False
    for idx, p in enumerate(paras):
        if idx in cover_range:
            keep.append(p)
            stats["cover"] += 1
            continue
        if not keep_cover and _looks_like_cover_date(du.text_of(p)):
            # 封面上的「二○二 年 月」：清空文字但保留段落（里面可能挂着分节符）
            for r in list(p.runs):
                r._r.getparent().remove(r._r)
            keep.append(p)
            stats["removed"] += 1
            continue
        if _has_sectpr(p):
            keep.append(p)
            continue
        if _has_toc_field(p):
            keep.append(p)
            toc_seen = True
            stats["toc"] += 1
            continue
        if drop_toc_entries and _is_toc_style(p):
            # 目录里旧的条目：留着 TOC 域即可，条目交给 Word 重新生成
            stats["removed"] += 1
            continue
        try:
            name = getattr(clf.match(p), "name", "") or ""
        except Exception:
            name = ""
        if name in STRUCT_ROLES:
            keep.append(p)
            if "标题" in name:
                stats["headings"] += 1
                if normalize_levels:
                    lvl = _heading_level(du.text_of(p))
                    role = profile.role(LEVEL_ROLE.get(lvl, "")) if lvl else None
                    if role is not None and role.style:
                        try:
                            st = du.ensure_style(doc, role.style, base_name="Normal")
                            if st is not None:
                                p.style = st
                                stats["relabeled"] += 1
                        except Exception:
                            pass
            continue
        stats["removed"] += 1

    keep_ids = {id(p._p) for p in keep}
    body = doc.element.body
    for p in paras:
        if id(p._p) in keep_ids:
            continue
        parent = p._p.getparent()
        if parent is not None:
            parent.remove(p._p)
    for tbl in list(body.findall(qn("w:tbl"))):
        body.remove(tbl)

    if pad:
        body_style = profile.role("正文")
        for p in list(keep):
            if _has_sectpr(p) or _has_toc_field(p):
                continue
            try:
                name = getattr(clf.match(p), "name", "") or ""
            except Exception:
                name = ""
            if "标题" not in name:
                continue
            blank = copy.deepcopy(p._p)
            for child in list(blank):
                if child.tag != qn("w:pPr"):
                    blank.remove(child)
            p._p.addnext(blank)
            np = Paragraph(blank, p._parent)
            try:
                if body_style and body_style.style:
                    st = du.ensure_style(doc, body_style.style, base_name="Normal")
                    if st is not None:
                        np.style = st
            except Exception:
                pass

    out = dst or os.path.splitext(src)[0] + "_结构稿.docx"
    doc.save(out)
    stats.update({"out": out, "kept": len(keep)})
    return out, stats

# ---------------------------------------------------------------- 标准大纲
OUTLINE_SECTIONS = ["引言", "文献综述", "研究方法", "结果与分析", "讨论", "结论与展望"]


def _add_para(doc, profile, text, role_name, base="Normal"):
    """按角色新建一个段落：样式 + 字体/段落格式都照角色写，和排版时一致。"""
    p = doc.add_paragraph()
    role = profile.role(role_name) if profile else None
    st = None
    if role is not None:
        if role.style:
            try:
                st = du.ensure_style(doc, role.style, base_name=base)
                if st is not None:
                    p.style = st
            except Exception:
                pass
        try:
            du.set_para(p, role.para, font_size_hint=role.font.size)
            if st is not None:
                du.set_font(st, role.font)
        except Exception:
            pass
    if text:
        run = p.add_run(text)
        if role is not None:
            try:
                du.set_font(run, role.font)
            except Exception:
                pass
    return p


def build_outline(src_template, dst=None, profile=None, keep_cover=False,
                  numbered=True, with_toc=True, with_abstract=True,
                  with_appendix=True):
    """按「通用论文大纲」生成结构稿：只到一级标题，不抄模板的小标题。

    结构：摘要 + 关键词 → 目录 → 1 引言 → 2 文献综述 → 3 研究方法 →
          4 结果与分析 → 5 讨论 → 6 结论与展望 → 致谢 → 附录 → 参考文献
    """
    from docx import Document

    doc = Document(src_template)
    body = doc.element.body
    for child in list(body):
        if child.tag == qn("w:sectPr"):
            continue
        body.remove(child)

    made = []

    if with_abstract:
        _add_para(doc, profile, "摘  要", "摘要标题")
        _add_para(doc, profile, "", "正文")
        _add_para(doc, profile, "关键词：", "关键词")

    if with_toc:
        _add_para(doc, profile, "目  录", "目录标题")
        toc_p = _add_para(doc, profile, "", "目录条目1")
        try:
            du.add_field(toc_p, 'TOC \\o "1-3" \\h \\z \\u', "（右键 → 更新域，生成目录）")
        except Exception:
            pass

    for i, name in enumerate(OUTLINE_SECTIONS, 1):
        _add_para(doc, profile, ("%d %s" % (i, name)) if numbered else name, "一级标题")
        _add_para(doc, profile, "", "正文")

    if with_appendix:
        _add_para(doc, profile, "致  谢", "致谢标题")
        _add_para(doc, profile, "", "正文")
        _add_para(doc, profile, "附  录", "附录标题")
        _add_para(doc, profile, "", "正文")

    _add_para(doc, profile, "参考文献", "参考文献标题")
    _add_para(doc, profile, "", "参考文献条目")

    out = dst or os.path.splitext(src_template)[0] + "_结构稿.docx"
    doc.save(out)
    return out, {"out": out, "sections": len(OUTLINE_SECTIONS), "toc": int(with_toc)}

# ---------------------------------------------------------------- 补齐缺失章节
END_SECTIONS = [("参考文献", "参考文献标题", "参考文献条目"),
                ("致  谢", "致谢标题", "正文"),
                ("附  录", "附录标题", "正文")]

HEAD_PATTERNS = {
    "参考文献": r"^\s*参\s*考\s*文\s*献",
    "致  谢": r"^\s*(致\s*谢|谢\s*辞|致\s*辞)",
    "附  录": r"^\s*附\s*录",
}


def ensure_sections(doc, profile):
    """把缺的「参考文献 / 致谢 / 附录」补在文末（内容留空），返回补了哪些。"""
    texts = [du.text_of(p).strip() for p in du.iter_paragraphs(doc)]
    joined = "\n".join(texts)
    added = []
    for title, title_role, body_role in END_SECTIONS:
        if re.search(HEAD_PATTERNS[title], joined, re.M):
            continue
        _add_para(doc, profile, title, title_role)
        _add_para(doc, profile, "", body_role)
        added.append(title.replace(" ", ""))
    return added


def new_formatted_doc(profile, dst, numbered=True, with_toc=True, with_appendix=True):
    """一步生成「排好格式的空白论文」：标准大纲 + 按方案排版。

    基准文档用方案里的 source_template（继承页面设置、页眉页脚、样式库）；
    找不到就用空白文档，样式由排版时按方案现建。
    """
    import tempfile

    from docx import Document

    from .apply import apply_profile

    base = getattr(profile, "source_template", None)
    if not base or not os.path.exists(base):
        base = None
        tmp_src = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        tmp_src.close()
        Document().save(tmp_src.name)
        base = tmp_src.name
        cleanup = True
    else:
        cleanup = False

    tmp_out = dst + ".骨架.docx"
    build_outline(base, tmp_out, profile, numbered=numbered, with_toc=with_toc,
                  with_appendix=with_appendix)
    try:
        report = apply_profile(tmp_out, dst, profile)
    finally:
        for f in (tmp_out, base if cleanup else None):
            try:
                if f and os.path.exists(f):
                    os.remove(f)
            except Exception:
                pass
    return dst, report
