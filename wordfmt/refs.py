# -*- coding: utf-8 -*-
"""参考文献表：从 RIS（EndNote 导出的那种）生成 GB/T 7714 条目并插进 Word。

这样不用在 EndNote 里折腾 Format Paper：EndNote 里选好文献导出 RIS，
到这边一键排成参考文献表，格式跟着方案里的「参考文献条目」样式走。
"""
import io
import re

from docx.oxml.ns import qn
from . import docxutil as du


def _read_text(path):
    """读文本：UTF-8 优先，失败就试 GBK / GB18030（知网导出的 RIS 常是 GBK）。"""
    data = open(path, "rb").read()
    for enc in ("utf-8-sig", "utf-8", "gb18030", "gbk"):
        try:
            text = data.decode(enc)
        except UnicodeDecodeError:
            continue
        if text.count("\ufffd") == 0:
            return text
    return data.decode("utf-8", errors="replace")


def parse_ris(path):
    """解析 RIS 文件，返回条目列表（每条是 dict）。"""
    raw = _read_text(path)
    entries, cur = [], {}
    for line in raw.replace("\r\n", "\n").split("\n"):
        if line.startswith("ER  -"):
            if cur:
                entries.append(cur)
            cur = {}
            continue
        m = re.match(r"^([A-Z][A-Z0-9])  - (.*)$", line)
        if not m:
            continue
        tag, val = m.group(1), m.group(2).strip()
        if tag in ("AU", "A1"):
            cur.setdefault("authors", []).append(val)
        elif tag == "TI" or tag == "T1":
            cur["title"] = val
        elif tag in ("JO", "JF", "T2", "JA"):
            cur.setdefault("journal", val)
        elif tag in ("PY", "Y1"):
            cur["year"] = val
        elif tag == "VL":
            cur["volume"] = val
        elif tag == "IS":
            cur["issue"] = val
        elif tag == "SP":
            cur["sp"] = val
        elif tag == "EP":
            cur["ep"] = val
        elif tag == "DO":
            cur["doi"] = val
        elif tag == "UR":
            cur.setdefault("url", val)
        elif tag == "TY":
            cur["type"] = val
    if cur:
        entries.append(cur)
    return entries


def _split_names(raw):
    """把一条作者串拆成多位作者；「姓, 名」这种英文写法不拆。"""
    raw = raw.strip()
    if not raw:
        return []
    if re.search(r"[，；;]", raw):
        return [x.strip() for x in re.split(r"[，；;]", raw) if x.strip()]
    if ", " in raw:
        parts = [x.strip() for x in raw.split(",") if x.strip()]
        if len(parts) == 2 and len(parts[0].split()) == 1:
            return [raw]          # 姓, 名 → 同一个人
        return parts
    return [raw]


def _authors_text(names):
    if not names:
        return ""
    flat = []
    for n in names:
        flat.extend(_split_names(n))
    names = flat
    if len(names) > 3:
        return ", ".join(names[:3]) + ", 等"
    return ", ".join(names)


def to_gbt(entry, number):
    """一条 GB/T 7714（顺序编码制）著录。"""
    kind = (entry.get("type") or "JOUR").upper()
    mark = "J" if kind in ("JOUR", "EJOUR", "MGZN") else "M" if kind == "BOOK" else "C" if kind == "CONF" else "J"
    authors = _authors_text(entry.get("authors"))
    title = (entry.get("title") or "").strip()
    journal = (entry.get("journal") or "").strip()
    year = (entry.get("year") or "").strip()
    vol = (entry.get("volume") or "").strip()
    issue = (entry.get("issue") or "").strip()
    pages = ""
    if entry.get("sp"):
        pages = entry["sp"]
        if entry.get("ep") and entry["ep"] != entry["sp"]:
            pages += "-" + entry["ep"]
    tail = year
    if vol:
        tail += ", %s" % vol
    if issue:
        tail += "(%s)" % issue
    if pages:
        tail += ": %s" % pages
    body = "%s. %s[%s]" % (authors, title, mark) if authors else "%s[%s]" % (title, mark)
    if journal:
        body += ". " + journal
    if tail.strip(", "):
        body += ", " + tail.strip(", ")
    body = body.rstrip(". ")
    doi = (entry.get("doi") or "").strip()
    if doi:
        body += ". DOI: %s" % doi
    text = "[%d] %s." % (number, body)
    return re.sub(r"\.{2,}", ".", text)


def ris_to_paper(entry):
    """RIS 字典 → 统一的 Paper 对象，方便套各种体例。"""
    from .litcite import Paper

    pages = ""
    if entry.get("sp"):
        pages = entry["sp"]
        if entry.get("ep") and entry["ep"] != entry["sp"]:
            pages += "-" + entry["ep"]
    return Paper(title=entry.get("title") or "", authors=entry.get("authors") or [],
                 year=entry.get("year"), journal=entry.get("journal") or "",
                 doi=entry.get("doi") or "", url=entry.get("url") or "",
                 volume=entry.get("volume") or "", issue=entry.get("issue") or "",
                 pages=pages)


def build_entries(ris_entries, style=None, template=None, include_doi=None):
    """按体例（或自定义模板）把 RIS 条目排成参考文献。"""
    from . import citestyle

    out = []
    for i, e in enumerate(ris_entries, 1):
        out.append(citestyle.format_entry(ris_to_paper(e), i, style, template, include_doi))
    return out


REF_HEAD_RE = re.compile(r"^\s*(参\s*考\s*文\s*献|references?)\s*$", re.I)
ENTRY_RE = re.compile(r"^\s*\[(\d+)\]")


def find_reference_heading(doc, profile=None):
    """找「参考文献」这个标题段落（先按文字，再按样式）。"""
    for p in du.iter_paragraphs(doc):
        if REF_HEAD_RE.match(du.text_of(p).strip()):
            return p
    if profile is not None:
        role = profile.role("参考文献标题")
        if role is not None and role.style:
            for p in du.iter_paragraphs(doc):
                if p.style is not None and p.style.name == role.style:
                    return p
    return None


def _is_stop_heading(para, profile):
    """参考文献区后面紧跟的「致谢 / 附录 / 一级标题」——到这儿就停。"""
    if profile is not None:
        for name in ("致谢标题", "附录标题", "一级标题", "目录标题"):
            role = profile.role(name)
            if role is not None and role.style and para.style is not None \
                    and para.style.name == role.style:
                return True
    text = du.text_of(para).strip()
    return bool(re.match(r"^\s*(致\s*谢|谢\s*辞|附\s*录)", text))


def _section_anchor(doc, profile):
    """返回 (参考文献标题段落 or None, 该区里最后一条条目段落 or None)。"""
    heading = find_reference_heading(doc, profile)
    if heading is None:
        return None, None
    last = None
    node = heading._p.getnext()
    parent = heading._parent
    from docx.oxml.ns import qn
    from docx.text.paragraph import Paragraph

    while node is not None and node.tag == qn("w:p"):
        para = Paragraph(node, parent)
        text = du.text_of(para).strip()
        if text and _is_stop_heading(para, profile):
            break
        if text:
            last = para
        node = node.getnext()
    return heading, last


def relocate_entries(doc, profile=None):
    """把落在「参考文献」区外面的条目（比如上次被塞到文末的）挪回标题下面。"""
    heading, last = _section_anchor(doc, profile)
    if heading is None:
        return 0

    # 先记下「参考文献」区里已有的段落
    in_zone = set()
    node = heading._p.getnext()
    while node is not None and node.tag == qn("w:p"):
        from docx.text.paragraph import Paragraph

        para = Paragraph(node, heading._parent)
        text = du.text_of(para).strip()
        if text and _is_stop_heading(para, profile):
            break
        in_zone.add(id(para._p))
        node = node.getnext()

    strays = [p for p in du.iter_paragraphs(doc)
              if id(p._p) not in in_zone and ENTRY_RE.match(du.text_of(p).strip())]
    if not strays:
        return 0

    entry_style = None
    if profile is not None:
        role = profile.role("参考文献条目")
        if role is not None and role.style:
            try:
                entry_style = du.ensure_style(doc, role.style, base_name="Normal")
            except Exception:
                entry_style = None

    anchor = last if last is not None else heading
    moved = 0
    for para in strays:
        el = para._p
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
        anchor._p.addnext(el)
        from docx.text.paragraph import Paragraph

        anchor = Paragraph(el, anchor._parent)
        if entry_style is not None:
            try:
                anchor.style = entry_style
            except Exception:
                pass
        moved += 1
    return moved


def insert_entries(doc, profile, texts, start_number=None, numbering=True):
    """把参考文献条目插到「参考文献」标题下面（不是文末）。"""
    from docx.text.paragraph import Paragraph

    try:
        relocate_entries(doc, profile)
    except Exception:
        pass
    heading, last = _section_anchor(doc, profile)
    role = profile.role("参考文献条目") if profile is not None else None
    entry_style = None
    if role is not None and role.style:
        try:
            entry_style = du.ensure_style(doc, role.style, base_name="Normal")
        except Exception:
            entry_style = None

    # 编号接在已有条目的最大号后面
    if start_number is None:
        start = 1
        for para in (p for p in [last] if p is not None):
            m = ENTRY_RE.match(du.text_of(para).strip())
            if m:
                start = int(m.group(1)) + 1
        for p in du.iter_paragraphs(doc):
            m = ENTRY_RE.match(du.text_of(p).strip())
            if m:
                start = max(start, int(m.group(1)) + 1)
        start_number = start

    anchor = last if last is not None else heading
    made = []
    for i, text in enumerate(texts):
        if numbering:
            body = re.sub(r"^\s*\[\d+\]\s*", "", text)
            text = "[%d] %s" % (start_number + i, body)
        try:
            if entry_style is not None:
                new_p = doc.add_paragraph(text, style=entry_style)
            else:
                new_p = doc.add_paragraph(text)
        except Exception:
            new_p = doc.add_paragraph(text)
        if anchor is not None:
            anchor._p.addnext(new_p._p)
            anchor = new_p
        made.append(new_p)
    return made


def append_reference_list(doc, profile, texts, start_number=None, numbering=True):
    """把条目接到「参考文献」小节的最后一条后面，沿用那条的样式。"""
    import copy as _copy

    from docx.text.paragraph import Paragraph

    if find_reference_heading(doc, profile) is not None:
        return insert_entries(doc, profile, texts, start_number, numbering)

    ref_role = profile.role("参考文献条目") if profile else None
    anchors = [p for p in du.iter_paragraphs(doc)
               if ref_role is not None and p.style is not None and p.style.name == ref_role.style
               and du.text_of(p).strip()]
    anchor = anchors[-1] if anchors else None
    if start_number is None:
        start = 1
        for t in texts:
            m = re.match(r"\s*\[(\d+)\]", t)
            if m:
                start = max(start, int(m.group(1)) + 1)
        start_number = start
    made = []
    for i, text in enumerate(texts):
        if numbering:
            body = re.sub(r"^\s*\[\d+\]\s*", "", text)   # 去掉原来的编号，统一重编
            text = "[%d] %s" % (start_number + i, body)
        if anchor is not None:
            el = _copy.deepcopy(anchor._p)
            for child in list(el):
                if child.tag != "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr":
                    el.remove(child)
            anchor._p.addnext(el)
            new_p = Paragraph(el, anchor._parent)
            new_p.add_run(text)
            anchor = new_p
        else:
            new_p = doc.add_paragraph(text)
            anchor = new_p
        made.append(new_p)
    return made


def papers_to_texts(papers):
    """把检索到的 Paper 列表排成 GB/T 7714 条目（不经过 RIS/EndNote）。"""
    from . import litcite

    return [litcite.gbt_entry(i, p) for i, p in enumerate(papers, 1)]


def from_papers_to_doc(papers, docx_path, profile, out_path=None, style=None, template=None,
                       include_doi=None):
    """检索结果 → 参考文献表 → 插进 docx，全程不碰 EndNote。"""
    from docx import Document

    if not papers:
        raise ValueError("没有可用文献")
    doc = Document(docx_path)
    start = 1
    for p in doc.paragraphs:
        m = re.match(r"\s*\[(\d+)\]", p.text or "")
        if m:
            start = max(start, int(m.group(1)) + 1)
    from . import citestyle

    texts = [citestyle.format_entry(p, i, style, template, include_doi)
             for i, p in enumerate(papers, 1)]
    from . import citestyle as _cs

    tpl = template or (_cs.STYLES.get(style or "") or {}).get("template", "")
    made = append_reference_list(doc, profile, texts, start_number=start,
                                 numbering=("{seq}" in (tpl or "")))
    out = out_path or docx_path
    doc.save(out)
    return {"count": len(made), "out": out, "entries": texts}


def from_ris_to_doc(ris_path, docx_path, profile, out_path=None, style=None, template=None,
                    include_doi=None):
    """一站式：RIS → 参考文献表 → 插进 docx。"""
    from docx import Document

    entries = parse_ris(ris_path)
    if not entries:
        raise ValueError("RIS 里没解析到条目：%s" % ris_path)
    doc = Document(docx_path)
    texts = build_entries(entries, style, template, include_doi)
    # 文档里已经有序号时接着往后编
    start = 1
    for p in doc.paragraphs:
        m = re.match(r"\s*\[(\d+)\]", p.text or "")
        if m:
            start = max(start, int(m.group(1)) + 1)
    from . import citestyle as _cs

    tpl = template or (_cs.STYLES.get(style or "") or {}).get("template", "")
    made = append_reference_list(doc, profile, texts, start_number=start,
                                 numbering=("{seq}" in (tpl or "")))
    out = out_path or docx_path
    doc.save(out)
    return {"count": len(made), "out": out, "entries": texts}
