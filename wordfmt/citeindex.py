# -*- coding: utf-8 -*-
"""引用编号自动排序：把正文里的 [n] 按出现顺序重排，参考文献表跟着重排。

思路（不依赖 Word 域，稳定可预测）：
  1. 按文档顺序扫正文，记下每个编号第一次出现的位置；
  2. 第一次出现的序号就是新编号（1,2,3…），同一编号后面再出现也换成新号；
  3. 参考文献条目按新编号重排、重写条目开头的 [n]；
  4. 正文里没引用到的条目排在最后，编号继续往后接（不会丢）。
"""
import re

from . import docxutil as du

MARK = re.compile(r"\[(\d+(?:\s*[,，]\s*\d+)*)\]")


def _para_text_runs(p):
    return [(r, r.text or "") for r in p.runs]


def _rewrite_marks(p, mapping):
    """把段落里整段落在单个 run 内的 [n] 换掉；返回改了几处。"""
    changed = 0
    for run, text in _para_text_runs(p):
        if not text:
            continue
        new_text = text
        for m in reversed(list(MARK.finditer(text))):
            old_nums = [int(x) for x in re.split(r"[,，]", m.group(1))]
            new_nums = [mapping.get(n, n) for n in old_nums]
            new_nums = sorted(set(new_nums))
            if new_nums != old_nums:
                new_text = new_text[:m.start()] + "[%s]" % ",".join(str(x) for x in new_nums) + new_text[m.end():]
        if new_text != text:
            run.text = new_text
            changed += 1
    return changed


def _scan_body(doc):
    """按顺序扫正文，返回 [(编号, 段落)]（不含参考文献表里的条目）。"""
    found = []
    for p in du.iter_paragraphs(doc):
        text = du.text_of(p)
        if not text.strip():
            continue
        if re.match(r"^\s*\[\d+\]", text.strip()) and p.style is not None \
                and p.style.name and "参考文献" in (p.style.name or ""):
            continue          # 参考文献表里的条目不算正文引用
        for m in MARK.finditer(text):
            for x in re.split(r"[,，]", m.group(1)):
                try:
                    found.append(int(x))
                except ValueError:
                    pass
    return found


def _ref_paragraphs(doc):
    """找出参考文献表的条目段落（以 [n] 开头、且样式名含「参考文献」）。"""
    out = []
    for p in du.iter_paragraphs(doc):
        t = du.text_of(p).strip()
        m = re.match(r"^\[(\d+)\]\s*(.*)$", t, re.S)
        if not m:
            continue
        style_name = (p.style.name if p.style is not None else "") or ""
        if "参考文献" in style_name or "参考" in style_name:
            out.append((int(m.group(1)), p, m.group(2)))
    return out


def _set_entry_number(p, new_no):
    """把条目开头的 [n] 换成新编号，正文其余部分不动。"""
    text = du.text_of(p)
    body = re.sub(r"^\s*\[\d+\]\s*", "", text)
    new_text = "[%d] %s" % (new_no, body)
    if new_text.strip() == text.strip():
        return
    font = None
    if p.runs:
        try:
            font = p.runs[0].font
        except Exception:
            font = None
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    run = p.add_run(new_text)
    if font is not None:
        try:
            run.font.size = font.size
            run.font.name = font.name
            run.font.bold = font.bold
        except Exception:
            pass


def renumber(doc, dry_run=False):
    """按正文出现顺序重排引用编号与参考文献表。

    返回 {map: 旧→新, cited: [旧号...], moved: 条目数, marks: 改了正文几处}
    """
    try:
        from . import refs as _refs

        _refs.relocate_entries(doc)
    except Exception:
        pass

    body_nums = _scan_body(doc)
    order = []
    for n in body_nums:
        if n not in order:
            order.append(n)
    mapping = {old: i + 1 for i, old in enumerate(order)}

    refs = _ref_paragraphs(doc)
    uncited = [r for r in refs if r[0] not in mapping]
    next_no = len(order) + 1
    for old, _p, _t in uncited:
        mapping[old] = next_no
        next_no += 1

    changed = 0
    for p in du.iter_paragraphs(doc):
        changed += _rewrite_marks(p, mapping)

    # 参考文献条目：按新编号排序，并重写条目开头的 [n]
    entries = sorted(refs, key=lambda x: mapping.get(x[0], 9999))
    moved = 0
    if entries and not dry_run:
        ref_ids = {id(p._p) for _o, p, _b in refs}
        anchor = entries[0][1]._p.getprevious()
        while anchor is not None and id(anchor) in ref_ids:
            anchor = anchor.getprevious()     # 别把锚点定在会被删掉的条目上
        body_el = entries[0][1]._p.getparent()
        for _old, p, _b in entries:
            body_el.remove(p._p)
        cur = anchor
        for old, p, _b in entries:
            new_no = mapping.get(old, old)
            _set_entry_number(p, new_no)
            if cur is None:
                body_el.insert(0, p._p)
                cur = p._p
            else:
                cur.addnext(p._p)
                cur = p._p
            moved += 1

    return {"map": mapping, "cited": order, "uncited": [r[0] for r in uncited],
            "marks": changed, "moved": moved}


def check(doc):
    """检查现有编号是否已按出现顺序（返回一串问题描述，空表示没问题）。"""
    body = _scan_body(doc)
    order, problems = [], []
    for i, n in enumerate(body):
        if n not in order:
            order.append(n)
            if n != len(order):
                problems.append("第 %d 处引用是 [%d]，按出现顺序应该是 [%d]" % (i + 1, n, len(order)))
    refs = {n for n, _p, _t in _ref_paragraphs(doc)}
    for n in order:
        if n not in refs:
            problems.append("正文引用了 [%d]，但参考文献表里没有这一条" % n)
    return problems
