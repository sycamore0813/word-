# -*- coding: utf-8 -*-
"""清洗文档里的隐形字符 / 特殊占位符。

Word 里从网页、PDF、知网复制来的文字常带这些东西：
  U+200B 零宽空格、U+200C/200D 连接符、U+FEFF BOM、U+2060 word joiner、
  U+00AD 软连字符、U+2028/2029 行分隔符、U+180E 蒙古文元音分隔符、
  U+200E/200F 方向标记、U+00A0 不间断空格（按需换成普通空格）。

它们肉眼看不见，却会让排版出现“莫名多一个空”“行距突然变大”“替换格式后仍留占位”等问题。
"""
from . import docxutil as du

ZERO_WIDTH = {"\u200b", "\u200c", "\u200d", "\ufeff", "\u2060",
              "\u180e", "\u200e", "\u200f"}
LINE_SEP = {"\u2028", "\u2029"}
SOFT_HYPHEN = {"\u00ad"}
NBSP = "\u00a0"

ALL_BAD = ZERO_WIDTH | LINE_SEP | SOFT_HYPHEN


WIDE_SPACE = "\u3000"          # 全角空格


def clean_text(text, drop_nbsp=False, wide_space_as_spaces=True):
    """返回 (清洗后的文字, 去掉的字符数)。

    wide_space_as_spaces：把全角空格 U+3000 换成两个普通空格。
    为什么要换：U+3000 属于「中西文两可」字符，Word 会拿样式里的西文字体去排，
    Times New Roman 没有这个字形，标题里就会出现小方框 □。
    两个普通空格在中文里正好约等于一个全角宽，看上去一样。
    """
    if not text:
        return text, 0
    removed = 0
    out = []
    for ch in text:
        if ch in ALL_BAD:
            removed += 1
            continue
        if ch == NBSP:
            out.append(" ")
            removed += 1
            continue
        if ch == WIDE_SPACE and wide_space_as_spaces:
            out.append("  ")
            removed += 1
            continue
        out.append(ch)
    return "".join(out), removed


def clean_document(doc, drop_nbsp=True, wide_space_as_spaces=True):
    """清掉全文（含表格、页眉页脚）里的隐形字符，返回统计。"""
    stats = {"chars": 0, "runs": 0}
    targets = []

    for p in du.iter_paragraphs(doc):
        targets.extend(p.runs)
    for t in getattr(doc, "tables", []) or []:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    targets.extend(p.runs)
    for sec in doc.sections:
        for part in (sec.header, sec.footer):
            try:
                for p in part.paragraphs:
                    targets.extend(p.runs)
            except Exception:
                pass

    for run in targets:
        try:
            if du.is_field_run(run):
                continue
        except Exception:
            pass
        text = run.text or ""
        if not text:
            continue
        if not any((ch in ALL_BAD) or (ch == NBSP and drop_nbsp)
                   or (ch == WIDE_SPACE and wide_space_as_spaces) for ch in text):
            continue
        new_text, removed = clean_text(text, drop_nbsp=drop_nbsp,
                                       wide_space_as_spaces=wide_space_as_spaces)
        if removed:
            run.text = new_text
            stats["chars"] += removed
            stats["runs"] += 1
    return stats
