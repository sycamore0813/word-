# -*- coding: utf-8 -*-
"""造一份「格式全乱」的测试稿：文字不动，只把字体/字号/行距/缩进/对齐搞乱，
再故意把一部分标题的样式改成正文，用来验证「识别 + 重排」是否真的能救回来。

用法：
    python tests/make_messy.py 原稿.docx 输出.docx [--seed 7]
"""
from __future__ import annotations

import argparse
import os
import random
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT = os.path.dirname(HERE)
WORKSPACE = os.path.dirname(PROJECT)
for p in (PROJECT, os.path.join(WORKSPACE, "_pylib")):
    if p not in sys.path:
        sys.path.insert(0, p)

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from wordfmt import docxutil as du  # noqa: E402

FONTS = ["宋体", "黑体", "仿宋", "楷体", "Calibri", "Arial"]
SIZES = [9, 10.5, 12, 14, 16, 18, 22]
LINES = ["240", "360", "480", "600"]
ALIGNS = ["left", "center", "right", "both"]


def mess_run(run, rnd: random.Random) -> None:
    if du.is_field_run(run):
        return
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        if rFonts.get(qn("w:" + attr)) is not None:
            del rFonts.attrib[qn("w:" + attr)]
    rFonts.set(qn("w:eastAsia"), rnd.choice(FONTS))
    rFonts.set(qn("w:ascii"), rnd.choice(FONTS))
    rFonts.set(qn("w:hAnsi"), rFonts.get(qn("w:ascii")))
    for tag in ("sz", "szCs"):
        el = rPr.find(qn("w:" + tag))
        if el is None:
            el = OxmlElement("w:" + tag)
            rPr.append(el)
        el.set(qn("w:val"), str(int(rnd.choice(SIZES) * 2)))
    if rnd.random() < 0.3:
        el = rPr.find(qn("w:color"))
        if el is None:
            el = OxmlElement("w:color")
            rPr.append(el)
        el.set(qn("w:val"), rnd.choice(["FF0000", "0000FF", "00AA00", "808080"]))


def mess_para(p, rnd: random.Random) -> None:
    pPr = p._p.get_or_add_pPr()
    for tag in ("spacing", "ind", "jc"):
        for el in pPr.findall(qn("w:" + tag)):
            pPr.remove(el)
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:line"), rnd.choice(LINES))
    sp.set(qn("w:lineRule"), "auto")
    sp.set(qn("w:before"), str(rnd.choice([0, 120, 240, 480])))
    sp.set(qn("w:after"), str(rnd.choice([0, 120, 240, 480])))
    pPr.append(sp)
    ind = OxmlElement("w:ind")
    ind.set(qn("w:firstLineChars"), str(rnd.choice([0, 100, 200, 400])))
    ind.set(qn("w:firstLine"), str(rnd.choice([0, 200, 420, 840])))
    pPr.append(ind)
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), rnd.choice(ALIGNS))
    pPr.append(jc)
    for run in p.runs:
        mess_run(run, rnd)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("--seed", type=int, default=7)
    ap.add_argument("--strip-style-ratio", type=float, default=0.4,
                    help="把多大比例的标题段落改成正文样式")
    args = ap.parse_args()

    rnd = random.Random(args.seed)
    doc = Document(args.src)
    paras = list(du.iter_paragraphs(doc))
    headings = 0
    for p in paras:
        text = du.text_of(p).strip()
        if not text:
            continue
        style = du.p_style_name(doc, p)
        is_heading = style.lower().startswith("heading") or "标题" in style
        # 故意把一部分标题的样式改成正文，考验内容规则
        if is_heading and rnd.random() < args.strip_style_ratio:
            try:
                p.style = doc.styles["Normal"]
                headings += 1
            except Exception:
                pass
        mess_para(p, rnd)
    doc.save(args.dst)
    print("已生成格式错乱测试稿：%s（打乱段落 %d，其中 %d 个标题被降级为正文）" %
          (args.dst, len(paras), headings))
    return 0


if __name__ == "__main__":
    sys.exit(main())
