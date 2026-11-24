# -*- coding: utf-8 -*-
"""封面功能自测。

1) 提取封面入库；
2) 造「格式乱 + 换人」的稿子 → 用封面库【重建】→ 检查文字是否填对；
3) 造「只把格式搞乱、内容不换」的稿子 → 用封面库【只改格式】→ 与原封面逐像素比对。
"""
import os
import random
import sys

sys.path.insert(0, r"D:\deepseek工作区\_pylib")
sys.path.insert(0, r"D:\deepseek工作区\Word格式工具")

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from wordfmt import cover, docxutil as du

ROOT = r"D:\deepseek工作区\Word格式工具"
SRC = os.path.join(ROOT, "示例", "论文样板_原稿.docx")
OUTDIR = os.path.join(ROOT, "输出", "_封面测试")
os.makedirs(OUTDIR, exist_ok=True)


def mess_formats(doc, rnd, retitle=False, rename=False):
    start, end = cover.cover_span(doc)
    for p in list(du.iter_paragraphs(doc))[start:end + 1]:
        t = du.text_of(p)
        if rename and "管贺文" in t:
            for run in p.runs:
                if "管贺文" in run.text:
                    run.text = run.text.replace("管贺文", "李小明")
        if rename and "20230812117" in t:
            for run in p.runs:
                if "20230812117" in run.text:
                    run.text = run.text.replace("20230812117", "20239999001")
        if retitle and "草莓果汁发酵前后感官品质差异及综合评价研究" in t:
            for run in p.runs:
                if "草莓果汁" in run.text:
                    run.text = run.text.replace(
                        "草莓果汁发酵前后感官品质差异及综合评价研究",
                        "基于深度学习的柑橘品质无损检测研究")
        pPr = p._p.get_or_add_pPr()
        for tag in ("spacing", "jc", "ind"):
            for el in pPr.findall(qn("w:" + tag)):
                pPr.remove(el)
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:line"), rnd.choice(["240", "480", "600"]))
        sp.set(qn("w:lineRule"), "auto")
        pPr.append(sp)
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), rnd.choice(["left", "right", "both"]))
        pPr.append(jc)
        for run in p.runs:
            if du.is_field_run(run):
                continue
            rPr = run._r.get_or_add_rPr()
            rf = rPr.find(qn("w:rFonts"))
            if rf is None:
                rf = OxmlElement("w:rFonts")
                rPr.insert(0, rf)
            rf.set(qn("w:eastAsia"), rnd.choice(["黑体", "楷体", "宋体"]))
            sz = rPr.find(qn("w:sz"))
            if sz is None:
                sz = OxmlElement("w:sz")
                rPr.append(sz)
            sz.set(qn("w:val"), str(int(rnd.choice([9, 12, 14, 20]) * 2)))


def main():
    # 1) 提取并入库
    doc = Document(SRC)
    cov = cover.extract_cover(doc, "南昌师范学院本科毕业论文封面", "论文样板.docx")
    lib = cover.save_to_library(cov, OUTDIR)
    print("入库：", lib)
    print("封面摘要：", cov.summary())
    print("识别到的值：", cover.extract_values(doc))

    # 2) 换人 + 乱格式 → 重建
    messy = Document(SRC)
    mess_formats(messy, random.Random(3), retitle=True, rename=True)
    p2 = os.path.join(OUTDIR, "封面格式乱_换人.docx")
    messy.save(p2)
    d2 = Document(p2)
    rep = cover.apply_cover(d2, cover.Cover.load(lib), mode="rebuild")
    out2 = os.path.join(OUTDIR, "封面_重建.docx")
    d2.save(out2)
    print("\n[重建]", rep)
    texts = [du.text_of(p) for p in list(du.iter_paragraphs(Document(out2)))[:22]]
    for t in texts:
        if any(k in t for k in ("李小明", "20239999001", "柑橘", "管贺文", "草莓果汁发酵")):
            print("   文字：", t.strip()[:70])

    # 3) 只乱格式 → 只改格式
    messy2 = Document(SRC)
    mess_formats(messy2, random.Random(11))
    p3 = os.path.join(OUTDIR, "封面格式乱_不换人.docx")
    messy2.save(p3)
    d3 = Document(p3)
    rep3 = cover.apply_cover(d3, cover.Cover.load(lib), mode="format")
    out3 = os.path.join(OUTDIR, "封面_只改格式.docx")
    d3.save(out3)
    print("\n[只改格式]", rep3)
    return out2, out3


if __name__ == "__main__":
    main()
