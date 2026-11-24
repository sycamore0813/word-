# -*- coding: utf-8 -*-
"""验证：封面库自带校徽，能套到「封面没图」的文档和「完全没有封面」的新文档上。"""
import os
import sys

sys.path.insert(0, r"D:\deepseek工作区\_pylib")
sys.path.insert(0, r"D:\deepseek工作区\Word格式工具")
from docx import Document
from docx.oxml.ns import qn

from wordfmt import cover, docxutil as du

ROOT = r"D:\deepseek工作区\Word格式工具"
LIB = os.path.join(ROOT, "covers", "南昌师范学院本科毕业论文_封面.yaml")
OUT = os.path.join(ROOT, "输出", "_封面测试")
os.makedirs(OUT, exist_ok=True)


def strip_logo(path, dst):
    d = Document(path)
    span = cover.cover_span(d)
    for p in list(du.iter_paragraphs(d))[span[0]:span[1] + 1]:
        for dr in p._p.findall(".//" + qn("w:drawing")):
            dr.getparent().remove(dr)
    d.save(dst)
    print("已去掉校徽：", dst)


def main():
    # 1) 有封面区但没图的文档
    no_logo = os.path.join(OUT, "封面但无校徽.docx")
    strip_logo(os.path.join(ROOT, "示例", "论文样板_原稿.docx"), no_logo)
    d = Document(no_logo)
    rep = cover.apply_cover(d, cover.Cover.load(LIB), mode="blank")
    out1 = os.path.join(OUT, "无图文档_套封面.docx")
    d.save(out1)
    print("[有封面区/无图]", rep)

    # 2) 完全没有封面的新文档
    fresh = Document()
    for line in ("这是正文第一段。", "第二段。"):
        fresh.add_paragraph(line)
    rep2 = cover.apply_cover(fresh, cover.Cover.load(LIB), mode="blank")
    out2 = os.path.join(OUT, "新文档_插封面.docx")
    fresh.save(out2)
    print("[无封面区]", rep2)

    # 检查两张结果的封面里都有图片
    for path in (out1, out2):
        doc = Document(path)
        n = len(doc.element.body.findall(".//" + qn("w:drawing")))
        print(os.path.basename(path), "图片数：", n)


if __name__ == "__main__":
    main()
