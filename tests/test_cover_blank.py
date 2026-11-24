# -*- coding: utf-8 -*-
"""验证「目标文档缺信息时字段留空」。"""
import sys

sys.path.insert(0, r"D:\deepseek工作区\_pylib")
sys.path.insert(0, r"D:\deepseek工作区\Word格式工具")
from docx import Document

from wordfmt import cover, docxutil as du

SRC = r"D:\deepseek工作区\Word格式工具\示例\论文样板_原稿.docx"
LIB = r"D:\deepseek工作区\Word格式工具\covers\南昌师范学院本科毕业论文_封面.yaml"

doc = Document(SRC)
start, end = cover.cover_span(doc)
for p in list(du.iter_paragraphs(doc))[start:end + 1]:
    for run in p.runs:
        if "20230812117" in run.text:
            run.text = run.text.replace("20230812117", "")
        if "管贺文" in run.text:
            run.text = run.text.replace("管贺文", "")
doc.save(r"D:\deepseek工作区\Word格式工具\输出\_封面测试\缺学号姓名.docx")

d2 = Document(r"D:\deepseek工作区\Word格式工具\输出\_封面测试\缺学号姓名.docx")
rep = cover.apply_cover(d2, cover.Cover.load(LIB), mode="rebuild")
print("结果：", rep)
d2.save(r"D:\deepseek工作区\Word格式工具\输出\_封面测试\缺信息_重建.docx")

for t in [du.text_of(p) for p in list(du.iter_paragraphs(Document(
        r"D:\deepseek工作区\Word格式工具\输出\_封面测试\缺信息_重建.docx")))[:22]]:
    if any(k in t for k in ("姓名", "学号", "指导教师", "职称")):
        print("  ", repr(t))
