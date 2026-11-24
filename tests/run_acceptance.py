# -*- coding: utf-8 -*-
"""一键验收：提取方案 → 排版原稿 → 排版「格式错乱」测试稿 → 逐段校验 → 渲染比对。

用法：
    python tests/run_acceptance.py 示例\论文样板_原稿.docx
需要本机装有 Word（用于更新域、导出 PDF）；没有 Word 时只做结构与格式校验。
"""
from __future__ import annotations

import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT = os.path.dirname(HERE)
WORKSPACE = os.path.dirname(PROJECT)
for p in (PROJECT, os.path.join(WORKSPACE, "_pylib")):
    if p not in sys.path:
        sys.path.insert(0, p)

import make_messy  # noqa: E402  (同目录脚本)
from wordfmt import wordio  # noqa: E402
from wordfmt.apply import apply_profile  # noqa: E402
from wordfmt.extract import build_profile  # noqa: E402
from wordfmt.verify import verify  # noqa: E402

OUT = os.path.join(PROJECT, "输出", "_验收")


def _export_pdf(docx: str) -> str:
    pdf = os.path.splitext(docx)[0] + ".pdf"
    tmp = docx + ".wordtmp.docx"
    open(tmp, "wb").write(open(docx, "rb").read())
    ok, msg = wordio.update_fields(tmp, pdf_path=pdf)
    os.remove(tmp)
    if not ok:
        print("  ! Word 导出失败：%s" % msg.strip().splitlines()[-1][:120])
        return ""
    return pdf


def _render(pdf: str, tag: str, dpi: int = 110):
    import pymupdf
    doc = pymupdf.open(pdf)
    paths = []
    for i in range(doc.page_count):
        p = os.path.join(OUT, "%s_p%02d.png" % (tag, i + 1))
        doc[i].get_pixmap(dpi=dpi).save(p)
        paths.append(p)
    doc.close()
    return paths


def _diff(a_png, b_png):
    import numpy as np
    from PIL import Image, ImageChops
    a = Image.open(a_png).convert("RGB")
    b = Image.open(b_png).convert("RGB")
    if a.size != b.size:
        return None
    d = np.asarray(ImageChops.difference(a, b))
    return float((d.max(axis=2) > 40).mean() * 100.0)


def main() -> int:
    src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        PROJECT, "示例", "论文样板_原稿.docx")
    os.makedirs(OUT, exist_ok=True)
    lines = []

    def say(s):
        print(s)
        lines.append(s)

    say("源文件：%s" % src)
    prof = build_profile(src)
    prof_path = os.path.join(OUT, "验收方案.yaml")
    prof.save(prof_path)
    say("提取方案：%d 个部分，页面 %.2f×%.2fcm，页眉「%s」" %
        (len(prof.roles), prof.page.width_cm, prof.page.height_cm, prof.page.header_text))

    # 1) 原稿重排
    out1 = os.path.join(OUT, "重排_原稿.docx")
    r1 = apply_profile(src, out1, prof)
    v1 = verify(out1, prof)
    say("\n[1] 原稿重排：%s（改写 %d 段）" % (r1.summary().splitlines()[1], r1.styled))
    say("    校验：%d/%d 通过（%.1f%%），页面问题 %d 条" %
        (v1.ok, v1.checked, v1.pass_rate, len(v1.page_problems)))

    # 2) 格式错乱稿重排
    messy = os.path.join(OUT, "格式错乱.docx")
    sys.argv_backup = list(sys.argv)
    sys.argv = ["make_messy.py", src, messy, "--seed", "7"]
    make_messy.main()
    sys.argv = sys.argv_backup
    out2 = os.path.join(OUT, "重排_格式错乱.docx")
    r2 = apply_profile(messy, out2, prof)
    v2 = verify(out2, prof)
    say("\n[2] 格式错乱稿重排：改写 %d 段，校验 %d/%d 通过（%.1f%%）" %
        (r2.styled, v2.ok, v2.checked, v2.pass_rate))
    if v2.issues:
        kinds = v2.by_kind()
        say("    未通过项：" + "、".join("%s×%d" % (k, n) for k, n in kinds.most_common(5)))

    # 3) 渲染比对
    if wordio.word_installed():
        say("\n[3] 渲染比对（与原始模板逐页对比）")
        pdf0 = _export_pdf(src)
        pdf1 = _export_pdf(out1)
        pdf2 = _export_pdf(out2)
        if pdf0 and pdf1 and pdf2:
            a = _render(pdf0, "基准")
            b = _render(pdf1, "重排")
            c = _render(pdf2, "混乱后重排")
            if len(a) != len(b) or len(a) != len(c):
                say("    页面数不同：基准 %d、重排 %d、混乱后 %d" % (len(a), len(b), len(c)))
            for label, pages in (("原稿重排", b), ("混乱稿重排", c)):
                diffs = [d for d in (_diff(a[i], pages[i]) for i in range(min(len(a), len(pages))))
                         if d is not None]
                if diffs:
                    body = [d for d in diffs[2:]]  # 前两页是封面
                    say("    %s：逐页差异 平均 %.2f%%，最大 %.2f%%；正文页平均 %.2f%%" %
                        (label, sum(diffs) / len(diffs), max(diffs),
                         (sum(body) / len(body)) if body else 0))
    else:
        say("\n[3] 未检测到 Word，跳过渲染比对")

    say("\n验收结论：%s" % ("通过" if (v1.pass_rate >= 99 and v2.pass_rate >= 99) else "需检查"))
    report = os.path.join(PROJECT, "输出", "验收报告.md")
    with open(report, "w", encoding="utf-8") as f:
        f.write("# 验收报告\n\n" + "\n".join(lines) + "\n")
    print("\n报告已写入：%s" % report)
    return 0


if __name__ == "__main__":
    sys.exit(main())
