# -*- coding: utf-8 -*-
"""提取方案后的「回头比对」：拿新方案把模板自己重排一遍，逐页和原模板比像素差异。

目的：提取完不要直接交付，先证明“用这个方案重排模板，能还原到原样”。
差异大于阈值（默认 3%）就提示检查，并列出差异最大的几页。
"""
import os
import shutil
import tempfile


def _export_pdf(docx_path, pdf_path):
    from . import wordio

    tmp = docx_path + ".rt.docx"
    shutil.copyfile(docx_path, tmp)
    try:
        ok, msg = wordio.update_fields(tmp, pdf_path=pdf_path)
    finally:
        try:
            os.remove(tmp)
        except Exception:
            pass
    return ok, msg


def _render(pdf_path, folder, tag, dpi=110):
    import pymupdf

    doc = pymupdf.open(pdf_path)
    paths = []
    for i in range(doc.page_count):
        p = os.path.join(folder, "%s_p%02d.png" % (tag, i + 1))
        doc[i].get_pixmap(dpi=dpi).save(p)
        paths.append(p)
    doc.close()
    return paths


def _diff_pct(a_png, b_png):
    import numpy as np
    from PIL import Image, ImageChops

    a = Image.open(a_png).convert("RGB")
    b = Image.open(b_png).convert("RGB")
    if a.size != b.size:
        return None
    d = np.asarray(ImageChops.difference(a, b))
    return float((d.max(axis=2) > 40).mean() * 100.0)


def roundtrip_check(template_path, profile, profile_path=None, workdir=None,
                    threshold=3.0):
    """用方案重排模板本身，与原模板逐页比对。

    返回 {"ok": bool, "pages": (原, 新), "mean": 平均差异%, "max": 最大差异%,
          "worst": [(页号, 差异%)...], "message": 说明}
    """
    from .apply import apply_profile

    out_dir = workdir or tempfile.mkdtemp(prefix="wordfmt_rt_")
    os.makedirs(out_dir, exist_ok=True)

    rebuilt = os.path.join(out_dir, "rebuilt.docx")
    try:
        apply_profile(template_path, rebuilt, profile)
    except Exception as e:
        return {"ok": False, "pages": (0, 0), "mean": None, "max": None, "worst": [],
                "message": "重排失败：%s" % e}

    pdf_a = os.path.join(out_dir, "base.pdf")
    pdf_b = os.path.join(out_dir, "rebuilt.pdf")
    ok_a, msg_a = _export_pdf(template_path, pdf_a)
    ok_b, msg_b = _export_pdf(rebuilt, pdf_b)
    if not (ok_a and ok_b):
        return {"ok": False, "pages": (0, 0), "mean": None, "max": None, "worst": [],
                "message": "Word 导出 PDF 失败（%s）" % (msg_a or msg_b).strip()[:120]}

    a = _render(pdf_a, out_dir, "base")
    b = _render(pdf_b, out_dir, "rebuilt")
    n = min(len(a), len(b))
    diffs = []
    for i in range(n):
        d = _diff_pct(a[i], b[i])
        if d is not None:
            diffs.append((i + 1, d))
    if not diffs:
        return {"ok": False, "pages": (len(a), len(b)), "mean": None, "max": None,
                "worst": [], "message": "页面尺寸不一致，无法比对"}

    mean = sum(d for _p, d in diffs) / len(diffs)
    worst = sorted(diffs, key=lambda x: -x[1])[:3]
    pages_same = len(a) == len(b)
    ok = pages_same and mean <= threshold
    if pages_same:
        msg = "页数一致（%d 页），逐页差异 平均 %.2f%%，最大 %.2f%%" % (
            len(a), mean, max(d for _p, d in diffs))
    else:
        msg = "页数不一致（原 %d 页，重排 %d 页）" % (len(a), len(b))
    if not ok:
        msg += "；差异最大的页：" + "、".join("第%d页 %.1f%%" % w for w in worst)
    return {"ok": ok, "pages": (len(a), len(b)), "mean": mean,
            "max": max(d for _p, d in diffs), "worst": worst, "message": msg}
