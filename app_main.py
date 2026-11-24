# -*- coding: utf-8 -*-
"""打包入口：无参数时开图形界面；带 --selftest 时跑一遍自检（用于验证 exe 是否完整）。

PyInstaller 用这个文件做入口：
    pyinstaller --onefile --windowed --name Word格式排版助手 app_main.py
"""
from __future__ import annotations

import os
import sys
import traceback

BASE = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False)
                                      else __file__))


def selftest() -> int:
    log_path = os.path.join(BASE, "selftest.log")
    lines = []
    ok = True
    try:
        import tkinter  # noqa: F401
        lines.append("tkinter OK")
        import tkinter as tk
        from tkinter import ttk
        root = tk.Tk()
        root.withdraw()
        ttk.Label(root, text="ok").pack()
        root.update()
        root.destroy()
        lines.append("Tk 窗口创建 OK")
        import docx
        lines.append("python-docx OK")
        import yaml
        lines.append("PyYAML OK")

        from wordfmt import paths
        from wordfmt.apply import apply_profile
        from wordfmt.profile import Profile, Page, Role, Font, Para
        from wordfmt.verify import verify
        lines.append("wordfmt OK  BASE=%s" % paths.BASE)

        import tempfile
        from docx import Document
        src = os.path.join(tempfile.gettempdir(), "wf_selftest_in.docx")
        dst = os.path.join(tempfile.gettempdir(), "wf_selftest_out.docx")
        d = Document()
        d.add_paragraph("1 引言")
        d.add_paragraph("这是一段用来验证排版引擎的正文文字。")
        d.save(src)

        prof_path = os.path.join(paths.CONFIG_DIR, "论文格式.yaml")
        if os.path.exists(prof_path):
            prof = Profile.load(prof_path)
            lines.append("载入方案：%s（%d 个部分）" % (prof.name, len(prof.roles)))
        else:
            prof = Profile(name="selftest")
            prof.roles = [Role(name="正文", style="正文",
                               font=Font(east_asia="宋体", size=12),
                               para=Para(align="justify", first_line_chars=2))]
            lines.append("未找到方案文件，使用内置最小方案")
        prof.page = Page()
        report = apply_profile(src, dst, prof)
        lines.append("排版引擎 OK：%s" % report.summary().splitlines()[1])
        res = verify(dst, prof)
        lines.append("校验 OK：%d/%d" % (res.ok, res.checked))

        from wordfmt.cover import apply_cover, list_covers, Cover
        covers = list_covers(paths.BASE)
        lines.append("封面库 OK：%d 个" % len(covers))
        if covers:
            cov = Cover.load(covers[0][1])
            doc2 = Document(dst)
            rep = apply_cover(doc2, cov, mode="rebuild")
            lines.append("封面套用 OK：%s，%d 行，填 %d 项，留空 %d 项" %
                         (cov.name, rep.get("lines", 0), len(rep.get("filled", [])),
                          len(rep.get("blank", []))))
            doc2.save(dst)

            try:
                from wordfmt.preview import draw_cover, page_geometry
                g = page_geometry(cov, 380)
                lines.append("封面预览 OK：版面 %.0f×%.0f px" % (g["width"], g["height"]))
            except Exception as e:
                lines.append("封面预览：绘制模块有问题（%s）" % e)
    except Exception:
        ok = False
        lines.append(traceback.format_exc())
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return 0 if ok else 1


def main() -> int:
    if "--selftest" in sys.argv:
        return selftest()
    from wordfmt.gui import main as gui_main
    gui_main()
    return 0


if __name__ == "__main__":
    sys.exit(main())
