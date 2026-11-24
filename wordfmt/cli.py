# -*- coding: utf-8 -*-
"""命令行入口。

例子：
  python -m wordfmt.cli extract 模板.docx
  python -m wordfmt.cli apply   论文.docx 方案.yaml --word
  python -m wordfmt.cli verify  论文_已排版.docx 方案.yaml
  python -m wordfmt.cli roles   论文.docx 方案.yaml
"""
from __future__ import annotations

import argparse
import os
import shutil
import sys
from datetime import datetime

from . import paths

paths.ensure_dirs()

from .apply import apply_profile  # noqa: E402
from .classify import Classifier  # noqa: E402
from .extract import build_profile  # noqa: E402
from .profile import Profile  # noqa: E402
from .verify import verify  # noqa: E402


def default_output(src: str) -> str:
    d, base = os.path.split(os.path.abspath(src))
    stem, ext = os.path.splitext(base)
    return os.path.join(d, stem + "_已排版" + (ext or ".docx"))


def backup_of(src: str) -> str:
    """固定一个备份文件：同一个文档多次排版只留一份备份，不再堆一桌面。"""
    d, base = os.path.split(os.path.abspath(src))
    stem, ext = os.path.splitext(base)
    return os.path.join(d, "%s_备份%s" % (stem, ext or ".docx"))


def _profile_path_for(template: str) -> str:
    stem = os.path.splitext(os.path.basename(template))[0]
    return os.path.join(paths.CONFIG_DIR, stem + ".yaml")


def cmd_extract(args) -> int:
    prof = build_profile(args.template, name=args.name)
    out = args.out or _profile_path_for(args.template)
    prof.save(out)
    print("已生成格式方案：%s" % out)
    print("识别到 %d 个部分：" % len(prof.roles))
    for r in prof.roles:
        print("  %-10s 样式=%-16s %s" % (r.name, r.style or "-", r.describe()))
    print("页面：%.2f×%.2fcm，边距 上%.2f 下%.2f 左%.2f 右%.2f" %
          (prof.page.width_cm, prof.page.height_cm, prof.page.margin_top_cm,
           prof.page.margin_bottom_cm, prof.page.margin_left_cm, prof.page.margin_right_cm))
    if prof.page.header_text:
        print("页眉：%s" % prof.page.header_text)

    if not getattr(args, "no_cover", False):
        _save_cover(args.template, args.name)

    if not getattr(args, "no_verify", False):
        try:
            from .roundtrip import roundtrip_check

            print("\n正在比对：用新方案把模板重排一遍，逐页和原模板比 …")
            rt = roundtrip_check(args.template, prof)
            print("  比对结果：%s" % rt.get("message", ""))
            print("  %s" % ("通过 ✅" if rt.get("ok") else "未通过 ⚠（看差异大的页）"))
        except Exception as e:
            print("  比对跳过：%s" % e)
    return 0


def _save_cover(doc_path: str, name: str = "") -> str:
    """从文档提取封面并入库，返回封面库路径（失败返回空串）。"""
    from docx import Document
    from .cover import extract_cover, save_to_library
    try:
        doc = Document(doc_path)
        stem = os.path.splitext(os.path.basename(doc_path))[0]
        cov = extract_cover(doc, name=(name or stem) + "封面", source=doc_path)
        if cov is None:
            print("没有识别到封面（缺分节符或摘要标题），已跳过")
            return ""
        path = save_to_library(cov, paths.BASE, filename=(name or stem) + "_封面")
        print("封面已入库：%s" % path)
        print("  " + cov.summary())
        return path
    except Exception as e:
        print("封面提取失败：%s" % e)
        return ""


def cmd_cover(args) -> int:
    path = _save_cover(args.doc, args.name or "")
    return 0 if path else 1


def cmd_apply(args) -> int:
    if not os.path.exists(args.input):
        print("找不到文件：%s" % args.input)
        return 2
    prof = Profile.load(args.profile)
    if getattr(args, "ensure_sections", False):
        prof.options["ensure_sections"] = True
    out = args.output or default_output(args.input)
    if os.path.abspath(out) == os.path.abspath(args.input):
        print("输出文件不能覆盖原文件，请换一个 --output")
        return 2

    if args.backup and prof.option("backup", True):
        bak = backup_of(args.input)
        if os.path.exists(bak):
            print("备份已存在，保留这一份：%s" % bak)
        else:
            shutil.copy2(args.input, bak)
            print("已备份原文件：%s" % bak)

    cover_obj = None
    if getattr(args, "cover", None):
        from .cover import Cover
        cover_obj = Cover.load(args.cover)
        print("封面库：%s（%s 模式）" % (cover_obj.name, args.cover_mode))

    report = apply_profile(args.input, out, prof, progress=lambda m: print("  " + m),
                           cover=cover_obj, cover_mode=args.cover_mode)
    print(report.summary())
    print("已生成：%s" % out)

    if args.verify:
        res = verify(out, prof)
        print("\n" + res.text_report())
        if not args.ignore_fail and res.issues:
            return 1

    if args.word:
        from . import wordio
        pdf_path = None
        if args.pdf is not None:
            pdf_path = args.pdf or (os.path.splitext(out)[0] + ".pdf")
        ok, msg = wordio.update_fields(out, pdf_path=pdf_path)
        print(("\nWord 更新页/目录：成功" if ok else "\nWord 更新页/目录：失败") + "\n" + msg)
        if pdf_path:
            print("PDF：%s" % pdf_path)
    return 0


def cmd_cite(args) -> int:
    """联网检索文献 → 在正文里插引用 → 导出 RIS。"""
    from . import litcite

    profile = Profile.load(args.profile)
    conf = {
        "count": args.count,
        "max_per_para": args.per_para,
        "style": args.style,
        "year_from": args.year_from,
        "require_doi": not args.no_doi,
        "topics": args.topics or "",
        "triggers": [t for t in (args.trigger or "").split(",") if t.strip()],
    }
    out = args.output or os.path.join(
        os.path.dirname(os.path.abspath(args.input)),
        os.path.splitext(os.path.basename(args.input))[0] + "_已引用.docx")
    res = litcite.cite_document(args.input, out, profile, conf, ris_dir=args.ris_dir,
                                progress=lambda m: print("  " + m))
    if not res["papers"]:
        print("没有找到合适的文献。可以加 --topics \"extrusion cooking snack food\" 再试。")
        return 1
    print("已插入 %d 处引用：" % len(res["papers"]))
    for para, end, n, sent, paper in res["picked"]:
        print("  [%d] %s" % (n, sent[:30]))
        print("       %s (%s) %s" % ((paper.title or "")[:70], paper.year or "", paper.doi or ""))
    print("文献清单：%s" % res["md"])
    print("RIS（EndNote 导入用）：%s" % res["ris"])
    print("已生成：%s" % out)
    return 0


def cmd_chart(args) -> int:
    """数据表 → 数据图（PNG），可顺手插进 Word 并配图题。"""
    from . import charts

    if args.data.lower().endswith((".xlsx", ".xlsm")):
        header, rows = charts.read_excel(args.data, args.sheet)
    else:
        import io as _io
        with _io.open(args.data, "r", encoding="utf-8-sig", errors="replace") as f:
            header, rows = charts.parse_table(f.read())
    if not rows:
        print("数据读不出来：%s" % args.data)
        return 1
    out_png = args.out or os.path.splitext(args.data)[0] + "_%s.png" % (
        [k for k, v in charts.CHART_TYPES.items() if v == args.kind] or [args.kind])[0]
    charts.render(header, rows, args.kind, out_png, {
        "title": args.title or "", "xlabel": args.xlabel or "", "ylabel": args.ylabel or "",
        "width_in": args.width_in, "height_in": args.height_in, "fontsize": args.fontsize,
        "xtick_rotation": args.rotate, "palette": args.palette,
        "colors": args.colors or "", "cmap": args.cmap,
    })
    print("已生成图：%s" % out_png)
    if args.insert:
        from docx import Document
        from .profile import Profile

        doc = Document(args.insert)
        profile = Profile.load(args.profile) if args.profile else None
        pic, cap, no, _t = charts.insert_figure(
            doc, out_png, args.caption, profile, width_cm=args.width_cm,
            anchor_text=args.anchor, after_section=args.after_section)
        out = args.output or (os.path.splitext(args.insert)[0] + "_插图.docx")
        doc.save(out)
        print("已插入：%s → %s" % (cap.text, out))
    return 0


def cmd_refs(args) -> int:
    """RIS（或联网检索）→ GB/T 7714 参考文献表 → 插进 Word。"""
    from .profile import Profile
    from . import refs

    profile = Profile.load(args.profile)
    if args.topic:
        from . import litcite

        papers = litcite.search(args.topic, args.count, args.year_from, True)[:args.count]
        if not papers:
            print("没检索到文献，换个 --topic 试试")
            return 1
        res = refs.from_papers_to_doc(papers, args.doc, profile, args.output,
                                      style=args.format, template=args.template,
                                      include_doi=args.with_doi)
    else:
        res = refs.from_ris_to_doc(args.ris, args.doc, profile, args.output,
                                   style=args.format, template=args.template,
                                   include_doi=args.with_doi)
    print("共 %d 条，已写入：%s" % (res["count"], res["out"]))
    for t in res["entries"][:3]:
        print("  " + t[:90])
    return 0


def cmd_analyze(args) -> int:
    """文献分析：统计 + 报告 + 年份分布图。"""
    from . import analysis, citestyle

    if args.list:
        print("内置参考文献格式：")
        for name in citestyle.STYLES:
            print("  -", name)
        return 0
    papers = []
    if args.topic:
        from . import litcite

        papers = litcite.search(args.topic, args.count, None, True)[:args.count]
    elif args.ris:
        from . import refs

        papers = [refs.ris_to_paper(e) for e in refs.parse_ris(args.ris)]
    if not papers:
        print("没有文献可分析：给 --ris 文件，或用 --topic 联网检索")
        return 1
    folder = args.out or os.path.join(os.path.dirname(os.path.abspath(args.ris)) if args.ris
                                      else os.path.join(os.getcwd(), "输出", "文献库"))
    md, png = analysis.export(papers, folder, "文献分析")
    print("共 %d 篇" % len(papers))
    print("报告：%s" % md)
    if png:
        print("年份分布图：%s" % png)
    return 0


def cmd_styles(args) -> int:
    """管理参考文献格式（内置只读，自定义可新建/改名/删除）。"""
    from . import citestyle

    if args.list or not (args.add or args.rename or args.delete):
        print("内置格式（只读）：")
        for name in citestyle.STYLES:
            print("  -", name)
        custom = citestyle.load_custom()
        print("\n自定义格式（可改）：" + ("无" if not custom else ""))
        for name, conf in custom.items():
            print("  - %s  %s" % (name, conf.get("template", "")[:60]))
        print("\n文件：%s" % citestyle.CUSTOM_FILE)
        return 0
    if args.add:
        citestyle.add_custom(args.add, args.template)
        print("已新建：%s" % args.add)
        return 0
    if args.rename:
        old, new = args.rename
        citestyle.rename_custom(old, new)
        print("已改名：%s → %s" % (old, new))
        return 0
    if args.delete:
        ok = citestyle.delete_custom(args.delete)
        print("已删除：%s" % args.delete if ok else "没找到：%s" % args.delete)
        return 0 if ok else 1
    return 0


def cmd_renumber(args) -> int:
    """按正文出现顺序重排引用编号，并同步参考文献表。"""
    from docx import Document

    from . import citeindex

    doc = Document(args.input)
    problems = citeindex.check(doc)
    if args.check:
        if problems:
            print("有 %d 处不规范：" % len(problems))
            for p in problems[:20]:
                print("  -", p)
            return 1
        print("编号规范：正文按出现顺序、且都能在参考文献表里找到。")
        return 0
    if not problems:
        print("已经是规范编号，无需重排。")
        return 0
    rep = citeindex.renumber(doc)
    out = args.output or os.path.splitext(args.input)[0] + "_已重排.docx"
    doc.save(out)
    print("重排完成：正文 %d 处、参考文献 %d 条" % (rep["marks"], rep["moved"]))
    pairs = ["[%d]→[%d]" % (o, n) for o, n in sorted(rep["map"].items()) if o != n]
    if pairs:
        print("编号变化：" + "，".join(pairs[:12]) + ("…" if len(pairs) > 12 else ""))
    if rep["uncited"]:
        print("正文未引用的条目（排在最后）：%s" % "、".join(str(x) for x in rep["uncited"]))
    print("已保存：%s" % out)
    return 0


def cmd_newdoc(args) -> int:
    """一步生成「排好格式的空白论文」。"""
    from . import skeleton

    prof = Profile.load(args.profile)
    out = args.out or os.path.join(OUTPUT_DIR, "空白论文_已排版.docx")
    path, report = skeleton.new_formatted_doc(prof, out)
    print("已生成排好格式的空白论文：%s" % path)
    print("  结构：摘要 + 关键词 → 目录 → 1 引言 / 2 文献综述 / 3 研究方法 / "
          "4 结果与分析 / 5 讨论 / 6 结论与展望 → 致谢 → 附录 → 参考文献")
    if report is not None:
        print("  已按方案「%s」排好版" % prof.name)
    return 0


def cmd_skeleton(args) -> int:
    """从模板生成空白结构稿（只留章节骨架）。"""
    from . import skeleton
    from .profile import Profile

    profile = Profile.load(args.profile) if args.profile else None
    if args.from_template:
        out, st = skeleton.build_skeleton(args.template, args.out, profile,
                                          keep_cover=args.keep_cover, pad=not args.no_pad)
        print("已生成空白结构稿（沿用模板章节）：%s" % out)
        print("  保留段落 %d 个（其中标题 %d 个，重排层级 %d 个）；删掉正文 %d 段"
              % (st["kept"], st["headings"], st["relabeled"], st["removed"]))
        if st["cover"]:
            print("  保留了封面 %d 行" % st["cover"])
    else:
        out, st = skeleton.build_outline(args.template, args.out, profile,
                                         with_toc=not args.no_toc)
        print("已生成空白结构稿（标准大纲）：%s" % out)
        print("  一级章节 %d 个：引言 / 文献综述 / 研究方法 / 结果与分析 / 讨论 / 结论与展望"
              % st["sections"])
        print("  另有：摘要 + 关键词、目录（TOC 域）、致谢、附录、参考文献")
    print("  下一步：打开它，在每个标题下直接写内容；写完再回「排版文档」页排一次。")
    return 0


def cmd_verify(args) -> int:
    prof = Profile.load(args.profile)
    res = verify(args.doc, prof)
    print(res.text_report())
    return 0 if not res.issues and not res.page_problems else 1


def cmd_roles(args) -> int:
    from docx import Document
    prof = Profile.load(args.profile)
    doc = Document(args.doc)
    clf = Classifier(doc, prof)
    plan = clf.plan()
    for p, role, reason in plan:
        text = "".join(t.text or "" for t in p._p.iter(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")).strip()
        print("%-10s %-6s %s" % (role.name if role else "-", reason or "-", text[:50]))
    print("\n统计：")
    for name, cnt in clf.stats(plan).most_common():
        print("  %-10s %d" % (name, cnt))
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="wordfmt", description="Word 文档格式规范化工具")
    sub = ap.add_subparsers(dest="cmd")

    e = sub.add_parser("extract", help="从模板 docx 提取格式方案（顺带把封面存进封面库）")
    e.add_argument("template")
    e.add_argument("--out", help="输出的 yaml 路径")
    e.add_argument("--name", help="方案名称")
    e.add_argument("--no-cover", action="store_true", help="不提取封面")
    e.add_argument("--no-verify", dest="no_verify", action="store_true",
                   help="提取后不做「重排模板并逐页比对」")
    e.set_defaults(func=cmd_extract)

    c = sub.add_parser("cover", help="只从文档提取封面，存进封面库")
    c.add_argument("doc")
    c.add_argument("--name", help="封面名称")
    c.set_defaults(func=cmd_cover)

    a = sub.add_parser("apply", help="按方案排版文档")
    a.add_argument("input")
    a.add_argument("profile")
    a.add_argument("--output", help="输出路径（默认 xxx_已排版.docx）")
    a.add_argument("--cover", help="封面库里的封面 yaml（covers/*.yaml）")
    a.add_argument("--cover-mode", choices=["format", "rebuild", "blank"], default="format",
                   help="format=只改封面格式（默认，最稳）；rebuild=按封面库重建并自动填姓名学号等；"
                        "blank=重建但全部留空（当空表用，谁用谁填）")
    a.add_argument("--no-backup", dest="backup", action="store_false")
    a.add_argument("--verify", action="store_true", help="改完立即校验")
    a.add_argument("--ignore-fail", action="store_true", help="校验有问题也返回 0")
    a.add_argument("--ensure-sections", dest="ensure_sections", action="store_true",
                   help="排版时把缺的「参考文献 / 致谢 / 附录」补在文末（内容留空）")
    a.add_argument("--word", action="store_true", help="用本机 Word 更新目录与页码域")
    a.add_argument("--pdf", nargs="?", const="", default=None,
                   help="同时导出 PDF（不给路径就放在输出文件旁边）")
    a.set_defaults(func=cmd_apply)

    ci = sub.add_parser("cite", help="联网找真实文献，在正文里插引用并导出 RIS（EndNote 可导入）")
    ci.add_argument("input")
    ci.add_argument("profile")
    ci.add_argument("--output", help="输出路径（默认 xxx_已引用.docx）")
    ci.add_argument("--count", type=int, default=2, help="插几处引用（默认 2）")
    ci.add_argument("--per-para", type=int, default=1, help="每段最多几处（默认 1）")
    ci.add_argument("--style", default="顺序编码-上标[n]", help="引用样式：" + "/".join(
        ["顺序编码-上标[n]", "顺序编码-普通[n]", "顺序编码-上标(n)", "作者-年份"]))
    ci.add_argument("--topics", help="检索主题（留空=按句子自动推断），英文效果最好")
    ci.add_argument("--trigger", help="只在含这些词的句子上加引用，逗号分隔")
    ci.add_argument("--year-from", dest="year_from", type=int, default=2010, help="最早年份")
    ci.add_argument("--no-doi", dest="no_doi", action="store_true", help="不要求文献有 DOI")
    ci.add_argument("--ris-dir", dest="ris_dir", help="RIS 输出目录（默认输出目录下 文献库）")
    ci.set_defaults(func=cmd_cite)

    ch = sub.add_parser("chart", help="把数据表画成数据图（PNG），可插进 Word 并配图题")
    ch.add_argument("data", help="数据文件：.xlsx / .csv / .txt（第一行表头）")
    ch.add_argument("--sheet", help="Excel 的工作表名（默认第一张）")
    ch.add_argument("--kind", default="bar", choices=[
        "bar", "grouped_bar", "stacked_bar", "line", "area", "pie", "donut",
        "scatter", "box", "barh", "heatmap", "radar"])
    ch.add_argument("--title")
    ch.add_argument("--xlabel")
    ch.add_argument("--ylabel")
    ch.add_argument("--width-in", dest="width_in", type=float, default=8)
    ch.add_argument("--height-in", dest="height_in", type=float, default=5)
    ch.add_argument("--fontsize", type=int, default=11)
    ch.add_argument("--rotate", type=int, default=0, help="横轴标签旋转角度")
    ch.add_argument("--palette", default="学术蓝", help="配色：学术蓝/暖色/绿色/灰度/彩虹")
    ch.add_argument("--colors", help="自定义颜色，逗号分隔，如 #1E6F5C,#E67E22")
    ch.add_argument("--cmap", default="YlGnBu", help="热力图色带：YlGnBu/Blues/Reds/viridis…")
    ch.add_argument("--out", help="PNG 输出路径")
    ch.add_argument("--insert", help="要插图的 docx")
    ch.add_argument("--profile", help="格式方案（给图题用）")
    ch.add_argument("--anchor", help="插在含这段文字的段落后面")
    ch.add_argument("--after-section", dest="after_section",
                    help="插在某个标题所属小节的末尾（填标题里的文字）")
    ch.add_argument("--caption", help="图题文字")
    ch.add_argument("--width-cm", dest="width_cm", type=float, default=14.0)
    ch.add_argument("--output", help="插图后的 docx")
    ch.set_defaults(func=cmd_chart)

    rf = sub.add_parser("refs", help="生成 GB/T 7714 参考文献表并插入（RIS 或联网检索，不用 EndNote）")
    rf.add_argument("ris", nargs="?", help="RIS 文件（用 --topic 时可以不给）")
    rf.add_argument("doc")
    rf.add_argument("profile")
    rf.add_argument("--output")
    rf.add_argument("--topic", help="不给 RIS 时，直接联网按这个主题检索")
    rf.add_argument("--count", type=int, default=5, help="检索几篇（配 --topic）")
    rf.add_argument("--year-from", dest="year_from", type=int, default=2010)
    rf.add_argument("--format", help="参考文献格式（默认 GB/T 7714），可用 analyze --list 查看")
    rf.add_argument("--template", help="自定义模板，如 \"[{seq}] {authors}. {title}. {journal}, {year}.\"")
    rf.add_argument("--with-doi", dest="with_doi", action="store_true",
                    help="条目末尾附 DOI（默认不附，学校模板一般不要求）")
    rf.set_defaults(func=cmd_refs)

    an = sub.add_parser("analyze", help="导出文献分析报告（年份分布/期刊/作者/DOI 比例）")
    an.add_argument("ris", nargs="?", help="RIS 文件（用 --topic 时可以不给）")
    an.add_argument("--topic", help="直接联网按主题检索")
    an.add_argument("--count", type=int, default=20)
    an.add_argument("--out", help="报告输出目录（默认 输出/文献库）")
    an.add_argument("--list", action="store_true", help="只列出内置的参考文献格式名")
    an.set_defaults(func=cmd_analyze)

    st = sub.add_parser("styles", help="管理参考文献格式（列出/新建/改名/删除自定义格式）")
    st.add_argument("--list", action="store_true", help="列出所有格式")
    st.add_argument("--add", metavar="NAME", help="新建一个自定义格式")
    st.add_argument("--template", help="配合 --add 使用的模板")
    st.add_argument("--rename", nargs=2, metavar=("OLD", "NEW"), help="重命名自定义格式")
    st.add_argument("--delete", metavar="NAME", help="删除自定义格式")
    st.set_defaults(func=cmd_styles)

    rn = sub.add_parser("renumber", help="引用编号自动排序：按正文出现顺序重排 [n] 并同步参考文献表")
    rn.add_argument("input")
    rn.add_argument("--output")
    rn.add_argument("--check", action="store_true", help="只检查，不改文件")
    rn.set_defaults(func=cmd_renumber)

    nd = sub.add_parser("newdoc", help="一步生成排好格式的空白论文（标准大纲）")
    nd.add_argument("--profile", required=True)
    nd.add_argument("--out", help="输出路径（默认 输出/空白论文_已排版.docx）")
    nd.set_defaults(func=cmd_newdoc)

    sk = sub.add_parser("skeleton", help="从模板生成空白结构稿（保留章节目录，正文留空）")
    sk.add_argument("template")
    sk.add_argument("--out", help="输出路径（默认 模板名_结构稿.docx）")
    sk.add_argument("--profile", help="格式方案（不给我就现提取一份）")
    sk.add_argument("--from-template", dest="from_template", action="store_true",
                    help="沿用模板里的章节标题（默认生成标准大纲）")
    sk.add_argument("--keep-cover", dest="keep_cover", action="store_true",
                    help="配合 --from-template：连封面一起保留")
    sk.add_argument("--no-toc", dest="no_toc", action="store_true", help="不要目录")
    sk.add_argument("--no-pad", dest="no_pad", action="store_true", help="标题后不补空段")
    sk.set_defaults(func=cmd_skeleton)

    v = sub.add_parser("verify", help="校验文档是否符合方案")
    v.add_argument("doc")
    v.add_argument("profile")
    v.set_defaults(func=cmd_verify)

    r = sub.add_parser("roles", help="只做识别，打印每段被判成什么，不写文件")
    r.add_argument("doc")
    r.add_argument("profile")
    r.set_defaults(func=cmd_roles)

    args = ap.parse_args(argv)
    if not getattr(args, "func", None):
        ap.print_help()
        return 2
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
