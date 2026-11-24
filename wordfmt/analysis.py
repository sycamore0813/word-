# -*- coding: utf-8 -*-
"""文献分析：给一批文献做统计，导出一份可读的报告（Markdown）+ 年份分布图。"""
import os
from collections import Counter


def _years(papers):
    out = []
    for p in papers:
        try:
            y = int(str(p.year)[:4])
        except Exception:
            continue
        if 1900 < y < 2100:
            out.append(y)
    return out


def stats(papers):
    years = _years(papers)
    journals = Counter((p.journal or "（未注明）").strip() for p in papers)
    authors = Counter()
    for p in papers:
        for a in (p.authors or []):
            authors[a.strip()] += 1
    dois = sum(1 for p in papers if (p.doi or "").strip())
    oa = sum(1 for p in papers if getattr(p, "is_oa", False))
    cited = [p.cited_by or 0 for p in papers]
    return {
        "count": len(papers),
        "years": years,
        "year_min": min(years) if years else None,
        "year_max": max(years) if years else None,
        "journals": journals,
        "authors": authors,
        "doi_ratio": dois / len(papers) if papers else 0,
        "oa_ratio": oa / len(papers) if papers else 0,
        "cited_total": sum(cited),
        "cited_max": max(cited) if cited else 0,
    }


def build_report(papers, title="文献分析报告", style=None, chart_path=None):
    st = stats(papers)
    lines = ["# %s" % title, ""]
    lines.append("- 文献数量：**%d 篇**" % st["count"])
    if st["year_min"]:
        lines.append("- 年份跨度：**%s – %s**" % (st["year_min"], st["year_max"]))
    lines.append("- 有 DOI 的比例：**%.0f%%**" % (st["doi_ratio"] * 100))
    lines.append("- 开放获取（可免费下载）比例：**%.0f%%**" % (st["oa_ratio"] * 100))
    lines.append("- 被引次数合计：%d（单篇最高 %d）" % (st["cited_total"], st["cited_max"]))
    lines.append("")

    if st["years"]:
        lines.append("## 年份分布")
        lines.append("")
        cnt = Counter(st["years"])
        for y in sorted(cnt):
            bar = "█" * max(1, int(cnt[y] * 20 / max(cnt.values())))
            lines.append("- %d：%s %d 篇" % (y, bar, cnt[y]))
        lines.append("")

    if st["journals"]:
        lines.append("## 来源期刊（前 10）")
        lines.append("")
        lines.append("| 期刊 | 篇数 |")
        lines.append("|---|---|")
        for name, n in st["journals"].most_common(10):
            lines.append("| %s | %d |" % (name, n))
        lines.append("")

    if st["authors"]:
        lines.append("## 高频作者（前 10）")
        lines.append("")
        lines.append("| 作者 | 出现次数 |")
        lines.append("|---|---|")
        for name, n in st["authors"].most_common(10):
            lines.append("| %s | %d |" % (name, n))
        lines.append("")

    lines.append("## 文献清单")
    lines.append("")
    lines.append("| # | 年份 | 标题 | 期刊 | DOI |")
    lines.append("|---|---|---|---|---|")
    from . import citestyle

    for i, p in enumerate(papers, 1):
        lines.append("| %d | %s | %s | %s | %s |" % (
            i, p.year or "", (p.title or "").replace("|", "/"),
            (p.journal or "").replace("|", "/"), p.doi or ""))
    lines.append("")
    if chart_path:
        lines.append("![年份分布](%s)" % os.path.basename(chart_path))
        lines.append("")
    lines.append("> 数据来自 OpenAlex / Crossref 的真实题录，DOI 可直接解析。")
    return "\n".join(lines), st


def export(papers, folder, stem="文献分析", make_chart=True):
    """写报告 md（可选同时画年份分布柱状图），返回文件路径。"""
    folder = os.path.abspath(folder)
    os.makedirs(folder, exist_ok=True)
    chart_path = ""
    st = stats(papers)
    if make_chart and st["years"]:
        try:
            from . import charts

            cnt = Counter(st["years"])
            header = ["年份", "篇数"]
            rows = [[str(y), str(cnt[y])] for y in sorted(cnt)]
            chart_path = os.path.join(folder, "%s_年份分布.png" % stem)
            charts.render(header, rows, "bar", chart_path,
                          {"title": "文献年份分布", "xlabel": "年份", "ylabel": "篇数",
                           "palette": "学术蓝"})
        except Exception:
            chart_path = ""
    report, _st = build_report(papers, title="文献分析报告（%s）" % stem, chart_path=chart_path)
    md_path = os.path.join(folder, "%s.md" % stem)
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(report)
    return md_path, chart_path
