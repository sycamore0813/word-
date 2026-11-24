# -*- coding: utf-8 -*-
"""数据图：把一张数据表画成论文里能直接用的图，并插进 Word（带图题）。

支持：柱状图、分组柱状图、折线图、面积图、饼图、散点图、箱线图、水平条形图。
默认 300 dpi、中文黑体标注，尺寸按论文常用的 8×5 英寸（约 20×12.7cm）。
"""
import csv
import io
import os
import re

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib import font_manager  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm  # noqa: E402

CHART_TYPES = {
    "柱状图": "bar",
    "分组柱状图": "grouped_bar",
    "堆叠柱状图": "stacked_bar",
    "折线图": "line",
    "面积图": "area",
    "饼图": "pie",
    "环形图": "donut",
    "散点图": "scatter",
    "箱线图": "box",
    "水平条形图": "barh",
    "热力图": "heatmap",
    "雷达图": "radar",
}

# 几套现成配色；也可以用「自定义」直接填颜色代码
PALETTES = {
    "学术蓝": ["#2E5C8A", "#C0504D", "#9BBB59", "#8064A2", "#4BACC6", "#F79646", "#7F7F7F"],
    "暖色": ["#B03A2E", "#E67E22", "#F1C40F", "#D35400", "#A04000", "#7E5109"],
    "绿色": ["#1E6F5C", "#289672", "#29BB89", "#6FCF97", "#A8E6CF"],
    "灰度": ["#2F2F2F", "#555555", "#7A7A7A", "#9E9E9E", "#C2C2C2"],
    "彩虹": ["#4E79A7", "#F28E2B", "#E15759", "#76B7B2", "#59A14F", "#EDC948", "#B07AA1"],
}

# 热力图用的渐变色带（matplotlib 标准名）
CMAPS = ["YlGnBu", "Blues", "Greens", "Reds", "Oranges", "Purples",
         "viridis", "RdYlGn", "coolwarm", "Greys"]

PALETTE = PALETTES["学术蓝"]

_FONT_CANDIDATES = ["Microsoft YaHei", "SimHei", "SimSun", "DengXian", "KaiTi"]


def _setup_font():
    """挑一个系统里有的中文字体，避免图上出现方块。"""
    have = {f.name for f in font_manager.fontManager.ttflist}
    for name in _FONT_CANDIDATES:
        if name in have:
            plt.rcParams["font.sans-serif"] = [name, "DejaVu Sans"]
            break
    plt.rcParams["axes.unicode_minus"] = False


def parse_table(text):
    """把粘贴的多行数据解析成 (表头, 行列表)。

    第一行是表头；列之间用 Tab / 逗号 / 分号 / 连续空格分隔。
    """
    lines = [ln for ln in (text or "").replace("\r", "").split("\n") if ln.strip()]
    if not lines:
        return [], []
    def split(ln):
        if "\t" in ln:
            parts = ln.split("\t")
        elif "," in ln:
            parts = ln.split(",")
        elif "；" in ln:
            parts = ln.split("；")
        elif ";" in ln:
            parts = ln.split(";")
        else:
            parts = re.split(r"\s{2,}|\s+", ln)
        return [p.strip() for p in parts if p.strip() != ""]
    header = split(lines[0])
    rows = []
    for ln in lines[1:]:
        cells = split(ln)
        if len(cells) < 2:
            continue
        rows.append(cells)
    return header, rows


def read_csv(path):
    with io.open(path, "r", encoding="utf-8-sig", newline="") as f:
        rows = [r for r in csv.reader(f) if any(c.strip() for c in r)]
    if not rows:
        return [], []
    return rows[0], rows[1:]


def excel_sheets(path):
    """列出 Excel 里的工作表名；读不了就返回空列表。"""
    try:
        import openpyxl

        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        names = list(wb.sheetnames)
        wb.close()
        return names
    except Exception:
        return []


def read_excel(path, sheet=None):
    """读 Excel 第一张表（或指定工作表）→ (表头, 行列表)。

    第一行当表头；空行、空列自动跳过；数字原样保留成文本，交给统一解析。
    """
    try:
        import openpyxl
    except Exception as e:
        raise RuntimeError("缺少 openpyxl，读不了 Excel：%s" % e)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet and sheet in wb.sheetnames:
            ws = wb[sheet]
        else:
            ws = wb[wb.sheetnames[0]]
        raw = []
        for row in ws.iter_rows(values_only=True):
            vals = []
            for v in row:
                if v is None:
                    vals.append("")
                elif isinstance(v, float) and v.is_integer():
                    vals.append(str(int(v)))
                else:
                    vals.append(str(v).strip())
            if any(x for x in vals):
                raw.append(vals)
    finally:
        wb.close()
    if not raw:
        return [], []
    width = max(len(r) for r in raw)
    raw = [r + [""] * (width - len(r)) for r in raw]
    header = raw[0]
    body = [r for r in raw[1:] if any(x for x in r[1:])]
    return header, body


def _to_float(v):
    try:
        return float(str(v).replace(",", "").strip())
    except Exception:
        return None


def render(header, rows, kind="bar", out_path="chart.png", opts=None):
    """画图并存成 PNG，返回图片路径。"""
    opts = opts or {}
    _setup_font()
    title = opts.get("title") or ""
    xlabel = opts.get("xlabel") or ""
    ylabel = opts.get("ylabel") or ""
    width = float(opts.get("width_in") or 8)
    height = float(opts.get("height_in") or 5)
    dpi = int(opts.get("dpi") or 300)
    grid = bool(opts.get("grid", True))
    fontsize = int(opts.get("fontsize") or 11)
    rotate = int(opts.get("xtick_rotation") or 0)
    palette = opts.get("palette") or "学术蓝"
    raw_colors = opts.get("colors")
    if isinstance(raw_colors, str) and raw_colors.strip():
        colors = [c.strip() for c in re.split(r"[,，;；\s]+", raw_colors) if c.strip()]
    elif isinstance(raw_colors, (list, tuple)) and raw_colors:
        colors = list(raw_colors)
    else:
        colors = PALETTES.get(palette, PALETTE)
    if not colors:
        colors = PALETTE
    cmap = opts.get("cmap") or "YlGnBu"
    annotate = bool(opts.get("annotate", True))
    import numpy as _np

    def clean(v):
        return str(v).replace("\t", " ").replace("\n", " ").replace("\r", " ").strip()

    header = [clean(h) for h in header]
    rows = [[clean(c) for c in r] for r in rows]
    title, xlabel, ylabel = clean(title), clean(xlabel), clean(ylabel)

    fig, ax = plt.subplots(figsize=(width, height), dpi=dpi)
    cats = [r[0] for r in rows] if rows else []
    series = []
    for ci in range(1, len(header)):
        vals = [_to_float(r[ci]) if ci < len(r) else None for r in rows]
        if any(v is not None for v in vals):
            series.append((header[ci] if ci < len(header) else "系列%d" % ci, vals))

    if kind == "pie" or kind == "donut":
        vals = [v or 0 for v in (series[0][1] if series else [])]
        kwargs = {"wedgeprops": {"width": 0.42}, "pctdistance": 0.79} if kind == "donut" else {}
        ax.pie(vals, labels=cats, autopct="%1.1f%%", startangle=90,
               colors=colors[:len(cats)], textprops={"fontsize": fontsize}, **kwargs)
        ax.axis("equal")
    elif kind == "heatmap":
        matrix = _np.array([[( _to_float(r[ci]) if _to_float(r[ci]) is not None else 0.0)
                             for ci in range(1, len(header))] for r in rows])
        if matrix.size == 0:
            matrix = _np.zeros((1, 1))
        im = ax.imshow(matrix, cmap=cmap, aspect="auto")
        ax.set_xticks(range(matrix.shape[1]))
        ax.set_xticklabels(header[1:])
        ax.set_yticks(range(matrix.shape[0]))
        ax.set_yticklabels(cats)
        cbar = fig.colorbar(im, ax=ax)
        cbar.ax.tick_params(labelsize=max(7, fontsize - 2))
        if annotate:
            mid = (matrix.max() + matrix.min()) / 2.0 if matrix.size else 0
            for i in range(matrix.shape[0]):
                for j in range(matrix.shape[1]):
                    ax.text(j, i, ("%g" % matrix[i, j]), ha="center", va="center",
                            fontsize=max(6, fontsize - 2),
                            color="white" if matrix[i, j] > mid else "#333333")
    elif kind == "radar":
        # 雷达图：表头各列是指标（轴），每行数据是一个对象
        labels = header[1:]
        n = len(labels)
        if n >= 3 and rows:
            angles = _np.linspace(0, 2 * _np.pi, n, endpoint=False).tolist()
            angles += angles[:1]
            ax.remove()
            ax = fig.add_subplot(111, polar=True)
            # 各指标单位不同（硬度 120、弹性 0.9），先按列归一化到 0~1 再画
            raw = []
            for r in rows:
                line = []
                for ci in range(1, len(labels) + 1):
                    v = _to_float(r[ci]) if ci < len(r) else None
                    line.append(v if v is not None else 0.0)
                raw.append(line)
            cols = list(zip(*raw)) if raw else []
            mins = [min(c) for c in cols]
            maxs = [max(c) for c in cols]
            for ri, line in enumerate(raw):
                vals = []
                for ci, v in enumerate(line):
                    span = (maxs[ci] - mins[ci]) or 1.0
                    vals.append((v - mins[ci]) / span)
                vals += vals[:1]
                ax.plot(angles, vals, linewidth=1.7,
                        color=colors[ri % len(colors)], label=cats[ri])
                ax.fill(angles, vals, alpha=0.15, color=colors[ri % len(colors)])
            ax.set_xticks(angles[:-1])
            ax.set_xticklabels(labels, fontsize=fontsize)
            ax.set_ylim(0, 1.05)
            ax.set_yticklabels([])
            ax.tick_params(labelsize=max(7, fontsize - 2))
            ax.grid(True, linestyle="--", linewidth=0.5, alpha=0.5)
            ax.legend(fontsize=fontsize - 1, frameon=False, loc="upper right",
                      bbox_to_anchor=(1.15, 1.1))
        else:
            ax.text(0.5, 0.5, "雷达图至少要 3 个指标（表头列）", ha="center", va="center")
    elif kind == "scatter":
        xs = [_to_float(c) for c in cats]
        if all(x is not None for x in xs):
            for si, (name, vals) in enumerate(series):
                ax.scatter(xs, [v or 0 for v in vals], s=36, label=name,
                           color=colors[si % len(colors)], alpha=0.85, edgecolors="white")
        else:
            for si, (name, vals) in enumerate(series):
                ax.scatter(range(len(cats)), [v or 0 for v in vals], s=36, label=name,
                           color=colors[si % len(colors)], alpha=0.85, edgecolors="white")
            ax.set_xticks(range(len(cats)))
            ax.set_xticklabels(cats)
    elif kind == "box":
        data = [[v or 0 for v in vals] for _n, vals in series]
        ax.boxplot(data, patch_artist=True,
                   boxprops={"facecolor": colors[0], "alpha": 0.6},
                   medianprops={"color": "#C0504D"})
        ax.set_xticks(range(1, len(series) + 1))
        ax.set_xticklabels([n for n, _v in series])
    elif kind == "line" or kind == "area":
        xs = range(len(cats))
        for si, (name, vals) in enumerate(series):
            ys = [v or 0 for v in vals]
            if kind == "area":
                ax.fill_between(list(xs), ys, alpha=0.35, color=colors[si % len(colors)], label=name)
            ax.plot(list(xs), ys, marker="o", linewidth=1.8, markersize=5,
                    color=colors[si % len(colors)], label=name)
        ax.set_xticks(list(xs))
        ax.set_xticklabels(cats)
    elif kind == "stacked_bar":
        base = [0.0] * len(cats)
        for si, (name, vals) in enumerate(series):
            ys = [v or 0 for v in vals]
            ax.bar(list(range(len(cats))), ys, bottom=base,
                   color=colors[si % len(colors)], label=name, width=0.62)
            base = [b + y for b, y in zip(base, ys)]
        ax.set_xticks(list(range(len(cats))))
        ax.set_xticklabels(cats)
    elif kind == "barh":
        ys = range(len(cats))
        for si, (name, vals) in enumerate(series):
            ax.barh([y + si * 0.8 / max(1, len(series)) for y in ys], [v or 0 for v in vals],
                    height=0.8 / max(1, len(series)), color=colors[si % len(colors)], label=name)
        ax.set_yticks(list(ys))
        ax.set_yticklabels(cats)
    else:  # bar / grouped_bar
        n = len(series)
        base = range(len(cats))
        for si, (name, vals) in enumerate(series):
            if kind == "bar" and n == 1:
                pos = list(base)
                ax.bar(pos, [v or 0 for v in vals], color=colors[0], width=0.62, label=name)
            else:
                offset = (si - (n - 1) / 2) * (0.78 / max(1, n))
                ax.bar([x + offset for x in base], [v or 0 for v in vals],
                       width=0.78 / max(1, n), color=colors[si % len(colors)], label=name)
        ax.set_xticks(list(base))
        ax.set_xticklabels(cats)
        if n == 1 and len(series) == 1:
            ax.legend_.remove() if ax.legend_ else None

    if kind not in ("pie", "donut", "heatmap") and len(series) > 1 and kind != "box":
        ax.legend(fontsize=fontsize - 1, frameon=False)
    if title:
        ax.set_title(title, fontsize=fontsize + 2, pad=12)
    if xlabel and kind not in ("heatmap", "radar"):
        ax.set_xlabel(xlabel, fontsize=fontsize)
    if ylabel and kind not in ("heatmap", "radar"):
        ax.set_ylabel(ylabel, fontsize=fontsize)
    if grid and kind not in ("pie", "donut", "heatmap", "radar"):
        ax.grid(True, axis="y" if kind not in ("barh",) else "x",
                linestyle="--", linewidth=0.6, alpha=0.45)
        ax.set_axisbelow(True)
    for side in ("top", "right"):
        if kind not in ("box", "heatmap", "radar"):
            ax.spines[side].set_visible(False)
    if rotate:
        plt.setp(ax.get_xticklabels(), rotation=rotate, ha="right" if rotate else "center")
    ax.tick_params(labelsize=fontsize - 1)

    fig.tight_layout()
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    fig.savefig(out_path, dpi=dpi, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return out_path


# ---------------------------------------------------------------- 插进 Word
def _caption_number(doc, prefix="图"):
    """算出下一个图号：全篇「图N」最大编号 +1。"""
    from . import docxutil as du

    top = 0
    for p in du.iter_paragraphs(doc):
        m = re.match(r"\s*%s\s*(\d+)(?:[-.．](\d+))?" % prefix, du.text_of(p))
        if m:
            top = max(top, int(m.group(1)))
    return top + 1


def _find_section_end(doc, heading_text):
    """找到标题所在小节的最后一段（下一个同级或更高级标题之前）。"""
    from . import docxutil as du

    paras = [p for p in du.iter_paragraphs(doc)]

    def is_toc(p):
        name = (p.style.name if p.style is not None else "") or ""
        return name.lower().startswith("toc")

    start = None
    for i, p in enumerate(paras):
        if is_toc(p):
            continue
        if heading_text in du.text_of(p):
            start = i
            break
    if start is None:
        return None
    level = du.outline_level(paras[start])
    if level is None:
        level = 1
    for p in paras[start + 1:]:
        if is_toc(p):
            continue
        lv = du.outline_level(p)
        if lv is not None and lv <= level:
            return p
    return None


def insert_figure(doc, image_path, caption=None, profile=None, width_cm=14.0,
                  anchor_text=None, prefix="图", after_section=None):
    """把图片插到文档里，后面跟一行图题（用方案里的图题样式）。

    定位优先级：anchor_text（含这段文字的段落之后）> after_section（该标题所属小节末尾）。
    给了定位却找不到就报错，免得把图丢到文末。
    """
    from . import docxutil as du

    number = _caption_number(doc, prefix)
    caption = caption or "（图题）"
    cap_text = "%s%d %s" % (prefix, number, caption)

    target = None
    before = False
    if anchor_text:
        for p in du.iter_paragraphs(doc):
            if anchor_text in du.text_of(p):
                target = p
                break
        if target is None:
            raise ValueError("文档里找不到含「%s」的段落，图和图题没有插入" % anchor_text)
    elif after_section:
        target = _find_section_end(doc, after_section)
        if target is None:
            raise ValueError("文档里找不到「%s」这个小节，图和图题没有插入" % after_section)
        before = True
    if target is None:
        target = doc.paragraphs[-1] if doc.paragraphs else None

    pic_p = doc.add_paragraph()
    pic_p.alignment = 1  # 居中
    run = pic_p.add_run()
    run.add_picture(image_path, width=Cm(width_cm))

    role = profile.role("图题") if profile else None
    cap_p = doc.add_paragraph(cap_text)
    if role is not None:
        try:
            st = du.ensure_style(doc, role.style) if role.style else None
            if st is not None:
                cap_p.style = st
            du.set_para(cap_p, role.para, font_size_hint=role.font.size)
            font = role.font
            if getattr(font, "char_spacing_pt", None):
                font.char_spacing_pt = None   # 图题不要字间距，否则「图 1」中间会空一格
            du.set_font(cap_p, font)
        except Exception:
            pass
    # 搬到 target 后面（小节末尾那种是搬到下一个标题前面）
    if target is not None:
        if before:
            target._p.addprevious(pic_p._p)
            target._p.addprevious(cap_p._p)
        else:
            target._p.addnext(cap_p._p)
            target._p.addnext(pic_p._p)
    return pic_p, cap_p, number, target
