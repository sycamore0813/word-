# -*- coding: utf-8 -*-
"""封面预览：在 tkinter 画布上把封面画出来。

不依赖 Word、不依赖第三方渲染库：封面库里有每行的字号/对齐/缩进/空行宽度和校徽图片，
按页边距和 1pt=2.54/72cm 换算成像素画一遍就够了，几毫秒出图，断网也行。
"""
from __future__ import annotations

import os
from typing import List, Optional

from .cover import Cover, CoverLine

# 中文字体名 → tkinter 能认的字体
FONT_MAP = {
    "宋体": "SimSun", "新宋体": "NSimSun", "黑体": "SimHei", "仿宋": "FangSong",
    "仿宋_GB2312": "FangSong", "楷体": "KaiTi", "楷体_GB2312": "KaiTi",
    "微软雅黑": "Microsoft YaHei", "等线": "DengXian", "幼圆": "YouYuan",
    "Times New Roman": "Times New Roman", "Arial": "Arial",
}

PT_TO_CM = 2.54 / 72.0


def _family(name: Optional[str]) -> str:
    if not name:
        return "SimSun"
    return FONT_MAP.get(name, name)


def _display_text(line: CoverLine) -> str:
    """没有填内容时，把「下划线空格」显示成 _，跟 Word 里看到的横线一致。"""
    out = []
    for text, u in line.render_segments({}):
        out.append("_" * len(text) if (u and not text.strip()) else text)
    return "".join(out)


def _line_metrics(line: CoverLine, scale: float, default_size: float = 12.0):
    """返回 (字号像素, 行高像素, 段前像素, 段后像素, 首行缩进像素, 左缩进像素)。"""
    size_pt = line.font.size or default_size
    size_px = size_pt * PT_TO_CM * scale
    spacing = line.para.line_spacing or 1.0
    if line.para.line_exact_pt:
        height = line.para.line_exact_pt * PT_TO_CM * scale
    else:
        height = size_px * 1.18 * spacing
    before = (line.para.space_before_pt or 0) * PT_TO_CM * scale
    after = (line.para.space_after_pt or 0) * PT_TO_CM * scale
    indent = (line.para.first_line_chars or 0) * size_px
    left = (line.para.left_chars or 0) * size_px
    return size_px, height, before, after, indent, left


def page_geometry(cover: Cover, width_px: int):
    pag = cover.page or {}
    width_cm = float(pag.get("width_cm") or 21.0)
    height_cm = float(pag.get("height_cm") or 29.7)
    scale = width_px / width_cm
    return {
        "scale": scale,
        "width": width_px,
        "height": int(round(height_cm * scale)),
        "left": float(pag.get("left_cm") or 3.17) * scale,
        "right": float(pag.get("right_cm") or 3.17) * scale,
        "top": float(pag.get("top_cm") or 2.54) * scale,
    }


def draw_format_preview(canvas, profile, width_px: int = 380) -> List[str]:
    """把「格式方案」画成一页示意：标题、正文、图题、表题、参考文献各是什么样。"""
    canvas.delete("all")
    page = getattr(profile, "page", None)
    pag = {
        "width_cm": getattr(page, "width_cm", 21.0) or 21.0,
        "height_cm": getattr(page, "height_cm", 29.7) or 29.7,
        "left_cm": getattr(page, "margin_left_cm", 3.17) or 3.17,
        "right_cm": getattr(page, "margin_right_cm", 3.17) or 3.17,
        "top_cm": getattr(page, "margin_top_cm", 2.54) or 2.54,
    }
    scale = width_px / pag["width_cm"]
    width = width_px
    height = int(round(pag["height_cm"] * scale))
    left = pag["left_cm"] * scale
    right = width - pag["right_cm"] * scale
    usable = right - left
    canvas.configure(width=width, height=height, bg="white",
                     highlightthickness=1, highlightbackground="#cccccc")

    samples = [
        (["论文题目"], "基于深度学习的柑橘品质无损检测研究"),
        (["摘要标题"], "摘  要"),
        (["关键词"], "关键词：柑橘；无损检测；深度学习；品质评价"),
        (["一级标题"], "1 引言"),
        (["二级标题"], "1.1 研究背景与意义"),
        (["三级标题"], "1.1.1 国内外研究现状"),
        (["正文"], "食品感官评价是依托人体感官器官对食品色泽、香气、滋味、质地等品质指标"
                 "进行综合判定的科学方法，是食品品质检测、工艺优化与产品研发的核心手段。"
                 "本研究以草莓汁为对象，比较发酵前后的感官品质差异。"),
        (["图题"], "图1-1 不同发酵时间感官评分变化"),
        (["表题"], "表1-1 两种草莓果汁样品感官评分统计表"),
        (["表内文字"], "评价指标        样品A        样品B        标准差"),
        (["参考文献标题"], "参考文献"),
        (["参考文献条目"], "[1] 张三, 李四. 草莓汁发酵工艺优化[J]. 食品科学, 2023, 44(5): 12-18."),
    ]
    roles = {r.name: r for r in getattr(profile, "roles", []) if getattr(r, "enabled", True)}

    y = pag["top_cm"] * scale
    drawn = 0
    for names, text in samples:
        role = next((roles[n] for n in names if n in roles), None)
        if role is None:
            continue
        size_pt = role.font.size or 12.0
        size_px = size_pt * PT_TO_CM * scale
        spacing = role.para.line_spacing or 1.0
        if role.para.line_exact_pt:
            line_h = role.para.line_exact_pt * PT_TO_CM * scale
        else:
            line_h = size_px * 1.18 * spacing
        before = (role.para.space_before_pt or 0) * PT_TO_CM * scale
        after = (role.para.space_after_pt or 0) * PT_TO_CM * scale
        indent = (role.para.first_line_chars or 0) * size_px
        align = (role.para.align or "left").lower()

        y += before
        if align == "center":
            x, anchor, wrap = width / 2, "n", int(usable)
        elif align == "right":
            x, anchor, wrap = right, "ne", int(usable)
        else:
            x, anchor, wrap = left + indent, "nw", int(usable - indent)
        canvas.create_text(x, y + max(0, (line_h - size_px) / 2), text=text, anchor=anchor,
                           font=(_family(role.font.east_asia or role.font.ascii),
                                 -max(6, int(round(size_px)))),
                           fill="#000000", width=wrap)
        lines = max(1, int(len(text) * size_px / max(wrap, 1)) + 1)
        y += line_h * lines + after
        drawn += 1
        if y > height - 10:
            break

    if drawn == 0:
        canvas.create_text(width / 2, 40, text="这个方案里没有可预览的部分", anchor="n")
    elif y < height - 30:
        canvas.create_text(width / 2, height - 24, text="（示意：按方案里的字号/对齐/行距画，实际以 Word 为准）",
                           anchor="n", fill="#999999", font=("Microsoft YaHei", -11))
    return []


def draw_cover(canvas, cover: Cover, width_px: int = 380) -> List[str]:
    """在 canvas 上画出封面，返回图片对象的引用列表（tkinter 要求保持引用）。"""
    canvas.delete("all")
    g = page_geometry(cover, width_px)
    canvas.configure(width=g["width"], height=g["height"], bg="white",
                     highlightthickness=1, highlightbackground="#cccccc")
    usable = g["width"] - g["left"] - g["right"]
    y = g["top"]
    keep_images = []

    for line in cover.lines:
        if line.is_picture:
            path = cover.image_path(line)
            size_px, height, before, after, _ind, _left = _line_metrics(line, g["scale"])
            if path and os.path.exists(path):
                try:
                    import tkinter as tk
                    img = tk.PhotoImage(file=path)
                    size = line.image or {}
                    cx = size.get("cx")
                    if cx:
                        want = (float(cx) / 360000.0) * g["scale"]   # EMU → cm → px
                        factor = max(1, int(round(img.width() / max(want, 1))))
                    else:
                        factor = max(1, int(round(img.width() / (usable * 0.6))))
                    if factor > 1:
                        img = img.subsample(factor, factor)
                    keep_images.append(img)
                    canvas.create_image(g["left"], y + before, anchor="nw", image=img)
                    y += before + img.height() + after
                except Exception:
                    y += before + size_px + after
            else:
                y += before + size_px + after
            continue

        size_px, height, before, after, indent, left = _line_metrics(line, g["scale"])
        text = _display_text(line)
        y += before
        if text.strip():
            align = (line.para.align or "left").lower()
            if align in ("center",):
                x, anchor = g["width"] / 2, "n"
            elif align in ("right",):
                x, anchor = g["width"] - g["right"] - left, "ne"
            else:
                x, anchor = g["left"] + left + indent, "nw"
            canvas.create_text(x, y + (height - size_px) / 2, text=text, anchor=anchor,
                               font=(_family(line.font.east_asia or line.font.ascii),
                                     -max(6, int(round(size_px)))),
                               fill="#000000", width=max(40, int(usable - left - indent)))
        y += height + after

    if y > g["height"]:
        canvas.configure(height=int(y + 4))
    return keep_images
