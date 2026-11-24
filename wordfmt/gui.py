# -*- coding: utf-8 -*-
"""图形界面：选文档 → 选方案 → 一键排版；也能从模板提取格式方案。

双击「启动.bat」即可打开，不需要命令行。
"""
from __future__ import annotations

import ctypes
import os
import queue
import sys
import threading
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import paths

paths.ensure_dirs()

from . import wordio  # noqa: E402
from .apply import apply_profile  # noqa: E402
from .cover import FIELD_NAMES as FIELD_CN  # noqa: E402
from .extract import build_profile  # noqa: E402
from .profile import Profile  # noqa: E402
from .verify import verify  # noqa: E402

APP_TITLE = "Word 格式排版助手"
DEFAULT_TEMPLATE = r"C:\Users\sycamore\DeskBox\安装包与压缩文件\论文格式\论文样板.docx"

# 封面套用方式：显示文字 → 内部模式
COVER_MODES = {
    "只改格式（文字不动）": "format",
    "重建并自动填（缺的留空）": "rebuild",
    "重建但全部留空（空表）": "blank",
}


class Tooltip:
    """鼠标停在「?」上时，弹一句解释；不占版面。"""

    def __init__(self, widget, text, wraplength=320):
        self.widget = widget
        self.text = text
        self.wraplength = wraplength
        self.tip = None
        widget.bind("<Enter>", self._show, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _show(self, _e=None):
        if self.tip is not None:
            return
        x = self.widget.winfo_rootx() + 12
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry("+%d+%d" % (x, y))
        tk.Label(self.tip, text=self.text, justify="left", background="#ffffe0",
                 relief="solid", borderwidth=1, wraplength=self.wraplength,
                 font=("Microsoft YaHei UI", 9), padx=8, pady=6).pack()

    def _hide(self, _e=None):
        if self.tip is not None:
            self.tip.destroy()
            self.tip = None


class ScrollFrame(ttk.Frame):
    """竖向可滚动的页签容器：内容再长也能滚到底，不会被窗口切掉。"""

    def __init__(self, parent):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, highlightthickness=0, bg=self._bg())
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.hsb = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")
        self.hsb.pack(side="bottom", fill="x")
        self.inner = ttk.Frame(self.canvas, padding=12)
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", self._on_inner)
        self.canvas.bind("<Configure>", self._on_canvas)
        self.canvas.bind("<Enter>", self._bind_wheel)
        self.canvas.bind("<Leave>", self._unbind_wheel)

    def _bg(self):
        try:
            return ttk.Style().lookup("TFrame", "background") or "#f0f0f0"
        except Exception:
            return "#f0f0f0"

    def _on_inner(self, _e=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas(self, e):
        # 内容比视口宽时保持原宽（由横向滚动条处理），否则铺满
        w = max(e.width, self.inner.winfo_reqwidth())
        self.canvas.itemconfigure(self._win, width=w)
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _bind_wheel(self, _e=None):
        self.canvas.bind_all("<MouseWheel>", self._on_wheel)

    def _unbind_wheel(self, _e=None):
        self.canvas.unbind_all("<MouseWheel>")

    def _on_wheel(self, e):
        self.canvas.yview_scroll(int(-e.delta / 120), "units")


def _enable_dpi_awareness():
    """让 Windows 按真实分辨率渲染，避免高分屏下字被拉糊。"""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)   # 每显示器 DPI 感知
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


class App(tk.Tk):
    def __init__(self):
        _enable_dpi_awareness()
        super().__init__()
        self._tune_fonts()
        self.title(APP_TITLE)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w = min(1240, max(980, int(sw * 0.72)))
        h = min(940, max(700, int(sh * 0.82)))
        self.geometry("%dx%d+%d+%d" % (w, h, max(0, (sw - w) // 2), max(0, (sh - h) // 3)))
        self.minsize(900, 640)
        self.q = queue.Queue()
        self.busy = False
        self._build_style()
        self._build_ui()
        self.after(120, self._poll)
        self.log("欢迎使用。第 1 步：在「排版文档」页选好文件；第 2 步：点「开始排版」。")
        self.log("第一次用某个模板，先去「提取模板格式」把模板里的格式读出来。")

    # ---------------------------------------- 外观
    def _tune_fonts(self):
        """统一界面字体：中文用微软雅黑 UI，等宽（日志/代码）用 Consolas。"""
        import tkinter.font as tkfont

        family = "Microsoft YaHei UI"
        pairs = [("TkDefaultFont", family, 10), ("TkTextFont", family, 10),
                 ("TkHeadingFont", family, 10), ("TkMenuFont", family, 10),
                 ("TkCaptionFont", family, 9), ("TkSmallCaptionFont", family, 9),
                 ("TkIconFont", family, 9), ("TkTooltipFont", family, 9),
                 ("TkFixedFont", "Consolas", 10)]
        for name, fam, size in pairs:
            try:
                tkfont.nametofont(name).configure(family=fam, size=size)
            except Exception:
                pass

    def _build_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except Exception:
            pass
        style.configure(".", font=("Microsoft YaHei UI", 10))
        # 表格：字号放大后默认行高会切字，这里显式给字体和行高
        style.configure("Treeview", font=("Microsoft YaHei UI", 10), rowheight=28)
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 10, "bold"))
        style.map("Treeview", background=[("selected", "#cfe3f7")],
                  foreground=[("selected", "#1a1a1a")])
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 14, "bold"))
        style.configure("Hint.TLabel", foreground="#666666")
        style.configure("Go.TButton", font=("Microsoft YaHei UI", 11, "bold"), padding=(16, 8))

    # ---------------------------------------- 界面
    def _build_ui(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=(10, 4))

        self.tab_apply = ttk.Frame(nb, padding=12)
        self.tab_lib = ttk.Frame(nb, padding=12)
        self.tab_fmt = ttk.Frame(nb, padding=12)
        self.tab_chart = ScrollFrame(nb)
        self.tab_lit = ScrollFrame(nb)
        self.tab_extract = ttk.Frame(nb, padding=12)
        self.tab_help = ttk.Frame(nb, padding=12)
        nb.add(self.tab_apply, text="  排版文档  ")
        nb.add(self.tab_lib, text="  封面库  ")
        nb.add(self.tab_fmt, text="  格式预览  ")
        nb.add(self.tab_chart, text="  图表  ")
        nb.add(self.tab_lit, text="  文献引用  ")
        nb.add(self.tab_extract, text="  提取模板格式  ")
        nb.add(self.tab_help, text="  帮助  ")
        self.nb = nb
        self._build_apply(self.tab_apply)
        self._build_library(self.tab_lib)
        self._build_format_preview(self.tab_fmt)
        self._build_chart(self.tab_chart.inner)
        self._build_lit(self.tab_lit.inner)
        self._build_extract(self.tab_extract)
        self._build_help(self.tab_help)
        try:
            self._refresh_profiles()
            self._refresh_covers()
            self._refresh_coverlib()
        except Exception:
            pass
        if os.environ.get("WF_START_TAB", "").lower() == "cover":
            self.nb.select(self.tab_lib)

        # ---- 功能说明条：随页签自动切换 ----
        self.helpbar = ttk.LabelFrame(self, text="功能说明（这一页是做什么的）", padding=(10, 4))
        self.helpbar.pack(fill="x", padx=10, pady=(0, 4))
        self.helpbar_text = tk.Text(self.helpbar, height=2, wrap="word",
                                    font=("Microsoft YaHei UI", 10), bg="#f7f9fc",
                                    relief="flat", borderwidth=0, padx=6, pady=2)
        self.helpbar_text.pack(fill="x", expand=True)
        self.helpbar_text.configure(state="disabled")
        row = ttk.Frame(self.helpbar)
        row.pack(fill="x")
        self.helpbar_row = row
        self.helpbar_open = True
        self.helpbar_btn = ttk.Button(row, text="收起说明", width=10, command=self._toggle_helpbar)
        self.helpbar_btn.pack(side="right")
        nb.bind("<<NotebookTabChanged>>", lambda e: self._update_helpbar())
        self._update_helpbar()

        bottom = ttk.Frame(self, padding=(10, 0, 10, 8))
        bottom.pack(fill="both", expand=False)
        ttk.Label(bottom, text="运行日志", style="Hint.TLabel").pack(anchor="w")
        self.logbox = tk.Text(bottom, height=4, wrap="word", font=("Consolas", 10),
                              bg="#1e1e1e", fg="#dcdcdc", insertbackground="#dcdcdc")
        self.logbox.pack(fill="both", expand=True)
        self.logbox.configure(state="disabled")
        self.status = tk.StringVar(value="就绪")
        self.status_label = ttk.Label(self, textvariable=self.status, relief="sunken", anchor="w",
                                      padding=(8, 3))
        self.status_label.pack(fill="x", side="bottom")

        # 重新按「从下往上」排一遍：先放底部固定高度的，最后让 Notebook 占满剩下的空间，
        # 否则 Notebook 会把功能说明条和日志区挤成 1 像素。
        for w in (self.nb, self.helpbar, bottom):
            w.pack_forget()
        self.status_label.pack_forget()
        self.status_label.pack(fill="x", side="bottom")
        bottom.pack(fill="x", side="bottom", padx=10, pady=(0, 6))
        self.helpbar.pack(fill="x", side="bottom", padx=10, pady=(0, 4))
        self.nb.pack(fill="both", expand=True, padx=10, pady=(10, 4))

    # ---------------------------------------- 功能说明条
    TAB_HELP = {
        "排版文档": (
            "把格式乱掉的论文，按选定方案重新排一遍。\n"
            "① 选要排版的文档 → ② 选格式方案（论文格式.yaml） → ③ 选输出位置 → ④ 需要的话选封面。\n"
            "点「开始排版」后，程序会：重设样式 → 逐段套用 → 清掉手工格式 → 更新目录页码 → 逐段校验。\n"
            "原文件不会被改，结果另存到「输出」文件夹；同一份文档反复跑也只留一份备份。"),
        "封面库": (
            "看封面、换封面。\n"
            "左边点一张封面，右边立刻画出样子（校徽、标题、填空线都能看到，不用开 Word）。\n"
            "点「用它排版 →」会跳回排版页并选好这张封面。\n"
            "封面上的横线是真的下划线：可以直接在线上打字，字会落在线上，线不会断开。"),
        "格式预览": (
            "看方案本身长什么样，不用真的排版。\n"
            "用一段示例文字把题目、摘要、一二三级标题、正文、图题、表题、参考文献各画一遍。\n"
            "字号大小、首行缩进、行距、居中还是两端对齐，一眼就能比较出来。\n"
            "改完 yaml 点「重新绘制」即可；排版页换方案时这一页也会跟着换。"),
        "图表": (
            "输入数据自动画图（12 种：柱状/分组柱/堆叠柱/折线/面积/饼/环形/散点/箱线/水平条形/热力图/雷达）。\n"
            "① 左边粘贴数据（第一行表头，Tab 或逗号分隔）或导入 CSV；\n"
            "② 选图类型、填标题与轴名、选配色（也可自定义颜色，热力图还能选色带）；\n"
            "③ 点「画图预览」右边出图；满意后填图题和小节名，点「插入到文档」，图号会自动接续。"),
        "文献引用": (
            "联网找真实论文 → 在正文句末插上标 [n] → 自动补参考文献表（不需要 EndNote）。\n"
            "① 选文档 → ② 填检索主题（英文最准，如 extrusion cooking snack） → ③ 设引用篇数/样式 → "
            "④ 点「检索并插入引用」，结果表格里能看每处插在哪、配了哪篇。\n"
            "参考文献格式内置 8 种、也支持自定义模板；还能「导出文献分析」看年份/期刊分布。\n"
            "详细图文步骤见程序目录下的《参考文献引用功能说明.md》。"),
        "提取模板格式": (
            "拿一份排好的范文（比如学校的论文样板），提取出它的格式，生成一份 .yaml 方案。\n"
            "以后排版别的文档时选这份方案，就能一篇篇套成同样的样子。\n"
            "提取时会顺便把封面存进封面库；识别不准的部分可以在 yaml 里手改 match_styles / match_regex。"),
        "帮助": (
            "完整的文字说明：使用步骤、yaml 怎么改、常见问题、命令行用法。\n"
            "新手建议先看「排版文档」和「文献引用」两节。"),
    }

    def _toggle_helpbar(self):
        self.helpbar_open = not getattr(self, "helpbar_open", True)
        if self.helpbar_open:
            self.helpbar_text.pack(fill="x", expand=True, before=self.helpbar_row)
            self.helpbar_btn.configure(text="收起说明")
        else:
            self.helpbar_text.pack_forget()
            self.helpbar_btn.configure(text="展开说明")

    def _update_helpbar(self):
        try:
            current = self.nb.tab(self.nb.select(), "text").strip()
        except Exception:
            current = "排版文档"
        text = self.TAB_HELP.get(current, "")
        if not text:
            return
        self.helpbar.configure(text="功能说明 · %s" % current)
        self.helpbar_text.configure(state="normal")
        self.helpbar_text.delete("1.0", "end")
        self.helpbar_text.insert("1.0", text)
        self.helpbar_text.configure(state="disabled")

    def _build_apply(self, f):
        ttk.Label(f, text="把一份格式混乱的文档，按选定的格式方案重新排好",
                  style="Title.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))

        ttk.Label(f, text="① 要排版的文档").grid(row=1, column=0, sticky="w", pady=6)
        self.in_path = tk.StringVar()
        ttk.Entry(f, textvariable=self.in_path, width=62).grid(row=1, column=1, sticky="we", padx=6)
        ttk.Button(f, text="浏览…", command=self._pick_input).grid(row=1, column=2)

        ttk.Label(f, text="② 格式方案（YAML）").grid(row=2, column=0, sticky="w", pady=6)
        self.profile_path = tk.StringVar()
        self.profile_box = ttk.Combobox(f, textvariable=self.profile_path, width=60)
        self.profile_box.grid(row=2, column=1, sticky="we", padx=6)
        self.profile_box.bind("<<ComboboxSelected>>", lambda e: self._show_profile())
        ttk.Button(f, text="浏览…", command=self._pick_profile).grid(row=2, column=2)

        ttk.Label(f, text="③ 输出到").grid(row=3, column=0, sticky="w", pady=6)
        self.out_path = tk.StringVar()
        ttk.Entry(f, textvariable=self.out_path, width=62).grid(row=3, column=1, sticky="we", padx=6)
        ttk.Button(f, text="另存为…", command=self._pick_output).grid(row=3, column=2)

        ttk.Label(f, text="④ 封面（可选）").grid(row=4, column=0, sticky="w", pady=6)
        self.cover_choice = tk.StringVar(value="不套封面")
        self.cover_box = ttk.Combobox(f, textvariable=self.cover_choice, width=60)
        self.cover_box.grid(row=4, column=1, sticky="we", padx=6)
        self.cover_box.bind("<<ComboboxSelected>>", lambda e: self._sync_lib_selection())
        ttk.Button(f, text="浏览…", command=self._pick_cover).grid(row=4, column=2)

        opt = ttk.LabelFrame(f, text="选项", padding=10)
        opt.grid(row=5, column=0, columnspan=3, sticky="we", pady=(12, 8))
        self.opt_backup = tk.BooleanVar(value=True)
        self.opt_word = tk.BooleanVar(value=True)
        self.opt_pdf = tk.BooleanVar(value=False)
        self.opt_verify = tk.BooleanVar(value=True)
        self.cover_mode_choice = tk.StringVar(value=list(COVER_MODES.keys())[0])
        ttk.Checkbutton(opt, text="先备份原文件", variable=self.opt_backup).grid(row=0, column=0, sticky="w", padx=(0, 18))
        ttk.Checkbutton(opt, text="用 Word 自动更新目录/页码", variable=self.opt_word).grid(row=0, column=1, sticky="w", padx=(0, 18))
        ttk.Checkbutton(opt, text="同时导出一份 PDF", variable=self.opt_pdf).grid(row=0, column=2, sticky="w")
        ttk.Checkbutton(opt, text="改完自动校验", variable=self.opt_verify).grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.opt_sections = tk.BooleanVar(value=False)
        chk_sec = ttk.Checkbutton(opt, text="补齐缺失章节（参考文献/致谢/附录）",
                                  variable=self.opt_sections)
        chk_sec.grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Label(opt, text="封面怎么套：").grid(row=1, column=1, sticky="e", pady=(6, 0))
        ttk.Combobox(opt, textvariable=self.cover_mode_choice, values=list(COVER_MODES.keys()),
                     state="readonly", width=32).grid(row=1, column=2, sticky="w", pady=(6, 0))

        btns = ttk.Frame(f)
        btns.grid(row=6, column=0, columnspan=3, sticky="we", pady=8)
        self.btn_apply = ttk.Button(btns, text="开始排版", style="Go.TButton", command=self._do_apply)
        self.btn_apply.pack(side="left")
        self.btn_newdoc = ttk.Button(btns, text="新建空白论文（已排版）",
                                     command=self._do_newdoc)
        self.btn_newdoc.pack(side="left", padx=(10, 0))
        q_new = ttk.Label(btns, text="?", foreground="#2E5C8A", cursor="question_arrow")
        q_new.pack(side="left", padx=(4, 0))
        Tooltip(q_new, "不用先有稿子：按当前「格式方案」直接生成一篇**排好格式的空白论文**。\n"
                       "结构：摘要 + 关键词 → 目录 → 1 引言 / 2 文献综述 / 3 研究方法 /\n"
                       "4 结果与分析 / 5 讨论 / 6 结论与展望 → 致谢 → 附录 → 参考文献。\n\n"
                       "生成后直接在里面写内容即可；写完不用再排（格式已经是好的）。")
        ttk.Button(btns, text="打开输出文件夹", command=lambda: self._open_dir(paths.OUTPUT_DIR)).pack(side="left", padx=10)
        ttk.Button(btns, text="打开方案文件夹", command=lambda: self._open_dir(paths.CONFIG_DIR)).pack(side="left")
        ttk.Button(btns, text="打开封面库", command=lambda: self._open_dir(os.path.join(paths.BASE, "covers"))).pack(side="left", padx=10)

        self.result = tk.Text(f, height=10, wrap="none", font=("Consolas", 9))
        self.result.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        sb = ttk.Scrollbar(f, orient="vertical", command=self.result.yview)
        sb.grid(row=7, column=3, sticky="ns", pady=(8, 0))
        self.result.configure(yscrollcommand=sb.set)
        f.columnconfigure(1, weight=1)
        f.rowconfigure(7, weight=1)

    def _build_library(self, f):
        ttk.Label(f, text="封面库：选中左边任何一个，右边就能看到它长什么样",
                  style="Title.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        left = ttk.Frame(f)
        left.grid(row=1, column=0, sticky="nsew", padx=(0, 12))
        cols = ("name", "lines", "fields")
        self.lib_tree = ttk.Treeview(left, columns=cols, show="headings", height=14)
        for c, t, w in (("name", "封面", 320), ("lines", "行数", 70), ("fields", "可填", 180)):
            self.lib_tree.heading(c, text=t)
            self.lib_tree.column(c, width=w, anchor="w")
        self.lib_tree.pack(fill="both", expand=True)
        self.lib_tree.bind("<<TreeviewSelect>>", lambda e: self._on_lib_select())
        self.lib_tree.bind("<Double-1>", lambda e: self._use_selected_cover())

        right = ttk.Frame(f)
        right.grid(row=1, column=1, sticky="nsew")
        self.lib_preview = tk.Canvas(right, width=300, height=420, bg="white",
                                     highlightthickness=1, highlightbackground="#cccccc")
        self.lib_preview.pack(side="top", anchor="n", pady=(0, 6))
        self._lib_imgs = []
        self.lib_info = tk.StringVar(value="")
        ttk.Label(right, textvariable=self.lib_info, style="Hint.TLabel",
                  wraplength=360).pack(side="top", anchor="w")

        bar = ttk.Frame(f)
        bar.grid(row=2, column=0, columnspan=2, sticky="we", pady=(10, 0))
        ttk.Button(bar, text="刷新", command=self._refresh_coverlib).pack(side="left")
        ttk.Button(bar, text="重新绘制预览", command=self._redraw_preview).pack(side="left", padx=8)
        ttk.Button(bar, text="用它排版 →", style="Go.TButton",
                   command=self._use_selected_cover).pack(side="left", padx=8)
        ttk.Button(bar, text="打开封面文件夹",
                   command=lambda: self._open_dir(os.path.join(paths.BASE, "covers"))).pack(side="left")
        ttk.Button(bar, text="删除这个封面", command=self._delete_selected_cover).pack(side="right")

        f.columnconfigure(0, weight=0)
        f.columnconfigure(1, weight=1)
        f.rowconfigure(1, weight=1)

    # ---------------------------------------- 封面库逻辑
    def _refresh_coverlib(self):
        from .cover import Cover, list_covers
        self._lib_covers = {}
        self.lib_tree.delete(*self.lib_tree.get_children())
        for name, path in list_covers(paths.BASE):
            try:
                cov = Cover.load(path)
                fields = []
                for line in cov.lines:
                    for f in line.needed_fields():
                        if f not in fields:
                            fields.append(f)
                self._lib_covers[path] = cov
                self.lib_tree.insert("", "end", iid=path,
                                     values=(name, len(cov.lines),
                                             "%d 项" % len(fields) if fields else "—"))
            except Exception as e:
                self.lib_tree.insert("", "end", iid=path, values=(name, "?", "读取失败：%s" % e))
        children = self.lib_tree.get_children()
        if children:
            self.lib_tree.selection_set(children[0])
            self._on_lib_select()

    def _selected_cover_path(self) -> str:
        sel = self.lib_tree.selection()
        return sel[0] if sel else ""

    def _on_lib_select(self):
        path = self._selected_cover_path()
        if not path:
            return
        cov = self._lib_covers.get(path)
        if cov is None:
            return
        fields = []
        pics = 0
        for line in cov.lines:
            if line.is_picture:
                pics += 1
            for f in line.needed_fields():
                if f not in fields:
                    fields.append(f)
        self.lib_info.set("%s\n%d 行，%d 处图片；可填：%s" %
                          (cov.name, len(cov.lines), pics,
                           "、".join(FIELD_CN.get(f, f) for f in fields) or "无"))
        self._draw_preview(cov)

    def _draw_preview(self, cov=None):
        if cov is None:
            path = self._selected_cover_path()
            cov = self._lib_covers.get(path)
        if cov is None:
            return
        try:
            from .preview import draw_cover
            self._lib_imgs = draw_cover(self.lib_preview, cov, width_px=300)
        except Exception as e:
            self.lib_preview.delete("all")
            self.lib_preview.create_text(10, 10, anchor="nw", text="预览绘制失败：%s" % e)

    def _redraw_preview(self):
        self._draw_preview()

    def _sync_lib_selection(self):
        """排版页选了哪个封面，封面库页就跳到哪个，两页是连着的。"""
        path = self.cover_paths.get(self.cover_choice.get(), "")
        if not path or path not in self._lib_covers:
            return
        self.lib_tree.selection_set(path)
        self.lib_tree.see(path)
        self._on_lib_select()

    def _use_selected_cover(self):
        path = self._selected_cover_path()
        cov = self._lib_covers.get(path)
        if cov is None:
            return
        if cov.name not in self.cover_paths:
            self.cover_paths[cov.name] = path
            self.cover_box["values"] = list(self.cover_paths.keys())
        self.cover_choice.set(cov.name)
        self.nb.select(self.tab_apply)
        self.log("已选用封面：%s，到「排版文档」页点开始排版即可。" % cov.name)

    def _delete_selected_cover(self):
        path = self._selected_cover_path()
        cov = self._lib_covers.get(path)
        if not path or cov is None:
            return
        if not messagebox.askyesno(APP_TITLE, "确定删除封面「%s」吗？（连同它的图片）" % cov.name):
            return
        try:
            img_dir = os.path.join(os.path.dirname(path),
                                   os.path.splitext(os.path.basename(path))[0] + "_images")
            os.remove(path)
            if os.path.isdir(img_dir):
                import shutil
                shutil.rmtree(img_dir, ignore_errors=True)
            from .preview import clear_thumbs
            clear_thumbs(path)
        except Exception as e:
            messagebox.showwarning(APP_TITLE, "删除失败：%s" % e)
        self._refresh_coverlib()
        self._refresh_covers()

    def _build_format_preview(self, f):
        ttk.Label(f, text="格式预览：看这套方案排出来是什么样（标题多大、正文多宽、行距多少）",
                  style="Title.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        left = ttk.Frame(f)
        left.grid(row=1, column=0, sticky="nsew", padx=(0, 12))
        ttk.Label(left, text="格式方案").pack(anchor="w")
        self.fp_choice = tk.StringVar()
        self.fp_box = ttk.Combobox(left, textvariable=self.fp_choice, width=34, state="readonly")
        self.fp_box.pack(anchor="w", pady=(2, 8))
        self.fp_box.bind("<<ComboboxSelected>>", lambda e: self._draw_format_preview())
        self.fp_info = tk.StringVar(value="")
        ttk.Label(left, textvariable=self.fp_info, style="Hint.TLabel",
                  wraplength=240, justify="left").pack(anchor="w")
        ttk.Button(left, text="打开方案文件夹",
                   command=lambda: self._open_dir(paths.CONFIG_DIR)).pack(anchor="w", pady=(12, 0))
        ttk.Button(left, text="重新绘制", command=self._draw_format_preview).pack(anchor="w", pady=(6, 0))

        right = ttk.Frame(f)
        right.grid(row=1, column=1, sticky="nsew")
        self.fp_canvas = tk.Canvas(right, width=320, height=460, bg="white",
                                   highlightthickness=1, highlightbackground="#cccccc")
        self.fp_canvas.pack(side="top", anchor="n")
        f.columnconfigure(1, weight=1)
        f.rowconfigure(1, weight=1)

    def _draw_format_preview(self):
        path = self.fp_choice.get()
        if not path or not os.path.exists(path):
            return
        try:
            prof = Profile.load(path)
        except Exception as e:
            self.fp_info.set("读取失败：%s" % e)
            return
        self.fp_info.set("%s\n%d 个部分\n页面 %.1f×%.1fcm\n边距 上%.2f 下%.2f 左%.2f 右%.2f" % (
            prof.name, len(prof.roles), prof.page.width_cm, prof.page.height_cm,
            prof.page.margin_top_cm, prof.page.margin_bottom_cm,
            prof.page.margin_left_cm, prof.page.margin_right_cm))
        try:
            from .preview import draw_format_preview
            self._fp_imgs = draw_format_preview(self.fp_canvas, prof, width_px=320)
        except Exception as e:
            self.fp_canvas.delete("all")
            self.fp_canvas.create_text(10, 10, anchor="nw", text="绘制失败：%s" % e)

    # ---------------------------------------- 图表
    def _build_chart(self, f):
        from . import charts

        ttk.Label(f, text="把数据粘进来，选一种图，点「画图预览」，满意了再插进文档",
                  style="Title.TLabel").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 8))

        left = ttk.Frame(f)
        left.grid(row=1, column=0, sticky="nsew", padx=(0, 12))
        lbl_data = ttk.Label(left, text="① 数据 (?)")
        lbl_data.grid(row=0, column=0, sticky="w", pady=(0, 6))
        Tooltip(lbl_data, "三种来源任选：\n· 直接粘贴（Tab 或逗号分隔）；\n· 「从 Excel 读取…」选 .xlsx，右边下拉选工作表；\n· 「从 CSV 读入…」选 .csv/.txt。\n第一列是横轴类别，第二列起每列一条数据。")
        self.chart_data = tk.Text(left, width=42, height=4, wrap="none", font=("Consolas", 10))
        self.chart_data.grid(row=1, column=0, sticky="we", pady=4)
        self.chart_data.insert("1.0", "月份\t销量\t返修率\n1月\t120\t2.1\n2月\t135\t1.8\n3月\t150\t1.5\n4月\t142\t2.0")
        read_row = ttk.Frame(left)
        read_row.grid(row=2, column=0, sticky="w")
        ttk.Button(read_row, text="从 Excel 读取…", command=self._chart_from_excel).pack(side="left")
        self.chart_sheet = tk.StringVar()
        self.chart_sheet_box = ttk.Combobox(read_row, textvariable=self.chart_sheet,
                                            width=12, state="readonly")
        self.chart_sheet_box.pack(side="left", padx=(6, 0))
        self.chart_sheet_box.bind("<<ComboboxSelected>>", lambda e: self._chart_load_sheet())
        ttk.Button(read_row, text="从 CSV 读入…", command=self._chart_from_csv).pack(side="left", padx=(6, 0))
        self._chart_xlsx = ""

        box = ttk.LabelFrame(f, text="图形设置", padding=6)
        box.grid(row=2, column=0, sticky="we", pady=(8, 0), padx=(0, 12))
        self.chart_kind = tk.StringVar(value="柱状图")
        self.chart_title = tk.StringVar()
        self.chart_xlabel = tk.StringVar()
        self.chart_ylabel = tk.StringVar()
        ttk.Label(box, text="图类型").grid(row=0, column=0, sticky="e")
        ttk.Combobox(box, textvariable=self.chart_kind, values=list(charts.CHART_TYPES),
                     state="readonly", width=12).grid(row=0, column=1, sticky="w", padx=(4, 12))
        ttk.Label(box, text="标题").grid(row=0, column=2, sticky="e")
        ttk.Entry(box, textvariable=self.chart_title, width=16).grid(row=0, column=3, sticky="w", padx=4)
        self.chart_palette = tk.StringVar(value="学术蓝")
        self.chart_colors = tk.StringVar()
        self.chart_cmap = tk.StringVar(value="YlGnBu")
        lbl_color = ttk.Label(box, text="配色 (?)")
        lbl_color.grid(row=2, column=0, sticky="e", pady=(6, 0))
        Tooltip(lbl_color, "现成配色：学术蓝/暖色/绿色/灰度/彩虹；\n选「自定义」后，在右边填颜色代码，如 #1E6F5C,#E67E22。\n热力图另外用「热力图色带」控制深浅色。")
        ttk.Combobox(box, textvariable=self.chart_palette,
                     values=list(charts.PALETTES) + ["自定义"],
                     state="readonly", width=10).grid(row=2, column=1, sticky="w", padx=(4, 12), pady=(6, 0))
        ttk.Entry(box, textvariable=self.chart_colors, width=20).grid(
            row=2, column=3, sticky="w", padx=4, pady=(6, 0))
        ttk.Label(box, text="热力图色带").grid(row=3, column=0, sticky="e", pady=(6, 0))
        ttk.Combobox(box, textvariable=self.chart_cmap, values=charts.CMAPS,
                     state="readonly", width=10).grid(row=3, column=1, sticky="w", padx=(4, 12), pady=(6, 0))

        ttk.Label(box, text="横轴").grid(row=1, column=0, sticky="e", pady=(6, 0))
        ttk.Entry(box, textvariable=self.chart_xlabel, width=10).grid(row=1, column=1, sticky="w", padx=(4, 12), pady=(6, 0))
        ttk.Label(box, text="纵轴").grid(row=1, column=2, sticky="e", pady=(6, 0))
        ttk.Entry(box, textvariable=self.chart_ylabel, width=16).grid(row=1, column=3, sticky="w", padx=4, pady=(6, 0))

        ttk.Button(left, text="画图预览", style="Go.TButton", command=self._chart_preview).grid(
            row=4, column=0, sticky="w", pady=8)

        right = ttk.Frame(f)
        right.grid(row=1, column=1, sticky="nsew")
        ttk.Label(right, text="② 预览").grid(row=0, column=0, sticky="w")
        self.chart_preview_lbl = ttk.Label(right, text="（还没有图）", anchor="center",
                                           relief="solid", borderwidth=1, width=38, padding=(2, 4))
        self.chart_preview_lbl.grid(row=1, column=0, sticky="nsew", pady=4)
        right.rowconfigure(1, weight=1)

        ins = ttk.LabelFrame(f, text="③ 插进文档（可选）", padding=6)
        ins.grid(row=2, column=1, sticky="nsew", pady=(8, 0))
        self.chart_doc = tk.StringVar()
        self.chart_caption = tk.StringVar()
        self.chart_anchor = tk.StringVar()
        ttk.Label(ins, text="文档").grid(row=0, column=0, sticky="e")
        ttk.Entry(ins, textvariable=self.chart_doc, width=16).grid(row=0, column=1, sticky="w", padx=4)
        ttk.Button(ins, text="浏览…", width=6,
                   command=self._chart_pick_doc).grid(row=0, column=2, padx=(4, 0))
        ttk.Label(ins, text="图题").grid(row=1, column=0, sticky="e", pady=(8, 0))
        ttk.Entry(ins, textvariable=self.chart_caption, width=16).grid(row=1, column=1, sticky="w", padx=4, pady=(6, 0))
        q1 = ttk.Label(ins, text="?", foreground="#2E5C8A", cursor="question_arrow")
        q1.grid(row=1, column=2, sticky="w", padx=(6, 0), pady=(6, 0))
        Tooltip(q1, "图题 = 图片下面那一行说明文字。只写内容部分，例如「各月产量变化」；\n"
                    "程序插入时会自动加上「图1 」，并接着文档里已有的图号往后编。")
        ttk.Label(ins, text="插在哪").grid(row=2, column=0, sticky="e", pady=(8, 0))
        ttk.Entry(ins, textvariable=self.chart_anchor, width=16).grid(row=2, column=1, sticky="w", padx=4, pady=(6, 0))
        q2 = ttk.Label(ins, text="?", foreground="#2E5C8A", cursor="question_arrow")
        q2.grid(row=2, column=2, sticky="w", padx=(6, 0), pady=(6, 0))
        Tooltip(q2, "插在哪 = 图要放到哪个小节的末尾。填标题里的几个字即可，\n"
                    "例如「3.1 工艺流程选择」；程序会找到这一节，把图和图题放在本节最后。")
        ttk.Button(ins, text="插入到文档", command=self._chart_insert).grid(row=3, column=1, sticky="w", pady=(8, 0))

        f.rowconfigure(1, weight=1)
        f.columnconfigure(0, weight=1)
        f.columnconfigure(1, weight=1)

    def _chart_from_csv(self):
        from . import charts

        p = filedialog.askopenfilename(title="选择 CSV / 文本数据",
                                       filetypes=[("数据文件", "*.csv *.txt"), ("所有文件", "*.*")])
        if not p:
            return
        try:
            import io as _io
            with _io.open(p, "r", encoding="utf-8-sig", errors="replace") as fh:
                text = fh.read()
        except Exception as e:
            messagebox.showerror(APP_TITLE, "读不了文件：%s" % e)
            return
        self.chart_data.delete("1.0", "end")
        self.chart_data.insert("1.0", text)
        self.log("已读入数据：%s" % p)

    def _chart_from_excel(self):
        from . import charts

        p = filedialog.askopenfilename(title="选择 Excel 文件",
                                       filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if not p:
            return
        sheets = charts.excel_sheets(p)
        if not sheets:
            messagebox.showerror(APP_TITLE, "这个 Excel 读不出来（只支持 .xlsx / .xlsm）")
            return
        self._chart_xlsx = p
        self.chart_sheet_box.configure(values=sheets)
        self.chart_sheet.set(sheets[0])
        self._chart_load_sheet()
        self.log("已读取 Excel：%s（工作表：%s）" % (p, "、".join(sheets)))

    def _chart_load_sheet(self):
        from . import charts

        path = getattr(self, "_chart_xlsx", "")
        if not path:
            return
        try:
            header, rows = charts.read_excel(path, self.chart_sheet.get() or None)
        except Exception as e:
            messagebox.showerror(APP_TITLE, "读取失败：%s" % e)
            return
        if not rows:
            messagebox.showwarning(APP_TITLE, "这张表里没读到数据（第一行要是表头）")
            return
        lines = ["\t".join(header)]
        for r in rows[:200]:
            lines.append("\t".join(r))
        self.chart_data.delete("1.0", "end")
        self.chart_data.insert("1.0", "\n".join(lines))
        self.log("已载入 %d 行数据（第一行是表头）" % len(rows))

    def _chart_render(self):
        from . import charts

        text = self.chart_data.get("1.0", "end")
        header, rows = charts.parse_table(text)
        if not rows:
            messagebox.showwarning(APP_TITLE, "数据读不出来：第一行是表头，后面每行一条数据，列之间用 Tab 或逗号隔开")
            return None
        kind = charts.CHART_TYPES.get(self.chart_kind.get(), "bar")
        out_dir = os.path.join(paths.OUTPUT_DIR, "图表")
        os.makedirs(out_dir, exist_ok=True)
        name = "%s_%s.png" % (self.chart_kind.get(), datetime.now().strftime("%H%M%S"))
        out = os.path.join(out_dir, name)
        charts.render(header, rows, kind, out, {
            "title": self.chart_title.get(), "xlabel": self.chart_xlabel.get(),
            "ylabel": self.chart_ylabel.get(), "dpi": 300,
            "palette": self.chart_palette.get(),
            "colors": self.chart_colors.get().strip(),
            "cmap": self.chart_cmap.get(),
        })
        return out

    def _chart_preview(self):
        out = self._chart_render()
        if not out:
            return
        self._chart_png = out
        try:
            self.chart_preview_img = tk.PhotoImage(file=out)
            w = self.chart_preview_img.width()
            h = self.chart_preview_img.height()
            factor = max(1, int(max(w / 420.0, h / 300.0) + 0.5))
            if factor > 1:
                self.chart_preview_img = self.chart_preview_img.subsample(factor, factor)
            self.chart_preview_lbl.configure(image=self.chart_preview_img, text="")
        except Exception as e:
            self.chart_preview_lbl.configure(text="预览失败：%s" % e)
        self.log("已画好图：%s" % out)

    def _chart_pick_doc(self):
        p = filedialog.askopenfilename(title="选择要插图的文档",
                                       filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")])
        if p:
            self.chart_doc.set(p)

    def _chart_insert(self):
        from docx import Document
        from . import charts

        out = getattr(self, "_chart_png", None) or self._chart_render()
        doc_path = self.chart_doc.get().strip()
        if not doc_path or not os.path.exists(doc_path):
            messagebox.showwarning(APP_TITLE, "先选一个要插图的 Word 文档")
            return
        if not out or not os.path.exists(out):
            messagebox.showwarning(APP_TITLE, "先点「画图预览」把图画出来")
            return
        prof_path = self.profile_path.get()
        profile = None
        if prof_path and os.path.exists(prof_path):
            try:
                profile = Profile.load(prof_path)
            except Exception:
                profile = None
        try:
            doc = Document(doc_path)
            pic, cap, no, _t = charts.insert_figure(
                doc, out, self.chart_caption.get().strip() or None, profile,
                anchor_text=None, after_section=self.chart_anchor.get().strip() or None)
        except ValueError as e:
            messagebox.showwarning(APP_TITLE, str(e))
            return
        stem = os.path.splitext(os.path.basename(doc_path))[0]
        target = os.path.join(paths.OUTPUT_DIR, "%s_插图.docx" % stem)
        try:
            doc.save(target)
        except PermissionError:
            messagebox.showwarning(APP_TITLE, "文档被占用了（可能在 Word 里开着），关掉再试")
            return
        self.log("已插入 %s → %s" % (cap.text, target))
        messagebox.showinfo(APP_TITLE, "已插入：%s\n文档：%s" % (cap.text, target))

    # ---------------------------------------- 文献引用
    def _build_lit(self, f):
        from . import litcite

        ttk.Label(f, text="联网检索真实论文 → 句末插引用 → 自动补参考文献表（不需要 EndNote）",
                  style="Title.TLabel").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))

        ttk.Label(f, text="① 要加引用的文档").grid(row=1, column=0, sticky="w", pady=4)
        self.cite_in = tk.StringVar()
        ttk.Entry(f, textvariable=self.cite_in, width=40).grid(row=1, column=1, columnspan=2, sticky="we", padx=6)
        ttk.Button(f, text="浏览…", command=self._pick_cite_input).grid(row=1, column=3)

        lbl_topic = ttk.Label(f, text="② 检索主题 (?)")
        lbl_topic.grid(row=2, column=0, sticky="w", pady=6)
        Tooltip(lbl_topic, "决定去学术库里搜什么词。\n· 填英文命中率最高，如 extrusion cooking snack food；\n· 留空则程序从你的句子里自动挑词，但常常不够准。\n搜不准就换个英文主题再点一次。")
        self.cite_topic = tk.StringVar()
        ttk.Entry(f, textvariable=self.cite_topic, width=40).grid(row=2, column=1, columnspan=3, sticky="we", padx=6)
        ttk.Label(f, text="例：extrusion cooking cereal snack food",
                  foreground="#666").grid(row=3, column=1, columnspan=3, sticky="w", padx=6)

        box = ttk.LabelFrame(f, text="引用设置", padding=10)
        box.grid(row=4, column=0, columnspan=4, sticky="we", pady=(10, 6))
        self.cite_count = tk.IntVar(value=2)
        self.cite_per_para = tk.IntVar(value=1)
        self.cite_year = tk.StringVar(value="2010")
        self.cite_style = tk.StringVar(value=list(litcite.CITE_STYLES)[0])
        self.cite_doi = tk.BooleanVar(value=True)
        self.cite_body = tk.BooleanVar(value=True)
        self.cite_skip_first = tk.BooleanVar(value=True)

        ttk.Label(box, text="引用几篇：").grid(row=0, column=0, sticky="e", pady=(0, 4))
        ttk.Spinbox(box, from_=1, to=10, width=5, textvariable=self.cite_count).grid(row=0, column=1, sticky="w", padx=(0, 16))
        ttk.Label(box, text="每段最多几处：").grid(row=0, column=2, sticky="e")
        ttk.Spinbox(box, from_=1, to=3, width=5, textvariable=self.cite_per_para).grid(row=0, column=3, sticky="w", padx=(0, 16))
        ttk.Label(box, text="最早年份：").grid(row=0, column=4, sticky="e")
        ttk.Entry(box, textvariable=self.cite_year, width=7).grid(row=0, column=5, sticky="w")

        ttk.Label(box, text="引用样式：").grid(row=1, column=0, sticky="e", pady=(10, 0))
        ttk.Combobox(box, textvariable=self.cite_style, values=list(litcite.CITE_STYLES),
                     state="readonly", width=22).grid(row=1, column=1, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Checkbutton(box, text="文献必须有 DOI", variable=self.cite_doi).grid(row=1, column=3, sticky="w", pady=(8, 0))
        ttk.Checkbutton(box, text="只插在正文", variable=self.cite_body).grid(row=1, column=4, sticky="w", pady=(8, 0))
        ttk.Checkbutton(box, text="跳过每节首段", variable=self.cite_skip_first).grid(row=1, column=5, sticky="w", pady=(8, 0))

        btns = ttk.Frame(f)
        btns.grid(row=5, column=0, columnspan=4, sticky="we", pady=8)
        self.btn_cite = ttk.Button(btns, text="检索并插入引用", style="Go.TButton", command=self._do_cite)
        self.btn_cite.pack(side="left")
        ttk.Button(btns, text="打开文献库文件夹",
                   command=lambda: self._open_dir(os.path.join(paths.OUTPUT_DIR, "文献库"))).pack(side="left", padx=10)

        cols = ("no", "where", "title", "year", "doi")
        self.cite_tree = ttk.Treeview(f, columns=cols, show="headings", height=6)
        for c, t, w in (("no", "序号", 45), ("where", "插在哪里", 210),
                        ("title", "文献标题", 300), ("year", "年份", 55), ("doi", "DOI", 170)):
            self.cite_tree.heading(c, text=t)
            self.cite_tree.column(c, width=w, anchor="w")
        self.cite_tree.grid(row=6, column=0, columnspan=4, sticky="nsew", pady=(6, 0))
        f.rowconfigure(6, weight=1)
        f.columnconfigure(1, weight=1)
        ris_box = ttk.LabelFrame(f, text="参考文献表", padding=8)
        ris_box.grid(row=7, column=0, columnspan=4, sticky="we", pady=(8, 0))
        self.ris_path = tk.StringVar()
        self.ris_doc = tk.StringVar()
        lbl_ris = ttk.Label(ris_box, text="RIS 文件 (?)")
        lbl_ris.grid(row=0, column=0, sticky="e")
        Tooltip(lbl_ris, "只有你想用 EndNote 时才需要：\nEndNote 里文件 → 导出（输出样式选 RefMan RIS）得到 .ris，\n在这里选它就能排成参考文献表。不用 EndNote 可留空。")
        ttk.Entry(ris_box, textvariable=self.ris_path, width=34).grid(row=0, column=1, sticky="w", padx=4)
        ttk.Button(ris_box, text="浏览…", command=self._pick_ris).grid(row=0, column=2)
        ttk.Button(ris_box, text="怎么引知网文章？", command=self._show_cnki_help).grid(
            row=0, column=3, sticky="w", padx=(6, 0))
        lbl_doc = ttk.Label(ris_box, text="目标文档 (?)")
        lbl_doc.grid(row=1, column=0, sticky="e", pady=(6, 0))
        Tooltip(lbl_doc, "参考文献表要插进哪个 Word 文档。\n建议选已经排好版的稿子，程序会把新条目接到\n「参考文献」最后一条后面，编号自动接续。")
        ttk.Entry(ris_box, textvariable=self.ris_doc, width=34).grid(row=1, column=1, sticky="w", padx=4, pady=(6, 0))
        ttk.Button(ris_box, text="浏览…", command=self._pick_ris_doc).grid(row=1, column=2, pady=(6, 0))
        ttk.Button(ris_box, text="生成参考文献表并插入", command=self._do_refs).grid(
            row=2, column=1, sticky="w", pady=(8, 0))
        ttk.Button(ris_box, text="用刚才检索到的文献生成（不用 EndNote）",
                   command=self._do_refs_from_papers).grid(row=3, column=1, sticky="w", pady=(6, 0))

        from . import citestyle

        lbl_fmt = ttk.Label(ris_box, text="参考文献格式 (?)")
        lbl_fmt.grid(row=4, column=0, sticky="e", pady=(8, 0))
        Tooltip(lbl_fmt, "内置 9 种体例（GB/T 7714、APA、MLA、Chicago、IEEE、\nVancouver、Harvard、Nature、ACS）；\n选中自定义格式时，下面的模板框会带出它的写法，\n可以直接改，再点「保存」。")
        self.ref_style = tk.StringVar(value="GB/T 7714（顺序编码）")
        self.ref_style_box = ttk.Combobox(ris_box, textvariable=self.ref_style,
                                          values=self._style_names(),
                                          state="readonly", width=22)
        self.ref_style_box.grid(row=4, column=1, sticky="w", padx=4, pady=(8, 0))
        self.ref_style_box.bind("<<ComboboxSelected>>", lambda e: self._on_style_change())
        ttk.Button(ris_box, text="导入 .ens 样式",
                   command=self._import_ens).grid(row=4, column=2, sticky="w", padx=(4, 0), pady=(8, 0))
        style_btns = ttk.Frame(ris_box)
        style_btns.grid(row=5, column=1, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Button(style_btns, text="新建…", width=6, command=self._style_new).pack(side="left")
        ttk.Button(style_btns, text="重命名…", width=8, command=self._style_rename).pack(side="left", padx=2)
        ttk.Button(style_btns, text="删除", width=5, command=self._style_delete).pack(side="left", padx=2)
        ttk.Button(style_btns, text="保存", width=5, command=self._style_save).pack(side="left", padx=2)
        ttk.Label(ris_box, text="自定义模板").grid(
            row=6, column=0, sticky="e", pady=(6, 0))
        self.ref_template = tk.StringVar()
        ttk.Entry(ris_box, textvariable=self.ref_template, width=40).grid(
            row=6, column=1, columnspan=2, sticky="w", padx=4, pady=(6, 0))
        tool_row = ttk.Frame(ris_box)
        tool_row.grid(row=7, column=1, columnspan=2, sticky="w", pady=(8, 0))
        self.ref_doi = tk.BooleanVar(value=False)
        ttk.Checkbutton(ris_box, text="条目里带 DOI（一般不用勾）", variable=self.ref_doi).grid(
            row=5, column=2, sticky="w", padx=(6, 0), pady=(6, 0))
        ttk.Button(tool_row, text="导出文献分析",
                   command=self._do_analyze).pack(side="left")
        q_an = ttk.Label(tool_row, text="?", foreground="#2E5C8A", cursor="question_arrow")
        q_an.pack(side="left", padx=(5, 18))
        Tooltip(q_an, "生成一份「文献分析报告」，输出到 输出\\文献库：\n"
                      "· 文献数量、年份跨度\n"
                      "· 有 DOI 的比例、能免费下载的比例、被引合计\n"
                      "· 年份分布（文字条形图 + 一张柱状图 PNG）\n"
                      "· 来源期刊前 10、高频作者前 10\n"
                      "· 完整清单\n\n"
                      "用途：写「研究现状 / 文献综述」时用数据说话，\n"
                      "也能自检这批文献是不是太旧、来源够不够硬。")
        ttk.Button(tool_row, text="重排引用编号",
                   command=self._do_renumber).pack(side="left")
        q_rn = ttk.Label(tool_row, text="?", foreground="#2E5C8A", cursor="question_arrow")
        q_rn.pack(side="left", padx=(5, 0))
        Tooltip(q_rn, "按正文出现顺序重新编号：第一次出现的引用是 [1]、第二次是 [2]……\n"
                      "参考文献条目跟着重排、编号一起改；正文没引用到的排在最后。\n\n"
                      "什么时候用：删了某处引用、编号断号，或想统一成“按出现顺序”。")

        ttk.Label(f, text="检索结果来自 OpenAlex（免费公开学术库），RIS 文件导出到「输出/文献库」，"
                          "在 EndNote 里 File → Import 选中它即可入库。",
                  foreground="#666", wraplength=980).grid(row=8, column=0, columnspan=4,
                                                          sticky="w", pady=(6, 0))
        self._attach_lit_tips()

    def _attach_lit_tips(self):
        """给「文献引用」页的每个控件挂上中文解释（按钮悬停即看，标签加 (?) 标记）。"""
        label_tips = {
            "① 要加引用的文档": "要插引用的论文。选好后程序**不会改原文件**，结果另存到「输出」文件夹。",
            "引用几篇：": "一共插几处引用（一处配一篇文献）。一般正文引 5～15 处；这里填 2 就插 2 处。",
            "每段最多几处：": "同一段落里最多插几处。填 1 表示一段只插一处，避免连着堆。",
            "最早年份：": "只找这一年之后发表的文献。想要经典老文献就改小（如 2000）。",
            "引用样式：": "正文里记号的样子：\n· 顺序编码-上标[n]：特点。⁽¹⁾  ← 多数理工科论文\n"
                        "· 顺序编码-普通[n]：特点。[1]\n· 上标(n)：特点。⁽¹⁾\n"
                        "· 作者-年份：（Tiwari, 2017）← 人文社科常用",
            "文献必须有 DOI": "勾上后只保留有 DOI 的文献。DOI 是论文的“身份证”，能保证可核查，\n建议一直勾着。",
            "只插在正文": "勾上后只在正文段落插引用，标题、摘要、关键词、图表题、参考文献都不碰。",
            "跳过每节首段": "每个大标题下的第一段通常只是引入语，勾上就跳过，不硬塞引用。",
            "自定义模板": "选「自定义模板（临时）」时用这个输入框写格式，占位符：\n"
                       "{seq} {authors} {title} {journal} {year} {volume} {issue} {pages} {doi} {type}\n"
                       "空字段会自动清掉多余标点。",
        }
        button_tips = {
            "检索并插入引用": "联网去 OpenAlex 找真实论文，在正文句末插上标 [n]，并在文末补参考文献条目。\n"
                          "可能要等几秒到十几秒。搜不到就换「② 检索主题」再试。",
            "打开文献库文件夹": "打开「输出\\文献库」，里面有每次检索生成的 .ris 和文献清单 .md。",
            "生成参考文献表并插入": "把你选的 RIS 文件排成 GB/T 7714 等体例的参考文献表，接到「参考文献」最后一条后面。",
            "用刚才检索到的文献生成（不用 EndNote）": "用上一步检索到的文献直接排参考文献表，全程不碰 EndNote。\n"
                                              "适合只想要文末列表、不想在正文插上标的情况。",
            "导入 .ens 样式": "读 EndNote 的样式文件（.ens），取出样式名和说明，自动挑最接近的内置格式。\n"
                            "要 100% 一致得复制 Bibliography → Templates 里的模板，用「自定义模板」。",
            "新建…": "把当前模板存成一个有名字的格式，以后在下拉里直接选。",
            "重命名…": "改自定义格式的名字（内置格式改不了，可以「新建…」另存一份）。",
            "删除": "删掉当前选中的自定义格式（内置格式删不掉）。",
            "保存": "把输入框里改过的模板存回当前格式；如果当前是内置格式，会提示你另存成新名字。",
            "导出文献分析": "生成一份文献分析报告：文献数量、年份跨度、有 DOI 比例、年份分布图、\n"
                        "来源期刊前 10、高频作者前 10、完整清单，另配一张年份分布柱状图。\n"
                        "适合写“研究现状”那一节时用数据说话。",
            "重排引用编号": "按正文出现顺序重新编号（[8][9] 可能变成 [1][2]），参考文献表跟着重排。\n"
                        "增删过引用、编号断号时点一下即可；正文没引用到的条目会排在最后。",
            "浏览…": "选文件。",
        }
        tree_tip = ("结果表格：\n· 序号：正文里会显示的编号\n· 插在哪里：被引用那句话的开头\n"
                    "· 文献标题 / 年份 / DOI：匹配到的真实论文，DOI 可到 doi.org 打开原文")

        def walk(w):
            for child in w.winfo_children():
                yield child
                for sub in walk(child):
                    yield sub

        for w in walk(self.tab_lit.inner):
            cls = w.winfo_class()
            try:
                text = str(w.cget("text"))
            except Exception:
                text = ""
            if cls in ("TButton",) and text in button_tips:
                Tooltip(w, button_tips[text])
            elif cls in ("TCheckbutton",) and text in label_tips:
                Tooltip(w, label_tips[text])
            elif cls in ("TLabel",) and text in label_tips:
                try:
                    w.configure(text=text + " (?)")
                except Exception:
                    pass
                Tooltip(w, label_tips[text])
        try:
            Tooltip(self.cite_tree, tree_tip, wraplength=360)
        except Exception:
            pass

    def _style_names(self):
        from . import citestyle

        return citestyle.all_names() + ["自定义模板（临时）"]

    def _refresh_style_box(self):
        try:
            self.ref_style_box.configure(values=self._style_names())
        except Exception:
            pass

    def _on_style_change(self):
        """选中自定义格式时，把它的模板显示到下面输入框里。"""
        from . import citestyle

        name = self.ref_style.get()
        custom = citestyle.load_custom()
        if name in custom:
            self.ref_template.set(custom[name].get("template", ""))

    def _style_new(self):
        from tkinter import simpledialog

        from . import citestyle

        name = simpledialog.askstring(APP_TITLE, "新格式叫什么名字？\n（例如：我的学校格式）", parent=self)
        if not name:
            return
        tpl = self.ref_template.get().strip() or citestyle.CUSTOM_DEFAULT_TEMPLATE
        try:
            citestyle.add_custom(name, tpl)
        except Exception as e:
            messagebox.showwarning(APP_TITLE, str(e))
            return
        self._refresh_style_box()
        self.ref_style.set(name)
        self.log("已新建参考文献格式：%s" % name)

    def _style_rename(self):
        from tkinter import simpledialog

        from . import citestyle

        old = self.ref_style.get()
        if old not in citestyle.load_custom():
            messagebox.showinfo(APP_TITLE, "只有自己新建的格式才能改名；内置格式请用「新建…」另存一份。")
            return
        new = simpledialog.askstring(APP_TITLE, "改成什么名字？", initialvalue=old, parent=self)
        if not new or new == old:
            return
        try:
            citestyle.rename_custom(old, new)
        except Exception as e:
            messagebox.showwarning(APP_TITLE, str(e))
            return
        self._refresh_style_box()
        self.ref_style.set(new)
        self.log("参考文献格式已改名：%s → %s" % (old, new))

    def _style_delete(self):
        from . import citestyle

        name = self.ref_style.get()
        if name not in citestyle.load_custom():
            messagebox.showinfo(APP_TITLE, "内置格式删不掉，只能删自己新建的。")
            return
        if not messagebox.askyesno(APP_TITLE, "确定删除格式「%s」？" % name):
            return
        citestyle.delete_custom(name)
        self._refresh_style_box()
        self.ref_style.set("GB/T 7714（顺序编码）")
        self.log("已删除参考文献格式：%s" % name)

    def _style_save(self):
        from tkinter import simpledialog

        from . import citestyle

        name = self.ref_style.get()
        tpl = self.ref_template.get().strip()
        if not tpl:
            messagebox.showwarning(APP_TITLE, "先把模板写进下面的输入框")
            return
        custom = citestyle.load_custom()
        if name not in custom:
            new_name = simpledialog.askstring(
                APP_TITLE, "这个模板存成什么名字？\n（内置格式不能直接改，存成新名字）",
                initialvalue="我的格式", parent=self)
            if not new_name:
                return
            name = new_name
        try:
            citestyle.add_custom(name, tpl)
        except Exception as e:
            messagebox.showwarning(APP_TITLE, str(e))
            return
        self._refresh_style_box()
        self.ref_style.set(name)
        self.log("已保存参考文献格式：%s" % name)

    def _style_args(self):
        name = self.ref_style.get()
        if name == "自定义模板（临时）":
            return None, self.ref_template.get().strip() or None
        return name, None

    def _import_ens(self):
        from . import citestyle

        p = filedialog.askopenfilename(
            title="选择 EndNote 输出的样式文件（.ens）",
            initialdir=r"D:\Styles" if os.path.isdir(r"D:\Styles") else "",
            filetypes=[("EndNote 样式", "*.ens"), ("所有文件", "*.*")])
        if not p:
            return
        try:
            info = citestyle.read_ens(p)
        except Exception as e:
            messagebox.showerror(APP_TITLE, "读不了这个样式文件：%s" % e)
            return
        guess = ""
        low = (info["name"] + " " + (info.get("description") or "")).lower()
        table = [("apa", "APA 7th"), ("ieee", "IEEE"), ("mla", "MLA 9th"),
                 ("chicago", "Chicago（作者-年份）"), ("vancouver", "Vancouver"),
                 ("harvard", "Harvard"), ("nature", "Nature"), ("gb", "GB/T 7714（顺序编码）")]
        for key, name in table:
            if key in low:
                guess = name
                break
        if guess:
            self.ref_style.set(guess)
        msg = "样式名：%s\n\n说明：%s\n\n" % (info["name"], info.get("description") or "（没有读到说明）")
        if guess:
            msg += "已自动选用内置格式：%s\n（内置格式是按常见体例写的，不完全等于这个样式）" % guess
        else:
            msg += ("没有识别出对应体例。\n\n"
                    "要复制的是这一条：\n"
                    "EndNote 菜单 编辑 → 输出样式 → 编辑「样式名」\n"
                    "  → 左栏点 **Bibliography（参考文献）→ Templates**\n"
                    "  → 右上方「参考文献类型」保持 **Journal Article（期刊论文）**\n"
                    "  → 中间那个大文本框里的整行文字，全选复制\n\n"
                    "⚠ 别复制 Citation → Templates 那一栏，那是正文里 [1] 的写法，不是文末条目。\n\n"
                    "拿到后：参考文献格式选「自定义模板（临时）」，粘到下面输入框，\n"
                    "把 Author / Title / Journal / Year / Volume / Pages 依次换成\n"
                    "{authors} {title} {journal} {year} {volume} {pages} 即可。")
        messagebox.showinfo(APP_TITLE, msg)

    def _do_analyze(self):
        from . import analysis

        papers = getattr(self, "_last_papers", None)
        if not papers:
            messagebox.showwarning(APP_TITLE, "还没有检索结果：先点上面的「检索并插入引用」")
            return
        folder = os.path.join(paths.OUTPUT_DIR, "文献库")
        try:
            md, png = analysis.export(papers, folder, "文献分析")
        except Exception as e:
            messagebox.showerror(APP_TITLE, "导出分析失败：%s" % e)
            return
        self.log("文献分析已导出：%s" % md)
        messagebox.showinfo(APP_TITLE, "已导出：\n%s\n%s" % (md, png or ""))

    def _do_renumber(self):
        """按正文出现顺序重排 [n]，参考文献表跟着排。"""
        from docx import Document

        from . import citeindex

        doc_path = self.ris_doc.get().strip() or self.cite_in.get().strip()
        if not doc_path or not os.path.exists(doc_path):
            messagebox.showwarning(APP_TITLE, "先选一个要重排的文档（下面「目标文档」或上面「①」都行）")
            return
        try:
            doc = Document(doc_path)
        except Exception as e:
            messagebox.showerror(APP_TITLE, "打不开文档：%s" % e)
            return
        problems = citeindex.check(doc)
        if not problems:
            messagebox.showinfo(APP_TITLE, "编号已经是规范的：正文按出现顺序，参考文献表也对得上。")
            return
        try:
            rep = citeindex.renumber(doc)
        except Exception as e:
            messagebox.showerror(APP_TITLE, "重排失败：%s" % e)
            return
        stem = os.path.splitext(os.path.basename(doc_path))[0]
        out = os.path.join(paths.OUTPUT_DIR, "%s_已重排.docx" % stem)
        try:
            doc.save(out)
        except PermissionError:
            messagebox.showwarning(APP_TITLE, "文档被占用了（可能在 Word 里开着），关掉再试")
            return
        pairs = ["[%d]→[%d]" % (o, n) for o, n in sorted(rep["map"].items()) if o != n]
        self.log("重排引用：正文 %d 处、参考文献 %d 条；%s" % (
            rep["marks"], rep["moved"], "，".join(pairs[:8])))
        msg = "已重排：正文 %d 处，参考文献 %d 条。\n\n%s\n\n文档：%s" % (
            rep["marks"], rep["moved"],
            "，".join(pairs[:10]) + ("…" if len(pairs) > 10 else ""), out)
        if rep["uncited"]:
            msg += "\n\n注意：正文没引用到的条目排在最后：%s" % "、".join(str(x) for x in rep["uncited"])
        messagebox.showinfo(APP_TITLE, msg)

    CNKI_HELP = """引用知网（CNKI）的文章：三条路，任选一条
========================================

【路 A · 推荐】知网导出 RIS → 程序排参考文献表
----------------------------------------
1. 知网打开这篇文献的详情页；
2. 点「导出 / 参考文献」（有的版本写「引用」）；
3. 格式选 RIS；
   （如果只有 EndNote 选项，也行，导出的 .enw 多数能读）
4. 下载得到一个 .txt 或 .ris 文件；
5. 回到本程序：
   · 「RIS 文件」→ 浏览… 选刚下载的文件；
   · 「目标文档」→ 选你要插参考文献的论文；
   · 点「生成参考文献表并插入」。

程序会自动识别编码（知网常是 GBK），中文不会乱码；
知网条目常常没有 DOI，缺了也没关系，那一段会自动省掉。

正文里的上标要自己插（打字输入 [1] 即可），插完点「重排引用编号」
让编号和文末顺序对齐。


【路 B】经 EndNote 转一道（装了 EndNote 才需要）
----------------------------------------
1. 知网导出选「EndNote」，得到一个 .enw 文件；
2. 打开 EndNote：文件 → 导入，选这个 .enw；
3. EndNote 里再 文件 → 导出，输出样式选 RefMan (RIS) Export；
4. 拿这个 .ris 回本程序，按路 A 的第 5 步做。

为什么要转一道？程序不直接读 .enw（字段标记和 RIS 略有差别），
转换成 RIS 最稳。


【路 C】完全手动（只引一两篇时最快）
----------------------------------------
1. 在论文末尾的参考文献里，照体例加一行，例如：
   [1] 张三, 李四. 题名[J]. 期刊名, 2023, 44(5): 12-18.
2. 在正文对应句子后面打上标 [1]；
3. 回到本程序点「重排引用编号」——按正文出现顺序统一编号，
   并同步文末条目。


【为什么不用程序的联网检索找知网文章？】
----------------------------------------
程序查的是 OpenAlex 和 Crossref 两个「开放」学术库，
对知网独家收录的中文期刊覆盖不全（常缺摘要、缺 DOI，相关性判断会打折）。

所以：
  · 中文文献（知网、万方）→ 用路 A 导出 RIS；
  · 英文文献 → 用程序的「检索并插入引用」最省事。


小提示
----------------------------------------
· 程序不会改你的原文件，结果都另存到「输出」文件夹；
· 知网可以勾选多篇一起导出，一个 RIS 里就有好几条；
· 想换参考文献体例（GB/T 7714、APA、ACS…）→ 用「参考文献格式」下拉。
"""

    def _show_cnki_help(self):
        """把「怎么引知网文章」直接写在程序里，不用去翻 md 文件。"""
        win = tk.Toplevel(self)
        win.title("怎么引用知网（CNKI）的文章")
        win.geometry("780x640")
        win.transient(self)
        ttk.Label(win, text="三条路任选一条，推荐第一条",
                  font=("Microsoft YaHei UI", 12, "bold")).pack(anchor="w", padx=14, pady=(12, 4))
        frame = ttk.Frame(win, padding=(14, 0, 14, 8))
        frame.pack(fill="both", expand=True)
        txt = tk.Text(frame, wrap="word", font=("Microsoft YaHei UI", 10),
                      bg="#fbfcfe", relief="solid", borderwidth=1, padx=10, pady=8)
        sb = ttk.Scrollbar(frame, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        txt.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        txt.insert("1.0", self.CNKI_HELP)
        txt.configure(state="disabled")
        btns = ttk.Frame(win, padding=(14, 0, 14, 12))
        btns.pack(fill="x")
        ttk.Button(btns, text="打开文献库文件夹",
                   command=lambda: self._open_dir(os.path.join(paths.OUTPUT_DIR, "文献库"))).pack(side="left")
        ttk.Button(btns, text="关闭", command=win.destroy).pack(side="right")
        win.bind("<Escape>", lambda e: win.destroy())

    def _do_refs_from_papers(self):
        """把上次检索到的文献直接排成参考文献表，全程不碰 EndNote。"""
        from . import refs

        papers = getattr(self, "_last_papers", None)
        if not papers:
            messagebox.showwarning(APP_TITLE, "还没有检索结果：先点上面的「检索并插入引用」")
            return
        doc = self.ris_doc.get().strip()
        if not doc or not os.path.exists(doc):
            messagebox.showwarning(APP_TITLE, "先在下面选一个目标文档")
            return
        prof_path = self.profile_path.get()
        if not prof_path or not os.path.exists(prof_path):
            messagebox.showwarning(APP_TITLE, "先在「排版文档」页选一个格式方案")
            return
        stem = os.path.splitext(os.path.basename(doc))[0]
        out = os.path.join(paths.OUTPUT_DIR, "%s_参考文献.docx" % stem)
        try:
            prof = Profile.load(prof_path)
            style, template = self._style_args()
            res = refs.from_papers_to_doc(papers, doc, prof, out, style=style, template=template,
                                          include_doi=bool(self.ref_doi.get()))
        except Exception as e:
            messagebox.showerror(APP_TITLE, "生成失败：%s" % e)
            return
        self.log("参考文献表已生成：%d 条 → %s" % (len(papers), out))
        messagebox.showinfo(APP_TITLE, "共 %d 条，已写入：\n%s" % (len(papers), out))

    def _pick_ris(self):
        p = filedialog.askopenfilename(title="选择 EndNote 导出的 RIS",
                                       filetypes=[("RIS 文献", "*.ris *.txt"), ("所有文件", "*.*")])
        if p:
            self.ris_path.set(p)

    def _pick_ris_doc(self):
        p = filedialog.askopenfilename(title="选择要插入参考文献表的文档",
                                       filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")])
        if p:
            self.ris_doc.set(p)

    def _do_refs(self):
        from . import refs

        ris = self.ris_path.get().strip()
        doc = self.ris_doc.get().strip()
        if not ris or not os.path.exists(ris):
            messagebox.showwarning(APP_TITLE, "先选一个 RIS 文件")
            return
        if not doc or not os.path.exists(doc):
            messagebox.showwarning(APP_TITLE, "先选一个目标文档")
            return
        prof_path = self.profile_path.get()
        if not prof_path or not os.path.exists(prof_path):
            messagebox.showwarning(APP_TITLE, "先在「排版文档」页选一个格式方案")
            return
        stem = os.path.splitext(os.path.basename(doc))[0]
        out = os.path.join(paths.OUTPUT_DIR, "%s_参考文献.docx" % stem)
        try:
            prof = Profile.load(prof_path)
            style, template = self._style_args()
            res = refs.from_ris_to_doc(ris, doc, prof, out, style=style, template=template,
                                       include_doi=bool(self.ref_doi.get()))
        except Exception as e:
            messagebox.showerror(APP_TITLE, "生成失败：%s" % e)
            return
        self.log("参考文献表已生成：%d 条 → %s" % (res["count"], out))
        messagebox.showinfo(APP_TITLE, "共 %d 条，已写入：\n%s" % (res["count"], out))

    def _pick_cite_input(self):
        p = filedialog.askopenfilename(title="选择要加引用的文档",
                                       filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")])
        if p:
            self.cite_in.set(p)

    def _do_cite(self):
        from . import litcite

        src = self.cite_in.get().strip()
        if not src or not os.path.exists(src):
            messagebox.showwarning(APP_TITLE, "先选一个要加引用的文档")
            return
        prof_path = self.profile_path.get()
        if not prof_path or not os.path.exists(prof_path):
            messagebox.showwarning(APP_TITLE, "先在「排版文档」页选一个格式方案")
            return
        conf = {
            "count": int(self.cite_count.get() or 2),
            "max_per_para": int(self.cite_per_para.get() or 1),
            "style": self.cite_style.get(),
            "year_from": int(self.cite_year.get() or 0) or None,
            "require_doi": bool(self.cite_doi.get()),
            "only_body": bool(self.cite_body.get()),
            "skip_first_para": bool(self.cite_skip_first.get()),
            "topics": self.cite_topic.get().strip(),
        }
        stem = os.path.splitext(os.path.basename(src))[0]
        out = os.path.join(paths.OUTPUT_DIR, "%s_已引用.docx" % stem)
        self.btn_cite.configure(state="disabled")
        self.config(cursor="watch")
        self.update()
        try:
            prof = Profile.load(prof_path)
            res = litcite.cite_document(src, out, prof, conf, progress=self.log)
        except Exception as e:
            messagebox.showerror(APP_TITLE, "检索/插入失败：%s" % e)
            self.log("文献引用失败：%s" % e)
            return
        finally:
            self.btn_cite.configure(state="normal")
            self.config(cursor="")
        for item in self.cite_tree.get_children():
            self.cite_tree.delete(item)
        self._last_papers = list(res["papers"])
        for para, end, n, sent, paper in res["picked"]:
            self.cite_tree.insert("", "end", values=(
                n, sent[:40], (paper.title or "")[:60], paper.year or "", paper.doi or ""))
        if res["papers"]:
            self.log("已插入 %d 处引用，文献写入：%s" % (len(res["papers"]), res["ris"]))
            self.log("已生成：%s" % out)
            messagebox.showinfo(APP_TITLE, "插入 %d 处引用\n\n文献库：%s\n文档：%s" %
                                (len(res["papers"]), res["ris"], out))
        else:
            self.log("没找到合适文献，可以填一个检索主题再试（例如 extrusion cooking snack food）")
            messagebox.showwarning(APP_TITLE, "没找到合适的文献。\n可以填一个英文检索主题再试一次。")

    def _build_extract(self, f):
        ttk.Label(f, text="从一份「格式正确」的模板文档里，自动读出每个部分的格式",
                  style="Title.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))

        ttk.Label(f, text="① 模板文档").grid(row=1, column=0, sticky="w", pady=6)
        self.tex_path = tk.StringVar(value=DEFAULT_TEMPLATE if os.path.exists(DEFAULT_TEMPLATE) else "")
        ttk.Entry(f, textvariable=self.tex_path, width=62).grid(row=1, column=1, sticky="we", padx=6)
        ttk.Button(f, text="浏览…", command=self._pick_template).grid(row=1, column=2)

        ttk.Label(f, text="② 方案保存为").grid(row=2, column=0, sticky="w", pady=6)
        self.tex_out = tk.StringVar()
        ttk.Entry(f, textvariable=self.tex_out, width=62).grid(row=2, column=1, sticky="we", padx=6)
        ttk.Button(f, text="浏览…", command=self._pick_extract_out).grid(row=2, column=2)

        btn_row = ttk.Frame(f)
        btn_row.grid(row=3, column=0, columnspan=3, sticky="w", pady=12)
        ttk.Button(btn_row, text="提取格式方案", style="Go.TButton",
                   command=self._do_extract).pack(side="left")
        ttk.Button(btn_row, text="生成空白结构稿",
                   command=self._do_skeleton).pack(side="left", padx=(12, 0))
        q_sk = ttk.Label(btn_row, text="?", foreground="#2E5C8A", cursor="question_arrow")
        q_sk.pack(side="left", padx=(5, 0))
        Tooltip(q_sk, "按模板生成一份「只有骨架」的 Word：\n"
                      "· 留下：摘要 / Abstract / 目录 / 各级标题 / 参考文献 / 致谢 等章节；\n"
                      "· 删掉：所有正文、表格、图片（内容留空）。\n"
                      "· 标题层级按 1 / 1.1 / 1.1.1 自动规整；每个标题下补一个空段，直接就能打字。\n\n"
                      "用法：拿到结构稿先写大纲和内容，写完再回「排版文档」页排一遍。")
        self.skel_template = tk.BooleanVar(value=False)
        ttk.Checkbutton(btn_row, text="沿用模板的章节标题", variable=self.skel_template).pack(
            side="left", padx=(16, 0))
        ttk.Label(f, text="提取结果可以直接在「方案文件夹」里用记事本修改：字体、字号、行距、"
                          "各部分怎么识别（样式名/正则）都在里面。",
                  style="Hint.TLabel", wraplength=820).grid(row=4, column=0, columnspan=3, sticky="w")

        cols = ("role", "style", "fmt", "how")
        self.tree = ttk.Treeview(f, columns=cols, show="headings", height=14)
        for c, t, w in (("role", "部分", 130), ("style", "套用样式", 150),
                        ("fmt", "格式", 410), ("how", "怎么识别", 380)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="w", minwidth=90)
        self.tree.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(10, 0))
        sb = ttk.Scrollbar(f, orient="vertical", command=self.tree.yview)
        sb.grid(row=5, column=3, sticky="ns", pady=(10, 0))
        hsb = ttk.Scrollbar(f, orient="horizontal", command=self.tree.xview)
        hsb.grid(row=6, column=0, columnspan=3, sticky="we")
        self.tree.configure(yscrollcommand=sb.set, xscrollcommand=hsb.set)
        f.columnconfigure(1, weight=1)
        f.rowconfigure(5, weight=1)

    def _build_help(self, f):
        txt = tk.Text(f, wrap="word", font=("Microsoft YaHei UI", 10), padx=8, pady=8)
        txt.pack(fill="both", expand=True)
        txt.insert("1.0", HELP_TEXT)
        txt.configure(state="disabled")

    # ---------------------------------------- 选文件
    def _pick_input(self):
        p = filedialog.askopenfilename(title="选择要排版的 Word 文档",
                                       filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")])
        if p:
            self.in_path.set(p)
            stem, ext = os.path.splitext(p)
            self.out_path.set(os.path.join(paths.OUTPUT_DIR, os.path.basename(stem) + "_已排版" + (ext or ".docx")))
            self._refresh_profiles()

    def _pick_profile(self):
        p = filedialog.askopenfilename(title="选择格式方案", initialdir=paths.CONFIG_DIR,
                                       filetypes=[("格式方案", "*.yaml *.yml"), ("所有文件", "*.*")])
        if p:
            self.profile_path.set(p)
            self._show_profile()

    def _pick_output(self):
        p = filedialog.asksaveasfilename(title="输出为", initialdir=paths.OUTPUT_DIR,
                                          defaultextension=".docx",
                                          filetypes=[("Word 文档", "*.docx")])
        if p:
            self.out_path.set(p)

    def _pick_template(self):
        p = filedialog.askopenfilename(title="选择模板文档",
                                       filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")])
        if p:
            self.tex_path.set(p)
            stem = os.path.splitext(os.path.basename(p))[0]
            self.tex_out.set(os.path.join(paths.CONFIG_DIR, stem + ".yaml"))

    def _do_skeleton(self):
        """按模板生成空白结构稿（只留章节骨架）。"""
        from . import skeleton

        src = self.tex_in.get().strip()
        if not src or not os.path.exists(src):
            messagebox.showwarning(APP_TITLE, "先在「① 模板文档」选一份排好的范文")
            return
        prof = None
        prof_path = self.profile_path.get()
        if prof_path and os.path.exists(prof_path):
            try:
                prof = Profile.load(prof_path)
            except Exception:
                prof = None
        out = self.tex_out.get().strip()
        if not out:
            stem = os.path.splitext(os.path.basename(src))[0]
            out = os.path.join(paths.OUTPUT_DIR, "%s_结构稿.docx" % stem)
        else:
            out = os.path.splitext(out)[0] + "_结构稿.docx"
        try:
            if self.skel_template.get():
                path, st = skeleton.build_skeleton(src, out, prof)
                detail = ("沿用模板章节：保留 %d 段（标题 %d 个，层级规整 %d 个），删掉正文 %d 段。"
                          % (st["kept"], st["headings"], st["relabeled"], st["removed"]))
            else:
                path, st = skeleton.build_outline(src, out, prof)
                detail = ("标准大纲：%d 个一级章节（引言/文献综述/研究方法/结果与分析/讨论/结论与展望）"
                          "+ 摘要 + 关键词 + 目录 + 致谢 + 附录 + 参考文献。" % st["sections"])
        except Exception as e:
            messagebox.showerror(APP_TITLE, "生成失败：%s" % e)
            return
        self.log("已生成空白结构稿：%s（%s）" % (path, detail))
        messagebox.showinfo(APP_TITLE,
                            "已生成空白结构稿：\n%s\n\n%s\n\n"
                            "下一步：打开它，在每个标题下直接写内容；写完再回「排版文档」页排一次。"
                            % (path, detail))

    def _pick_extract_out(self):
        p = filedialog.asksaveasfilename(title="方案保存为", initialdir=paths.CONFIG_DIR,
                                          defaultextension=".yaml",
                                          filetypes=[("格式方案", "*.yaml")])
        if p:
            self.tex_out.set(p)

    def _open_dir(self, d):
        os.makedirs(d, exist_ok=True)
        os.startfile(d)

    # ---------------------------------------- 封面库
    def _refresh_covers(self):
        from .cover import list_covers
        self.cover_paths = {"不套封面": ""}
        for name, path in list_covers(paths.BASE):
            if name in self.cover_paths:
                continue
            self.cover_paths[name] = path
        self.cover_box["values"] = list(self.cover_paths.keys())
        if self.cover_choice.get() not in self.cover_paths:
            self.cover_choice.set("不套封面")

    def _pick_cover(self):
        d = os.path.join(paths.BASE, "covers")
        os.makedirs(d, exist_ok=True)
        p = filedialog.askopenfilename(title="选择封面库里的封面", initialdir=d,
                                       filetypes=[("封面", "*.yaml *.yml"), ("所有文件", "*.*")])
        if not p:
            return
        try:
            from .cover import Cover
            cov = Cover.load(p)
        except Exception as e:
            messagebox.showwarning(APP_TITLE, "封面文件读不出来：%s" % e)
            return
        self.cover_paths[cov.name] = p
        self.cover_box["values"] = list(self.cover_paths.keys())
        self.cover_choice.set(cov.name)

    def _refresh_profiles(self):
        try:
            files = sorted(f for f in os.listdir(paths.CONFIG_DIR) if f.lower().endswith((".yaml", ".yml")))
        except Exception:
            files = []
        values = [os.path.join(paths.CONFIG_DIR, f) for f in files]
        self.profile_box["values"] = values
        if hasattr(self, "fp_box"):
            self.fp_box["values"] = values
        if files and not self.profile_path.get():
            # 优先选带「论文」的方案
            pick = next((f for f in files if "论文" in f), files[0])
            self.profile_path.set(os.path.join(paths.CONFIG_DIR, pick))
            self._show_profile()
        if hasattr(self, "fp_box") and values and not self.fp_choice.get():
            self.fp_choice.set(self.profile_path.get() or values[0])
            self._draw_format_preview()

    def _show_profile(self):
        p = self.profile_path.get()
        if not p or not os.path.exists(p):
            return
        try:
            prof = Profile.load(p)
        except Exception as e:
            self.log("方案读取失败：%s" % e)
            return
        # 格式预览页跟着当前方案走
        if hasattr(self, "fp_choice"):
            self.fp_choice.set(p)
            self._draw_format_preview()
        lines = ["方案：%s" % prof.name]
        if prof.note:
            lines.append("说明：%s" % prof.note)
        lines.append("页面：%.2f×%.2fcm  边距 上%.2f 下%.2f 左%.2f 右%.2f" %
                     (prof.page.width_cm, prof.page.height_cm, prof.page.margin_top_cm,
                      prof.page.margin_bottom_cm, prof.page.margin_left_cm, prof.page.margin_right_cm))
        if prof.page.header_text:
            lines.append("页眉：%s" % prof.page.header_text)
        lines.append("共 %d 个部分：" % len(prof.roles))
        for r in prof.roles:
            lines.append("  · %-10s 套用样式「%s」 %s" % (r.name, r.style, r.describe()))
        self._set_result("\n".join(lines))
        self._fill_tree(prof)

    def _fill_tree(self, prof):
        self.tree.delete(*self.tree.get_children())
        for r in prof.roles:
            how = "样式名：" + "、".join(r.match_styles) if r.match_styles else ""
            if r.match_regex:
                how += ("；" if how else "") + "内容规则 %d 条" % len(r.match_regex)
            self.tree.insert("", "end", values=(r.name, r.style, r.describe(), how))

    # ---------------------------------------- 输出
    def log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self.logbox.configure(state="normal")
        self.logbox.insert("end", "[%s] %s\n" % (ts, msg))
        self.logbox.see("end")
        self.logbox.configure(state="disabled")

    def _set_result(self, text: str):
        self.result.delete("1.0", "end")
        self.result.insert("1.0", text)

    def _poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    self.log(payload)
                elif kind == "status":
                    self.status.set(payload)
                elif kind == "done":
                    self._finish(*payload)
                elif kind == "error":
                    self._finish(False, payload, "")
        except queue.Empty:
            pass
        self.after(120, self._poll)

    def _set_busy(self, busy: bool):
        self.busy = busy
        state = "disabled" if busy else "normal"
        self.btn_apply.configure(state=state)
        self.configure(cursor="watch" if busy else "")

    def _finish(self, ok: bool, summary: str, detail: str):
        self._set_busy(False)
        self.status.set("完成" if ok else "失败")
        if detail:
            self._set_result(detail)
        if not ok:
            messagebox.showerror(APP_TITLE, summary)
        else:
            self.log("完成。")

    # ---------------------------------------- 动作
    def _do_newdoc(self):
        """一步生成排好格式的空白论文。"""
        if self.busy:
            return
        prof_path = self.profile_path.get().strip()
        if not prof_path or not os.path.exists(prof_path):
            messagebox.showwarning(APP_TITLE, "先选一个格式方案（② 格式方案 YAML）")
            return
        default_name = os.path.join(paths.OUTPUT_DIR, "空白论文_已排版.docx")
        out = filedialog.asksaveasfilename(
            title="新论文存到哪里", defaultextension=".docx",
            initialfile=os.path.basename(default_name),
            initialdir=paths.OUTPUT_DIR,
            filetypes=[("Word 文档", "*.docx")])
        if not out:
            return
        self._set_busy(True)
        self.status.set("正在生成空白论文…")

        def work():
            try:
                from . import skeleton

                prof = Profile.load(prof_path)
                path, report = skeleton.new_formatted_doc(prof, out)
                self.q.put(("log", "已生成空白论文：%s" % path))
                self.q.put(("status", "已生成：%s" % os.path.basename(path)))
                self.q.put(("done", (True,
                                     "已生成排好格式的空白论文。",
                                     "已生成空白论文：\n%s\n\n"
                                     "直接打开写内容即可；目录在 Word 里右键「更新域」即可生成。" % path)))
            except Exception as e:
                self.q.put(("done", (False, "生成失败：%s" % e, "")))

        threading.Thread(target=work, daemon=True).start()

    def _do_apply(self):
        if self.busy:
            return
        src = self.in_path.get().strip()
        prof_path = self.profile_path.get().strip()
        if not src or not os.path.exists(src):
            messagebox.showwarning(APP_TITLE, "请先选择要排版的文档")
            return
        if not prof_path or not os.path.exists(prof_path):
            messagebox.showwarning(APP_TITLE, "请先选择格式方案（.yaml）")
            return
        out = self.out_path.get().strip() or os.path.join(
            paths.OUTPUT_DIR, os.path.splitext(os.path.basename(src))[0] + "_已排版.docx")
        if os.path.abspath(out) == os.path.abspath(src):
            messagebox.showwarning(APP_TITLE, "输出文件不能和原文件同名")
            return
        self._set_busy(True)
        self.status.set("正在排版…")
        self._set_result("")
        cover_path = self.cover_paths.get(self.cover_choice.get(), "")
        cover_mode = COVER_MODES.get(self.cover_mode_choice.get(), "format")
        threading.Thread(target=self._apply_worker,
                         args=(src, prof_path, out, self.opt_backup.get(),
                               self.opt_word.get(), self.opt_pdf.get(), self.opt_verify.get(),
                               cover_path, cover_mode, self.opt_sections.get()),
                         daemon=True).start()

    def _apply_worker(self, src, prof_path, out, backup, use_word, pdf, do_verify,
                      cover_path="", cover_mode="format", ensure_sections=False):
        try:
            prof = Profile.load(prof_path)
            try:
                prof.options["ensure_sections"] = bool(ensure_sections)
            except Exception:
                pass
            os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
            if backup:
                bak = os.path.join(os.path.dirname(os.path.abspath(src)),
                                   os.path.splitext(os.path.basename(src))[0] + "_备份.docx")
                if os.path.exists(bak):
                    self.q.put(("log", "备份已存在，保留这一份：%s" % bak))
                else:
                    import shutil
                    shutil.copy2(src, bak)
                    self.q.put(("log", "已备份原文件：%s" % bak))

            cover_obj = None
            if cover_path and os.path.exists(cover_path):
                from .cover import Cover
                cover_obj = Cover.load(cover_path)
                self.q.put(("log", "套用封面：%s（%s）" % (cover_obj.name, cover_mode)))

            report = apply_profile(src, out, prof, progress=lambda m: self.q.put(("log", "  " + m)),
                                   cover=cover_obj, cover_mode=cover_mode)
            for line in report.summary().splitlines():
                self.q.put(("log", line))

            detail = report.summary()
            if use_word:
                self.q.put(("log", "调用 Word 更新目录与页码…"))
                pdf_path = (os.path.splitext(out)[0] + ".pdf") if pdf else None
                ok, msg = wordio.update_fields(out, pdf_path=pdf_path)
                self.q.put(("log", "Word：" + ("成功" if ok else "失败——" + msg[-300:])))
                if not ok:
                    detail += "\n\n提示：目录/页码没能自动更新，打开文档后按 Ctrl+A 再按 F9 即可。"

            if do_verify:
                self.q.put(("log", "校验结果…"))
                res = verify(out, prof)
                detail += "\n\n—— 校验 ——\n" + res.text_report()
                for line in res.text_report().splitlines()[:6]:
                    self.q.put(("log", line))

            self.q.put(("done", (True, "排版完成", detail)))
            self.q.put(("log", "已生成：%s" % out))
            try:
                os.startfile(os.path.dirname(os.path.abspath(out)))
            except Exception:
                pass
        except Exception:
            tb = traceback.format_exc()
            self.q.put(("log", tb))
            self.q.put(("error", ("处理失败：\n" + tb.strip().splitlines()[-1], "")))

    def _do_extract(self):
        if self.busy:
            return
        src = self.tex_path.get().strip()
        if not src or not os.path.exists(src):
            messagebox.showwarning(APP_TITLE, "请先选择模板文档")
            return
        out = self.tex_out.get().strip()
        if not out:
            out = os.path.join(paths.CONFIG_DIR,
                               os.path.splitext(os.path.basename(src))[0] + ".yaml")
            self.tex_out.set(out)
        self._set_busy(True)
        self.status.set("正在读取模板格式…")
        threading.Thread(target=self._extract_worker, args=(src, out), daemon=True).start()

    def _extract_worker(self, src, out):
        try:
            prof = build_profile(src)
            prof.save(out)
            self.q.put(("log", "已生成格式方案：%s" % out))
            for r in prof.roles:
                self.q.put(("log", "  · %-10s 样式「%s」 %s" % (r.name, r.style, r.describe())))

            cover_msg = ""
            try:
                from .cover import extract_cover, save_to_library
                from docx import Document
                stem = os.path.splitext(os.path.basename(src))[0]
                cov = extract_cover(Document(src), name=stem + "封面", source=src)
                if cov is not None:
                    path = save_to_library(cov, paths.BASE, filename=stem + "_封面")
                    self.q.put(("log", "封面已入库：%s" % path))
                    self.q.put(("log", "  " + cov.summary()))
                    cover_msg = "\n\n封面已存进封面库（%d 行），排版时可以在「④ 封面」里选它。" % len(cov.lines)
                else:
                    self.q.put(("log", "没识别到封面（缺分节符或摘要标题），跳过"))
            except Exception as e:
                self.q.put(("log", "封面提取失败：%s" % e))

            # 提取完先做「回头比对」：拿新方案把模板重排一遍，逐页和原模板比
            verify_msg = ""
            try:
                from .roundtrip import roundtrip_check

                self.q.put(("status", "正在比对：用新方案重排模板…"))
                rt = roundtrip_check(src, prof)
                if rt.get("ok"):
                    self.q.put(("log", "比对通过：%s" % rt["message"]))
                    verify_msg = "\n\n比对：%s ✅" % rt["message"]
                else:
                    self.q.put(("log", "比对未通过：" + rt.get("message", "")))
                    verify_msg = ("\n\n比对：%s ⚠\n"
                                  "说明方案和模板没完全对上，建议看一下差异大的页。" %
                                  rt.get("message", ""))
            except Exception as e:
                self.q.put(("log", "比对跳过：%s" % e))

            self.q.put(("done", (True, "提取完成",
                                 "格式方案已保存：\n%s\n\n识别到 %d 个部分，可在「方案文件夹」里用记事本微调。%s%s" %
                                 (out, len(prof.roles), cover_msg, verify_msg))))
            self.profile_path.set(out)
            self._refresh_profiles()
            self.profile_path.set(out)
            self._fill_tree(prof)
            self._refresh_covers()
            self._refresh_coverlib()
            self._sync_lib_selection()
        except Exception:
            tb = traceback.format_exc()
            self.q.put(("log", tb))
            self.q.put(("error", ("提取失败：\n" + tb.strip().splitlines()[-1], "")))


HELP_TEXT = """使用步骤

一、先准备「格式方案」
  1. 切到「提取模板格式」页；
  2. 「模板文档」选一份格式正确的文件（例如你的论文样板.docx）；
  3. 点「提取格式方案」，会生成一个 .yaml 文件，保存在 configs 文件夹里。
     同一份模板的封面也会自动存进「封面库」（covers 文件夹）。

二、再排版文档
  1. 回到「排版文档」页，选要改的文档；
  2. 格式方案选上一步生成的 .yaml（没有就先去提取）；
  3. 「④ 封面」可选一个封面库里的封面，下面「封面怎么套」三选一：
       · 只改格式 → 只改封面的字体/字号/对齐/间距，文字一个字不动；
       · 重建并自动填 → 按封面库重排，自动填本文档里的姓名/学号/题目等，缺的留空；
       · 重建但全部留空 → 生成一张空白封面（标签 + 下划线），打印手写或自己往横线上打。
  4. 点「开始排版」。输出文件默认放在「输出」文件夹，文件名后面加「_已排版」。

三、封面库怎么来、怎么用
  在「提取模板格式」页提取模板时，封面会自动入库；也可以单独放一份封面文档，
  在「排版文档 → ④ 封面 → 浏览…」里直接选它的 yaml。
  库里存的是**空白表**（标签 + 横线，不含具体题目姓名，但校徽校名图会一起存），
  封面库文件在 covers 文件夹。

  「封面库」页可以直接看每一张封面：
    · 左边列表点一下，右边立刻画出这张封面的样子（不用开 Word，秒出）；
    · 点「用它排版 →」会跳回排版页并把它选上；
    · 排版页选了哪张，封面库页也会自动跳到哪张，两页是连着的。

  「格式预览」页看的是格式方案本身：用一段示例文字把题目、摘要、一二三级标题、
    正文、图题、表题、参考文献各画一遍，字号、缩进、行距、对齐一眼就能看出来。
    排版页换了方案，这一页会跟着换。

  「图表」页：把数据粘进去 → 选图类型（12 种：柱状/分组柱状/堆叠柱状/折线/面积/饼/环形/
   散点/箱线/水平条形/热力图/雷达图），配色可选现成 5 套或自定义颜色
    → 点「画图预览」立刻出图 → 填图题和「插在哪」→ 一键插进 Word，图号自动接续。

  「文献引用」页：联网找真实论文（OpenAlex），导出 RIS 给 EndNote，
    再在正文句末按你选的样式插上标 [n]，参考文献表自动补条目。
    能调：引用几篇、每段最多几处、样式（上标[n]/普通[n]/上标(n)/作者-年份）、
    检索主题、最早年份、是否必须有 DOI。已有 [1]…[7] 会接着编 [8][9]，不重号。
    只插正文陈述句，标题/摘要/图表题/参考文献不碰。

四、想改格式怎么办
  .yaml 就是格式清单，用记事本打开即可修改，常见改法：

    font:
      east_asia: 仿宋        # 中文字体
      ascii: Times New Roman # 西文字体
      size: 14               # 字号（磅），14=四号，10.5=五号
      bold: false
    para:
      align: justify         # left/center/right/justify
      line_spacing: 1.5      # 倍数行距；要固定行距改用 line_exact_pt: 20
      first_line_chars: 2    # 首行缩进几个字符
      space_before_pt: 0
      space_after_pt: 0

  每个「部分」有两套识别依据：
    match_styles: 文档里用这些样式名的段落，直接归到该部分；
    match_regex:  样式名对不上时，用内容正则兜底（如 ^图\\s*\\d+）。
  两者都不写，就只按结构规则（表格内、目录行）或兜底角色处理。

四、几个约定
  · 封面默认不动（cover_keep: true）——封面是自由排版，硬套样式往往更难看。
  · 默认不改页面已有的页码格式（封面无页码、摘要用罗马数字、正文用阿拉伯数字这种常见做法会保留）。
  · 只改格式，不改一个字：文字、图片、表格内容、EndNote 域都不碰。
  · 目录和页码是「域」，需要 Word 打开后更新才能显示新页码；工具会自动调 Word 更新，失败时打开文档按 Ctrl+A、F9 即可。

五、命令行（可选）
  python -m wordfmt.cli extract 模板.docx
  python -m wordfmt.cli apply 论文.docx configs\\论文格式.yaml --word --verify
  python -m wordfmt.cli verify 论文_已排版.docx configs\\论文格式.yaml
"""


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
