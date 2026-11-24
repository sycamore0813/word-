# -*- coding: utf-8 -*-
"""封面：识别、存进「封面库」、再套用到别的文档。

封面和正文不一样——每一行字号、对齐、间距都不同，所以库里存的不是「一个样式」，
而是**逐行的格式 + 文字骨架**：

    lines:
      - text: "题   目  草莓果汁发酵前后感官品质差异及综合评价研究"
        parts: [{text: "题   目  "}, {field: title}]
        font: {east_asia: 宋体, size: 18, bold: true}
        para: {align: center, line_exact_pt: 30}

套用时按 parts 拼文字：固定部分照抄，字段部分填目标文档里识别到的值；
目标文档没有这个信息就留空，方便本人再填。
"""
from __future__ import annotations

import copy
import os
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

import yaml
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from . import docxutil as du
from .profile import Font, Para

# 封面上常见的信息栏：字段名 → 标签正则（标签里允许夹空格，冒号中英文均可）
FIELD_LABELS = [
    ("secret", r"密\s*级\s*[：:]"),
    ("title", r"题\s*目\s*[：:]?"),
    ("college", r"学\s*院\s*[：:]"),
    ("major", r"专\s*业\s*[：:]"),
    ("class_name", r"班\s*级\s*[：:]"),
    ("name", r"(?:学生)?姓\s*名\s*[：:]"),
    ("sid", r"学\s*号\s*[：:]"),
    ("teacher", r"指\s*导\s*教\s*师\s*[：:]"),
    ("rank", r"职\s*称\s*[：:]"),
]

FIELD_NAMES = {
    "secret": "密级", "title": "题目", "college": "学院", "major": "专业",
    "class_name": "班级", "name": "姓名", "sid": "学号", "teacher": "指导教师",
    "rank": "职称", "date": "日期",
}

# 这几项属于「信息栏」，要统一左缩进、两列对齐
INFO_FIELDS = {"college", "major", "class_name", "name", "sid", "teacher", "rank"}


def _is_info_line(line: "CoverLine") -> bool:
    fields = set(line.needed_fields())
    return bool(fields) and fields <= INFO_FIELDS


def _strip_leading(segments):
    out = []
    first = True
    for text, u in segments:
        if first:
            text = text.lstrip()
            if not text:
                continue
            first = False
        out.append((text, u))
    return out


@dataclass
class CoverLine:
    text: str = ""
    parts: Optional[List[Dict[str, str]]] = None
    field: Optional[str] = None
    font: Optional[Font] = None
    para: Optional[Para] = None
    is_picture: bool = False
    image: Optional[Dict[str, object]] = None   # {file, cx, cy}，图片存在封面库旁边的 _images 里
    image_blob: Optional[bytes] = None          # 提取时暂存，入库后清空，不进 YAML
    has_tab: bool = False                       # 两列信息行：第二列用制表位对齐

    def __post_init__(self):
        if self.parts is None:
            self.parts = []
        if self.font is None:
            self.font = Font()
        if self.para is None:
            self.para = Para()

    def render(self, values: Dict[str, str]) -> str:
        return "".join(t for t, _u in self.render_segments(values))

    def render_segments(self, values: Dict[str, str]) -> List[Tuple[str, bool]]:
        """按「文字 + 是否下划线」切段——封面上的填空横线是下划线空格，不能丢。

        两列信息行（学生姓名…学号…）用制表符把第二列顶到统一的制表位，
        这样填进去的字多长都不会把右边那列挤歪。
        """
        if not self.parts:
            return [(self.text, False)]
        total_fields = sum(1 for p in self.parts if "field" in p)
        out: List[Tuple[str, bool]] = []
        seen = 0
        tabbed = False
        for part in self.parts:
            if "field" in part:
                seen += 1
                text = str(values.get(part["field"], "") or "").strip()
                if not text:
                    width = int(part.get("width") or 0)
                    text = " " * width if width else ""
                    u = bool(part.get("u", True))
                else:
                    u = bool(part.get("u", False))
            else:
                text = part.get("text", "")
                u = bool(part.get("u"))
                # 第一列填完之后、第二列标签之前的空白 → 换成制表符
                if (self.has_tab and not tabbed and seen == 1 and total_fields > 1
                        and text.strip() == ""):
                    out.append(("\t", False))
                    tabbed = True
                    continue
            if not text:
                continue
            if out and out[-1][1] == u:
                out[-1] = (out[-1][0] + text, u)
            else:
                out.append((text, u))
        return out or [("", False)]

    def skeleton(self) -> str:
        """只留标签和空行，不含具体内容（入库时存这个，免得把别人的题目姓名存进去）。"""
        if not self.parts:
            return self.text
        out = []
        for part in self.parts:
            if "field" in part:
                out.append(" " * int(part.get("width") or 0))
            else:
                out.append(part.get("text", ""))
        return "".join(out)

    def needed_fields(self) -> List[str]:
        return [p["field"] for p in self.parts if "field" in p]


@dataclass
class Cover:
    name: str = "未命名封面"
    source: str = ""
    has_picture: bool = False
    lines: List[CoverLine] = field(default_factory=list)
    base_dir: str = ""      # 封面 yaml 所在目录（图片相对它找）
    page: Dict[str, float] = field(default_factory=dict)   # 纸张/页边距，插到新文档时用
    info_tab_cm: float = 0.0   # 两列信息行的第二列制表位（相对左边距，cm）
    info_indent_chars: float = 0.0   # 信息栏统一的首行缩进（字符），保证「学院/专业/姓名/教师」左对齐

    # ---------------- 读写 ----------------
    def save(self, path: str) -> None:
        self.save_images(path)
        data = {"name": self.name, "source": self.source,
                "has_picture": self.has_picture,
                "page": self.page or {},
                "info_tab_cm": round(self.info_tab_cm, 2) if self.info_tab_cm else 0,
                "info_indent_chars": round(self.info_indent_chars, 2) if self.info_indent_chars else 0,
                "lines": [self._line_to_dict(l) for l in self.lines]}
        with open(path, "w", encoding="utf-8") as f:
            yaml.safe_dump(data, f, allow_unicode=True, sort_keys=False,
                           default_flow_style=False, width=200)

    @staticmethod
    def _line_to_dict(line: CoverLine) -> Dict:
        d: Dict = {"text": line.text}
        if line.field:
            d["field"] = line.field
        if line.parts and any("field" in p for p in line.parts):
            d["parts"] = line.parts
        if line.is_picture:
            d["is_picture"] = True
        if line.has_tab:
            d["has_tab"] = True
        if line.image:
            d["image"] = line.image
        fonts = {k: v for k, v in _asdict(line.font).items() if v is not None}
        paras = {k: v for k, v in _asdict(line.para).items() if v is not None}
        if fonts:
            d["font"] = fonts
        if paras:
            d["para"] = paras
        return d

    @classmethod
    def load(cls, path: str) -> "Cover":
        with open(path, encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
        cov = cls(name=data.get("name", "未命名封面"),
                  source=data.get("source", ""),
                  has_picture=bool(data.get("has_picture")),
                  base_dir=os.path.dirname(os.path.abspath(path)),
                  page=data.get("page") or {},
                  info_tab_cm=float(data.get("info_tab_cm") or 0.0),
                  info_indent_chars=float(data.get("info_indent_chars") or 0.0))
        for item in data.get("lines") or []:
            cov.lines.append(CoverLine(
                text=item.get("text", ""),
                parts=item.get("parts") or [],
                field=item.get("field"),
                font=Font(**(item.get("font") or {})),
                para=Para(**(item.get("para") or {})),
                is_picture=bool(item.get("is_picture")),
                image=item.get("image"),
                has_tab=bool(item.get("has_tab"))))
        return cov

    def image_path(self, line: CoverLine) -> str:
        if not line.image or not line.image.get("file"):
            return ""
        return os.path.join(self.base_dir or "", str(line.image["file"]))

    def save_images(self, path: str) -> None:
        """把提取时暂存的图片写到 yaml 旁边的 <名字>_images/ 里。"""
        stem = os.path.splitext(os.path.basename(path))[0]
        img_dir = os.path.join(os.path.dirname(os.path.abspath(path)), stem + "_images")
        idx = 0
        for line in self.lines:
            if not line.image_blob:
                continue
            os.makedirs(img_dir, exist_ok=True)
            idx += 1
            ext = (line.image or {}).get("ext") or "png"
            fname = "%s_%d.%s" % (stem, idx, ext)
            with open(os.path.join(img_dir, fname), "wb") as f:
                f.write(line.image_blob)
            line.image = dict(line.image or {})
            line.image["file"] = os.path.join(stem + "_images", fname)
            line.image_blob = None
        self.base_dir = os.path.dirname(os.path.abspath(path))

    def summary(self) -> str:
        fields = []
        for line in self.lines:
            fields.extend(line.needed_fields())
        uniq = []
        for f in fields:
            if f not in uniq:
                uniq.append(f)
        return "%s（%d 行，可填字段：%s）" % (
            self.name, len(self.lines),
            "、".join(FIELD_NAMES.get(f, f) for f in uniq) or "无")


def _asdict(obj) -> Dict:
    import dataclasses
    return {f.name: getattr(obj, f.name) for f in dataclasses.fields(obj)}


# ---------------------------------------------------------------- 定位封面
# 这些段落属于「正文内容」，绝不能算进封面，否则重建封面时会被删掉
CONTENT_HEAD_RE = re.compile(
    r"^\s*(摘\s*要|英文摘要|关键词|关键字|abstract|key\s*words?|目\s*录|引\s*言|前\s*言)",
    re.I)


def cover_span(doc) -> Optional[Tuple[int, int]]:
    """返回封面段落的下标区间 [start, end]（含）。

    判定方式：从开头数到「摘要/关键词/Abstract/目录」或第一段长正文之前，
    再取这一段里最后一个带分节符的段落作为封面终点（封面通常单独一节）。
    找不到明确的封面就返回 None，宁可不套封面，也不能把正文删掉。
    """
    paras = list(du.iter_paragraphs(doc))
    if not paras:
        return None

    content = None
    for i, p in enumerate(paras):
        if du.in_table(p):
            continue
        text = du.text_of(p).strip()
        if not text:
            continue
        if CONTENT_HEAD_RE.match(text) or len(text) > 160:
            content = i
            break
    if content is None or content == 0:
        return None

    # 封面终点：content 之前最后一个带 sectPr 的段落；没有分节符就取 content 前一段
    end = content - 1
    for j in range(content - 1, -1, -1):
        pPr = paras[j]._p.find(qn("w:pPr"))
        if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
            end = j
            break

    # 封面区里如果有长段落，说明这不是封面，别碰
    for p in paras[:end + 1]:
        if len(du.text_of(p).strip()) > 160:
            return None
    # 摘要前面常有一个和封面重复的论文题目：没有分节符时它属于正文，不算封面
    while end > 0:
        text = du.text_of(paras[end]).strip()
        if not text:
            end -= 1
            continue
        earlier = {du.text_of(q).strip() for q in paras[:end]}
        if (len(text) > 10 and text in earlier) or _title_like(text):
            end -= 1
            continue
        break
    if end < 0:
        return None
    return 0, end


def _title_like(text: str) -> bool:
    """像「论文题目」的一行：短、没有信息栏标签、不是日期也不是「本科毕业论文」这种固定字。"""
    if not text or len(text) > 80:
        return False
    if re.search(r"二\s*[○〇0零].{0,6}年", text):
        return False
    if any(k in text for k in ("本科毕业论文", "毕业设计", "课程论文", "论文")):
        return False
    for _name, pat in FIELD_LABELS:
        if re.search(pat, text):
            return False
    return True


# ---------------------------------------------------------------- 文字拆分
def _split_parts(text: str, u_spans: Optional[List[Tuple[int, int]]] = None):
    """把一行拆成「固定文字 + 字段」交替的片段，并记录每段是否带下划线。

    封面的填空横线其实是「加下划线的空格」，所以下划线必须按字符区间保留，
    否则套用后横线就没了。
    """
    u_spans = u_spans or []

    def char_u(i: int) -> bool:
        return any(a <= i < b for a, b in u_spans)

    def emit_fixed(parts, s: int, e: int) -> None:
        if s >= e:
            return
        start, cur = s, char_u(s)
        for i in range(s + 1, e):
            u = char_u(i)
            if u != cur:
                parts.append({"text": text[start:i], "u": cur})
                start, cur = i, u
        parts.append({"text": text[start:e], "u": cur})

    if not text.strip():
        parts = []
        emit_fixed(parts, 0, len(text))
        return parts, None, []

    hits = []
    for name, pat in FIELD_LABELS:
        for m in re.finditer(pat, text):
            hits.append((m.start(), m.end(), name))
    if not hits:
        parts = []
        emit_fixed(parts, 0, len(text))
        return parts, None, []
    hits.sort()

    clean: List[Tuple[int, int, str]] = []
    for s, e, name in hits:
        if clean and s < clean[-1][1]:
            continue
        clean.append((s, e, name))

    parts = []
    fields: List[str] = []
    cursor = 0
    for idx, (s, e, name) in enumerate(clean):
        emit_fixed(parts, cursor, e)
        next_start = clean[idx + 1][0] if idx + 1 < len(clean) else len(text)
        value = text[e:next_start].rstrip()
        # 记住原来留了多宽的空：库里存成「标签 + 这么宽的下划线空格」，
        # 具体内容不存，谁用谁填。
        width = len(value.strip())
        if width:
            parts.append({"field": name, "width": width, "u": True})
            fields.append(name)
        else:
            parts.append({"field": name})
        cursor = e + len(value)
    emit_fixed(parts, cursor, len(text))
    return parts, (clean[0][2] if clean else None), fields


# ---------------------------------------------------------------- 提取
def _underline_spans(p) -> List[Tuple[int, int]]:
    """按 w:t 的顺序，找出哪些字符区间是加下划线的（封面的填空横线）。"""
    spans = []
    pos = 0
    for r in p._p.iter(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        u = False
        if rPr is not None:
            el = rPr.find(qn("w:u"))
            if el is not None:
                u = (el.get(qn("w:val")) or "single") not in ("0", "none", "false")
        text = "".join(t.text or "" for t in r.iter(qn("w:t")))
        if u and text:
            spans.append((pos, pos + len(text)))
        pos += len(text)
    return spans


def _extract_image(doc, p) -> Optional[Dict]:
    """把段落里的图片抠出来：二进制 + 原始扩展名 + 尺寸（EMU）。"""
    blip = p._p.find(".//" + qn("a:blip"))
    if blip is None:
        return None
    rid = blip.get(qn("r:embed"))
    part = doc.part.related_parts.get(rid) if rid else None
    if part is None:
        return None
    name = str(getattr(part, "partname", ""))
    ext = os.path.splitext(name)[1].lstrip(".").lower() or "png"
    if ext == "jpeg":
        ext = "jpg"
    cx = cy = None
    extent = p._p.find(".//" + qn("wp:extent"))
    if extent is not None:
        try:
            cx = int(extent.get("cx"))
            cy = int(extent.get("cy"))
        except (TypeError, ValueError):
            pass
    return {"blob": part.blob, "ext": ext, "cx": cx, "cy": cy}


def extract_cover(doc, name: str = "", source: str = "") -> Optional[Cover]:
    """从文档里把封面抽成 Cover（逐行格式 + 文字骨架）。"""
    from .extract import measure_paragraph

    span = cover_span(doc)
    if span is None:
        return None
    start, end = span
    paras = list(du.iter_paragraphs(doc))
    picked = paras[start:end + 1]
    if not any(du.text_of(p).strip() for p in picked):
        return None

    cover = Cover(name=name or "未命名封面", source=source)
    try:
        sec = doc.sections[0]
        cover.page = {
            "width_cm": round(sec.page_width.cm, 2),
            "height_cm": round(sec.page_height.cm, 2),
            "top_cm": round(sec.top_margin.cm, 2),
            "bottom_cm": round(sec.bottom_margin.cm, 2),
            "left_cm": round(sec.left_margin.cm, 2),
            "right_cm": round(sec.right_margin.cm, 2),
        }
    except Exception:
        cover.page = {}
    for p in picked:
        font, para = measure_paragraph(doc, p)
        parts, primary, _fields = _split_parts(du.text_of(p), _underline_spans(p))
        line = CoverLine(text="", parts=parts, field=primary,
                         font=font, para=para, is_picture=du.has_picture(p))
        # 库里只存「标签 + 空行」，不存具体题目/姓名/学号
        line.text = line.skeleton()
        if line.is_picture:
            cover.has_picture = True
            img = _extract_image(doc, p)
            if img:
                line.image = {"cx": img["cx"], "cy": img["cy"], "ext": img["ext"]}
                line.image_blob = img["blob"]
        # 两列信息行（姓名…学号…）：记下第二列大概从哪开始，套用时用制表位对齐
        fields = [q for q in parts if "field" in q]
        if len(fields) >= 2:
            line.has_tab = True
            size_cm = (font.size or 14.0) * 2.54 / 72.0
            indent_cm = (para.first_line_chars or 0) * size_cm
            x_cm = _second_column_cm(du.text_of(p), font.size or 14.0) + indent_cm
            if x_cm:
                cover.info_tab_cm = max(cover.info_tab_cm, x_cm)
        # 信息栏统一缩进：学院/专业靠段落缩进，姓名/教师靠行首空格，两种写法要归一
        names = {q["field"] for q in fields}
        if names and names <= INFO_FIELDS:
            raw = du.text_of(p)
            lead = len(raw) - len(raw.lstrip(" "))
            indent_chars = (para.first_line_chars or 0) + lead * 0.5
            cover.info_indent_chars = max(cover.info_indent_chars, indent_chars)
        cover.lines.append(line)
    if cover.info_tab_cm and cover.page:
        usable = (cover.page.get("width_cm", 21.0) - cover.page.get("left_cm", 3.17)
                  - cover.page.get("right_cm", 3.17))
        cover.info_tab_cm = round(min(cover.info_tab_cm, max(6.0, usable - 5.0)), 2)
    return cover if cover.lines else None


def _text_width_cm(text: str, size_pt: float) -> float:
    """粗算一段文字的宽度（cm）：中日韩字符算 1 个字宽，ASCII 算半个。"""
    em = 0.0
    for ch in text:
        em += 1.0 if ord(ch) > 0x2E80 else 0.5
    return em * size_pt * 2.54 / 72.0


def _second_column_cm(text: str, size_pt: float) -> float:
    """第二列标签（学号/职称…）在这行里从哪个位置开始（cm）。"""
    starts = []
    for _name, pat in FIELD_LABELS:
        for m in re.finditer(pat, text):
            starts.append(m.start())
    starts = sorted(set(starts))
    if len(starts) >= 2:
        return _text_width_cm(text[:starts[1]], size_pt)
    return 0.0


def extract_values(doc, cover: Optional[Cover] = None) -> Dict[str, str]:
    """从文档封面里把「姓名/学号/题目…」的值抠出来。"""
    span = cover_span(doc)
    texts: List[str] = []
    if span is not None:
        paras = list(du.iter_paragraphs(doc))
        texts = [du.text_of(p) for p in paras[span[0]:span[1] + 1]]
    else:
        # 没有封面区：从头几段里找
        texts = [du.text_of(p) for p in list(du.iter_paragraphs(doc))[:12]]

    values: Dict[str, str] = {}
    for line in texts:
        hits = []
        for name, pat in FIELD_LABELS:
            for m in re.finditer(pat, line):
                hits.append((m.start(), m.end(), name))
        hits.sort()
        clean = []
        for s, e, name in hits:
            if clean and s < clean[-1][1]:
                continue
            clean.append((s, e, name))
        for idx, (s, e, name) in enumerate(clean):
            nxt = clean[idx + 1][0] if idx + 1 < len(clean) else len(line)
            val = line[e:nxt]
            # 去掉填空用的下划线、零宽字符和多余空白，不然会把「____」「\u200b」当成内容抄进去
            val = re.sub(r"[_＿\u200b\u200c\u200d\ufeff]+", "", val)
            val = re.sub(r"[\s\u3000]+", " ", val).strip()
            if len(val) > 40:
                val = val[:40]
            if val and name not in values:
                values[name] = val
    return values


# ---------------------------------------------------------------- 套用
def _has_sectpr(el) -> bool:
    pPr = el.find(qn("w:pPr"))
    return pPr is not None and pPr.find(qn("w:sectPr")) is not None


def _has_pic(el) -> bool:
    return el.find(".//" + qn("w:drawing")) is not None or \
        el.find(".//" + qn("w:pict")) is not None


def _set_mark_font(para: Paragraph, font: Font) -> None:
    """设置「段落标记」的字体字号——空段落的高度就是由它决定的。"""
    pPr = para._p.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    if font.east_asia:
        rf.set(qn("w:eastAsia"), font.east_asia)
    if font.ascii:
        rf.set(qn("w:ascii"), font.ascii)
        rf.set(qn("w:hAnsi"), font.ascii)
    if font.size:
        for tag in ("sz", "szCs"):
            el = rPr.find(qn("w:" + tag))
            if el is None:
                el = OxmlElement("w:" + tag)
                rPr.append(el)
            el.set(qn("w:val"), str(int(round(font.size * 2))))


def _set_tab_stop(para: Paragraph, cm: float) -> None:
    """在段落上设一个左对齐制表位（cm，相对左边距）。"""
    if not cm or cm <= 0:
        return
    pPr = para._p.get_or_add_pPr()
    tabs = pPr.find(qn("w:tabs"))
    if tabs is None:
        tabs = OxmlElement("w:tabs")
        pPr.append(tabs)
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "left")
    tab.set(qn("w:pos"), str(int(round(cm * 567))))
    tabs.append(tab)


def _new_paragraph(doc, segments, font: Font, para: Para, tab_cm: float = 0.0):
    p = Paragraph(OxmlElement("w:p"), doc)
    _set_mark_font(p, font)
    du.set_para(p, para, font_size_hint=font.size)
    _add_segments(p, segments, font)
    _set_tab_stop(p, tab_cm)
    return p._p


def _add_segments(para: Paragraph, segments, font: Font) -> None:
    """按片段写 run。下划线空白用「空格 + 下划线」，末尾补一个零宽空格。

    Word 不会给行尾的空格画下划线；后面跟一个零宽空格它就会画，而且看不见，
    所以打印出来是一条连续的线，往线上打字时字也自带下划线，不会把线断开。
    """
    last_u = False
    for text, u in segments:
        if not text:
            continue
        run = para.add_run(text)
        f = copy.copy(font)
        if u:
            f.underline = True
        du.set_font(run, f)
        last_u = u
    if last_u:
        # 末尾补一个「白色 1 磅句点」：肉眼看不见，但能让 Word 给行尾的下划线空格画线。
        # 之前用零宽空格，有些字体/视图会显示成小方块；用制表符会把前面的空格吃掉。
        from docx.shared import Pt, RGBColor
        dot = para.add_run(".")
        dot.font.size = Pt(1)
        dot.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        dot.font.underline = False


def _ensure_trailing_underline(doc) -> None:
    """让 Word 给「行尾的下划线空格」画线。

    封面的填空横线就是加下划线的空格，Word 默认不画行尾空格的下划线，
    需要文档兼容设置里打开 underlineTrailingSpaces（并处于 Word 2013+ 模式）。
    """
    try:
        settings = doc.settings.element
        compat = settings.find(qn("w:compat"))
        if compat is None:
            compat = OxmlElement("w:compat")
            settings.append(compat)
        existing = {cs.get(qn("w:name")): cs for cs in compat.findall(qn("w:compatSetting"))}
        for name, val in (("compatibilityMode", "15"), ("underlineTrailingSpaces", "1")):
            cs = existing.get(name)
            if cs is None:
                cs = OxmlElement("w:compatSetting")
                cs.set(qn("w:name"), name)
                cs.set(qn("w:uri"), "http://schemas.microsoft.com/office/word")
                compat.append(cs)
            cs.set(qn("w:val"), val)
    except Exception:
        pass


def apply_cover(doc, cover: Cover, mode: str = "rebuild",
                values: Optional[Dict[str, str]] = None) -> Dict:
    """把封面套到文档上。

    mode = "rebuild"：按封面库重建封面（固定文字照抄，字段填目标文档里的值，缺的留空）
    mode = "blank"  ：按封面库重建，但**全部留空**（当空表用，谁用谁填）
    mode = "format" ：只把现成封面的每行格式改成库里那样，文字一个字不动
    """
    _ensure_trailing_underline(doc)
    span = cover_span(doc)
    report = {"mode": mode, "lines": 0, "filled": [], "blank": [], "warning": ""}
    body = doc.element.body
    if mode == "blank":
        values = {}
    elif values is None:
        values = extract_values(doc, cover) if span is not None else {}

    if span is not None and mode == "format":
        return _apply_format_only(doc, cover, span, report)

    if span is None:
        # 文档里没有封面区（新文档/只有正文）：直接把封面插到最前面
        return _insert_cover_at_start(doc, cover, values, report)

    start, end = span
    body_paras = body.findall(qn("w:p"))

    # ---------- 重建 ----------
    # 封面上的图片（校徽校名）优先用封面库里的原图；库里没有才退回目标文档那张。
    els = [p._p for p in list(du.iter_paragraphs(doc))[start:end + 1]]
    sect_el = els[-1] if _has_sectpr(els[-1]) else None
    pics = [el for el in els if _has_pic(el)]
    library_pics = sum(1 for l in cover.lines if l.is_picture)
    reused_pics = 0
    placed_pics = 0

    lines = list(cover.lines)
    # 最后一行（通常是日期行）写回带分节符的那一段，避免多出一个空行
    tail = None
    if sect_el is not None and lines and not lines[-1].is_picture:
        tail = lines.pop()

    new_els = []
    pic_i = 0
    for line in lines:
        if line.is_picture:
            path = cover.image_path(line)
            if path and os.path.exists(path):
                new_els.append(_new_picture_paragraph(doc, path, line))
                placed_pics += 1
                report["lines"] += 1
            elif pic_i < len(pics):
                new_els.append(pics[pic_i])
                pic_i += 1
                reused_pics += 1
            continue
        segs, spec = _segments_and_spec(cover, line, values)
        new_els.append(_new_paragraph(doc, segs, line.font, spec,
                                      tab_cm=cover.info_tab_cm if line.has_tab else 0.0))
        report["lines"] += 1
        _count_fields(line, values, report)

    anchor = sect_el if sect_el is not None else (els[0] if els else None)
    for el in new_els:
        if anchor is not None:
            anchor.addprevious(el)
        else:
            sect = body.find(qn("w:sectPr"))
            if sect is not None:
                sect.addprevious(el)
            else:
                body.append(el)

    if tail is not None:
        segs, spec = _segments_and_spec(cover, tail, values)
        _fill_paragraph(doc, Paragraph(sect_el, doc), segs, tail.font,
                        spec, tab_cm=cover.info_tab_cm if tail.has_tab else 0.0)
        report["lines"] += 1
        _count_fields(tail, values, report)

    # 清掉没用上的旧封面段
    keep_ids = {id(el) for el in new_els}
    for el in els:
        if id(el) in keep_ids or el is sect_el:
            continue
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)

    # 没有分节符时：清掉封面后面残留的空段（否则会多出一页空白），并让正文另起一页
    if sect_el is None:
        cur = new_els[-1].getnext() if new_els else body.find(qn("w:p"))
        while cur is not None and cur.tag == qn("w:p"):
            if _has_sectpr(cur):
                break
            txt = "".join(t.text or "" for t in cur.iter(qn("w:t"))).strip()
            if txt:
                break
            nxt = cur.getnext()
            body.remove(cur)
            cur = nxt
        if cur is not None and cur.tag == qn("w:p"):
            pPr = cur.get_or_add_pPr()
            if pPr.find(qn("w:pageBreakBefore")) is None:
                pPr.append(OxmlElement("w:pageBreakBefore"))

    missing = library_pics - placed_pics - reused_pics
    if missing > 0:
        report["warning"] = "封面库里有 %d 处图片没找到文件，已跳过" % missing
    if len(pics) > pic_i and len(pics) > reused_pics:
        report["warning"] = (report["warning"] + "；" if report["warning"] else "") + \
            "目标文档封面多出 %d 处图片，已移除" % (len(pics) - pic_i)
    return report


def _new_picture_paragraph(doc, path: str, line: CoverLine):
    """按封面库里的原图和尺寸插一张图（校徽校名那种）。"""
    from docx.shared import Emu

    p = Paragraph(OxmlElement("w:p"), doc)
    _set_mark_font(p, line.font)
    du.set_para(p, line.para, font_size_hint=line.font.size)
    run = p.add_run()
    size = line.image or {}
    kwargs = {}
    if size.get("cx") and size.get("cy"):
        kwargs["width"] = Emu(int(size["cx"]))
        kwargs["height"] = Emu(int(size["cy"]))
    run.add_picture(path, **kwargs)
    return p._p


def _insert_cover_at_start(doc, cover: Cover, values: Dict[str, str], report: Dict) -> Dict:
    """文档没有封面区时，把整张封面插到最前面，并让正文另起一页。"""
    from docx.shared import Cm

    if cover.page:
        try:
            sec = doc.sections[0]
            if cover.page.get("width_cm"):
                sec.page_width = Cm(cover.page["width_cm"])
            if cover.page.get("height_cm"):
                sec.page_height = Cm(cover.page["height_cm"])
            for key, setter in (("top_cm", "top_margin"), ("bottom_cm", "bottom_margin"),
                                ("left_cm", "left_margin"), ("right_cm", "right_margin")):
                if cover.page.get(key):
                    setattr(sec, setter, Cm(cover.page[key]))
        except Exception:
            pass

    body = doc.element.body
    first = body.find(qn("w:p"))
    for line in cover.lines:
        if line.is_picture:
            path = cover.image_path(line)
            if not (path and os.path.exists(path)):
                continue
            el = _new_picture_paragraph(doc, path, line)
            report["lines"] += 1
        else:
            segs, spec = _segments_and_spec(cover, line, values)
            el = _new_paragraph(doc, segs, line.font, spec,
                                tab_cm=cover.info_tab_cm if line.has_tab else 0.0)
            report["lines"] += 1
            _count_fields(line, values, report)
        if first is not None:
            first.addprevious(el)
        else:
            sect = body.find(qn("w:sectPr"))
            if sect is not None:
                sect.addprevious(el)
            else:
                body.append(el)
    if first is not None:
        pPr = first.get_or_add_pPr()
        if pPr.find(qn("w:pageBreakBefore")) is None:
            pPr.append(OxmlElement("w:pageBreakBefore"))
    report["warning"] = "文档里没有现成封面，已把封面插在最前面"
    return report


def _segments_and_spec(cover: Cover, line: CoverLine, values: Dict[str, str]):
    """信息栏统一缩进 + 去掉行首空格；其它行原样。"""
    segs = line.render_segments(values)
    spec = line.para
    if _is_info_line(line):
        segs = _strip_leading(segs)
        if cover.info_indent_chars:
            spec = copy.copy(line.para)
            spec.first_line_chars = cover.info_indent_chars
    return segs, spec


def _count_fields(line: CoverLine, values: Dict[str, str], report: Dict) -> None:
    for f in line.needed_fields():
        if (values.get(f) or "").strip():
            if f not in report["filled"]:
                report["filled"].append(f)
        else:
            if f not in report["blank"]:
                report["blank"].append(f)


def _fill_paragraph(doc, para: Paragraph, segments, font: Font, spec: Para,
                    tab_cm: float = 0.0) -> None:
    """把现成段落的内容换成 text，并按 font/spec 设格式（保留 pPr 里的 sectPr）。"""
    for child in list(para._p):
        if child.tag != qn("w:pPr"):
            para._p.remove(child)
    _set_mark_font(para, font)
    du.set_para(para, spec, font_size_hint=font.size)
    _add_segments(para, segments, font)
    _set_tab_stop(para, tab_cm)


def _apply_format_only(doc, cover: Cover, span, report: Dict) -> Dict:
    """只改格式，一个字不动：先按字段对应，剩下的按位置对应。

    关键点是「不动下划线」——封面上的填空横线是加下划线的空格，
    清格式时只能清字体/字号/行距，不能把 u 也清掉。
    """
    start, end = span
    paras = list(du.iter_paragraphs(doc))[start:end + 1]
    pairs = _map_paras_to_lines(paras, cover.lines)
    for p, target in pairs:
        if target.is_picture or du.has_picture(p):
            continue  # 图片段不动，免得把校徽压扁
        pPr = p._p.find(qn("w:pPr"))
        if pPr is not None:
            # 只清「库里明确给了值」的那几项，其余保持原样
            if target.para.align:
                for el in pPr.findall(qn("w:jc")):
                    pPr.remove(el)
            if any(v is not None for v in (target.para.first_line_chars, target.para.left_chars,
                                           target.para.right_chars)):
                for el in pPr.findall(qn("w:ind")):
                    pPr.remove(el)
            if any(v is not None for v in (target.para.line_spacing, target.para.line_exact_pt,
                                           target.para.space_before_pt, target.para.space_after_pt)):
                for el in pPr.findall(qn("w:spacing")):
                    pPr.remove(el)
        _set_mark_font(p, target.font)
        du.set_para(p, target.para, font_size_hint=target.font.size)
        for run in p.runs:
            if du.is_field_run(run):
                continue
            du.clear_run_direct(run, "font")
            du.set_font(run, target.font)
        report["lines"] += 1
    return report


def _map_paras_to_lines(paras, lines) -> List[Tuple]:
    """把文档里的段落和封面库里的行对上。

    段数一致时按位置一一对应（最可靠，版面不会乱）；
    段数不一致时先按字段（姓名/学号…）配对，剩下的按顺序补齐。
    """
    if len(paras) == len(lines):
        return list(zip(paras, lines))

    pairs = []
    pool = list(lines)
    rest_paras = []
    for p in paras:
        text = du.text_of(p).strip()
        hit = None
        if text:
            for line in pool:
                if not line.field:
                    continue
                pat = dict(FIELD_LABELS).get(line.field)
                if pat and re.search(pat, text):
                    hit = line
                    break
        if hit is not None:
            pool.remove(hit)
            pairs.append((p, hit))
        else:
            rest_paras.append(p)
    for p, line in zip(rest_paras, pool):
        pairs.append((p, line))
    return pairs


# ---------------------------------------------------------------- 封面库
def covers_dir(base: str) -> str:
    d = os.path.join(base, "covers")
    os.makedirs(d, exist_ok=True)
    return d


def list_covers(base: str) -> List[Tuple[str, str]]:
    """返回 [(显示名, 路径)]。"""
    d = covers_dir(base)
    out = []
    for fn in sorted(os.listdir(d)):
        if not fn.lower().endswith((".yaml", ".yml")):
            continue
        path = os.path.join(d, fn)
        try:
            with open(path, encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
            out.append((data.get("name") or os.path.splitext(fn)[0], path))
        except Exception:
            out.append((os.path.splitext(fn)[0], path))
    return out


def save_to_library(cover: Cover, base: str, filename: str = "") -> str:
    d = covers_dir(base)
    stem = filename or re.sub(r"[\\/:*?\"<>|]", "_", cover.name)
    path = os.path.join(d, stem + ".yaml")
    cover.save(path)
    return path
