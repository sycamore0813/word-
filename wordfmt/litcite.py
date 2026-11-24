# -*- coding: utf-8 -*-
"""文献引用：联网检索真实论文 → 生成 RIS（可导入 EndNote）→ 在 Word 里按规则插入引用。

设计原则：
  * 只连公开、免密钥的 OpenAlex，检索结果一定带标题，尽量带 DOI；
  * 引用只加在「陈述性正文句子」上，标题/摘要/图表题/参考文献一律跳过；
  * 引用格式、数量、条件都能调（界面和 YAML 里都能改）；
  * 不替用户改 EndNote 库：生成标准 RIS，用户双击导入即可（EndNote 会弹导入向导）。
"""
import json
import os
import re
import urllib.parse
import urllib.request

from docx.oxml.ns import qn

from . import docxutil as du

OPENALEX = "https://api.openalex.org/works"
MAILTO = "wordfmt-user@example.com"      # OpenAlex 的礼貌池：带上联系方式不容易被限流
UA = "wordfmt-litcite/1.0 (mailto:%s)" % MAILTO

# EndNote 里常见的几种引用样式（名字 → 渲染方式）
CITE_STYLES = {
    "顺序编码-上标[n]": {"numeric": True, "superscript": True, "template": "[%d]"},
    "顺序编码-普通[n]": {"numeric": True, "superscript": False, "template": "[%d]"},
    "顺序编码-上标(n)": {"numeric": True, "superscript": True, "template": "(%d)"},
    "作者-年份": {"numeric": False, "superscript": False, "template": "（%s, %d）"},
}

CITE_DEFAULTS = {
    "style": "顺序编码-上标[n]",
    "count": 2,              # 一共插几篇
    "max_per_para": 1,       # 每段最多几处
    "min_sentence_len": 20,  # 句子太短不值得引
    "year_from": 2010,
    "require_doi": True,
    "only_body": True,       # 只插在正文段落
    "skip_first_para": True, # 跳过每节第一段（通常是引入语）
    "triggers": [],          # 空 = 自动挑最像陈述句的句子
    "search_english": True,  # 中文查不到时用英文再试
    "only_oa": False,        # 只要能免费下载全文的（Open Access）
}

TRIGGER_WORDS = ["研究表明", "研究发现", "研究显示", "数据显示", "文献报道", "据统计",
                 "已有研究", "大量研究", "国内外研究", "学者", "据调查", "研究指出",
                 "results show", "studies have shown"]

# 常见学术词的中英对照（只用来补英文检索，不求全）
TERM_EN = {
    "挤压膨化": "extrusion cooking", "挤压": "extrusion", "膨化": "expansion",
    "膨化食品": "extruded snack", "挤压机": "extruder", "淀粉糊化": "starch gelatinization",
    "糊化": "gelatinization", "剪切": "shear", "休闲食品": "snack food",
    "谷物": "cereal", "淀粉": "starch", "大米": "rice",
    "玉米": "maize", "小麦": "wheat", "蛋白质": "protein", "膳食纤维": "dietary fiber",
    "食品": "food", "食品安全": "food safety", "工厂设计": "plant design",
    "工艺设计": "process design", "工艺流程": "process flow", "物料衡算": "material balance",
    "设备选型": "equipment selection", "车间": "workshop", "车间布置": "plant layout",
    "清洁生产": "cleaner production", "环境保护": "environmental protection",
    "废水": "wastewater", "能耗": "energy consumption", "质量控制": "quality control",
    "感官评价": "sensory evaluation", "质构": "texture", "水分": "moisture",
    "温度": "temperature", "烘焙": "baking", "干燥": "drying", "杀菌": "sterilization",
    "包装": "packaging", "自动化": "automation", "生产线": "production line",
    "深度学习": "deep learning", "机器学习": "machine learning", "神经网络": "neural network",
}

# 领域词表：让 jieba 把「休闲食品」「挤压膨化」这类术语当成一个词
DOMAIN_WORDS = [
    "休闲食品", "挤压膨化", "挤压膨化食品", "谷物", "谷物粉", "淀粉糊化", "糊化度",
    "工厂设计", "工艺设计", "工艺流程", "物料衡算", "设备选型", "车间布置", "厂区总平面",
    "清洁生产", "食品安全", "质量控制", "危害分析", "感官评价", "质构特性",
    "水分含量", "膨化度", "酥脆性", "复水性", "膳食纤维", "蛋白质", "氨基酸",
    "挤压机", "双螺杆", "单螺杆", "螺杆转速", "机筒温度", "模孔", "剪切力",
    "生产线", "自动化控制", "三废处理", "废水处理", "能耗", "技术经济分析",
    "深度学习", "机器学习", "神经网络", "卷积神经网络", "无损检测",
]

STOPWORDS = set("""的 了 是 在 和 与 及 等 对 为 以 能 可 将 会 把 被 从 而且 或 但 也 就 都 还 再 很 更 最
这 那 其 之 一个 一些 进行 通过 采用 具有 实现 用于 以及 不同 主要 相关 上述 如下 见表 如图
我们 本文 本设计 该 其中 由于 因此 所以 同时 另外 此外 目前 现在 一般 通常 需要 要求 可以
提高 降低 增加 减少 影响 作用 特点 优势 问题 情况 方面 方式 方法 过程 条件 因素
食品 工业 产品 设计 技术 研究 分析 系统 结构 材料 数据 结果 内容 部分 时间 空间
""".split())

# 这类词最能决定「这段话在讲什么」，检索时排最前
CORE_TERMS = {"挤压膨化", "挤压", "膨化", "膨化食品", "挤压机", "淀粉糊化", "糊化",
              "休闲食品", "淀粉", "谷物", "蛋白质", "膳食纤维", "工厂设计", "清洁生产"}

# 只有这些词算「够具体」，可以用它单独决定相关性；谷物/淀粉这种太泛的不算
SPECIFIC_TERMS = {"挤压膨化", "挤压", "膨化", "膨化食品", "挤压机", "淀粉糊化", "糊化",
                  "休闲食品", "工厂设计", "清洁生产", "车间布置", "设备选型", "物料衡算",
                  "工艺设计", "车间", "生产线", "食品安全", "膳食纤维", "感官评价"}

_jieba_ready = False


def _jieba_init():
    global _jieba_ready
    if _jieba_ready:
        return
    try:
        import jieba
        for w in DOMAIN_WORDS:
            jieba.add_word(w, freq=2000)
    except Exception:
        pass
    _jieba_ready = True

SPLIT_SENT = re.compile(r"(?<=[。！？；])")


class Paper:
    __slots__ = ("title", "authors", "year", "journal", "doi", "url", "abstract",
                 "cited_by", "query", "openalex_id", "volume", "issue", "pages",
                 "is_oa", "pdf_url")

    def __init__(self, **kw):
        for k in self.__slots__:
            setattr(self, k, kw.get(k))

    def __repr__(self):
        return "<%s %s>" % (self.year, (self.title or "")[:40])

    def author_text(self, sep=", "):
        return sep.join(self.authors or [])

    def short_author(self):
        if not self.authors:
            return "佚名"
        return self.authors[0]

    def key(self):
        return (self.doi or self.title or "").lower()


# ---------------------------------------------------------------- 检索
def _get(url, tries=3):
    """带重试的 GET：OpenAlex 偶尔 429，等两秒再来。"""
    import time as _time

    last = None
    for i in range(tries):
        req = urllib.request.Request(url, headers={"User-Agent": UA})
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                return json.load(r)
        except Exception as e:
            last = e
            if "429" in str(e) and i < tries - 1:
                _time.sleep(2 + i * 2)
                continue
            break
    if last is not None:
        raise last
    return {}


def _parse_work(w, query=""):
    inv = w.get("abstract_inverted_index") or {}
    abstract = ""
    if inv:
        pos = {}
        for word, idxs in inv.items():
            for i in idxs:
                pos[i] = word
        abstract = " ".join(pos[i] for i in sorted(pos))
    authors = []
    for a in w.get("authorships") or []:
        name = (a.get("author") or {}).get("display_name")
        if name:
            authors.append(name)
    loc = (w.get("primary_location") or {}).get("source") or {}
    biblio = w.get("biblio") or {}
    oa = w.get("open_access") or {}
    best = w.get("best_oa_location") or {}
    return Paper(
        is_oa=bool(oa.get("is_oa")),
        pdf_url=best.get("pdf_url") or oa.get("oa_url") or "",
        title=(w.get("title") or w.get("display_name") or "").strip(),
        authors=authors,
        year=w.get("publication_year"),
        journal=loc.get("display_name") or "",
        doi=(w.get("doi") or "").replace("https://doi.org/", ""),
        url=w.get("doi") or (w.get("id") or ""),
        abstract=abstract[:1200],
        cited_by=w.get("cited_by_count"),
        query=query,
        openalex_id=(w.get("id") or "").rsplit("/", 1)[-1],
        volume=biblio.get("volume") or "",
        issue=biblio.get("issue") or "",
        pages="%s-%s" % (biblio.get("first_page"), biblio.get("last_page"))
        if biblio.get("first_page") and biblio.get("last_page") else "",
    )


def search(query, limit=5, year_from=None, require_doi=False, types=("article", "review"),
           only_oa=False):
    """在 OpenAlex 里按标题/摘要检索。"""
    if not query or not query.strip():
        return []
    filt = "title_and_abstract.search:" + query.strip()
    if year_from:
        filt += ",from_publication_date:%d-01-01" % int(year_from)
    if require_doi:
        filt += ",has_doi:true"
    if only_oa:
        filt += ",open_access.is_oa:true"
    if types:
        filt += ",type:" + "|".join(types)
    url = "%s?%s" % (OPENALEX, urllib.parse.urlencode(
        {"filter": filt, "per_page": str(max(1, min(limit * 3, 50))), "mailto": MAILTO}))
    try:
        data = _get(url)
    except Exception:
        # OpenAlex 今天额度用完了 → 换 Crossref 顶一下
        alt = search_crossref(query, limit, year_from, require_doi)
        return [_with_query(p, query) for p in alt]
    if not (data.get("results") or []):
        alt = search_crossref(query, limit, year_from, require_doi)
        if alt:
            return [_with_query(p, query) for p in alt]
    out = []
    for w in data.get("results") or []:
        p = _parse_work(w, query)
        if not p.title:
            continue
        out.append(p)
    return out


CROSSREF = "https://api.crossref.org/works"


def search_crossref(query, limit=5, year_from=None, require_doi=True):
    """OpenAlex 限流/不可用时的后备：Crossref（免费、无需密钥）。"""
    if not query or not query.strip():
        return []
    params = {
        "query.bibliographic": query.strip(),
        "rows": str(max(1, min(limit * 2, 40))),
        "mailto": MAILTO,
    }
    filters = ["type:journal-article"]
    if year_from:
        filters.append("from-pub-date:%d-01-01" % int(year_from))
    params["filter"] = ",".join(filters)
    url = "%s?%s" % (CROSSREF, urllib.parse.urlencode(params))
    try:
        data = _get(url)
    except Exception:
        return []
    out = []
    for w in (data.get("message") or {}).get("items") or []:
        title = " ".join(w.get("title") or []).strip()
        if not title:
            continue
        doi = w.get("DOI") or ""
        if require_doi and not doi:
            continue
        authors = []
        for a in w.get("author") or []:
            name = " ".join(x for x in (a.get("given"), a.get("family")) if x).strip()
            if name:
                authors.append(name)
        parts = (w.get("issued") or {}).get("date-parts") or [[None]]
        year = (parts[0] or [None])[0]
        journal = " ".join(w.get("container-title") or []).strip()
        out.append(Paper(
            title=title, authors=authors, year=year, journal=journal, doi=doi,
            url=("https://doi.org/" + doi) if doi else "",
            abstract=_strip_tags(w.get("abstract") or "")[:1200],
            cited_by=w.get("is-referenced-by-count"), query=query,
            openalex_id="", volume=w.get("volume") or "", issue=w.get("issue") or "",
            pages=(w.get("page") or ""), is_oa=False, pdf_url=""))
    return out


def _with_query(paper, query):
    paper.query = query
    return paper


def _strip_tags(text):
    import re as _re

    return _re.sub(r"<[^>]+>", " ", text or "")


def _relevance(paper, zh_terms, en_terms, core=None):
    """标题命中权重更高；命中「具体术语」再加分。"""
    if paper.doi and paper.doi.startswith("10.5281"):   # Zenodo 数据集不是论文
        return 0
    title = (paper.title or "").lower()
    abstract = (paper.abstract or "").lower()
    terms = [t.lower() for t in list(zh_terms) + list(en_terms) if t]
    t_hits = sum(1 for t in terms if t in title)
    a_hits = sum(1 for t in terms if t in abstract)
    core_hits = sum(1 for c in (core or []) if c and c.lower() in title)
    return t_hits * 3 + a_hits + core_hits * 3


def _filter_relevant(papers, zh_terms, en_terms, need=None, core=None):
    """相关度筛选：句子里有核心术语时，论文标题必须命中核心术语。"""
    total = len([t for t in zh_terms if t]) + len([t for t in en_terms if t])
    if need is None:
        need = 5 if total >= 2 else 1     # 标题命中具体术语，或标题命中两个泛词
    scored = []
    for p in papers:
        title = (p.title or "").lower()
        if core:
            hits = [c for c in core if c and c.lower() in title]
            if not hits:
                continue
        scored.append((_relevance(p, zh_terms, en_terms, core), p))
    keep = [(s, p) for s, p in scored if s >= need]
    keep.sort(key=lambda x: (-x[0], -(x[1].cited_by or 0)))
    return [p for _s, p in keep]


def _seg_words(text):
    """切词，挑出名词性词。没装 jieba 时退回按标点切。"""
    _jieba_init()
    words = []
    try:
        import jieba.posseg as pseg
        for w, flag in pseg.cut(text):
            w = w.strip()
            if len(w) < 2 or w in STOPWORDS:
                continue
            keep = (flag.startswith(("n", "l", "j")) or flag in ("nz", "vn", "eng", "an")
                    or (flag in ("v", "a") and (w in TERM_EN or w in DOMAIN_WORDS)))
            if keep:
                words.append(w)
    except Exception:
        for w in re.split(r"[，。、；：（）\s]+", text):
            if 2 <= len(w) <= 8 and w not in STOPWORDS:
                words.append(w)
    return words


def keywords(text, limit=4):
    """从一句话里挑检索词：领域词优先，其次长词。"""
    words = _seg_words(text)
    uniq = []
    for w in words:
        if w not in uniq:
            uniq.append(w)
    rank = {w: (0 if w in CORE_TERMS else 1 if w in TERM_EN else 2, -len(w), uniq.index(w))
            for w in uniq}
    seen, picked = [], []
    for w in sorted(uniq, key=lambda x: rank[x]):
        if any(w in s or s in w for s in seen):
            continue
        picked.append(w)
        seen.append(w)
        if len(picked) >= limit:
            break
    return picked


def to_english(words):
    return [TERM_EN[w] for w in words if w in TERM_EN]


def search_for_text(text, limit=5, year_from=None, require_doi=True, english=True,
                    only_oa=False):
    """一句话 → 论文列表（中文检索 + 英文检索，按核心术语与相关度筛）。"""
    kws = keywords(text, 4)
    en = to_english(kws) if english else []
    core = [w for w in kws if w in SPECIFIC_TERMS or len(w) >= 4]
    core += [TERM_EN[w] for w in kws if w in SPECIFIC_TERMS and w in TERM_EN]
    found = []
    if kws:
        found = search(" ".join(kws), limit * 3, year_from, require_doi, only_oa=only_oa)
    good = _filter_relevant(found, kws, en, core=core)
    if len(good) < limit and en:
        more = search(" ".join(en), limit * 3, year_from, require_doi, only_oa=only_oa)
        seen = {p.key() for p in found}
        more = [p for p in more if p.key() not in seen]
        good += _filter_relevant(more, kws, en, core=core or None)
    seen, out = set(), []
    for p in good:
        if p.key() in seen:
            continue
        seen.add(p.key())
        out.append(p)
    return out[:limit]


# ---------------------------------------------------------------- RIS
def _ris_author(name):
    parts = name.replace(",", " ").split()
    if len(parts) == 1:
        return parts[0]
    return "%s, %s" % (parts[-1], " ".join(parts[:-1]))


def build_ris(papers):
    lines = []
    for p in papers:
        lines.append("TY  - JOUR")
        lines.append("TI  - %s" % (p.title or "").strip())
        for a in p.authors or []:
            lines.append("AU  - %s" % _ris_author(a))
        if p.journal:
            lines.append("JO  - %s" % p.journal)
        if p.year:
            lines.append("PY  - %s" % p.year)
        if p.volume:
            lines.append("VL  - %s" % p.volume)
        if p.issue:
            lines.append("IS  - %s" % p.issue)
        if p.pages:
            lines.append("SP  - %s" % p.pages.split("-")[0])
            if "-" in p.pages:
                lines.append("EP  - %s" % p.pages.split("-")[-1])
        if p.doi:
            lines.append("DO  - %s" % p.doi)
        if p.url:
            lines.append("UR  - %s" % p.url)
        if p.abstract:
            lines.append("AB  - %s" % p.abstract.replace("\n", " "))
        lines.append("ER  - ")
        lines.append("")
    return "\n".join(lines)


def save_ris(papers, folder, stem="引用"):
    folder = os.path.abspath(folder)
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, "%s.ris" % stem)
    text = build_ris(papers)
    with open(path, "w", encoding="utf-8-sig", newline="\r\n") as f:
        f.write(text)
    # 附一份人看的清单
    md = os.path.join(folder, "%s_文献清单.md" % stem)
    with open(md, "w", encoding="utf-8") as f:
        f.write("# 自动检索到的文献（%s）\n\n" % stem)
        f.write("| # | 年份 | 标题 | 期刊 | DOI |\n|---|---|---|---|---|\n")
        for i, p in enumerate(papers, 1):
            f.write("| %d | %s | %s | %s | %s |\n" % (
                i, p.year or "", p.title or "", p.journal or "", p.doi or ""))
        f.write("\n> 数据来源：OpenAlex（https://openalex.org），DOI 可直接解析。\n")
    return path, md


# ---------------------------------------------------------------- 引文格式
def format_mark(number, style, paper=None):
    conf = CITE_STYLES.get(style) or CITE_STYLES["顺序编码-上标[n]"]
    if conf["numeric"]:
        return conf["template"] % number, conf["superscript"]
    author = paper.short_author() if paper else ""
    year = paper.year if paper else ""
    return conf["template"] % (author, year or "n.d."), conf["superscript"]


def _insert_run_after(para, run, text, font=None, superscript=False):
    from docx.oxml import OxmlElement
    from docx.text.run import Run

    new_r = OxmlElement("w:r")
    if font is not None:
        du.set_font(_wrap_run(new_r, para), font)
    run._r.addnext(new_r)
    new_run = Run(new_r, para)
    if superscript:
        new_run.font.superscript = True
    new_run.text = text
    return new_run


def _wrap_run(r_el, para):
    from docx.text.run import Run
    return Run(r_el, para)


def _runs(para):
    return [r for r in para.runs if r.text]


def _split_at(para, offset):
    """把段落按字符偏移切开，返回该偏移所在的 run（切分后 offset 之前的 run）。"""
    acc = 0
    for run in _runs(para):
        n = len(run.text)
        if acc + n > offset:
            cut = offset - acc
            if cut <= 0:
                return run
            import copy as _copy
            tail_r = _copy.deepcopy(run._r)
            run._r.addnext(tail_r)
            from docx.text.run import Run
            tail = Run(tail_r, para)
            head_text, tail_text = run.text[:cut], run.text[cut:]
            run.text = head_text
            tail.text = tail_text
            return run
        acc += n
    return run if _runs(para) else None


# ---------------------------------------------------------------- 位置规划
def _is_body_role(role_name):
    return role_name in ("正文", "正文续", "列表项", "表内文字") or "正文" in role_name


def plan_spots(doc, profile, classify_fn=None, opts=None):
    """找出该加引用的句子：返回 [(段落, 句子, 句末偏移)]。"""
    conf = dict(CITE_DEFAULTS)
    conf.update(opts or {})
    spots = []
    triggers = conf["triggers"]
    started = not conf.get("start_after_heading", True)
    for p in du.iter_paragraphs(doc):
        if du.in_table(p) or du.has_picture(p):
            continue
        role = classify_fn(p) if classify_fn is not None else None
        role_name = getattr(role, "name", "") or ""
        if not started:
            if role_name in ("一级标题", "Heading 1"):
                started = True
            continue
        if role_name in ("一级标题", "二级标题", "三级标题", "Heading 1", "Heading 2", "Heading 3"):
            continue
        text = du.text_of(p)
        if not text.strip() or len(text.strip()) < conf["min_sentence_len"]:
            continue
        if re.search(r"\[\d+\]", text):
            continue   # 这一段已经有引用标记了，别重复堆
        if classify_fn is not None and conf["only_body"]:
            if not _is_body_role(role_name):
                continue
        sents = []
        start = 0
        for m in SPLIT_SENT.finditer(text):
            sents.append((start, m.end()))
            start = m.end()
        if start < len(text):
            sents.append((start, len(text)))
        para_spots = []
        for s, e in sents:
            sent = text[s:e].strip()
            body = sent.rstrip("。；！？")
            if len(body) < conf["min_sentence_len"]:
                continue
            # 列标准/法规的句子（《GB …》）是在引标准，不适合配研究论文
            if "《" in body or re.search(r"GB\s*\d|GB/T", body):
                continue
            if triggers:
                if not any(t in body for t in triggers):
                    continue
            else:
                # 自动模式：句子里得有领域术语，或者像在陈述结论
                has_term = any(t in body for t in TERM_EN) or any(t in body for t in TRIGGER_WORDS)
                has_number = bool(re.search(r"\d", body))
                if not (has_term or has_number):
                    continue
            para_spots.append((p, body, s + len(sent.rstrip())))
        if conf["skip_first_para"] and len(para_spots) > 1:
            para_spots = para_spots[1:]
        spots.extend(para_spots[:conf["max_per_para"]])
        if len(spots) >= conf["count"]:
            break
    return spots[:conf["count"]]


# ---------------------------------------------------------------- 参考文献
def gbt_entry(number, paper, style=None, template=None):
    """按体例（默认 GB/T 7714）排一条参考文献。"""
    from . import citestyle

    return citestyle.format_entry(paper, number, style, template)


def _gbt_entry_legacy(number, paper):
    """GB/T 7714 顺序编码制的一条（旧写法，留作对照）。"""
    authors = ", ".join(paper.authors[:3]) if paper.authors else ""
    if paper.authors and len(paper.authors) > 3:
        authors += ", 等"
    bits = [authors, paper.title or ""]
    if paper.journal:
        bits.append(paper.journal)
    tail = ""
    if paper.year:
        tail += str(paper.year)
    if paper.volume:
        tail += ", %s" % paper.volume
    if paper.issue:
        tail += "(%s)" % paper.issue
    if paper.pages:
        tail += ": %s" % paper.pages
    tail = tail.strip(", ")
    text = "[%d] %s[J]. %s" % (number, ". ".join(b for b in bits if b), tail)
    if paper.doi:
        text += ". DOI: %s" % paper.doi
    return text.rstrip(". ") + "."


def append_references(doc, profile, papers, start_number=1, style_entry=gbt_entry):
    """把文献按顺序补到「参考文献」那一段后面，格式沿用该段。"""
    from docx.oxml.ns import qn as _qn

    try:
        from . import refs as _refs

        if _refs.find_reference_heading(doc, profile) is not None:
            texts = [style_entry(start_number + i, paper) for i, paper in enumerate(papers)]
            return _refs.insert_entries(doc, profile, texts, start_number=start_number,
                                        numbering=False)
    except Exception:
        pass

    ref_role = profile.role("参考文献条目") if profile else None
    entries = [p for p in du.iter_paragraphs(doc)
               if du.text_of(p).strip() and (ref_role and p.style is not None
                                            and p.style.name == ref_role.style)]
    anchor = entries[-1] if entries else None
    added = []
    for i, paper in enumerate(papers):
        text = style_entry(start_number + i, paper)
        if anchor is not None:
            new_p = _clone_after(anchor, text)
            anchor = new_p
        else:
            new_p = doc.add_paragraph(text)
            anchor = new_p
        added.append(new_p)
    return added


def _clone_after(para, text):
    import copy as _copy

    from docx.text.paragraph import Paragraph

    el = _copy.deepcopy(para._p)
    for child in list(el):
        if child.tag != qn("w:pPr"):
            el.remove(child)
    para._p.addnext(el)
    new_p = Paragraph(el, para._parent)
    new_p.add_run(text)
    return new_p


# ---------------------------------------------------------------- 总流程
def cite_document(src, dst, profile, opts=None, ris_dir=None, progress=None):
    """检索 → 在正文句子末尾插引用 → 补参考文献 → 导出 RIS。"""
    import os as _os

    from docx import Document

    from .classify import Classifier

    conf = dict(CITE_DEFAULTS)
    conf.update(opts or {})
    doc = Document(src)
    clf = Classifier(doc, profile)

    want = int(conf["count"])
    # 文档里已经有序号 [1]…[7] 时，接着往后编，不能从 1 重来
    start_no = 1
    for p in du.iter_paragraphs(doc):
        m = re.match(r"\s*\[(\d+)\]", du.text_of(p))
        if m:
            start_no = max(start_no, int(m.group(1)) + 1)
    spots = plan_spots(doc, profile, classify_fn=lambda p: clf.match(p),
                       opts=dict(conf, count=max(want * 5, 8), max_per_para=conf["max_per_para"]))

    topics = conf.get("topics") or ""
    picked = []      # (段落, 句末偏移, 序号, 句子, Paper)
    used = []
    for para, sent, end in spots:
        if len(picked) >= want:
            break
        if progress:
            progress("检索：%s" % sent[:24])
        if topics.strip():
            # 用户指定主题：按主题检索，再挑和这句话最贴近的那篇
            kw = [t for t in re.split(r"[\s,，;；]+", topics.strip()) if t]
            found = search(" ".join(kw), 20, conf["year_from"], conf["require_doi"],
                           only_oa=bool(conf.get("only_oa")))
            papers = _filter_relevant(found, [], kw)
            seen_key, uniq = set(), []
            for _p in papers:
                if _p.key() in seen_key:
                    continue
                seen_key.add(_p.key())
                uniq.append(_p)
            papers = uniq
            # 标题里至少得有一个主题词，避免摘要里偶然出现 food 的无关文章混进来
            low_kw = [k.lower() for k in kw] + [k.lower() for k in to_english(keywords(topics, 6))]
            papers = [p for p in papers
                      if any(k in (p.title or "").lower() for k in low_kw if k)]
            sent_kw = keywords(sent, 4) + to_english(keywords(sent, 4))
            if sent_kw:
                papers.sort(key=lambda p: -_relevance(p, [], sent_kw))
                # 和这句话完全不相干的就不要了，宁可不引
                papers = [p for p in papers if _relevance(p, [], sent_kw) >= 3]
        else:
            papers = search_for_text(sent, 3, conf["year_from"], conf["require_doi"],
                                     only_oa=bool(conf.get("only_oa")))
        keys = {p.key() for p in used}
        paper = next((p for p in papers if p.key() not in keys), None)
        if paper is None:
            continue
        used.append(paper)
        picked.append((para, end, start_no + len(used) - 1, sent, paper))

    # 从后往前插，前面句子的偏移才不会被影响
    by_para = {}
    for para, end, num, sent, paper in picked:
        by_para.setdefault(id(para), (para, []))[1].append((end, num, paper))
    for para, items in by_para.values():
        for end, num, paper in sorted(items, key=lambda x: -x[0]):
            mark, sup = format_mark(num, conf["style"], paper)
            run = _split_at(para, end)
            if run is not None:
                _insert_run_after(para, run, mark, font=None, superscript=sup)

    if used:
        append_references(doc, profile, used, start_no)
    doc.save(dst)

    ris_path = md_path = ""
    if used:
        folder = ris_dir or _os.path.join(_os.path.dirname(_os.path.abspath(dst)), "文献库")
        stem = _os.path.splitext(_os.path.basename(dst))[0] + "_引用"
        ris_path, md_path = save_ris(used, folder, stem)
    return {"papers": used, "picked": picked, "ris": ris_path, "md": md_path}
