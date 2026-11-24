# -*- coding: utf-8 -*-
"""参考文献格式：内置多种常见体例，也支持自定义模板 + 从 EndNote 样式文件取名字和说明。

模板占位符（花括号里写这些）：
    {seq}      序号             {authors}  作者（按体例自动排）
    {title}    题名             {journal}  刊名
    {year}     年份             {volume}   卷
    {issue}    期               {pages}    页码
    {doi}      DOI              {url}      网址
    {type}     文献类型代码（J/M/C/D…）
    {publisher} 出版者          {edition}  版次
    {accessed} 访问日期
"""
import os
import re

from . import paths

CUSTOM_FILE = os.path.join(paths.CONFIG_DIR, "参考文献格式.yaml")
CUSTOM_DEFAULT_TEMPLATE = "[{seq}] {authors}. {title}[{type}]. {journal}, {year}, {volume}({issue}): {pages}."

# 作者排法
AUTHOR_STYLES = {
    "gbt": "gbt",        # 张三, 李四, 等
    "apa": "apa",        # Zhang, S., & Li, S.
    "ieee": "ieee",      # S. Zhang
    "nature": "nature",  # Zhang S
    "plain": "plain",    # 原样
}

# 内置体例：名称 → 模板
STYLES = {
    "GB/T 7714（顺序编码）": {
        "template": "[{seq}] {authors}. {title}[{type}]. {journal}, {year}, {volume}({issue}): {pages}. DOI: {doi}.",
        "authors": "gbt",
    },
    "APA 7th": {
        "template": "{authors} ({year}). {title}. {journal}, {volume}({issue}), {pages}. https://doi.org/{doi}",
        "authors": "apa",
    },
    "MLA 9th": {
        "template": "{authors}. \"{title}.\" {journal}, vol. {volume}, no. {issue}, {year}, pp. {pages}.",
        "authors": "plain",
    },
    "Chicago（作者-年份）": {
        "template": "{authors}. {year}. \"{title}.\" {journal} {volume}, no. {issue}: {pages}.",
        "authors": "plain",
    },
    "IEEE": {
        "template": "[{seq}] {authors}, \"{title},\" {journal}, vol. {volume}, no. {issue}, pp. {pages}, {year}, doi: {doi}.",
        "authors": "ieee",
    },
    "Vancouver": {
        "template": "{authors}. {title}. {journal}. {year};{volume}({issue}):{pages}.",
        "authors": "gbt",
    },
    "Harvard": {
        "template": "{authors} ({year}) '{title}', {journal}, {volume}({issue}), pp. {pages}.",
        "authors": "plain",
    },
    "Nature": {
        "template": "{authors}. {title}. {journal} {volume}, {pages} ({year}).",
        "authors": "nature",
    },
    "ACS": {
        # 作者用分号隔开，年份后接卷、页码；期刊名一般用缩写
        "template": "{authors}. {title}. {journal} {year}, {volume}, {pages}. DOI: {doi}.",
        "authors": "acs",
    },
}

# 参考文献条目里默认不写 DOI（学校模板一般不要求）；要写就传 include_doi=True
DEFAULT_INCLUDE_DOI = False

TYPE_MARK = {"JOUR": "J", "BOOK": "M", "CONF": "C", "THES": "D", "RPRT": "R",
             "ELEC": "EB", "CHAP": "M", "GEN": "Z"}


def _split_person(name):
    """把「Zhang San」或「张三」拆成 (姓, 名)。"""
    name = (name or "").strip()
    if not name:
        return "", ""
    if "," in name:
        last, first = name.split(",", 1)
        return last.strip(), first.strip()
    parts = name.split()
    if len(parts) == 1:
        # 中文名：第一个字当姓
        return name[:1], name[1:]
    return parts[-1], " ".join(parts[:-1])


def format_authors(names, style="gbt"):
    names = [n for n in (names or []) if n and n.strip()]
    if not names:
        return "佚名"
    if style == "apa":
        out = []
        for n in names:
            last, first = _split_person(n)
            initials = " ".join(p[0] + "." for p in re.split(r"[\s\-]+", first) if p)
            out.append(("%s, %s" % (last, initials)).strip(", "))
        if len(out) > 1:
            return ", ".join(out[:-1]) + ", & " + out[-1]
        return out[0]
    if style == "ieee":
        out = []
        for n in names:
            last, first = _split_person(n)
            initials = " ".join(p[0] + "." for p in re.split(r"[\s\-]+", first) if p)
            out.append((initials + " " + last).strip())
        return ", ".join(out)
    if style == "acs":
        out = []
        for n in names[:10]:
            last, first = _split_person(n)
            initials = " ".join(p[0] + "." for p in re.split(r"[\s\-]+", first) if p)
            out.append(("%s, %s" % (last, initials)).strip(", "))
        if len(names) > 10:
            out.append("et al.")
        return "; ".join(out)
    if style == "nature":
        out = []
        for n in names:
            last, first = _split_person(n)
            initials = " ".join(p[0] for p in re.split(r"[\s\-]+", first) if p)
            out.append((last + " " + initials).strip())
        return ", ".join(out)
    # gbt / plain：超过 3 位用「等」
    shown = names if len(names) <= 3 else names[:3]
    text = ", ".join(shown)
    if len(names) > 3:
        text += ", 等"
    return text


def _clean(text):
    """去掉多余标点（模板里某个字段为空时会留下 ", ," 这种）。"""
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"\s*,\s*(?=,|\.|;|:)", "", text)
    text = re.sub(r"\.\s*\.", ".", text)
    text = re.sub(r",\s*\)", ")", text)
    text = re.sub(r"\(\s*\)", "", text)
    text = re.sub(r"\s+([,.;:])", r"\1", text)
    text = re.sub(r"([,;:])\s*(?=[,;:.])", "", text)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip(" ,;:")


# ---------------------------------------------------------------- 自定义格式
def _read_yaml(path):
    try:
        import yaml
    except Exception:
        return {}
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


def _write_yaml(path, data):
    import yaml

    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        yaml.safe_dump(data, f, allow_unicode=True, sort_keys=False)


def load_custom():
    """读取用户自己新建的格式：{名字: {template, authors}}。"""
    data = _read_yaml(CUSTOM_FILE)
    out = {}
    for name, conf in data.items():
        if isinstance(conf, str):
            out[name] = {"template": conf, "authors": "gbt"}
        elif isinstance(conf, dict) and conf.get("template"):
            out[name] = {"template": conf["template"],
                         "authors": conf.get("authors") or "gbt"}
    return out


def save_custom(mapping):
    _write_yaml(CUSTOM_FILE, mapping)


def add_custom(name, template, authors="gbt"):
    """新建（或覆盖同名）自定义格式。"""
    name = (name or "").strip()
    template = (template or "").strip()
    if not name:
        raise ValueError("格式名不能为空")
    if not template:
        raise ValueError("模板不能为空")
    if name in STYLES:
        raise ValueError("「%s」是内置格式，换个名字" % name)
    data = load_custom()
    data[name] = {"template": template, "authors": authors}
    save_custom(data)
    return name


def rename_custom(old, new):
    """重命名自定义格式。"""
    old = (old or "").strip()
    new = (new or "").strip()
    data = load_custom()
    if old not in data:
        raise ValueError("「%s」不是自定义格式，改不了名" % old)
    if not new:
        raise ValueError("新名字不能为空")
    if new in STYLES or (new in data and new != old):
        raise ValueError("「%s」已经存在了，换个名字" % new)
    items = []
    for name, conf in data.items():
        items.append((new if name == old else name, conf))
    save_custom(dict(items))
    return new


def delete_custom(name):
    data = load_custom()
    if name in data:
        del data[name]
        save_custom(data)
        return True
    return False


def resolve(name):
    """按名字找格式：内置优先，再找自定义。"""
    if name in STYLES:
        return STYLES[name]
    custom = load_custom()
    if name in custom:
        conf = dict(custom[name])
        conf.setdefault("authors", "gbt")
        return conf
    return STYLES["GB/T 7714（顺序编码）"]


def all_names():
    """界面下拉里要显示的全部格式名（内置 + 自定义）。"""
    return list(STYLES) + list(load_custom())


def entry_fields(paper, seq):
    kind = "J"
    if getattr(paper, "doi", "") and str(paper.doi).startswith("10.5281"):
        kind = "D"
    pages = getattr(paper, "pages", "") or ""
    return {
        "seq": str(seq),
        "title": (paper.title or "").strip(),
        "journal": (paper.journal or "").strip(),
        "year": str(paper.year or ""),
        "volume": (paper.volume or "").strip(),
        "issue": (paper.issue or "").strip(),
        "pages": pages.strip(),
        "doi": (paper.doi or "").strip(),
        "url": (paper.url or "").strip(),
        "type": kind,
        "publisher": getattr(paper, "publisher", "") or "",
        "edition": getattr(paper, "edition", "") or "",
        "accessed": "",
    }


_DOI_PAT = re.compile(r"\s*\.?\s*(DOI|doi)[:：]\s*\S+\.?\s*$")
_DOI_URL = re.compile(r"\s*https?://doi\.org/\S+\s*")


def strip_doi(text):
    """把条目末尾的 DOI 段去掉（连同它前面的句号），保持以句号结尾。"""
    t = _DOI_URL.sub("", text or "").strip()
    t = _DOI_PAT.sub("", t).strip()
    t = re.sub(r"[\s,，;；:：.]+$", "", t)
    return (t + ".") if t else t


def format_entry(paper, seq, style=None, template=None, include_doi=None):
    """按体例或自定义模板排一条参考文献。"""
    conf = resolve(style or "GB/T 7714（顺序编码）")
    tpl = template or conf.get("template") or CUSTOM_DEFAULT_TEMPLATE
    fields = entry_fields(paper, seq)
    fields["authors"] = format_authors(getattr(paper, "authors", []) or [], conf.get("authors", "gbt"))
    out = tpl
    for key, val in fields.items():
        out = out.replace("{%s}" % key, val or "")
    out = _clean(out)
    want_doi = DEFAULT_INCLUDE_DOI if include_doi is None else bool(include_doi)
    if not want_doi:
        out = strip_doi(out)
    return out


# ---------------------------------------------------------------- EndNote 样式
def read_ens_strings(path):
    """EndNote 的 .ens 是二进制，只能尽力抽出可读的字符串。"""
    with open(path, "rb") as f:
        data = f.read()
    out = []
    for m in re.finditer(rb"(?:[\x20-\x7e]\x00){4,}", data):
        try:
            s = m.group().decode("utf-16-le").strip()
        except Exception:
            continue
        if s:
            out.append(s)
    for m in re.finditer(rb"[\x20-\x7e]{8,}", data):
        s = m.group().decode("latin-1").strip()
        if s:
            out.append(s)
    seen, uniq = set(), []
    for s in out:
        if s in seen:
            continue
        seen.add(s)
        uniq.append(s)
    return uniq


def read_ens(path):
    """读取 EndNote 样式文件，返回 {name, description, samples}。"""
    strings = read_ens_strings(path)
    name = os.path.splitext(os.path.basename(path))[0]
    desc = ""
    samples = []
    for s in strings:
        if not desc and len(s) > 40 and s.count(" ") > 5 and "." in s:
            desc = s
        if any(k in s for k in ("et al", "eds", "Ed.", "pp.", "vol.")):
            if len(s) < 60:
                samples.append(s)
    if strings:
        cand = strings[0]
        if len(cand) < 60 and not desc:
            name = cand or name
    return {"name": name, "description": desc, "samples": samples[:8], "path": path}
