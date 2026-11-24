# -*- coding: utf-8 -*-
"""混合识别：样式名 + 内容正则，按「可信度打分」取最合适的角色。

打分规则：
  · 整段完全命中某个标题/题注正则（如「参考文献」「摘要」「图1-1 xxx」）→ 40 分
  · 段落样式正好是某角色的样式 → 30 分
  · 正则只是部分命中 → 22 分
  · 命中「正文」这类兜底样式 → 10 分（所以样式没设对时，内容规则还能救回来）
分数相同则按 YAML 里的先后顺序取靠前的。
表格内文字、带 PAGEREF 域的目录行另有结构规则优先处理。
"""
from __future__ import annotations

import re
from collections import Counter
from typing import Dict, List, Optional, Tuple

from . import docxutil as du
from .profile import Profile, Role

from docx.oxml.ns import qn

S_FULL_TITLE = 45   # 整段命中「摘要/目录/参考文献」这类独占一行的标题特征
S_STYLE = 30        # 样式名正好对得上
S_REGEX = 22        # 内容正则命中
S_GENERIC = 10      # 只命中了「正文」这类兜底样式

# 这些角色独占一行，整段命中时可信度最高
TITLE_ROLES = {"摘要标题", "英文摘要标题", "目录标题", "参考文献标题",
               "致谢标题", "附录标题"}


class Classifier:
    def __init__(self, doc, profile: Profile):
        self.doc = doc
        self.profile = profile
        self.roles: List[Role] = profile.enabled_roles()
        self.fallback_name = profile.option("fallback_role", "正文")
        self._rx: List[Tuple[Role, List["re.Pattern"]]] = []
        for r in self.roles:
            pats = []
            for pat in r.match_regex or []:
                try:
                    pats.append(re.compile(pat, re.I))
                except re.error as e:
                    raise ValueError("角色「%s」的正则写错了：%s（%s）" % (r.name, pat, e))
            self._rx.append((r, pats))
        self.by_name = {du.norm_key(r.name): r for r in self.roles}

    # ------------------------------------------------------------ 主匹配
    def match(self, p) -> Optional[Role]:
        style_key = du.norm_key(du.p_style_name(self.doc, p))
        text = du.text_of(p).strip()

        # 1) 表格里的文字统一归「表内文字」
        if du.in_table(p):
            r = self.by_name.get(du.norm_key("表内文字"))
            if r is not None:
                return r

        # 2) 目录行（带 PAGEREF 域）按目录条目处理
        instrs = [i.strip().upper() for i in du.field_instrs(p)]
        if any(i.startswith("PAGEREF") for i in instrs):
            for role, _pats in self._rx:
                if role.name.startswith("目录条目") and role.match_styles and \
                        any(du.norm_key(s) == style_key for s in role.match_styles):
                    return role
            for role, _pats in self._rx:
                if role.name.startswith("目录条目"):
                    return role

        # 3) 打分
        best: Optional[Role] = None
        best_score = -1
        for role, pats in self._rx:
            score = -1
            if role.match_styles and any(du.norm_key(s) == style_key for s in role.match_styles):
                score = S_GENERIC if role.name == self.fallback_name else S_STYLE
            if pats and text:
                for pat in pats:
                    if pat.fullmatch(text):
                        score = max(score, S_FULL_TITLE if role.name in TITLE_ROLES else S_REGEX)
                        break
                    if (role.regex_mode or "search").lower() != "fullmatch" and pat.search(text):
                        score = max(score, S_REGEX)
            if score > best_score:
                best, best_score = role, score

        if best is not None:
            return best

        # 4) 结构兜底
        if not text and du.has_picture(p):
            r = self.by_name.get(du.norm_key("图片"))
            if r is not None:
                return r
        if self.fallback_name:
            return self.by_name.get(du.norm_key(self.fallback_name))
        return None

    # ------------------------------------------------------------ 封面
    def cover_count(self) -> int:
        """封面占前多少个段落。

        优先用「分节符」判断：封面通常单独一节，节结束处就是封面终点。
        这样摘要页上方的题目不会被误当成封面而漏掉。
        没有分节符时退回「摘要/Abstract/目录 前面」的正则判断。
        """
        if not self.profile.option("cover_keep", True):
            return 0
        pat_text = self.profile.option("cover_end_regex", "")
        if not pat_text:
            return 0
        try:
            pat = re.compile(pat_text, re.I)
        except re.error:
            return 0

        paras = [p for p in du.iter_paragraphs(self.doc)]
        anchor = None
        for i, p in enumerate(paras):
            if du.in_table(p):
                continue
            if pat.search(du.text_of(p).strip()):
                anchor = i
                break
        if anchor is None:
            return 0
        # 往前找最近的分节符段落，封面到此为止
        for j in range(anchor - 1, -1, -1):
            pPr = paras[j]._p.find(qn("w:pPr"))
            if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
                return j + 1
        return anchor

    # ------------------------------------------------------------ 计划
    def plan(self) -> List[Tuple[object, Optional[Role], str]]:
        skip = self.cover_count()
        out = []
        for i, p in enumerate(du.iter_paragraphs(self.doc)):
            if i < skip:
                out.append((p, None, "封面-跳过"))
                continue
            role = self.match(p)
            reason = ""
            if role is not None:
                style = du.p_style_name(self.doc, p)
                if any(du.norm_key(s) == du.norm_key(style) for s in role.match_styles):
                    reason = "样式"
                elif du.in_table(p):
                    reason = "表格"
                else:
                    reason = "内容"
            out.append((p, role, reason))
        self._mark_front_title(out, skip)
        return out

    def _mark_front_title(self, entries, skip: int) -> None:
        """摘要/Abstract 前紧挨着的短行 = 论文题目（原文里常是手工居中，没有样式）。"""
        anchor_roles = {"摘要标题": "论文题目", "英文摘要标题": "英文论文题目"}
        for idx, (_p, role, _r) in enumerate(entries):
            if role is None or role.name not in anchor_roles:
                continue
            title_role = self.by_name.get(du.norm_key(anchor_roles[role.name])) or \
                self.by_name.get(du.norm_key("论文题目"))
            if title_role is None:
                continue
            for j in range(idx - 1, max(skip - 1, idx - 5), -1):
                p, r, _reason = entries[j]
                text = du.text_of(p).strip()
                if not text:
                    continue
                if r is not None and r.name in anchor_roles:
                    break
                if len(text) > 120 or text.endswith(("。", "；", ";")):
                    break
                if r is None or r.name == self.fallback_name:
                    entries[j] = (p, title_role, "结构")
                break

    def stats(self, plan) -> "Counter":
        c = Counter()
        for _p, role, _r in plan:
            c[role.name if role else "封面/跳过"] += 1
        return c
