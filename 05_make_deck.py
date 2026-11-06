#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
outline.md  ->  pptx（开智风格 · 多版式）

用法：
    python 05_make_deck.py outline.md -o deck.pptx

outline.md 语法
---------------
    ## P12 · 页面标题 {layout}
    - 主条目
      - 小字说明（缩进两个空格）
    > 讲稿：写进备注

layout（不写会自动判断）：
    cover      封面
    statement  一句话大页（可用 [[ ]] 加描边框）
    section    章节页（大编号 + 标题 + 引导语）
    list       编号列表（1-4 条，可带小字）
    cards      三栏卡片
    twocol     左观点 + 右三条说明
    stats      三个数据块
    timeline   横向时间轴（4 节点）
    compare    左右对照
    quote      引言页
    dense      两栏论述页（左正文 + 右小字）
    closing    结尾页

标题或条目里用「」包住的文字会变成暖橙色强调。
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

import kz

# ---------------- 解析 ----------------

HEAD_RE = re.compile(r"^##\s*P?(\d+)?\s*[·.]?\s*(.*?)\s*(?:\{(\w[\w-]*)\})?\s*$")
NOTE_RE = re.compile(r"^>\s*讲稿[:：]\s*(.*)$")
BULLET_RE = re.compile(r"^(\s*)[-*]\s+(.*)$")
META_RE = re.compile(r"^[-*]\s*(总页数|来源)[:：]")


class Item:
    def __init__(self, text: str):
        self.text = text
        self.subs: list[str] = []


class Page:
    def __init__(self, no: int, title: str, layout: str = ""):
        self.no = no
        self.title = title
        self.layout = layout
        self.items: list[Item] = []
        self.notes = ""

    @property
    def bullets(self) -> list[str]:
        return [i.text for i in self.items]


def strip_md(text: str) -> str:
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"`(.+?)`", r"\1", text)
    return text.strip()


def parse(path: Path) -> tuple[str, list[Page]]:
    deck_title = path.stem
    pages: list[Page] = []
    cur: Page | None = None

    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.rstrip()
        if not line or line.strip() == "---":
            continue

        if line.startswith("# ") and not line.startswith("## "):
            t = strip_md(line[2:].strip())
            for pref in ("幻灯片大纲：", "幻灯片大纲:", "大纲：", "大纲:"):
                if t.startswith(pref):
                    t = t[len(pref):]
            deck_title = t.strip() or deck_title
            continue

        m = HEAD_RE.match(line)
        if m:
            no = int(m.group(1)) if m.group(1) else len(pages) + 1
            cur = Page(no, strip_md(m.group(2)), (m.group(3) or "").strip())
            pages.append(cur)
            continue

        if META_RE.match(line) or cur is None:
            continue

        m = NOTE_RE.match(line)
        if m:
            cur.notes = (cur.notes + " " + m.group(1)).strip()
            continue

        m = BULLET_RE.match(line)
        if m:
            indent = len(m.group(1).replace("\t", "  "))
            body = strip_md(m.group(2))
            if not body:
                continue
            if indent >= 2 and cur.items:
                cur.items[-1].subs.append(body)
            else:
                cur.items.append(Item(body))
            continue

        body = strip_md(line)
        if body:
            if cur.items and line.startswith(" "):
                cur.items[-1].subs.append(body)
            else:
                cur.items.append(Item(body))

    return deck_title, pages


# ---------------- 文本工具 ----------------

def rich_runs(p, text: str, *, size, color=kz.INK, emph=kz.ORANGE,
              bold=True, italic=False, latin=kz.CN_SANS):
    """把「」包住的片段渲染成强调色。"""
    for seg in re.split(r"(「[^」]*」)", text):
        if not seg:
            continue
        r = p.add_run()
        r.text = seg
        kz.set_font(r, size=size, latin=latin, ea=latin, bold=bold,
                    italic=italic,
                    color=emph if seg.startswith("「") else color)


def rich_para(tf, text, *, size, first=False, line=1.2, color=kz.INK,
              bold=True, space_before=0):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.line_spacing = line
    p.space_before = Pt(space_before)
    rich_runs(p, text, size=size, color=color, bold=bold)
    return p


def plain_para(tf, text, *, size, first=False, line=1.35, color=kz.GRAY,
               italic=False, space_before=0):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.line_spacing = line
    p.space_before = Pt(space_before)
    r = p.add_run()
    r.text = text
    kz.set_font(r, size=size, latin=kz.CN_SANS, ea=kz.CN_SANS,
                color=color, italic=italic)
    return p


def is_framed(t: str) -> bool:
    return t.startswith("[[") and t.endswith("]]")


def unframe(t: str) -> str:
    return t[2:-2].strip()


def head(slide, title: str, *, y=Inches(1.75), size=kz.SIZE_MAIN, width=Inches(11.2)):
    """标准页头：橙线 + 标题。"""
    kz.rule(slide, kz.MARGIN, Inches(1.25), Inches(0.7), color=kz.ORANGE)
    tb = kz.add_box(slide, kz.MARGIN, y, width, Inches(1.0))
    rich_para(tb.text_frame, title, size=size, first=True)
    return tb


# ---------------- 版式 ----------------

def L_cover(prs, page: Page, deck_title: str):
    s = kz.new_slide(prs)
    kz.page_label(s, "PRESENTATION")
    kz.rule(s, kz.MARGIN, Inches(1.25), Inches(1.1))
    tb = kz.add_box(s, kz.MARGIN, Inches(2.75), Inches(10.8), Inches(2.6))
    rich_para(tb.text_frame, page.title or deck_title, size=kz.SIZE_COVER,
              first=True, line=1.1)
    for i, it in enumerate(page.items[:2]):
        tb2 = kz.add_box(s, kz.MARGIN, Inches(4.45) + Inches(0.55) * i,
                         Inches(10.8), Inches(0.6))
        ascii_ratio = sum(c.isascii() for c in it.text) / max(len(it.text), 1)
        if ascii_ratio > 0.8:
            plain_para(tb2.text_frame, it.text, size=24, first=True, italic=True)
        else:
            plain_para(tb2.text_frame, it.text, size=17, first=True)
    return s


def L_statement(prs, page: Page, *, size=38):
    s = kz.new_slide(prs)
    kz.rule(s, kz.MARGIN, Inches(1.25), Inches(0.7), color=kz.ORANGE)
    t = page.title
    if is_framed(t):
        kz.emphasis_frame(s, unframe(t), x=kz.MARGIN, y=Inches(2.9),
                          size=size, color=kz.ORANGE, style="full", animate=True)
    else:
        tb = kz.add_box(s, kz.MARGIN, Inches(2.9), Inches(10.9), Inches(2.0))
        rich_para(tb.text_frame, t, size=size, first=True, line=1.3)
    if page.items:
        tb2 = kz.add_box(s, kz.MARGIN, Inches(4.9), Inches(10.4), Inches(1.0))
        for i, it in enumerate(page.items[:2]):
            plain_para(tb2.text_frame, it.text, size=14, first=(i == 0),
                       space_before=0 if i == 0 else 8)
    return s


def L_section(prs, page: Page):
    s = kz.new_slide(prs)
    kz.rule(s, kz.MARGIN, Inches(1.25), Inches(0.7), color=kz.ORANGE)
    nb = kz.add_box(s, kz.MARGIN, Inches(2.2), Inches(2.4), Inches(1.4))
    r = nb.text_frame.paragraphs[0].add_run()
    r.text = f"{page.no:02d}"
    kz.set_font(r, size=88, latin=kz.EN_SERIF, ea=kz.CN_SERIF, color=kz.GOLD)
    tb = kz.add_box(s, kz.MARGIN, Inches(3.85), Inches(9.6), Inches(1.2))
    rich_para(tb.text_frame, page.title, size=40, first=True)
    if page.items:
        tb2 = kz.add_box(s, kz.MARGIN, Inches(5.15), Inches(9.6), Inches(1.2))
        for i, it in enumerate(page.items[:2]):
            if i:
                plain_para(tb2.text_frame, it.text, size=15, space_before=10)
            else:
                plain_para(tb2.text_frame, it.text, size=15, first=True)
    return s


def L_list(prs, page: Page):
    s = kz.new_slide(prs)
    head(s, page.title)
    items = page.items[:4]
    gap = Inches(1.45) if len(items) <= 3 else Inches(1.15)
    y = Inches(3.25)
    for i, it in enumerate(items, 1):
        nb = kz.add_box(s, kz.MARGIN, y, Inches(0.75), Inches(0.45))
        p = nb.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = f"{i:02d}"
        kz.set_font(r, size=13, latin=kz.EN_SERIF, ea=kz.CN_SERIF, color=kz.ORANGE)
        tb = kz.add_box(s, kz.MARGIN + Inches(0.8), y - Inches(0.06),
                        Inches(10.2), Inches(1.1))
        tf = tb.text_frame
        rich_para(tf, it.text, size=24, first=True, line=1.15)
        for sub in it.subs[:2]:
            plain_para(tf, sub, size=13, space_before=6)
        y += gap
    return s


def L_cards(prs, page: Page):
    s = kz.new_slide(prs)
    head(s, page.title, y=Inches(1.7))
    items = page.items[:3]
    w, gapx = Inches(3.4), Inches(0.46)
    for i, it in enumerate(items):
        x = kz.MARGIN + (w + gapx) * i
        y = Inches(3.05)
        kz.line(s, x, y - Inches(0.18), x + w, y - Inches(0.18),
                color=kz.ORANGE, width=Pt(2.0))
        tb = kz.add_box(s, x, y, w, Inches(0.7))
        rich_para(tb.text_frame, it.text, size=21, first=True, line=1.15)
        if it.subs:
            tb2 = kz.add_box(s, x, y + Inches(0.85), w, Inches(2.2))
            for j, sub in enumerate(it.subs[:3]):
                plain_para(tb2.text_frame, sub, size=13, first=(j == 0),
                           space_before=0 if j == 0 else 10, line=1.45)
    return s


def L_twocol(prs, page: Page):
    s = kz.new_slide(prs)
    kz.rule(s, kz.MARGIN, Inches(1.25), Inches(0.7), color=kz.ORANGE)
    tb = kz.add_box(s, kz.MARGIN, Inches(2.15), Inches(5.6), Inches(2.4))
    rich_para(tb.text_frame, page.title, size=32, first=True, line=1.25)
    x = kz.MARGIN + Inches(6.2)
    y = Inches(2.15)
    for it in page.items[:3]:
        tb2 = kz.add_box(s, x, y, Inches(4.6), Inches(1.4))
        tf = tb2.text_frame
        rich_para(tf, it.text, size=18, first=True, line=1.2)
        for sub in it.subs[:2]:
            plain_para(tf, sub, size=12, space_before=5, line=1.4)
        y += Inches(1.55)
    return s


def L_stats(prs, page: Page):
    s = kz.new_slide(prs)
    head(s, page.title, y=Inches(1.65))
    y = Inches(2.85)
    for it in page.items[:3]:
        label = it.subs[0] if it.subs else ""
        desc = " ".join(it.subs[1:]) if len(it.subs) > 1 else ""
        nb = kz.add_box(s, kz.MARGIN, y - Inches(0.12), Inches(2.6), Inches(0.9))
        p = nb.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = it.text
        kz.set_font(r, size=40, latin=kz.EN_SERIF, ea=kz.CN_SERIF,
                    bold=True, color=kz.ORANGE)
        lb = kz.add_box(s, kz.MARGIN + Inches(2.7), y, Inches(2.6), Inches(0.6))
        rich_para(lb.text_frame, label, size=16, first=True)
        if desc:
            db = kz.add_box(s, kz.MARGIN + Inches(5.45), y + Inches(0.02),
                            Inches(5.7), Inches(0.9))
            plain_para(db.text_frame, desc, size=13, first=True, line=1.4)
        y += Inches(1.35)
    return s


def L_timeline(prs, page: Page):
    s = kz.new_slide(prs)
    head(s, page.title, y=Inches(1.65))
    items = page.items[:4]
    n = max(len(items), 1)
    x0 = kz.MARGIN + Inches(0.85)
    span = Inches(10.0)
    step = span / max(n - 1, 1)
    yline = Inches(3.9)
    kz.line(s, x0, yline, x0 + span, yline, color=kz.GRAY, width=Pt(1.0))
    for i, it in enumerate(items):
        x = x0 + step * i
        dot = s.shapes.add_shape(1, x - Inches(0.07), yline - Inches(0.07),
                                 Inches(0.14), Inches(0.14))
        dot.fill.solid()
        dot.fill.fore_color.rgb = kz.ORANGE
        dot.line.fill.background()
        dot.shadow.inherit = False
        tb = kz.add_box(s, x - Inches(0.9), yline - Inches(0.95),
                        Inches(1.8), Inches(0.7))
        p = tb.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        rich_runs(p, it.text, size=17, bold=True)
        if it.subs:
            tb2 = kz.add_box(s, x - Inches(1.1), yline + Inches(0.35),
                             Inches(2.2), Inches(1.6))
            for j, sub in enumerate(it.subs[:2]):
                p2 = (tb2.text_frame.paragraphs[0] if j == 0
                      else tb2.text_frame.add_paragraph())
                p2.alignment = PP_ALIGN.CENTER
                p2.line_spacing = 1.4
                r2 = p2.add_run()
                r2.text = sub
                kz.set_font(r2, size=11, latin=kz.CN_SANS, ea=kz.CN_SANS,
                            color=kz.GRAY)
    return s


def L_compare(prs, page: Page):
    s = kz.new_slide(prs)
    head(s, page.title, y=Inches(1.65))
    kz.line(s, kz.MARGIN + Inches(5.35), Inches(2.75),
            kz.MARGIN + Inches(5.35), Inches(6.4), color=kz.GRAY, width=Pt(0.75))
    mid = max(len(page.items) // 2, 1)
    groups = [page.items[:mid], page.items[mid:]]
    tags = ["现在", "目标"]
    for gi, group in enumerate(groups):
        x = kz.MARGIN + Inches(5.75) * gi
        tb = kz.add_box(s, x, Inches(2.55), Inches(4.9), Inches(0.4))
        p = tb.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = tags[gi]
        kz.set_font(r, size=11, latin=kz.EN_SERIF, ea=kz.CN_SERIF, color=kz.GOLD)
        y = Inches(3.15)
        for it in group[:3]:
            tb2 = kz.add_box(s, x, y, Inches(4.9), Inches(1.2))
            tf = tb2.text_frame
            rich_para(tf, it.text, size=18, first=True, line=1.2)
            for sub in it.subs[:1]:
                plain_para(tf, sub, size=12, space_before=5)
            y += Inches(1.25)
    return s


def L_quote(prs, page: Page):
    s = kz.new_slide(prs)
    kz.rule(s, kz.MARGIN, Inches(1.25), Inches(1.1))
    tb = kz.add_box(s, kz.MARGIN, Inches(2.5), Inches(10.6), Inches(2.4))
    rich_para(tb.text_frame, page.title, size=34, first=True, line=1.4)
    if page.items:
        tb2 = kz.add_box(s, kz.MARGIN, Inches(5.15), Inches(10.6), Inches(1.2))
        for i, it in enumerate(page.items[:2]):
            plain_para(tb2.text_frame, it.text, size=13, first=(i == 0),
                       space_before=0 if i == 0 else 8)
    return s


def L_dense(prs, page: Page):
    s = kz.new_slide(prs)
    kz.rule(s, kz.MARGIN, Inches(1.25), Inches(0.7), color=kz.ORANGE)
    tb = kz.add_box(s, kz.MARGIN, Inches(1.75), Inches(6.4), Inches(0.9))
    rich_para(tb.text_frame, page.title, size=27, first=True, line=1.2)
    y = Inches(2.85)
    for it in page.items[:2]:
        tb2 = kz.add_box(s, kz.MARGIN, y, Inches(6.4), Inches(1.9))
        tf = tb2.text_frame
        rich_para(tf, it.text, size=15, first=True, line=1.55, bold=False,
                  color=kz.INK)
        y += Inches(2.0)
    nb = kz.add_box(s, kz.MARGIN + Inches(7.0), Inches(2.3), Inches(4.1), Inches(3.6))
    for j, it in enumerate(page.items[2:5]):
        tf = nb.text_frame
        rich_para(tf, it.text, size=13, first=(j == 0), line=1.5,
                  space_before=0 if j == 0 else 12)
        for sub in it.subs[:3]:
            plain_para(tf, sub, size=11, space_before=3, line=1.45)
    return s


def L_closing(prs, page: Page):
    s = kz.new_slide(prs)
    kz.rule(s, kz.MARGIN, Inches(1.25), Inches(1.1))
    t = page.title
    if is_framed(t):
        kz.emphasis_frame(s, unframe(t), x=kz.MARGIN, y=Inches(2.85),
                          size=34, color=kz.ORANGE, style="full", animate=True)
    else:
        tb = kz.add_box(s, kz.MARGIN, Inches(2.85), Inches(10.6), Inches(1.6))
        rich_para(tb.text_frame, t, size=38, first=True, line=1.2)
    rest = list(page.items)
    if rest and is_framed(rest[0].text):
        kz.emphasis_frame(s, unframe(rest[0].text), x=kz.MARGIN, y=Inches(4.55),
                          size=30, color=kz.ORANGE, style="full", animate=True)
        rest = rest[1:]
    if rest:
        tb2 = kz.add_box(s, kz.MARGIN, Inches(5.75), Inches(10.6), Inches(0.8))
        plain_para(tb2.text_frame, rest[0].text, size=16, first=True)
    return s


LAYOUTS = {
    "cover": L_cover, "statement": L_statement, "section": L_section,
    "list": L_list, "cards": L_cards, "twocol": L_twocol, "stats": L_stats,
    "timeline": L_timeline, "compare": L_compare, "quote": L_quote,
    "dense": L_dense, "closing": L_closing,
}


def pick_layout(page: Page) -> str:
    if page.layout in LAYOUTS:
        return page.layout
    n = len(page.items)
    if n == 0:
        return "statement"
    if n <= 4:
        return "list"
    return "dense"


# ---------------- 主流程 ----------------

def build(md: Path, out: Path) -> None:
    deck_title, pages = parse(md)
    if not pages:
        sys.exit(f"没解析出任何页：{md}")

    prs = Presentation()
    prs.slide_width, prs.slide_height = kz.W, kz.H

    total = len(pages)
    for i, page in enumerate(pages):
        layout = pick_layout(page)
        if layout == "cover":
            slide = L_cover(prs, page, deck_title)
        elif layout == "closing" or (i == total - 1 and layout == "list"):
            slide = L_closing(prs, page)
        else:
            slide = LAYOUTS[layout](prs, page)

        kz.page_no(slide, i + 1)
        kz.notes(slide, page.notes)
        kz.fade_transition(slide)

    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out)
    print(f"[OK] {out}  ({len(prs.slides)} pages)")


def main() -> None:
    ap = argparse.ArgumentParser(description="outline.md -> 开智风格 pptx")
    ap.add_argument("md", help="输入 outline.md")
    ap.add_argument("-o", "--out", help="输出 .pptx（默认同名）")
    ap.add_argument("--no-anim", action="store_true", help="关闭描边动画")
    args = ap.parse_args()

    if args.no_anim:
        kz.ANIMATE = False

    md = Path(args.md)
    if not md.exists():
        sys.exit(f"找不到 {md}")
    out = Path(args.out) if args.out else md.with_suffix(".pptx")
    build(md, out)


if __name__ == "__main__":
    main()
