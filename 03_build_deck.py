#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
开智风格 PPTX 样张 v1（5 页）

样式全部来自 kz.py，改颜色/字号只改那边。
用法：python 03_build_deck.py
输出：demo/kaizhi_style_v1.pptx
"""
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

import kz


def build() -> Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = kz.W, kz.H

    # ---------- P1 封面 ----------
    s = kz.new_slide(prs)
    kz.page_label(s, "AI AGENT  ·  2026")
    kz.rule(s, kz.MARGIN, Inches(1.25), Inches(1.1))
    tb = kz.add_box(s, kz.MARGIN, Inches(2.75), Inches(10.5), Inches(2.4))
    kz.para(tb.text_frame, "把 AI 变成同事", size=kz.SIZE_COVER, latin=kz.CN_SANS,
            ea=kz.CN_SANS, bold=True, line=1.1, first=True)
    tb2 = kz.add_box(s, kz.MARGIN, Inches(4.35), Inches(10.5), Inches(0.6))
    kz.para(tb2.text_frame, "Turning AI into a coworker, not a chat box.",
            size=26, latin=kz.EN_SERIF, ea=kz.CN_SERIF, italic=True,
            color=kz.GRAY, first=True)
    kz.page_no(s, 1)
    kz.notes(s, "开场一句话：我们想做的不是又一个聊天框。")
    kz.fade_transition(s)

    # ---------- P2 观点页 ----------
    s = kz.new_slide(prs)
    kz.page_label(s, "01  ·  WHAT CHANGED")
    kz.rule(s, kz.MARGIN, Inches(1.25), Inches(0.7), color=kz.ORANGE)
    y = Inches(2.55)
    x, _, w, h = kz.emphasis_frame(s, "能自己检查结果", x=kz.MARGIN, y=y,
                                   size=kz.SIZE_MAIN, color=kz.ORANGE, style="full")
    tb = kz.add_box(s, kz.MARGIN + w + Inches(0.02), y + Inches(0.15),
                    Inches(4.2), h)
    kz.para(tb.text_frame, "，才叫 Agent", size=kz.SIZE_MAIN, latin=kz.CN_SANS,
            ea=kz.CN_SANS, bold=True, first=True, line=1.0)
    tb2 = kz.add_box(s, kz.MARGIN, Inches(4.15), Inches(11.1), Inches(0.9))
    kz.para(tb2.text_frame, "An agent is not a smarter chat box — it checks its own work.",
            size=kz.SIZE_EN, latin=kz.EN_SERIF, ea=kz.CN_SERIF, italic=True,
            color=kz.GRAY, first=True)
    kz.page_no(s, 2)
    kz.notes(s, "判断标准只有一条：它会不会回头检查自己做的事。")
    kz.fade_transition(s)

    # ---------- P3 三条目 ----------
    s = kz.new_slide(prs)
    kz.page_label(s, "02  ·  THREE LINES")
    kz.rule(s, kz.MARGIN, Inches(1.25), Inches(0.7), color=kz.ORANGE)
    tb = kz.add_box(s, kz.MARGIN, Inches(2.05), Inches(11.1), Inches(0.8))
    kz.para(tb.text_frame, "三条落地路径", size=kz.SIZE_MAIN, latin=kz.CN_SANS,
            ea=kz.CN_SANS, bold=True, first=True)
    items = [
        ("01", "直接用现成产品", "Dify / Coze，先跑通流程"),
        ("02", "开源自部署", "OpenHands，本地可控"),
        ("03", "自己写", "LangGraph，最后再考虑"),
    ]
    y = Inches(3.45)
    for num, cn, note in items:
        nb = kz.add_box(s, kz.MARGIN, y, Inches(1.2), Inches(0.5))
        kz.para(nb.text_frame, num, size=14, latin=kz.EN_SERIF, ea=kz.CN_SERIF,
                color=kz.ORANGE, first=True)
        tb = kz.add_box(s, kz.MARGIN + Inches(1.2), y - Inches(0.06),
                        Inches(9.6), Inches(0.95))
        kz.para(tb.text_frame, cn, size=22, latin=kz.CN_SANS, ea=kz.CN_SANS,
                bold=True, first=True)
        kz.para(tb.text_frame, note, size=15, latin=kz.CN_SANS, ea=kz.CN_SANS,
                color=kz.GRAY, space_before=4, line=1.1)
        y += Inches(1.25)
    kz.page_no(s, 3)
    kz.notes(s, "路径从轻到重，别一上来就自建。")
    kz.fade_transition(s)

    # ---------- P4 引用页 ----------
    s = kz.new_slide(prs)
    kz.page_label(s, "03  ·  NOTE")
    kz.rule(s, kz.MARGIN, Inches(1.25), Inches(0.7), color=kz.ORANGE)
    tb = kz.add_box(s, kz.MARGIN, Inches(2.5), Inches(10.8), Inches(1.0))
    kz.para(tb.text_frame, "它不会喊累，", size=36, latin=kz.CN_SANS,
            ea=kz.CN_SANS, bold=True, first=True)
    kz.emphasis_frame(s, "只会一直错下去。", x=kz.MARGIN, y=Inches(3.25),
                      size=36, color=kz.ORANGE, style="full")
    tb2 = kz.add_box(s, kz.MARGIN, Inches(5.05), Inches(10.8), Inches(0.8))
    kz.para(tb2.text_frame, "It never gets tired — it just keeps being wrong.",
            size=kz.SIZE_EN, latin=kz.EN_SERIF, ea=kz.CN_SERIF, italic=True,
            color=kz.GRAY, first=True)
    kz.page_no(s, 4)
    kz.notes(s, "所以人工兜底不是保险，是设计的一部分。")
    kz.fade_transition(s)

    # ---------- P5 结尾 ----------
    s = kz.new_slide(prs)
    kz.page_label(s, "04  ·  NEXT")
    kz.rule(s, kz.MARGIN, Inches(1.25), Inches(1.1))
    tb = kz.add_box(s, kz.MARGIN, Inches(2.9), Inches(10.5), Inches(1.8))
    kz.para(tb.text_frame, "这周，先跑通一个", size=40, latin=kz.CN_SANS,
            ea=kz.CN_SANS, bold=True, first=True)
    tb2 = kz.add_box(s, kz.MARGIN, Inches(4.5), Inches(10.5), Inches(0.6))
    kz.para(tb2.text_frame, "Ship one, then scale.", size=24, latin=kz.EN_SERIF,
            ea=kz.CN_SERIF, italic=True, color=kz.GRAY, first=True)
    kz.page_no(s, 5)
    kz.notes(s, "收尾落到行动：一个能演示的东西。")
    kz.fade_transition(s)

    out = Path(__file__).parent / "demo" / "kaizhi_style_v1.pptx"
    out.parent.mkdir(exist_ok=True)
    prs.save(out)
    print(f"[OK] {out}  ({len(prs.slides)} pages)")
    return out


if __name__ == "__main__":
    build()
