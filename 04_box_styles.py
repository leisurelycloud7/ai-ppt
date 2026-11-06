#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
强调框 · 三种做法对比样张
输出：demo/box_styles.pptx  （2 页：对比页 + 应用页）
"""
from pathlib import Path

import kz
from pptx import Presentation
from pptx.util import Inches, Pt

SAMPLE = "能自己检查结果"


def slide_compare(prs) -> None:
    s = kz.new_slide(prs)
    kz.page_label(s, "STYLE TEST  ·  EMPHASIS")
    kz.rule(s, kz.MARGIN, Inches(1.25), Inches(0.7), color=kz.ORANGE)

    tb = kz.add_box(s, kz.MARGIN, Inches(1.9), Inches(11), Inches(0.8))
    kz.para(tb.text_frame, "强调的三种框法", size=32, latin=kz.CN_SANS,
            ea=kz.CN_SANS, bold=True, first=True)

    rows = [
        ("A · 四边闭合细框", "full", kz.ORANGE),
        ("B · 四角标记（日式引号感）", "corner", kz.GOLD),
        ("C · 「」日式引号（不加框）", "quote", kz.ORANGE),
    ]
    y = Inches(2.95)
    for tag, style, color in rows:
        lb = kz.add_box(s, kz.MARGIN, y + Inches(0.16), Inches(2.35), Inches(0.6))
        kz.para(lb.text_frame, tag, size=12, latin=kz.CN_SERIF, ea=kz.CN_SERIF,
                color=kz.GOLD, first=True)
        kz.emphasis_frame(s, SAMPLE, x=kz.MARGIN + Inches(2.55), y=y,
                          size=kz.SIZE_MAIN, color=color, style=style)
        y += Inches(1.15)

    nb = kz.add_box(s, kz.MARGIN, Inches(6.5), Inches(11), Inches(0.4))
    kz.para(nb.text_frame, "A 最稳、最像「被圈出来的重点」；B 轻，适合一页只圈一处；C 用于句子中间的关键词。",
            size=13, latin=kz.CN_SANS, ea=kz.CN_SANS, color=kz.GRAY, first=True)

    kz.page_no(s, 1)
    kz.notes(s, "三种框法对比，选一个作为全片主用。")
    kz.fade_transition(s)


def slide_applied(prs) -> None:
    s = kz.new_slide(prs)
    kz.page_label(s, "APPLIED  ·  HOW IT READS")
    kz.rule(s, kz.MARGIN, Inches(1.25), Inches(0.7), color=kz.ORANGE)

    lb = kz.add_box(s, kz.MARGIN, Inches(1.75), Inches(6), Inches(0.3))
    kz.para(lb.text_frame, "整行强调 → 加框", size=12, latin=kz.CN_SERIF,
            ea=kz.CN_SERIF, color=kz.GRAY, first=True)

    tb = kz.add_box(s, kz.MARGIN, Inches(2.25), Inches(10), Inches(0.8))
    kz.para(tb.text_frame, "它不会喊累，", size=32, latin=kz.CN_SANS,
            ea=kz.CN_SANS, bold=True, first=True)
    kz.emphasis_frame(s, "只会一直错下去。", x=kz.MARGIN, y=Inches(3.05),
                      size=kz.SIZE_MAIN, color=kz.ORANGE, style="full")

    kz.line(s, kz.MARGIN, Inches(4.35), kz.MARGIN + Inches(3.0),
            Inches(4.35), color=kz.GRAY, width=Pt(0.75))

    lb2 = kz.add_box(s, kz.MARGIN, Inches(4.6), Inches(6), Inches(0.3))
    kz.para(lb2.text_frame, "句中强调 → 「」", size=12, latin=kz.CN_SERIF,
            ea=kz.CN_SERIF, color=kz.GRAY, first=True)

    tb2 = kz.add_box(s, kz.MARGIN, Inches(5.05), Inches(11), Inches(1.0))
    p = tb2.text_frame.paragraphs[0]
    for txt, col in (("能自己", kz.INK), ("「检查结果」", kz.ORANGE),
                     ("，才叫 Agent", kz.INK)):
        r = p.add_run()
        r.text = txt
        kz.set_font(r, size=32, latin=kz.CN_SANS, ea=kz.CN_SANS,
                    bold=True, color=col)

    kz.page_no(s, 2)
    kz.notes(s, "同一页里两种强调的读感差别。")
    kz.fade_transition(s)


def main() -> None:
    prs = Presentation()
    prs.slide_width, prs.slide_height = kz.W, kz.H
    slide_compare(prs)
    slide_applied(prs)
    out = Path(__file__).parent / "demo" / "box_styles.pptx"
    out.parent.mkdir(exist_ok=True)
    prs.save(out)
    print(f"[OK] {out}")


if __name__ == "__main__":
    main()
