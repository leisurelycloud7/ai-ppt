# -*- coding: utf-8 -*-
"""
开智风格 · 设计层（被 03/04 等脚本复用）

改样式只改这个文件，不要改各个出片脚本。
"""
from lxml import etree
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# ================= 设计令牌 =================
BG = RGBColor(0x0B, 0x12, 0x20)      # 深蓝黑
INK = RGBColor(0xF2, 0xF4, 0xF8)     # 近白
ORANGE = RGBColor(0xE8, 0x76, 0x3A)  # 暖橙（强调）
GOLD = RGBColor(0xC9, 0xA4, 0x5C)    # 金色衬线（点缀）
GRAY = RGBColor(0x7C, 0x86, 0x98)    # 辅助灰

CN_SANS = "Noto Sans SC"    # 思源黑体
CN_SERIF = "Noto Serif SC"  # 思源宋体
EN_SERIF = "Georgia"        # 英文衬线斜体

SIZE_MAIN = 32      # 中文主字幕
SIZE_EN = 19        # 英文副字幕（≈60%）
SIZE_COVER = 44
SIZE_LABEL = 11

W, H = Inches(13.333), Inches(7.5)
MARGIN = Inches(1.1)

FRAME_PAD = Inches(0.20)   # 框内边距
FRAME_W = Pt(1.5)          # 框线粗细
CORNER_LEN = Inches(0.30)  # 四角标记长度

ANIMATE = True             # 逐页描边动画总开关（--no-anim 会关掉）


# ================= 基础工具 =================

def set_font(run, *, size, latin, ea=None, bold=False, italic=False, color=INK):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.name = latin
    f.color.rgb = color
    if ea:
        rPr = run._r.get_or_add_rPr()
        el = rPr.find(qn("a:ea"))
        if el is None:
            el = etree.SubElement(rPr, qn("a:ea"))
        el.set("typeface", ea)


def add_box(slide, x, y, w, h, *, anchor=MSO_ANCHOR.TOP, wrap=True):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    return box


def para(tf, text, *, size, latin, ea=None, bold=False, italic=False,
         color=INK, space_before=0, space_after=0, line=1.15, first=False,
         align=PP_ALIGN.LEFT):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    p.line_spacing = line
    r = p.add_run()
    r.text = text
    set_font(r, size=size, latin=latin, ea=ea, bold=bold,
             italic=italic, color=color)
    return p


def new_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = BG
    return s


def line(slide, x1, y1, x2, y2, *, color=GOLD, width=Pt(1.2)):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    ln.line.color.rgb = color
    ln.line.width = width
    return ln


def rule(slide, x, y, w, *, color=GOLD, width=Pt(1.2)):
    return line(slide, x, y, x + w, y, color=color, width=width)


def page_label(slide, text):
    box = add_box(slide, MARGIN, Inches(0.75), Inches(7), Inches(0.35))
    para(box.text_frame, text, size=SIZE_LABEL, latin=EN_SERIF, ea=CN_SERIF,
         color=GOLD, first=True)


def page_no(slide, n):
    box = add_box(slide, W - MARGIN - Inches(1.0), H - Inches(0.9),
                  Inches(1.0), Inches(0.3))
    p = box.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    r.text = f"{n:02d}"
    set_font(r, size=SIZE_LABEL, latin=EN_SERIF, ea=CN_SERIF, color=GRAY)


def fade_transition(slide):
    tr = etree.SubElement(slide.element, qn("p:transition"))
    tr.set("spd", "med")
    etree.SubElement(tr, qn("p:fade"))


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# ================= 描边动画 =================

def _p(parent, tag, **attrs):
    e = etree.SubElement(parent, qn("p:" + tag))
    for k, v in attrs.items():
        e.set(k, str(v))
    return e


def add_draw_animation(slide, segments, *, dur_ms=350, stagger_ms=280):
    """
    让若干条线依次"画"出来（擦除进入动画，与上一动画同时 + 延迟）。

    segments: [(shape, filter), ...]，按绘制顺序；filter 取 wipe(right/down/left/up)。
    """
    timing = _p(slide.element, "timing")
    tnLst = _p(timing, "tnLst")
    par1 = _p(tnLst, "par")
    c1 = _p(par1, "cTn", id=1, dur="indefinite", restart="never", nodeType="tmRoot")
    ch1 = _p(c1, "childTnLst")
    seq = _p(ch1, "seq", concurrent="1", nextAc="seek")
    c2 = _p(seq, "cTn", id=2, dur="indefinite", nodeType="mainSeq")
    ch2 = _p(c2, "childTnLst")
    par3 = _p(ch2, "par")
    c3 = _p(par3, "cTn", id=3, fill="hold")
    st3 = _p(c3, "stCondLst")
    _p(st3, "cond", delay="indefinite")
    ch3 = _p(c3, "childTnLst")
    par4 = _p(ch3, "par")
    c4 = _p(par4, "cTn", id=4, fill="hold")
    st4 = _p(c4, "stCondLst")
    _p(st4, "cond", delay="0")
    ch4 = _p(c4, "childTnLst")

    nid = 4
    for i, (shape, filt) in enumerate(segments):
        nid += 1
        parE = _p(ch4, "par")
        node = "clickEffect" if i == 0 else "withEffect"
        cE = _p(parE, "cTn", id=nid, presetID=22, presetClass="entr",
                presetSubtype=0, fill="hold", grpId=0, nodeType=node)
        stE = _p(cE, "stCondLst")
        _p(stE, "cond", delay=str(i * stagger_ms))
        chE = _p(cE, "childTnLst")

        setel = _p(chE, "set")
        cbh = _p(setel, "cBhvr")
        nid += 1
        cset = _p(cbh, "cTn", id=nid, dur="1", fill="hold")
        st = _p(cset, "stCondLst")
        _p(st, "cond", delay="0")
        tgt = _p(cbh, "tgtEl")
        _p(tgt, "spTgt", spid=shape.shape_id)
        anl = _p(cbh, "attrNameLst")
        _p(anl, "attrName").text = "style.visibility"
        to = _p(setel, "to")
        _p(to, "strVal", val="visible")

        ae = _p(chE, "animEffect", transition="in", filter=filt)
        cbh2 = _p(ae, "cBhvr")
        nid += 1
        _p(cbh2, "cTn", id=nid, dur=str(dur_ms))
        tgt2 = _p(cbh2, "tgtEl")
        _p(tgt2, "spTgt", spid=shape.shape_id)


# ================= 强调框 =================

def est_width(text: str, size_pt: float) -> Emu:
    """粗估文本宽度（CJK=1em，西文≈0.5em）。只为排版，不追求像素级。"""
    em = size_pt / 72.0
    w = 0.0
    for ch in text:
        if ord(ch) > 0x2E80:
            w += 1.0
        elif ch == " ":
            w += 0.28
        elif ch in "mwMW@%":
            w += 0.75
        else:
            w += 0.52
    return Inches(w * em)


def emphasis_frame(slide, text, *, x, y, size=SIZE_MAIN, color=ORANGE,
                   text_color=INK, style="full", animate=False):
    """
    画一个强调单元，返回外框矩形 (left, top, width, height)。

    style:
      "full"   四边闭合细框
      "corner" 只留四个角的 L 形标记（日式引号意象）
      "quote"  不加框，文字两侧加 「 」
    """
    tw = est_width(text, size)
    if style == "quote":
        tw = est_width(f"「{text}」", size)
        w = tw * 1.05
        h = Inches(size / 72.0 * 1.5)
        box = add_box(slide, x, y, w + Inches(0.2), h, wrap=False)
        tf = box.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        # 「」用强调色，正文用白色
        r1 = p.add_run(); r1.text = "「"
        set_font(r1, size=size, latin=CN_SANS, ea=CN_SANS, bold=True, color=color)
        r2 = p.add_run(); r2.text = text
        set_font(r2, size=size, latin=CN_SANS, ea=CN_SANS, bold=True, color=text_color)
        r3 = p.add_run(); r3.text = "」"
        set_font(r3, size=size, latin=CN_SANS, ea=CN_SANS, bold=True, color=color)
        return (x, y, w, h)

    w = tw * 1.06 + FRAME_PAD * 2
    h = Inches(size / 72.0) + FRAME_PAD * 2

    if style == "full":
        if animate and ANIMATE:
            # 拆成四条线，顺时针依次"画"出来
            segs = [
                (line(slide, x, y, x + w, y, color=color, width=FRAME_W), "wipe(right)"),
                (line(slide, x + w, y, x + w, y + h, color=color, width=FRAME_W), "wipe(down)"),
                (line(slide, x + w, y + h, x, y + h, color=color, width=FRAME_W), "wipe(left)"),
                (line(slide, x, y + h, x, y, color=color, width=FRAME_W), "wipe(up)"),
            ]
            add_draw_animation(slide, segs)
        else:
            shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
            shp.fill.background()
            shp.line.color.rgb = color
            shp.line.width = FRAME_W
            shp.shadow.inherit = False
    elif style == "corner":
        L = CORNER_LEN
        c = color
        wd = Pt(2.0)
        # 左上
        line(slide, x, y, x + L, y, color=c, width=wd)
        line(slide, x, y, x, y + L, color=c, width=wd)
        # 右上
        line(slide, x + w - L, y, x + w, y, color=c, width=wd)
        line(slide, x + w, y, x + w, y + L, color=c, width=wd)
        # 左下
        line(slide, x, y + h - L, x, y + h, color=c, width=wd)
        line(slide, x, y + h, x + L, y + h, color=c, width=wd)
        # 右下
        line(slide, x + w - L, y + h, x + w, y + h, color=c, width=wd)
        line(slide, x + w, y + h - L, x + w, y + h, color=c, width=wd)
    else:
        raise ValueError(style)

    tb = add_box(slide, x + FRAME_PAD, y + FRAME_PAD - Inches(0.05),
                 w - FRAME_PAD * 2, h, wrap=False)
    para(tb.text_frame, text, size=size, latin=CN_SANS, ea=CN_SANS,
         bold=True, color=text_color, first=True, line=1.0)
    return (x, y, w, h)
