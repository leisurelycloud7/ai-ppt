#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PPTX -> Markdown 提取器（标题 / 正文 / 演讲者备注）

用法：
  python 02_extract_pptx.py deck.pptx
  python 02_extract_pptx.py deck.pptx -o outline.md
  python 02_extract_pptx.py deck.pptx --outline     # 只输出"页号+标题"清单
  python 02_extract_pptx.py deck.pptx --sheet       # 输出给 open-slide 用的紧凑版

依赖：python-pptx
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
except ImportError:
    sys.exit("缺少依赖：先运行  pip install python-pptx")


# ---------- 工具函数 ----------

def iter_shapes(shapes):
    """递归展开组合形状。"""
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(sh.shapes)
        else:
            yield sh


def clean(text: str) -> str:
    """去掉多余空白、空段落、零宽字符。"""
    text = text.replace("\u200b", "").replace("\xa0", " ")
    lines = [re.sub(r"[ \t]+", " ", ln).strip() for ln in text.splitlines()]
    return "\n".join(ln for ln in lines if ln)


def slide_title(slide) -> str:
    """优先取占位符标题，取不到就用第一段非空文本。"""
    try:
        if slide.shapes.title is not None:
            t = clean(slide.shapes.title.text_frame.text)
            if t:
                return t.splitlines()[0]
    except Exception:
        pass
    for sh in iter_shapes(slide.shapes):
        if getattr(sh, "has_text_frame", False):
            t = clean(sh.text_frame.text)
            if t:
                return t.splitlines()[0]
    return ""


def slide_body(slide, title: str) -> list[str]:
    """正文：除标题外的所有文本框，按出现顺序。"""
    out: list[str] = []
    for sh in iter_shapes(slide.shapes):
        if not getattr(sh, "has_text_frame", False):
            continue
        try:
            if sh.is_placeholder and sh.placeholder_format.idx == 0:
                continue  # 跳过标题占位符
        except Exception:
            pass
        t = clean(sh.text_frame.text)
        if not t:
            continue
        out.extend(t.splitlines())
    # 第一行常常就是标题的重复，去掉
    out = [ln for ln in out if ln != title]
    # 去重但保序
    seen, uniq = set(), []
    for ln in out:
        if ln not in seen:
            seen.add(ln)
            uniq.append(ln)
    return uniq


def slide_notes(slide) -> str:
    """演讲者备注（你的"字幕/讲稿句"通常在这里）。"""
    try:
        if not slide.has_notes_slide:
            return ""
        return clean(slide.notes_slide.notes_text_frame.text)
    except Exception:
        return ""


# ---------- 输出 ----------

def to_markdown(prs, src: Path, clean: bool = False) -> str:
    total = len(prs.slides)
    buf = [f"# 幻灯片大纲：{src.stem}", "", f"- 总页数：{total}", f"- 来源：{src.name}", "", "---", ""]

    # 先收集每页数据，方便做跨页去重
    raw = []
    for slide in prs.slides:
        title = slide_title(slide) or ""
        body = slide_body(slide, title)
        raw.append({"title": title, "body": body, "notes": slide_notes(slide)})

    if clean:
        from collections import Counter

        freq = Counter()
        for page in raw:
            freq.update(set(page["body"]))
        threshold = max(2, int(total * 0.6))

        def keep(s: str) -> bool:
            s = s.strip()
            if not s:
                return False
            if re.fullmatch(r"\d{1,2}\s*[/／]\s*\d{1,2}", s):    # 02 / 09
                return False
            if re.fullmatch(r"\d{1,2}", s):                       # 01 02
                return False
            if re.fullmatch(r"[「」『』\"'“”‘’\-—·。，、]+", s):   # 纯符号
                return False
            if freq[s] >= threshold:                              # 页脚重复文案
                return False
            return True

        for page in raw:
            page["body"] = [ln for ln in page["body"] if keep(ln)]

    for i, page in enumerate(raw, 1):
        title = page["title"] or f"（第 {i} 页无标题）"
        buf.append(f"## P{i} · {title}")
        buf.append("")
        for ln in page["body"]:
            buf.append(f"- {ln}")
        if page["body"]:
            buf.append("")
        if page["notes"]:
            for ln in page["notes"].splitlines():
                buf.append(f"> 讲稿：{ln}")
            buf.append("")
        buf.append("---")
        buf.append("")

    return "\n".join(buf)


def to_outline(prs) -> str:
    lines = []
    for i, slide in enumerate(prs.slides, 1):
        lines.append(f"- P{i}  {slide_title(slide) or '（无标题）'}")
    return "\n".join(lines)


def to_sheet(prs) -> str:
    """紧凑版：一页一段，方便整段贴进 AI 对话框。"""
    chunks = []
    for i, slide in enumerate(prs.slides, 1):
        title = slide_title(slide) or f"第{i}页"
        body = slide_body(slide, title)
        notes = slide_notes(slide)
        part = [f"### {i}. {title}"]
        part += [f"- {ln}" for ln in body]
        if notes:
            part.append(f"- 讲稿: {notes}")
        chunks.append("\n".join(part))
    return "\n\n".join(chunks)


# ---------- 入口 ----------

def main() -> None:
    ap = argparse.ArgumentParser(description="PPTX -> Markdown 提取器")
    ap.add_argument("pptx", help="输入 .pptx 文件")
    ap.add_argument("-o", "--out", help="输出 .md 路径（默认同名）")
    ap.add_argument("--outline", action="store_true", help="只输出标题清单")
    ap.add_argument("--sheet", action="store_true", help="输出紧凑版（贴给 AI 用）")
    ap.add_argument("--clean", action="store_true",
                    help="清洗模式：去掉页脚重复文案、页码、栏目编号")
    args = ap.parse_args()

    src = Path(args.pptx)
    if not src.exists():
        sys.exit(f"找不到文件：{src}")

    prs = Presentation(str(src))

    if args.outline:
        text = to_outline(prs)
        suffix = "_outline.md"
    elif args.sheet:
        text = to_sheet(prs)
        suffix = "_sheet.md"
    else:
        text = to_markdown(prs, src, clean=args.clean)
        suffix = ".md"

    out = Path(args.out) if args.out else src.with_name(src.stem + suffix)
    out.write_text(text, encoding="utf-8")

    print(f"[OK] 已提取 {len(prs.slides)} 页 -> {out}")
    print(f"     字数约 {len(text)} 字")


if __name__ == "__main__":
    main()
