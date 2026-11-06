#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
outline.md -> WPS AI 可直接粘贴的纯文本大纲

用法：
    python 07_wps_outline.py outline.md            # 生成 .txt（层级式）
    python 07_wps_outline.py outline.md --flat     # 每页压成一行

为什么要转：WPS AI 的「粘贴大纲生成 PPT」吃的是纯文本层级，
markdown 的 ## / - / > 它不一定认，转一遍命中率更高。
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

PAGE_RE = re.compile(r"^##\s*P?(\d+)?\s*[·.]?\s*(.*)$")
NOTE_RE = re.compile(r"^>\s*讲稿[:：]\s*(.*)$")
BULLET_RE = re.compile(r"^[-*]\s+(.*)$")
META_RE = re.compile(r"^[-*]\s*(总页数|来源)[:：]")

CN_NUM = "零一二三四五六七八九十"


def cn_index(n: int) -> str:
    if n <= 10:
        return CN_NUM[n]
    if n < 20:
        return "十" + CN_NUM[n - 10]
    return str(n)


def strip_md(t: str) -> str:
    t = re.sub(r"\*\*(.+?)\*\*", r"\1", t)
    t = re.sub(r"\[\[(.+?)\]\]", r"\1", t)
    t = re.sub(r"「([^」]*)」", r"\1", t)
    t = re.sub(r"`(.+?)`", r"\1", t)
    t = re.sub(r"^\s*[-*]\s*", "", t)
    return t.strip()


def parse(path: Path):
    title = path.stem
    pages = []
    cur = None
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.rstrip()
        if not line or line.strip() == "---":
            continue
        if line.startswith("# ") and not line.startswith("## "):
            t = strip_md(line[2:].strip())
            for pref in ("幻灯片大纲：", "幻灯片大纲:", "大纲：", "大纲:"):
                if t.startswith(pref):
                    t = t[len(pref):]
            title = t.strip()
            continue
        m = PAGE_RE.match(line)
        if m:
            cur = {"title": strip_md(m.group(2)), "bullets": [], "notes": ""}
            pages.append(cur)
            continue
        if META_RE.match(line) or cur is None:
            continue
        m = NOTE_RE.match(line)
        if m:
            cur["notes"] = (cur["notes"] + " " + m.group(1)).strip()
            continue
        m = BULLET_RE.match(line)
        if m:
            b = strip_md(m.group(1))
            if b:
                cur["bullets"].append(b)
            continue
        b = strip_md(line)
        if b:
            cur["bullets"].append(b)
    return title, pages


def build(title: str, pages, flat: bool = False) -> str:
    out = [f"主题：{title}", ""]
    if flat:
        for i, p in enumerate(pages, 1):
            bits = [p["title"]] + p["bullets"][:3]
            out.append(f"第{i}页：" + "｜".join(bits))
        return "\n".join(out)

    for i, p in enumerate(pages, 1):
        out.append(f"{cn_index(i)}、{p['title']}")
        for j, b in enumerate(p["bullets"][:3], 1):
            out.append(f"    {j}. {b}")
        out.append("")
    return "\n".join(out)


def main() -> None:
    ap = argparse.ArgumentParser(description="outline.md -> WPS AI 可粘贴大纲")
    ap.add_argument("md", help="输入 outline.md")
    ap.add_argument("-o", "--out", help="输出 .txt")
    ap.add_argument("--flat", action="store_true", help="每页压成一行")
    args = ap.parse_args()

    md = Path(args.md)
    if not md.exists():
        sys.exit(f"找不到 {md}")

    title, pages = parse(md)
    text = build(title, pages, flat=args.flat)
    out = Path(args.out) if args.out else md.with_name(md.stem + ("_wps_flat.txt" if args.flat else "_wps.txt"))
    out.write_text(text, encoding="utf-8")
    print(f"[OK] {out}  ({len(pages)} 页)")
    print("-" * 40)
    print(text[:600])


if __name__ == "__main__":
    main()
