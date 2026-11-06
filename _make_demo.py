# -*- coding: utf-8 -*-
"""生成一份测试用 PPTX（模拟 WorkBuddy 导出的成品），仅用于验证提取脚本。"""
from pathlib import Path
from pptx import Presentation
from pptx.util import Pt

DEMO = [
    ("AI Agent 是什么", ["从「聊天」到「干活」", "模型 + 工具 + 记忆 + 循环"], "AI Agent 不是一个更聪明的聊天框，而是一个能自己调用工具、自己检查结果的系统。"),
    ("为什么现在能做", ["模型能力跨过临界点", "工具调用标准化（MCP）", "部署成本下降 90%"], "过去三年，这三件事同时发生，才让 Agent 从 demo 变成生产可用。"),
    ("三条落地路径", ["SaaS 直接用：Dify / Coze", "开源自部署：OpenHands", "自己写：LangGraph"], "给普通人的建议是先走第一条，把流程跑通再谈自建。"),
    ("常见坑", ["权限给太大", "没有人工兜底", "成本失控"], "Agent 出错的方式和人不一样，它不会喊累，只会一直错下去。"),
    ("本周要做的事", ["装 Git 和 Docker", "跑通一个开源 Agent", "写一份部署笔记"], "目标不是学会，是有一个能演示的东西。"),
]


def main() -> None:
    prs = Presentation()
    for i, (title, bullets, notes) in enumerate(DEMO):
        if i == 0:
            layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(layout)
            slide.shapes.title.text = title
            slide.placeholders[1].text = bullets[0]
        else:
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = title
            tf = slide.placeholders[1].text_frame
            tf.text = bullets[0]
            for b in bullets[1:]:
                p = tf.add_paragraph()
                p.text = b
        slide.notes_slide.notes_text_frame.text = notes
    out = Path(__file__).parent / "demo" / "demo_deck.pptx"
    out.parent.mkdir(exist_ok=True)
    prs.save(out)
    print(f"saved {out}")


if __name__ == "__main__":
    main()
