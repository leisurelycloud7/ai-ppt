#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
生成《生涯发展报告》（成长赛道）融合版 docx

内容来源：职业杠杆 50 页大纲 + 大赛通用框架（自我认知 / 职业认知 / 职业决策 / 计划实施 / 评估调整）
隐私项一律留【请填写：…】，由本人补。
"""
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

CN = "宋体"
CN_BOLD = "黑体"


def set_run_font(run, name=CN, size=11, bold=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        from docx.oxml import OxmlElement
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), name)


def para(doc, text, *, size=11, bold=False, font=CN, align=None,
         space_before=0, space_after=6, indent=True, color=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = 1.5
    if indent and align is None:
        pf.first_line_indent = Pt(22)
    r = p.add_run(text)
    set_run_font(r, font, size, bold, color)
    return p


def h1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run(text)
    set_run_font(r, CN_BOLD, 15, True, RGBColor(0x0B, 0x12, 0x20))
    return p


def h2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    set_run_font(r, CN_BOLD, 12, True, RGBColor(0xE8, 0x76, 0x3A))
    return p


def table(doc, header, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(header):
        cell = t.rows[0].cells[i]
        cell.text = ""
        r = cell.paragraphs[0].add_run(h)
        set_run_font(r, CN_BOLD, 10, True)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            r = cells[i].paragraphs[0].add_run(v)
            set_run_font(r, CN, 10)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Pt(w)
    doc.add_paragraph()
    return t


def build(out: Path) -> None:
    doc = Document()

    # ---------- 封面信息 ----------
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("生涯发展报告")
    set_run_font(r, CN_BOLD, 26, True)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("成长赛道 · 把 AI 变成我的职业杠杆")
    set_run_font(r, CN, 13, False, RGBColor(0x50, 0x55, 0x60))

    doc.add_paragraph()
    table(doc, ["项目", "内容"], [
        ["姓名", "【请填写：姓名】"],
        ["学校 / 学院", "【请填写：学校、学院】"],
        ["专业 / 年级", "【请填写：专业】，大一"],
        ["联系方式", "【请填写：手机号 / 邮箱】"],
        ["目标职业", "AI 应用产品经理"],
        ["报告日期", "2026 年 11 月"],
    ], widths=[110, 330])

    # ---------- 一句话主张 ----------
    h1(doc, "写在前面")
    para(doc, "我不是来交一份计划的，我是来交一套可以被验证的成长路径。这份报告回答三个问题："
              "我现在是谁、我要去哪里、我打算怎么去。凡是无法被检验的目标，我都不会写进来。")

    # ---------- 一、自我认知 ----------
    h1(doc, "一、自我认知")

    h2(doc, "1.1 我是什么样的人")
    para(doc, "计算机专业大一学生，入学第一年就确定方向：不做「什么都学一点」的人，先找到一个能落地的切口。"
              "性格偏内向但不怕沟通，喜欢把复杂问题拆成步骤；面对陌生工具时，第一反应是动手跑通，而不是先看完整教程。"
              "价值观上更看重「做出来的东西有人用」，而不是「学过了多少门课」。")
    para(doc, "【请填写：补充你的兴趣、性格测评结果（如 MBTI / 霍兰德）、以及 1—2 个能体现性格的具体例子】")

    h2(doc, "1.2 能力盘点")
    table(doc, ["能力项", "具体证据", "自评水平"], [
        ["技术落地", "独立部署过开源 AI Agent：环境搭建、模型接入、流程跑通", "入门偏上"],
        ["工具学习", "一周内能完成一个新工具从了解到可用", "较强"],
        ["工程判断", "能判断一个技术方案在工程上是否可行、成本大致多少", "入门"],
        ["需求洞察", "缺少真实用户调研经验", "待补"],
        ["效果量化", "未系统做过指标定义与前后测", "待补"],
        ["项目推动", "【请填写：是否带过小组作业 / 社团项目】", "待补"],
    ], widths=[80, 260, 90])

    h2(doc, "1.3 我的短板")
    para(doc, "① 没进入过真实业务场景：项目停在「技术验证」，不知道谁在为什么发愁。"
              "② 没做过用户调研：习惯从技术出发，而不是从问题出发。"
              "③ 没量化过效果：做完说不清带来了什么改变。")
    para(doc, "这三条我不打算回避。它们不是能力上限问题，是经验结构问题——可以被项目覆盖，这也正是成长赛道的意义。")

    # ---------- 二、职业认知 ----------
    h1(doc, "二、职业认知")

    h2(doc, "2.1 行业环境：评价标准变了")
    para(doc, "2023 年，企业问的是「这东西能不能做出来」；2026 年，企业问的是「做完能省多少钱、能多赚多少钱」。"
              "同一项能力，评价口径完全变了。变化淘汰的不是技术弱的人，而是只会交付技术、不会交付价值的人。"
              "不少模型能力正在变成基础设施，会调用 API 不再是稀缺技能。")
    table(doc, ["观察", "数据", "对我们的含义"], [
        ["AI 岗位增长", "三年增长 119%（科锐国际 2026 报告）", "机会在「用 AI」一侧"],
        ["企业落地卡点", "57.4% 的企业不敢推进规模化（406 家企业调研）", "卡点在落地能力，不在技术"],
        ["收益实现", "仅 8% 的企业实现财务收益", "能定义价值的人稀缺"],
    ], widths=[110, 230, 160])
    para(doc, "注：以上三项为公开报道数据，正式提交前请核对原文出处；如无法核实，可改为定性表述。", indent=False, size=9)

    h2(doc, "2.2 目标职业分析：AI 应用产品经理")
    para(doc, "这个岗位位于「模型能力」与「业务价值」之间，核心工作是把模糊的用户需求翻译成团队可执行的方案，"
              "并在上线后用指标证明价值。它同时要求技术理解与用户洞察，恰好对应我的强项与缺口。")
    table(doc, ["维度", "岗位要求", "我的现状"], [
        ["职责", "需求定义、方案推进、效果验证", "了解，未独立做过"],
        ["技术", "能判断方案在工程上是否可行", "具备基础"],
        ["业务", "能理解行业场景与成本结构", "待补"],
        ["协作", "能跨部门推动项目上线", "待补"],
        ["产出", "可量化提升的产品指标", "待补"],
    ], widths=[60, 240, 160])

    h2(doc, "2.3 企业真正缺的三样东西")
    para(doc, "把需求翻译成方案、把效果量化出来、把项目推到上线。这三样都不是技术问题，但都卡在技术上——"
              "企业不缺做演示的人，缺能让演示进生产的人。")

    h2(doc, "2.4 胜任力差距")
    table(doc, ["胜任力", "企业要求", "我的差距", "补齐方式"], [
        ["需求洞察", "能独立完成调研并形成需求文档", "零经验", "20 人深度访谈"],
        ["效果量化", "能定义指标并做前后测", "零经验", "每个项目一张指标卡"],
        ["项目推动", "能按周推进并暴露风险", "弱", "周交付 + 周报制度"],
        ["技术判断", "能评估可行性", "基本具备", "保持，补成本判断"],
    ], widths=[80, 200, 100, 120])

    # ---------- 三、职业决策 ----------
    h1(doc, "三、职业决策")

    h2(doc, "3.1 目标：三年后独立负责一条 AI 产品线")
    para(doc, "选这个目标，是因为它同时覆盖我想练的三件事：定义需求、推动交付、对结果负责。"
              "更重要的是它有天然的检验方式——这条线上线后有没有人用、指标有没有变好，数据会说话。"
              "目标不是终点，是判断工具：以后每次纠结，我就问一句「这件事能不能让我更接近负责一条线」。")

    h2(doc, "3.2 决策依据")
    para(doc, "① 能力匹配：技术底子够用但不够深，做纯技术岗没有优势；做纯运营岗又缺少判断训练。"
              "产品经理是两者的接口，正好用得上我的长板、补得上我的短板。"
              "② 门槛匹配：不需要读研或大算力，靠项目与实习就能积累。"
              "③ 时间匹配：大一开始，比大三大四多两年试错窗口。")

    h2(doc, "3.3 备选方案")
    para(doc, "如果一年内发现强项其实在工程侧，转向 AI 解决方案工程师；如果行业降温，转向传统软件产品岗。"
              "这两条路与主路径共享「技术理解 + 表达翻译」的底层能力，切换成本低。"
              "备选不是退路，是同一棵能力树上的另一根枝。")

    # ---------- 四、计划实施 ----------
    h1(doc, "四、计划实施")

    h2(doc, "4.1 短期（本学期）")
    table(doc, ["目标", "行动", "交付物", "衡量标准"], [
        ["跑通一个 Agent 项目", "选真实场景，从需求出发做到可用", "能演示的成品", "别人能上手用一次"],
        ["完成需求调研", "深度访谈 20 位真实用户", "访谈记录 + 一页需求定义", "非技术者能看懂"],
        ["建立量化习惯", "每个项目定义 1 个指标并前后测", "指标卡", "能说清改变量"],
        ["沉淀部署能力", "整理环境、命令、坑与修法", "可复用部署笔记", "他人照抄能跑通"],
    ], widths=[110, 190, 120, 110])

    h2(doc, "4.2 中期（一年内）")
    para(doc, "进入实验室参与横向项目，或拿到一份 AI 产品相关实习。目标不是履历好看，是被真实需求推翻一次——"
              "只有在真实场景里，才能知道自己的判断错在哪。")
    para(doc, "【请填写：目标实验室 / 目标企业名单、投递时间点、需要的材料】")

    h2(doc, "4.3 长期（三年内）")
    para(doc, "独立负责一条 AI 产品线：需求自己定义，方案自己推进，指标自己兜底。"
              "把「做完」变成「做对」，完成从执行者到负责人的转变。")

    h2(doc, "4.4 时间与精力分配")
    table(doc, ["投入方向", "占比", "说明"], [
        ["真实项目", "50%", "不刷课，以项目带动学习"],
        ["见用户", "30%", "访谈、观察、复盘"],
        ["补基础", "20%", "只补项目卡住的部分"],
    ], widths=[120, 60, 300])

    # ---------- 五、评估调整 ----------
    h1(doc, "五、评估调整")

    h2(doc, "5.1 评估指标")
    table(doc, ["维度", "指标", "检查频率"], [
        ["产出", "每周是否有可演示的产出", "每周"],
        ["用户", "本月新增真实用户对话数", "每月"],
        ["量化", "在跑项目是否都有指标卡", "每项目"],
        ["能力", "能否独立写出被评审通过的需求文档", "每季度"],
        ["外部反馈", "导师 / 学长 / 实习主管的反馈意见", "每季度"],
    ], widths=[90, 260, 100])

    h2(doc, "5.2 调整机制")
    para(doc, "每季度做一次复盘，只回答三个问题：哪个假设被证伪了、哪个动作没产生效果、下一步删掉什么。"
              "如果连续两个季度没有可演示产出，视为方向或方法出了问题，触发重新评估。")

    h2(doc, "5.3 风险与预案")
    table(doc, ["风险", "触发信号", "预案"], [
        ["拿不到实习", "一年内无面试机会", "先做开源贡献补履历"],
        ["强项在工程侧", "连续项目更享受实现而非定义", "转向 AI 解决方案岗"],
        ["行业降温", "目标岗位招聘量明显下滑", "转向传统软件产品岗"],
    ], widths=[100, 200, 180])

    # ---------- 六、已落地证据 ----------
    h1(doc, "六、已落地证据")
    para(doc, "计划谁都会写，证据不会骗人。以下动作已经发生，可当场核验：")
    para(doc, "① GitHub 建仓并持续提交，提交记录公开可查。"
              "② 本地完整部署过一个开源 AI Agent，踩过依赖、模型接入与网络环境的坑。"
              "③ 用 AI 辅助完成了从选题到成片的整套演示，本报告配套 PPT 即是产出之一。")
    para(doc, "【请填写：你的 GitHub 主页链接、项目仓库链接、可展示的作品清单】", indent=True)

    # ---------- 附：待补充清单 ----------
    h1(doc, "附：提交前需要补齐的私人信息")
    for item in [
        "姓名、学校、学院、专业、年级、联系方式",
        "性格 / 兴趣测评结果（如有）",
        "成绩排名、已修核心课程",
        "竞赛、社团、志愿服务经历",
        "具体项目名称、时间、你的角色与产出",
        "目标企业与岗位名称（建议 2—3 家）",
        "指导老师姓名与意见",
    ]:
        p = doc.add_paragraph(style="List Bullet")
        r = p.add_run(item)
        set_run_font(r, CN, 10.5)

    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    print(f"[OK] {out}")


if __name__ == "__main__":
    build(Path(__file__).parent / "out" / "生涯发展报告_融合版.docx")
