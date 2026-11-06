# AI PPT 流水线：WorkBuddy → Markdown → open-slide

> 目标：用 WorkBuddy 出内容，用 open-slide 重排版式，"内容"和"设计"分开管。
> 环境：Windows 11 / PowerShell。所有命令在 PowerShell 里逐行粘贴执行。

---

## 0. 这台机器的现状（2026-11-05 实测）

| 需要的东西 | 现状 | 怎么处理 |
|---|---|---|
| Node.js 18+ | ❌ 没装 | `winget install OpenJS.NodeJS.LTS` |
| pnpm | ❌ 没装 | 装完 Node 后 `corepack enable` |
| Git | ❌ 没装 | `winget install --id Git.Git -e` |
| Python 3.14 | ✅ 有 | 够了 |
| python-pptx | ❌ 没装 | `python -m pip install python-pptx` |
| winget | ✅ 有 | 装软件用得上 |

**一句话：第 3 步（open-slide）现在跑不了，因为没 Node。先做第 0 步。**

---

## 1. 装环境（约 15 分钟）

```powershell
# 1.1 Node.js LTS（会带 npm；装完重开一个 PowerShell 窗口）
winget install OpenJS.NodeJS.LTS

# 1.2 验证（必须重开窗口，否则 PATH 没生效）
node --version      # 期望 v20 或 v22
npm --version

# 1.3 启用 pnpm
corepack enable
corepack prepare pnpm@latest --activate
pnpm --version

# 1.4 Git
winget install --id Git.Git -e
git --version

# 1.5 给系统 Python 补 pptx 库
python -m pip install python-pptx
```

> 报错自查：`node : 无法将"node"项识别为...` = 没重开窗口，不是装失败。

---

## 2. WorkBuddy 出大纲（约 30 分钟）

**关键认知：不要一上来就让它出 PPT。** 先出大纲，改到你满意，再出片。
否则 10 页里 9 页要重做。

### 2.1 提示词模板

见同目录 `01_workbuddy_提示词.md`，复制进去改主题即可。

### 2.2 必须检查的一件事：演讲者备注

你后面要"提取每页标题 + 字幕句"，**字幕句就存在 PPT 的"演讲者备注（Notes）"里**。
WorkBuddy 生成时如果只把讲稿写在正文，提取出来就是一堆散句。

所以提示词里必须明确要求：

> 每一页的"演讲者备注"里写一句 40 字以内的讲稿（口语，能直接念）。
> 正文只放 3 条以内的短句，不要整段文字。

### 2.3 验收标准

- [ ] 页数 8–12 页
- [ ] 每页正文 ≤ 3 条，每条 ≤ 15 字
- [ ] 每页备注都有 1 句讲稿
- [ ] 导出为 `.pptx`（不要 PNG / 不要 PDF）

---

## 3. 导出 PPTX → 提取 Markdown（约 5 分钟）

WorkBuddy 导出 `.pptx`，丢进本目录，然后：

```powershell
# 完整版：标题 + 正文 + 讲稿，给人看
python 02_extract_pptx.py 你的文件.pptx

# 只要标题清单：用来核对结构对不对
python 02_extract_pptx.py 你的文件.pptx --outline

# 紧凑版：一页一段，整段粘贴给 AI 用
python 02_extract_pptx.py 你的文件.pptx --sheet
```

输出三份 md，文件名同目录。抽取逻辑：
1. 标题 → 取占位符标题，取不到就取第一个非空文本框首行
2. 正文 → 除标题外的所有文本框，去重保序
3. 讲稿 → 演讲者备注

> 已验证：用 `_make_demo.py` 生成的 5 页测试稿，三份输出都正常。

---

## 4. open-slide 视觉重置（约 2–4 小时，这是最重的一步）

**先说清楚它是什么**：open-slide 不是"上传 PPT 自动美化"的网站。
它是一个 **React 幻灯片框架**——每页幻灯片 = 一个 React 组件（画布 1920×1080），
由 AI 编码助手（Claude Code / Codex / Cursor / DSH）写代码生成。
所以"喂 markdown"的真实含义是：**把 markdown 和风格规范一起交给 AI 助手，让它写页面代码。**

### 4.1 初始化项目

```powershell
# 在你的工作目录（不要放在 Atom 安装目录里）
npx @open-slide/cli init my-slide
cd my-slide
pnpm install
pnpm dev          # 打开 http://localhost:5173
```

### 4.2 把 Markdown 喂进去

打开你的 AI 编码助手，指向 `my-slide` 目录，发这一段：

```
这是我的一份 PPT 内容（Markdown），请用 open-slide 的 slide-authoring 规范，
把它重做成 10 页幻灯片。要求：
1. 每页只呈现一个核心观点，标题用大字号，正文最多 3 条；
2. 风格：白底 #FAFAF8、标题 #111 思源宋体加粗、正文 #333 思源黑体；
   只用一种强调色 #C0392B，禁止渐变、阴影、圆角卡片；
3. 大量留白，页边距 ≥ 120px，右下角小字页码；
4. 讲稿不上海报，放进页面组件的 note 里。

<粘贴 md 全文>
```

### 4.3 迭代（open-slide 的精髓）

浏览器里点任意元素 → 写评论（"这个标题太大了"）→ 回到 AI 助手执行 `/apply-comments`。
评论会存成源码里的 `@slide-comment` 标记，AI 批量改。**比来回截图讲需求快 10 倍。**

### 4.4 导出

```powershell
pnpm build        # 静态站点，dist/ 目录
pnpm export:pdf   # 可打印 PDF
```

### 4.5 部署（可选，跟"学 GitHub"合并练手）

```powershell
git init
git add .
git commit -m "first slide"
git remote add origin https://github.com/你的用户名/my-slide.git
git push -u origin main
# 然后到 vercel.com 导入仓库，构建命令 pnpm build，输出目录 dist
```

---

## 5. 时间预算

| 步骤 | 顺利 | 不顺利 |
|---|---|---|
| 装环境 | 20 min | 1.5 h（权限/网络） |
| WorkBuddy 大纲 | 30 min | 1 h |
| 提取 Markdown | 5 min | 20 min |
| open-slide 重排 | 2 h | 4 h+（第一次装依赖慢） |
| 部署 | 30 min | 1 h |

**总：顺利半天，不顺利一整天。别指望两小时搞定。**

---

## 6. 兜底方案（重要）

如果 **Node 装不上** 或 **open-slide 卡住**，别死磕，换这条：

> WorkBuddy 出大纲 → 提取 markdown → **用 python-pptx 套固定模板生成 PPTX**

好处：不依赖 Node，产出是真 `.pptx`（能交作业/发领导）。
坏处：版式是模板化的，不如 open-slide 自由。

这条路径本目录已有基础（`02_extract_pptx.py` 反向思路），需要的话我可以补一个 `03_build_pptx.py`。

---

## 7. 已确认的决定（2026-11-05）

1. **交付格式 = .pptx** → open-slide 退为可选分支，**主线改成第 6 节的兜底方案**，
   已经落地成脚本：[03_build_deck.py](03_build_deck.py)，样张见 `demo/kaizhi_style_v1.pptx`。
2. **"开智风格"定义**（口述固化，写在 `03_build_deck.py` 顶部的设计令牌里）：

   | 项 | 值 |
   |---|---|
   | 背景 | `#0B1220` 深蓝黑，纯色无渐变 |
   | 中文主字幕 | 思源黑体 Bold，32pt（封面 44pt） |
   | 英文副字幕 | Georgia Italic，19pt（≈ 主字幕 60%），灰色 `#7C8698` |
   | 强调色 | 暖橙 `#E8763A`（细线、序号） |
   | 金色衬线 | `#C9A45C`（小标签、页码、短横线） |
   | 版式 | 左对齐、大留白、每页一个观点、禁止居中和满屏小字 |
   | 转场 | 仅淡出（已写入 XML），禁止弹跳/翻转 |

3. **动画**：已实现。描边框拆成四条线，顺时针擦除（与上一动画同时 + 延迟），
   代码在 `kz.add_draw_animation()`，总时长约 1.1 秒。
   用 `--no-anim` 可关掉（PowerPoint 报错时兜底）。
   完整落地步骤见 [06_实施手册.md](06_实施手册.md)。
4. **强调框**：整行重点 → 描边框；句中关键词 → 「」。
   实现在 [kz.py](kz.py) 的 `emphasis_frame()`，三种样式：
   - `full` 四边闭合细框（橙 `#E8763A`）
   - `corner` 只留四个角（金 `#C9A45C`）
   - `quote` 「」日式引号（橙色括号 + 白字）

   对比样张：`demo/box_styles.pptx`（2 页）。
   **为什么要分两种**：PowerPoint 没有"给段落里某几个字加边框"的能力，
   所以行内重点只能用引号，整行重点才能加框。

5. **代码结构**：`kz.py` 是唯一的设计层（颜色/字号/字体/框），
   `03_build_deck.py` 和 `04_box_styles.py` 只写内容，改样式不要碰它们。

## 8. 还没做 / 待确认

- 主字幕 32pt 在成片里够不够大（视频可能要 40pt+）
- 画幅 16:9 还是 9:16
- 英文用 Georgia 还是思源宋体（中文字体做斜体是假斜，效果差）
