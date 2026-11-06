---
name: kaizhi-slide
description: 用 open-slide 制作"开智风格"幻灯片（深蓝黑底 / 暖橙强调 / 金色衬线 / 描边动画）。当用户要求做开智风格 PPT、要求"框画出来"的动画、要求把 outline.md 转成幻灯片时使用。
---

# 开智风格幻灯片（open-slide）

把 `outline.md` 变成 open-slide 网页幻灯片，再导出 PDF / 静态站。

## 前置条件（先检查，缺了就装）

```bash
node --version    # 需要 18+
pnpm --version    # 没有就 corepack enable
```

Windows 安装：`winget install OpenJS.NodeJS.LTS`，然后**重开终端**再检查。

## 工作流

### 1. 初始化（只做一次）

```bash
npx @open-slide/cli init my-deck
cd my-deck
pnpm install
```

### 2. 读规范

脚手架里有 `AGENTS.md` 或 `CLAUDE.md`，**先读它**，里面有画布尺寸、字号比例、目录结构等硬规则。
后续每页都遵守它，不要自创布局体系。

### 3. 生成页面

每页 = `slides/<id>/index.tsx` 里的一个 React 组件，画布固定 **1920 × 1080**。
按下面的设计规范写，不要用模板、不要用渐变。

### 4. 预览与迭代

```bash
pnpm dev      # http://localhost:5173
```

用 Agent Browser 技能打开页面截图自检；用户点元素留评论后，按评论改代码。

### 5. 导出

```bash
pnpm export:pdf     # 可打印 PDF
pnpm build          # 静态站，dist/
```

## 设计规范（硬约束，不要改）

| 项 | 值 |
|---|---|
| 画布 | 1920 × 1080，页边距 158px |
| 背景 | `#0B1220` 纯色，禁止渐变 |
| 主文字 | `#F2F4F8` |
| 强调色 | `#E8763A` 暖橙 |
| 金色 | `#C9A45C`（小标签、页码、细线） |
| 辅助灰 | `#7C8698` |
| 中文字体 | `"Noto Sans SC", sans-serif`，主标题 700 |
| 英文字体 | `"Georgia", serif`，斜体 |
| 主字幕字号 | 64px（≈PPT 的 32pt） |
| 英文副字幕 | 主字幕的 60%，即 38px，灰色 |
| 小标签 | 22px，Georgia 斜体，金色 |
| 版式 | 左对齐、大留白、每页一个观点 |
| 禁止 | 居中大字、满屏小字、渐变背景、圆角卡片、阴影 |

## 描边动画（这是本技能的核心）

用 SVG + `stroke-dasharray` 实现"有人拿笔围起来"的效果，**不要用 CSS 假装**：

```tsx
// 一条线：pathLength 归一化后，dashoffset 从 100 走到 0 就是"画出来"
const DrawPath = ({ d, delay = 0, dur = 0.42, color = "#E8763A" }) => (
  <path
    d={d}
    fill="none"
    stroke={color}
    strokeWidth={3}
    pathLength={100}
    style={{
      strokeDasharray: 100,
      strokeDashoffset: 100,
      animation: `draw ${dur}s linear ${delay}s forwards`,
    }}
  />
);
// keyframes: from { stroke-dashoffset: 100 } to { stroke-dashoffset: 0 }
```

四条边按顺时针依次延迟（0 / 0.28 / 0.56 / 0.84 秒），总时长约 1.3 秒：

```tsx
<svg viewBox="0 0 1000 260">
  <DrawPath d="M0,0 H1000"            delay={0}    />  {/* 上 */}
  <DrawPath d="M1000,0 V260"          delay={0.28} />  {/* 右 */}
  <DrawPath d="M1000,260 H0"          delay={0.56} />  {/* 下 */}
  <DrawPath d="M0,260 V0"             delay={0.84} />  {/* 左 */}
</svg>
```

强调整句放框里；句中关键词用「」包住并设为 `#E8763A`。

## 常见错误

| 现象 | 原因 | 修法 |
|---|---|---|
| `pnpm: command not found` | 没装或没重开终端 | `corepack enable` 后重开 |
| 页面全白 | 组件没 `export default [...]` | 按 AGENTS.md 的导出格式写 |
| 中文变成雅黑 | 字体未装 | 装 Noto Sans SC / Noto Serif SC |
| 动画不动 | 用了 transition 而非 keyframes | 用 `@keyframes` + `forwards` |
| 导出报错 | 依赖没装全 | `pnpm install` 重跑 |
