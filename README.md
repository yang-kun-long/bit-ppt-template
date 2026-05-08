# BIT PPT 模板生成器

这是一个无需 LaTeX 的北京理工大学风格 PPTX 生成器。它读取 YAML
内容描述，输出可编辑的 PowerPoint 文件。

项目目标不是把幻灯片截图塞进 PPT，而是尽量生成原生 Office 对象：
文本框、形状、表格、图表、图片和 Office Math 公式都应保持可编辑。

## 特性

- Node.js ESM CLI
- YAML 输入，PPTX 输出
- 基于 `pptxgenjs` 生成可编辑 PPTX
- 基于 `jszip` 做 OpenXML 后处理
- 基于 `latex-to-omml` 将公式转换为原生 Office Math
- 支持渐进式 guide，方便 Codex / Claude Code 等本地 agent 查询能力
- 支持 `--json` 和 `--strict`，便于脚本、CI 和 AI agent 自动调用
- 支持图片尺寸读取和 `imageText` 自动排版

## 安装

```powershell
npm install
```

本地开发可以直接使用 Node 入口：

```powershell
node bin/bit-ppt.mjs --help
```

也可以链接为全局命令：

```powershell
npm link
bit-ppt --help
```

## 快速开始

生成示例 PPT：

```powershell
npm run build:ppt
npm run build:body-layouts
npm run build:charts
```

输出文件位于：

```text
output/example.pptx
output/body-layout-test.pptx
output/chart-flow-test.pptx
```

手动指定输入和输出：

```powershell
node bin/bit-ppt.mjs generate content/example.yaml output/example.pptx
```

先检查再生成：

```powershell
node bin/bit-ppt.mjs check content/example.yaml --json
node bin/bit-ppt.mjs generate content/example.yaml output/example.pptx --json
```

严格模式会把 warning 也视为失败：

```powershell
node bin/bit-ppt.mjs check content/formula-test.yaml --json --strict
node bin/bit-ppt.mjs generate content/example.yaml output/example.pptx --json --strict
```

## CLI 命令

```text
bit-ppt generate <input.yaml> <output.pptx> [--json] [--strict] [font options]
bit-ppt check <input.yaml> [--json] [--strict]
bit-ppt list-layouts [--json]
bit-ppt guide [topic] [name] [--json]
bit-ppt doctor [--json]
```

常用命令：

```powershell
node bin/bit-ppt.mjs list-layouts
node bin/bit-ppt.mjs list-layouts --json
node bin/bit-ppt.mjs doctor
node bin/bit-ppt.mjs doctor --json
```

`doctor` 会检查 Node 版本、依赖、关键素材、示例 deck、输出目录可写性。

## 给 AI Agent 的渐进式 Guide

不要一次性把完整文档塞给 AI。推荐让 agent 按需查询：

```powershell
node bin/bit-ppt.mjs guide
node bin/bit-ppt.mjs guide workflow --json
node bin/bit-ppt.mjs guide layouts
node bin/bit-ppt.mjs guide layout imageText
node bin/bit-ppt.mjs guide schema chart --json
node bin/bit-ppt.mjs guide example flowchart --json
node bin/bit-ppt.mjs guide speaker-notes
node bin/bit-ppt.mjs guide image-placeholder
node bin/bit-ppt.mjs guide writing-rules
node bin/bit-ppt.mjs guide all --json
```

推荐 agent 工作流：

```text
1. guide / list-layouts 查询能力
2. guide schema <layout> 获取结构化字段
3. 生成 YAML deck
4. check --json --strict
5. 根据 repairPrompt 修改 YAML
6. generate --json --strict
```

目前结构化 guide 重点覆盖：

- `imageText`
- `chart`
- `flowchart`
- `table`
- `formula`

通用字段可通过 `guide speaker-notes` 渐进式查询；`guide schema <layout> --json`
也会在 `commonFields` 中返回 `speakerNotes`。

暂无图片时可通过 `guide image-placeholder` 查询占位图写法。

完整写作约束见 [AI_CONTENT_GUIDE.md](AI_CONTENT_GUIDE.md)。

## YAML 输入结构

输入文件包含两个顶层字段：

- `meta`：标题、副标题、作者、日期、字体等元信息
- `slides`：按顺序排列的幻灯片定义

示例：

```yaml
meta:
  title: 示例报告
  subtitle: BIT PPT Generator
  author: Your Name
  date: 2026

slides:
  - layout: title

  - layout: bullets
    title: 核心观点
    bullets:
      - 输出是可编辑 PPTX。
      - YAML 更适合 AI 生成和修复。
      - check 命令会返回 repairPrompt。

  - layout: closing
    title: 谢谢
```

每页都可以添加 PowerPoint/WPS 演讲者备注。备注不会出现在页面画布上，会写入 PPT 的备注区：

```yaml
- layout: bullets
  title: 核心观点
  bullets:
    - 输出是可编辑 PPTX。
  speakerNotes: |
    这一页先解释为什么选择 YAML 到 PPTX 的路线。
    公式暂时按普通文本保留，例如 $L(\theta)$。
```

也可以用字符串数组：

```yaml
speakerNotes:
  - 第一段演讲稿。
  - 第二段演讲稿。
```

## 字体

默认字体：

- 中文：`微软雅黑`
- 中文 Light：`微软雅黑 Light`
- 英文：`Arial`
- 衬线：`SimSun`
- 代码：`Consolas`

可以在 YAML 中设置：

```yaml
meta:
  fonts:
    cn: 微软雅黑
    cnLight: 微软雅黑 Light
    en: Arial
    serif: SimSun
    code: Consolas
```

也可以用 CLI 覆盖：

```powershell
node bin/bit-ppt.mjs generate input.yaml output.pptx --font-cn "Noto Sans CJK SC" --font-code "Cascadia Mono"
```

注意：项目不会打包微软雅黑 TTF 文件，避免字体授权问题。

## 支持布局

查看当前支持布局：

```powershell
node bin/bit-ppt.mjs list-layouts
```

当前支持：

- `title`
- `agenda`
- `section`
- `bullets`
- `claim`
- `twoColumn`
- `cards`
- `table`
- `comparison`
- `timeline`
- `process`
- `architecture`
- `ablation`
- `caseStudy`
- `imageGrid`
- `code`
- `appendix`
- `flowchart`
- `chart`
- `problemSolution`
- `painOpportunity`
- `experimentDesign`
- `resultAnalysis`
- `riskMitigation`
- `contribution`
- `summary`
- `metrics`
- `matrix`
- `quote`
- `formula`
- `references`
- `imageText`
- `closing`

## 图片排版

图片路径默认相对项目根目录。

`imageText` 支持字符串路径：

```yaml
- layout: imageText
  title: 图文说明
  image: assets/bit-campus-photo.png
  text:
    - 说明一。
    - 说明二。
```

也支持对象写法：

```yaml
- layout: imageText
  title: 图文说明
  image:
    path: assets/bit-campus-photo.png
    placement: auto # auto | top | side
    fit: contain    # contain | cover
  caption: 图片说明。
  text:
    - 图像展示关键现象。
    - 文本保持简洁。
```

自动规则：

- 明显超宽图自动采用上图下字
- 普通横图、方图、竖图默认保留侧图侧字
- `placement: top` / `placement: side` 可手动覆盖
- `fit: contain` 避免裁切，`fit: cover` 填满图片框

项目包含一个图片排版示例：

```powershell
node bin/bit-ppt.mjs generate content/image-layout-demo.yaml output/image-layout-demo.pptx
```

暂无图片时，可以让 AI 生成图片描述并写入可编辑占位框：

```yaml
- layout: imageText
  title: 系统架构示意
  image:
    mode: placeholder
    aspectRatio: "16:9" # 也可省略；未知比例会生成候选页
    placement: top
    prompt: 展示 YAML 输入、结构校验、PPTX 生成、OMML 后处理的流程图。
  text:
    - 用户后续可在 WPS / PowerPoint 中替换占位框。
```

如果 `imageText` 的占位图没有明确 `aspectRatio`，预检会自动生成两页候选：

- 横图方案：上图下文
- 侧图方案：左图右文

占位图也支持 `caseStudy` 和 `imageGrid`。示例：

```powershell
node bin/bit-ppt.mjs check content/placeholder-image-demo.yaml --json
node bin/bit-ppt.mjs generate content/placeholder-image-demo.yaml output/placeholder-image-demo.pptx
```

## 公式

`formula` 布局会将 LaTeX 风格公式转换为原生 Office Math / OMML：

```yaml
- layout: formula
  title: 优化目标
  formula:
    latex: "\\mathcal{L}=\\sum_i (y_i-\\hat{y}_i)^2"
  explanation:
    - 目标函数衡量预测误差。
```

普通文本中也支持 inline math：

```yaml
bullets:
  - 使用 $E=mc^2$ 作为示例公式。
```

公式不是图片，生成后仍尽量保持 Office 可编辑。

## 校验与修复

`check` 会返回：

- `validation.errors`：阻止生成的硬错误
- `validation.warnings`：可能影响版面的风险
- `repairPrompt`：可直接反馈给 AI 的修复提示
- `actions`：预检自动动作，例如拆分长列表或长表格

示例：

```powershell
node bin/bit-ppt.mjs check content/invalid-deck-test.yaml --json
```

生成时遇到 errors 会停止。默认情况下 warnings 不阻止生成；使用
`--strict` 后 warnings 也会导致失败。

## 测试

```powershell
npm test
npm run check:ppt
npm run check:body-layouts
npm run check:charts
```

打包检查：

```powershell
npm pack --dry-run
```

## 项目结构

```text
bin/bit-ppt.mjs          CLI 入口
src/generate.mjs         核心生成器
src/layout-guides.mjs    面向 AI 的结构化 guide
content/                 示例 YAML 和测试 fixture
assets/                  BIT 风格素材
output/                  本地生成结果，默认不提交
test/                    node:test 测试
AI_CONTENT_GUIDE.md      完整内容写作指南
AGENTS.md                给 coding agent 的项目交接说明
```

## 设计原则

- 不依赖 Pandoc、LaTeX、Beamer
- 不把整页幻灯片渲染成截图
- 优先生成可编辑 PPTX 对象
- CLI、未来 MCP 和其他集成都复用同一套核心实现
- 对 AI 友好：结构化 guide、JSON 输出、strict 模式、repairPrompt
