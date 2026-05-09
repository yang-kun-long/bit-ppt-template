# BIT PPT 模板生成器

这是一个无需 LaTeX 的北京理工大学风格 PPTX 生成器。它读取 YAML
内容描述，输出可编辑的 PowerPoint 文件。

项目目标不是把幻灯片截图塞进 PPT，而是尽量生成原生 Office 对象：
文本框、形状、表格、图表、图片和 Office Math 公式都应保持可编辑。

## 项目来源与相关路线

本项目的视觉风格和最初需求来自 TeXPage 上的北京理工大学 LaTeX 幻灯片模板：
[BIThesis Beamer Slide Template](https://www.texpage.com/zh/template/18f4a5d2-2167-47b8-b193-85e2475e7a06)。
原模板适合熟悉 LaTeX / Beamer 的用户，但 LaTeX 工具链较重，输出也通常是 PDF。
PDF 用 WPS 播放比较方便；如果没有 WPS，常见的 PDF 转 PPT 流程又容易造成格式错乱。

因此这里提供另一条路线：用 YAML 描述内容，直接生成可编辑 PPTX，尽量保留文本、
表格、图表、形状和 Office Math 公式的可编辑性。

如果你的需求是“已有 PDF 幻灯片，只想在线展示或播放”，可以使用另一个项目
[pptView](https://github.com/yang-kun-long/pptView)。在线演示：
[https://ppt.yangkunlong.top](https://ppt.yangkunlong.top)。

代码使用 MIT License。北京理工大学名称、标识和视觉参考归其各自权利人所有。

## 特性

- Node.js ESM CLI
- YAML 输入，PPTX 输出
- 基于 `pptxgenjs` 生成可编辑 PPTX
- 基于 `jszip` 做 OpenXML 后处理
- 基于 `latex-to-omml` 将公式转换为原生 Office Math
- 支持渐进式 guide，方便 Codex / Claude Code 等本地 agent 查询能力
- 支持 `--json` 和 `--strict`，便于脚本、CI 和 AI agent 自动调用
- 支持图片尺寸读取和 `imageText` 自动排版

## 三种使用方式

同一个 npm 包同时提供 Web UI、CLI 和 MCP 三个入口。普通用户优先使用 Web UI；
脚本和 CI 使用 CLI；Codex、Claude Code 等本地 agent 使用 MCP。

### 1. 普通用户：打开本地网页

无需安装，直接启动：

```powershell
npx bit-ppt-generator
```

然后打开命令行输出的本地地址，默认是：

```text
http://127.0.0.1:3000/
```

本地网页默认不需要登录；只有部署者配置了鉴权环境变量时才会显示登录区。

也可以全局安装后启动：

```powershell
npm install -g bit-ppt-generator
bit-ppt-generator
```

### 2. 命令行用户：检查和生成 PPTX

```powershell
npm install -g bit-ppt-generator
bit-ppt check content/example.yaml --json
bit-ppt generate content/example.yaml output/example.pptx
```

常用 CLI：

```text
bit-ppt generate <input.yaml> <output.pptx> [--json] [--strict] [font options]
bit-ppt check <input.yaml> [--json] [--strict]
bit-ppt list-layouts [--json]
bit-ppt guide [topic] [name] [--json]
bit-ppt doctor [--json]
```

### 3. Agent 用户：配置 MCP

全局安装后，MCP 客户端可以直接启动 npm 包提供的 `bit-ppt-mcp` 命令：

```powershell
npm install -g bit-ppt-generator
bit-ppt-mcp --help
```

MCP 客户端配置示例：

```json
{
  "mcpServers": {
    "bit-ppt": {
      "command": "bit-ppt-mcp",
      "args": []
    }
  }
}
```

如果 MCP 客户端不继承系统 PATH，改用 `npx` 启动：

```json
{
  "mcpServers": {
    "bit-ppt": {
      "command": "npx",
      "args": ["-y", "--package", "bit-ppt-generator", "bit-ppt-mcp"]
    }
  }
}
```

不再需要写类似 `D:/atuodl/presentation-slide/bit-ppt-template/bin/bit-ppt-mcp.mjs`
这样的本机源码路径。那种写法只适合仓库开发阶段；发布到 npm 后应使用包里的
`bit-ppt-mcp` 命令。

## 仓库开发

克隆仓库后安装依赖：

```powershell
npm install
```

启动本地网页：

```powershell
npm run serve
```

生成示例 PPT：

```powershell
npm run build:ppt
npm run build:body-layouts
npm run build:charts
```

源码入口：

```powershell
node bin/bit-ppt.mjs --help
node bin/bit-ppt-http.mjs --help
node bin/bit-ppt-mcp.mjs --help
```

本仓库采用单主干、多入口策略。Web UI、CLI、MCP 和 Node HTTP API 都复用同一套核心生成逻辑。

## npm 包入口

- Web UI：默认入口为 `bit-ppt-generator`，适合 npm 用户本地打开网页
- CLI：已支持，入口为 `bit-ppt` 或 `node bin/bit-ppt.mjs`
- MCP：已支持，入口为 `bit-ppt-mcp` 或 `node bin/bit-ppt-mcp.mjs`
- Node HTTP API：已支持，入口为 `bit-ppt-http` 或 `node bin/bit-ppt-http.mjs`

一个 npm release 同时包含 Web UI、CLI 和 MCP 三个入口，不拆成三个包。

## CLI 命令

```text
bit-ppt-generator
bit-ppt generate <input.yaml> <output.pptx> [--json] [--strict] [font options]
bit-ppt check <input.yaml> [--json] [--strict]
bit-ppt list-layouts [--json]
bit-ppt guide [topic] [name] [--json]
bit-ppt doctor [--json]
bit-ppt-mcp
bit-ppt-http
```

常用命令：

```powershell
bit-ppt list-layouts
bit-ppt list-layouts --json
bit-ppt doctor
bit-ppt doctor --json
```

`doctor` 会检查 Node 版本、依赖、关键素材、示例 deck、输出目录可写性。

## MCP Server

项目同时提供 stdio MCP 入口，供 Codex、Claude Code 等本地 agent 调用。
MCP 只是 adapter，仍复用 CLI 背后的同一套生成和校验函数。

本地运行：

```powershell
bit-ppt-mcp
```

从源码运行：

```powershell
node bin/bit-ppt-mcp.mjs
```

MCP 客户端配置示例，优先使用 npm 包命令：

```json
{
  "mcpServers": {
    "bit-ppt": {
      "command": "bit-ppt-mcp",
      "args": []
    }
  }
}
```

提供的 MCP tools：

- `list_layouts`
- `get_guide`
- `validate_deck`
- `preflight_deck`
- `get_repair_prompt`
- `generate_pptx`

`validate_deck`、`preflight_deck` 和 `get_repair_prompt` 支持 `inputPath`
或 `deckYaml`。`generate_pptx` 使用文件路径：

```json
{
  "inputPath": "content/example.yaml",
  "outputPath": "output/example-from-mcp.pptx",
  "strict": true
}
```

## Node HTTP API

项目提供一个轻量 Node HTTP 服务，用于上传 YAML 并下载生成的 PPTX。

普通本地用户启动网页：

```powershell
npx bit-ppt-generator
```

全局安装后：

```powershell
bit-ppt-generator
```

仓库开发启动：

```powershell
npm run serve
```

或指定地址：

```powershell
node bin/bit-ppt-http.mjs --host 127.0.0.1 --port 3000
```

本地模式未设置鉴权环境变量时不需要登录，网页也不会显示登录框。公网部署时建议启用鉴权和生成并发上限：

```powershell
$env:BIT_PPT_AUTH_SECRET="change-this-long-random-secret"
$env:BIT_PPT_SESSION_TTL_SECONDS="604800"
$env:BIT_PPT_TOKEN="change-this-token"
$env:BIT_PPT_MAX_GENERATE_CONCURRENCY="1"
npm run serve
```

Linux systemd 可设置同名环境变量。

- 设置 `BIT_PPT_AUTH_SECRET` 后，网页会显示北理工登录，登录成功后使用签名 token。
- 设置 `BIT_PPT_TOKEN` 后，HTTP API 也接受固定 Bearer token。
- 两者都不设置时，适合本机使用，`/check` 和 `/generate` 不需要鉴权。

固定 token 调用示例：

```http
Authorization: Bearer change-this-token
```

端点：

- `GET /`
- `GET /health`
- `POST /auth/bit-login`
- `POST /auth/verify`
- `POST /check`
- `POST /generate`

请求体支持 raw YAML，或 JSON：

```json
{
  "deckYaml": "slides:\n  - layout: bullets\n    title: Demo\n    bullets:\n      - Upload YAML\n      - Download PPTX\n",
  "outputName": "demo"
}
```

PowerShell 示例：

```powershell
Invoke-RestMethod `
  -Method Post `
  -Uri http://127.0.0.1:3000/check `
  -ContentType "application/json" `
  -Headers @{ Authorization = "Bearer change-this-token" } `
  -Body '{"deckYaml":"slides:\n  - layout: bullets\n    title: Demo\n    bullets:\n      - Upload YAML\n"}'

Invoke-WebRequest `
  -Method Post `
  -Uri http://127.0.0.1:3000/generate `
  -ContentType "application/json" `
  -Headers @{ Authorization = "Bearer change-this-token" } `
  -Body '{"deckYaml":"slides:\n  - layout: bullets\n    title: Demo\n    bullets:\n      - Upload YAML\n      - Download PPTX\n","outputName":"demo"}' `
  -OutFile output/demo.pptx
```

YAML 语法错误会返回结构化诊断，适合网页直接展示给用户复制回大模型：

```json
{
  "error": "deckYaml syntax error at line 2, column 5: ...",
  "syntax": {
    "errors": [
      {
        "level": "error",
        "code": "MULTILINE_IMPLICIT_KEY",
        "message": "Implicit keys need to be on a single line",
        "line": 2,
        "column": 5,
        "context": "1 | slides:\n2 |   - layout bullets\n3 |     title: Broken",
        "pointer": "    ^"
      }
    ]
  },
  "repairPrompt": "deckYaml syntax error at line 2, column 5: ...\nFix the YAML syntax first, then keep the deck schema unchanged."
}
```

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

## 发布

npm 包名为 `bit-ppt-generator`。后续版本通过 GitHub Release 触发
`.github/workflows/publish.yml`，并使用 npm Trusted Publisher 发布，不需要在
GitHub Secrets 保存 npm token。

Trusted Publisher 配置：

```text
Publisher: GitHub Actions
Organization or user: yang-kun-long
Repository: bit-ppt-template
Workflow filename: publish.yml
Environment name: 留空
```

发布新版本：

```powershell
npm version patch
git push
git push --tags
```

随后在 GitHub 创建并发布对应 tag 的 Release。

## 项目结构

```text
bin/bit-ppt.mjs          CLI 入口
bin/bit-ppt-http.mjs     Web UI / Node HTTP 入口，也是 bit-ppt-generator 默认命令
bin/bit-ppt-mcp.mjs      MCP 入口
src/core/                纯校验和预检 core
src/generate.mjs         核心生成器
src/http-server.mjs      Node HTTP 服务
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
- CLI、MCP、Node HTTP 和其他集成都复用同一套核心实现
- 对 AI 友好：结构化 guide、JSON 输出、strict 模式、repairPrompt
