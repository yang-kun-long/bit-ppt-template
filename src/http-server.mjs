import fs from "node:fs";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import { pathToFileURL } from "node:url";
import YAML from "yaml";
import { parseDeckYaml } from "./core/yaml-parse.mjs";
import {
  checkDeck,
  generateDeckFile,
} from "./generate.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const DEFAULT_BODY_LIMIT = 10 * 1024 * 1024;
const DEFAULT_GENERATE_CONCURRENCY = 1;
const BIT_CAS_TICKETS_URL = "https://sso.bit.edu.cn/cas/v1/tickets";
const DEFAULT_SESSION_TTL_SECONDS = 7 * 24 * 60 * 60;
const WEB_APP_HTML = String.raw`<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>BIT PPT Generator</title>
  <style>
    :root { color-scheme: light; --red: #9d1d22; --ink: #20242a; --muted: #68707a; --line: #d8dde3; --bg: #f5f7fa; --panel: #ffffff; }
    * { box-sizing: border-box; }
    body { margin: 0; font-family: "Microsoft YaHei", "Segoe UI", Arial, sans-serif; color: var(--ink); background: var(--bg); }
    header { background: var(--red); color: #fff; padding: 18px 24px; }
    header h1 { margin: 0; font-size: 22px; font-weight: 650; letter-spacing: 0; }
    main { max-width: 1320px; margin: 0 auto; padding: 22px; display: grid; gap: 16px; grid-template-columns: 320px 1fr; }
    section { background: var(--panel); border: 1px solid var(--line); border-radius: 8px; padding: 16px; }
    h2 { margin: 0 0 14px; font-size: 16px; }
    h3 { margin: 16px 0 8px; font-size: 14px; }
    label { display: block; margin: 12px 0 6px; color: var(--muted); font-size: 13px; }
    input, textarea { width: 100%; border: 1px solid var(--line); border-radius: 6px; padding: 10px 11px; font: inherit; background: #fff; color: var(--ink); }
    textarea { min-height: 520px; resize: vertical; font-family: Consolas, "Cascadia Mono", monospace; font-size: 13px; line-height: 1.45; }
    .mini { min-height: 132px; max-height: 220px; font-size: 12px; }
    .row { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
    button { border: 1px solid var(--red); background: var(--red); color: #fff; border-radius: 6px; padding: 9px 12px; font: inherit; cursor: pointer; }
    button.secondary { background: #fff; color: var(--red); }
    button:disabled { opacity: .55; cursor: not-allowed; }
    a { color: var(--red); text-decoration: none; }
    a:hover { text-decoration: underline; }
    .intro { margin: 0 0 14px; color: var(--muted); font-size: 13px; line-height: 1.6; }
    .repo-link { display: inline-flex; align-items: center; gap: 6px; margin-bottom: 12px; font-size: 13px; font-weight: 650; }
    .repo-icon { width: 16px; height: 16px; fill: currentColor; flex: 0 0 auto; }
    .guide { display: grid; gap: 10px; margin: 10px 0 14px; }
    .guide-step { border-left: 3px solid #d9b6b8; padding: 2px 0 2px 10px; }
    .guide-step strong { display: block; margin-bottom: 4px; font-size: 13px; }
    .guide-step p { margin: 0; color: var(--muted); font-size: 12px; line-height: 1.55; }
    .yaml-section { min-width: 0; }
    .yaml-workspace { display: grid; gap: 16px; grid-template-columns: minmax(0, 1fr) 290px; align-items: start; }
    .yaml-workspace textarea { min-height: 640px; }
    .side-guide { border-left: 1px solid var(--line); padding-left: 16px; }
    .status { min-height: 42px; white-space: pre-wrap; font-family: Consolas, "Cascadia Mono", monospace; font-size: 12px; color: var(--muted); }
    .ok { color: #137333; }
    .err { color: #b3261e; }
    .hidden { display: none; }
    @media (max-width: 1060px) { .yaml-workspace { grid-template-columns: 1fr; } .side-guide { border-left: 0; border-top: 1px solid var(--line); padding: 14px 0 0; } }
    @media (max-width: 860px) { main { grid-template-columns: 1fr; padding: 14px; } textarea, .yaml-workspace textarea { min-height: 420px; } }
  </style>
</head>
<body>
  <header><h1>BIT PPT Generator</h1></header>
  <main>
    <section>
      <div id="authSection" class="hidden">
        <h2>北理工登录</h2>
        <label for="username">学号</label>
        <input id="username" autocomplete="username" />
        <label for="password">密码</label>
        <input id="password" type="password" autocomplete="current-password" />
        <div class="row" style="margin-top: 14px;">
          <button id="loginBtn">登录</button>
          <button id="logoutBtn" class="secondary">退出</button>
        </div>
        <div id="authStatus" class="status"></div>
      </div>
      <a class="repo-link" href="https://github.com/yang-kun-long/bit-ppt-template" target="_blank" rel="noreferrer">
        <svg class="repo-icon" viewBox="0 0 16 16" aria-hidden="true">
          <path d="M8 0C3.58 0 0 3.64 0 8.13c0 3.59 2.29 6.63 5.47 7.7.4.07.55-.18.55-.39 0-.19-.01-.83-.01-1.51-2.01.38-2.53-.5-2.69-.96-.09-.23-.48-.96-.82-1.15-.28-.15-.68-.52-.01-.53.63-.01 1.08.59 1.23.83.72 1.23 1.87.88 2.33.67.07-.53.28-.88.51-1.09-1.78-.2-3.64-.9-3.64-4.01 0-.89.31-1.61.82-2.18-.08-.2-.36-1.03.08-2.15 0 0 .67-.22 2.2.83A7.5 7.5 0 0 1 8 3.92c.68 0 1.36.09 2 .27 1.53-1.05 2.2-.83 2.2-.83.44 1.12.16 1.95.08 2.15.51.57.82 1.29.82 2.18 0 3.12-1.87 3.81-3.65 4.01.29.25.54.74.54 1.5 0 1.09-.01 1.97-.01 2.24 0 .21.15.46.55.39A8.03 8.03 0 0 0 16 8.13C16 3.64 12.42 0 8 0Z"/>
        </svg>
        GitHub 仓库
      </a>
      <p class="intro">这是一个把 YAML 生成可编辑 PPTX 的网页入口。推荐先让 AI 产出 YAML，再在这里检查和生成；报错时把检查结果交给 AI 修改。</p>
      <h2>输出</h2>
      <label for="outputName">文件名</label>
      <input id="outputName" value="bit-ppt" />
      <div class="row" style="margin-top: 14px;">
        <button id="checkBtn" class="secondary">检查</button>
        <button id="generateBtn">生成 PPTX</button>
      </div>
      <div id="runStatus" class="status"></div>
      <h2>AI 辅助</h2>
      <div class="row">
        <button id="copyPromptBtn" class="secondary">复制提示词</button>
        <button id="copyRulesBtn" class="secondary">复制语法规则</button>
        <button id="copyWorkflowBtn" class="secondary">复制使用教程</button>
        <button id="copyErrorHelpBtn" class="secondary">复制报错求助</button>
        <button id="insertExampleBtn" class="secondary">插入最小示例</button>
        <button id="insertFullExampleBtn" class="secondary">插入完整示例</button>
      </div>
      <div id="copyStatus" class="status"></div>
      <h3>提示词</h3>
      <textarea id="aiPrompt" class="mini" readonly>你是一个 PPT 内容规划助手。请根据我给出的主题和材料，生成可直接用于 BIT PPT Generator 的 YAML。

硬性要求：
1. 只输出 YAML 正文，不要 Markdown 代码块，不要解释。
2. 顶层必须包含 meta 和 slides。
3. slides 中每一页必须包含 layout。
4. 使用北理工学术汇报风格：标题明确、要点简短、层次清楚、避免营销化表达。
5. 每页内容不要过满。bullets 建议 3-5 条，长内容拆成多页。
6. 不要编造本地图片路径。没有图片时使用 image.mode: placeholder，并写清 image.prompt。
7. 需要演讲稿时使用 speakerNotes；备注写演讲提示，不要重复整页正文。
8. 公式使用 LaTeX。行内公式写在正文中，例如 $p_{\theta}(x)$。展示公式使用 formula.latex。
9. YAML 中 LaTeX 推荐不用双引号；可写 plain scalar，例如 latex: \mathcal{L}=...，或用单引号。
10. 生成后应能通过 check：不要使用未知 layout，不要输出过长表格或过长参考文献。

优先使用的 layout：
- title: 封面
- agenda: 目录
- section: 章节分隔
- bullets: 要点页
- claim: 单页核心结论
- twoColumn: 双栏对比
- cards: 多卡片观点
- table: 表格
- comparison: 左右方案对比
- timeline: 时间线
- process: 流程步骤
- metrics: 指标页
- matrix: 判断矩阵
- quote: 引用页
- formula: 展示公式
- chart: 原生图表
- flowchart: 可编辑流程图
- architecture: 架构页
- experimentDesign: 实验设计
- resultAnalysis: 结果分析
- imageText: 图文页或图片占位
- imageGrid: 多图页
- code: 代码/伪代码
- references: 参考文献
- closing: 结束页

请根据材料自行选择合适布局，不必每种都用。输出要完整、可生成。

主题：

材料：</textarea>
      <h3>语法速查</h3>
      <textarea id="syntaxRules" class="mini" readonly>基本结构：
meta:
  title: 标题
  author: 作者
  date: 2026.05
slides:
  - layout: title
    title: 主标题
    subtitle: 副标题
    speakerNotes: |
      这里写演讲者备注。
      备注不会出现在幻灯片画布上。

常用页面：
- layout: bullets
  title: 页面标题
  lead: 页面导语，可选
  bullets:
    - 要点一
    - 要点二
  speakerNotes:
    - 备注也可以写成数组。
    - 每条是一段提示。

- layout: claim
  title: 核心结论
  claim: 一句话结论，可包含 $p_{\theta}(x)$。
  evidence:
    - 证据一
    - 证据二

- layout: twoColumn
  title: 双栏
  left:
    title: 左栏
    bullets: [A, B]
  right:
    title: 右栏
    bullets: [C, D]

- layout: cards
  title: 卡片页
  cards:
    - title: 卡片一
      text: 简短说明
    - title: 卡片二
      text: 简短说明

- layout: table
  title: 表格
  columns: [符号, 含义, 示例公式]
  rows:
    - ['$\theta$', 模型参数, '$p_\theta(x_i)$']
    - ['$x_i$', 第 i 个样本, '$\mathcal{L}_i=-y_i\log p_\theta(x_i)$']

- layout: comparison
  title: 方案对比
  left:
    label: 方案 A
    title: 传统做法
    bullets: [问题一, 问题二]
  right:
    label: 方案 B
    title: 推荐做法
    bullets: [优势一, 优势二]

- layout: timeline
  title: 路线图
  items:
    - date: 阶段 1
      title: 准备
      text: 收集材料
    - date: 阶段 2
      title: 生成
      text: 导出 PPTX

- layout: process
  title: 流程
  steps:
    - title: 输入
      text: 整理材料
    - title: 输出
      text: 生成 PPTX

- layout: metrics
  title: 指标页
  metrics:
    - value: 92%
      label: 准确率
      note: 实验集 A
    - value: 0
      label: 错误数
      note: 校验通过

- layout: matrix
  title: 判断矩阵
  cells:
    - title: 低成本 / 高控制
      text: 推荐路线
    - title: 高成本 / 低控制
      text: 不推荐

- layout: formula
  title: 展示公式
  formula:
    latex: \mathcal{L}(\theta)=-\sum_i y_i\log p_\theta(x_i)
  caption: 公式说明
  explanation:
    - '$\theta$ 表示模型参数。'
    - '$p_\theta(x_i)$ 表示预测概率。'

- layout: chart
  title: 图表
  type: bar
  categories: [A, B, C]
  series:
    - name: Baseline
      values: [80, 82, 84]
    - name: Ours
      values: [85, 87, 90]
  caption: 原生 PowerPoint 图表。

- layout: flowchart
  title: 流程图
  nodes:
    - id: input
      text: 输入
      note: 材料
    - id: output
      text: 输出
      note: PPTX
  edges:
    - from: input
      to: output

- layout: architecture
  title: 架构
  layers:
    - title: 输入层
      components: [论文, 数据, 图片]
      note: 材料整理
    - title: 生成层
      components: [校验, 公式, PPTX]
      note: 模板生成

- layout: experimentDesign
  title: 实验设计
  dataset:
    - 数据来源
  variables:
    - 自变量
  metrics:
    - 评价指标
  baselines:
    - 对照方法
  procedure: [准备数据, 运行实验, 分析结果]

- layout: resultAnalysis
  title: 结果分析
  finding: 一句话发现。
  metrics:
    - value: 18
      label: OMath 对象
      note: 公式可编辑
  analysis:
    - 分析一
    - 分析二

- layout: imageText
  title: 图片占位
  image:
    mode: placeholder
    aspectRatio: "16:9"
    placement: top
    prompt: 描述后续要替换的图片。
  text:
    - 说明一
    - 说明二

- layout: imageGrid
  title: 多图占位
  images:
    - mode: placeholder
      aspectRatio: "1:1"
      prompt: 输入截图
      caption: 输入
    - mode: placeholder
      aspectRatio: "1:1"
      prompt: 输出截图
      caption: 输出

- layout: code
  title: 伪代码
  language: Algorithm
  code: |
    Input: draft D
    1. parse D
    2. select layout
    3. export PPTX
  noteTitle: 关键约束
  notes:
    - 只输出结构化 YAML。

- layout: references
  title: 参考文献
  items:
    - Microsoft. Office Open XML File Formats.
    - PptxGenJS documentation.

YAML 与公式注意事项：
- 不要把 LaTeX 放进 YAML 双引号，除非把反斜杠写成双反斜杠。
- 推荐：latex: \mathcal{L}=...
- 推荐：'$p_\theta(x_i)$ 表示预测概率。'
- 表格中有公式时，建议用单引号包住单元格。
- 图片没有真实路径时使用 placeholder，不要写不存在的 assets 路径。
- 每页文字保持短；如果 check 返回 warnings，按 repairPrompt 压缩或拆页。</textarea>
      <textarea id="workflowGuide" hidden>请按这个流程和 AI 协作生成 PPT：

1. 先发“提示词”
- 告诉 AI：只输出 YAML，不要 Markdown 代码块，不要解释。
- 补充你的任务：主题、页数、受众、用途、是否需要演讲稿。

2. 再发“语法规则”
- 让 AI 使用支持的 layout 和字段。
- 告诉 AI：没有真实图片路径时必须用 image.mode: placeholder。
- 告诉 AI：公式用 LaTeX，生成后要能通过 check。

3. 最后发材料
- 可以贴论文摘要、章节结构、实验结果、表格数据、图片说明、参考文献。
- 要求 AI 生成完整 YAML。

4. 在网页里检查
- 把 YAML 粘到网页的 YAML 输入框，先点“检查”。
- 没有 errors 再点“生成 PPTX”。

5. 如果检查有问题
- 把 check 返回的 errors、warnings、repairPrompt 和当前 YAML 一起发给 AI。
- 要求 AI 只修改 YAML，不要解释，不要改变核心内容。</textarea>
      <textarea id="errorHelpPrompt" hidden>下面是 BIT PPT Generator 的检查/生成报错。请你只输出修复后的完整 YAML，不要 Markdown 代码块，不要解释。

修复要求：
1. 保留原始内容意图和页面顺序。
2. 修复未知 layout、字段结构错误、YAML 语法错误。
3. 如果 warnings 说明文字过长，请压缩文字或拆成多页。
4. 如果图片路径不存在，请改成 image.mode: placeholder，并补充具体 prompt。
5. 如果公式导致 YAML 转义问题，请避免双引号，优先使用 plain scalar 或单引号。
6. 输出必须能重新通过 check。

网页报错 / check 结果：

当前 YAML：
</textarea>
      <textarea id="exampleYaml" hidden>meta:
  title: 北理工风格汇报
  author: BIT PPT Generator
slides:
  - layout: title
    title: 北理工风格汇报
    subtitle: YAML 到可编辑 PPTX
  - layout: agenda
    title: 目录
    items:
      - 背景与问题
      - 方法设计
      - 实验结果
      - 总结展望
  - layout: bullets
    title: 背景与问题
    bullets:
      - 传统 PPT 制作成本高，格式一致性难以保证。
      - 目标是用 YAML 生成可编辑、可复用的 PPTX。
      - 生成结果保留文本、图表和公式的可编辑性。
  - layout: twoColumn
    title: 方案设计
    columns:
      - title: 输入
        bullets:
          - YAML 描述页面结构
          - 使用 layout 选择版式
          - 支持公式、图表和图片占位
      - title: 输出
        bullets:
          - 可编辑 PPTX
          - 北理工风格视觉规范
          - 适合答辩、组会和项目汇报
  - layout: formula
    title: 公式示例
    formula:
      latex: \mathcal{L}=\sum_i (y_i-f(x_i))^2
    caption: 均方误差目标函数
    notes:
      - 公式会转换为 Office Math，而不是截图。
  - layout: chart
    title: 图表示例
    type: bar
    categories: [方法A, 方法B, 本方案]
    series:
      - name: 效率
        values: [62, 74, 91]
    caption: 原生 PowerPoint 图表，可继续编辑数据。
  - layout: closing
    title: 谢谢
    subtitle: 欢迎批评指正</textarea>
      <textarea id="fullExampleYaml" hidden>meta:
  title: BIT PPT Generator 完整功能示例
  subtitle: 版式、公式、图表、备注与图片占位
  author: BIT PPT Generator
  date: 2026.05
slides:
  - layout: title
    title: BIT PPT Generator 完整功能示例
    subtitle: YAML 到可编辑 PPTX
    speakerNotes: |
      开场说明：这份示例用于展示主要能力。
      备注不会出现在幻灯片画布上，只会写入 PowerPoint/WPS 备注区。

  - layout: agenda
    title: 目录
    items:
      - 基础文本与行内公式
      - 表格、图表与流程图
      - 研究语义版式
      - 图片占位、代码与参考文献
    speakerNotes:
      - 这一页展示 speakerNotes 的数组写法。
      - 每一条会进入备注区，适合写演讲提示。

  - layout: bullets
    title: 要点页与行内公式
    lead: 普通正文中可以写 $\theta$ 和 $p_{\theta}(x)$。
    bullets:
      - 参数 $\theta$ 通过梯度下降更新。
      - 损失 $L_{train}$ 与误差 $E_{val}$ 可写正文。
      - 长列表会在 preflight 中被检查，必要时拆页。

  - layout: claim
    title: 单页结论
    claim: 当 $\mathcal{L}(\theta)$ 下降时，模型拟合能力通常增强。
    evidence:
      - 模板负责视觉规则和版式稳定性。
      - 生成器负责 OMML 公式、表格和图表写入。
      - AI 只需要输出结构化 YAML。

  - layout: twoColumn
    title: 两栏对比
    left:
      title: 传统做法
      bullets:
        - AI 直接写 PPT 代码。
        - 版式容易漂移。
        - 长文本容易溢出。
    right:
      title: 模板化做法
      bullets:
        - AI 输出 YAML。
        - 模板统一控制视觉。
        - 结果是可编辑 PPTX。

  - layout: cards
    title: 卡片页
    cards:
      - title: 内容结构
        text: 用 YAML 限定字段，减少异常结构。
      - title: 公式支持
        text: 行内公式 $p_{\theta}(x_i)$ 和展示公式都写为 OMML。
      - title: 自动检查
        text: check 接口返回 errors、warnings 和 repairPrompt。

  - layout: table
    title: 表格与单元格公式
    columns: [符号, 含义, 示例公式]
    rows:
      - ['$\theta$', 模型参数, '$p_\theta(x_i)$']
      - ['$x_i$', 第 i 个样本, '$\mathcal{L}_i=-y_i\log p_\theta(x_i)$']
      - ['$\eta$', 学习率, '$\theta_{t+1}=\theta_t-\eta g_t$']

  - layout: formula
    title: 展示公式
    formula:
      latex: \mathcal{L}(\theta) = -\frac{1}{n}\sum_{i=1}^{n}\sum_{k=1}^{K} y_{ik}\log p_{\theta,k}(x_i)
    caption: 多分类交叉熵目标函数
    explanation:
      - '$n$ 表示样本数量，$K$ 表示类别数量。'
      - '$p_{\theta,k}(x_i)$ 表示模型预测概率。'

  - layout: chart
    title: 原生柱状图
    type: bar
    categories: [Dataset A, Dataset B, Dataset C]
    valueAxisTitle: Accuracy
    series:
      - name: Baseline
        values: [81.2, 83.5, 84.1]
      - name: Ours
        values: [85.6, 87.2, 88.0]
    caption: 图表使用 PowerPoint 原生 chart XML。

  - layout: flowchart
    title: 可编辑流程图
    nodes:
      - id: input
        text: 输入材料
        note: 论文 / 数据 / 图片
      - id: plan
        text: 规划页面
        note: 选择 layout
      - id: render
        text: 生成 PPTX
        note: 形状 / 表格 / OMML
      - id: check
        text: 校验修复
        note: errors / warnings
        emphasis: true
    edges:
      - from: input
        to: plan
      - from: plan
        to: render
      - from: render
        to: check

  - layout: architecture
    title: 架构页
    layers:
      - title: 输入层
        components: [论文草稿, 实验结果, 图片素材]
        note: 将非结构化材料整理为内容块。
      - title: 规划层
        components: [章节规划, 页型选择, 字段填充]
        note: AI 只选择 layout 并填写 YAML。
      - title: 生成层
        components: [溢出检测, 公式 OMML, PPTX 写入]
        note: 模板侧保证字号、位置和颜色。

  - layout: experimentDesign
    title: 实验设计
    dataset:
      - 使用真实北理工 PPT 风格作为视觉参考。
      - 使用多种 AI 生成内容作为压力测试。
    variables:
      - 版式类型
      - 文本长度
      - 公式复杂度 $\mathcal{L}$
    metrics:
      - validation errors
      - preflight actions
      - PPTX 可编辑性
    baselines:
      - SVG 公式方案
      - 纯文本公式方案
    procedure: [生成 YAML, preflight, 导出 PPTX, WPS 打开验证]

  - layout: resultAnalysis
    title: 结果分析
    finding: OMML 方案解决了公式图片变形和行内公式割裂问题。
    metrics:
      - value: "0"
        label: SVG 公式
        note: 公式不再图片化
      - value: "18"
        label: OMath 对象
        note: 来自行内公式测试
      - value: "0"
        label: Marker 残留
        note: 后处理替换完成
    analysis:
      - 表格公式可以正常显示。
      - 公式容错可处理 $p_\theta$ 和 $x_i$ 等常见写法。

  - layout: imageText
    title: 图片占位
    image:
      mode: placeholder
      aspectRatio: "16:9"
      placement: top
      prompt: 一张展示 YAML 输入、结构校验、PPTX 生成和人工替换图片的流程图。
    text:
      - 用户暂时没有图片时，也可以先生成完整 PPT 草稿。
      - 后续可在 WPS 或 PowerPoint 中替换占位框。

  - layout: imageGrid
    title: 多图占位
    images:
      - mode: placeholder
        aspectRatio: "1:1"
        prompt: 输入样本截图
        caption: 输入
      - mode: placeholder
        aspectRatio: "1:1"
        prompt: 中间结果截图
        caption: 中间
      - mode: placeholder
        aspectRatio: "1:1"
        prompt: 输出 PPT 页面截图
        caption: 输出

  - layout: code
    title: 伪代码页
    language: Algorithm
    code: |
      Input: draft D, template layouts T
      1. parse D into claims, evidence and assets
      2. select layout t in T for each content block
      3. generate YAML fields with concise text
      4. run preflight and formula conversion
      5. export editable PPTX
    noteTitle: 关键约束
    notes:
      - 模型只产出结构化 YAML。
      - 公式使用 OMML 写入，避免图片变形。

  - layout: references
    title: 参考文献
    items:
      - Microsoft. Office Open XML File Formats specification.
      - PptxGenJS project documentation.
      - 北京理工大学学术答辩 PPT 模板视觉风格参考。
      - AI content generation workflow: planner, slide writer, validator and repair loop.

  - layout: closing
    title: 谢谢
    subtitle: 敬请批评指正
    speakerNotes: |
      收尾时提示听众：生成结果是可编辑 PPTX。
      后续可以在 PowerPoint 或 WPS 中继续调整。</textarea>
    </section>
    <section class="yaml-section">
      <h2>YAML</h2>
      <div class="yaml-workspace">
      <textarea id="yaml" spellcheck="false">meta:
  title: 北理工风格 PPT
  author: BIT PPT Generator
slides:
  - layout: title
    title: 北理工风格 PPT
    subtitle: YAML 到可编辑 PPTX
  - layout: bullets
    title: 示例页面
    bullets:
      - 使用北理工账号登录后生成
      - 输出为可编辑 PPTX
      - 支持公式、图表和多种版式
</textarea>
      <aside class="side-guide">
        <h3>使用教程</h3>
        <div class="guide">
          <div class="guide-step">
            <strong>1. 先复制提示词</strong>
            <p>发给 AI，说明它的角色、输出格式和页面限制。再补一句你的主题，例如“请做一份 10 页组会汇报”。</p>
          </div>
          <div class="guide-step">
            <strong>2. 再复制语法规则</strong>
            <p>继续发给 AI，让它按支持的 layout 和字段写 YAML。最后把论文摘要、实验结果、表格数据、图片说明等材料贴给 AI。</p>
          </div>
          <div class="guide-step">
            <strong>3. 粘贴 YAML 后先检查</strong>
            <p>把 AI 输出粘到左侧输入框，点“检查”。如果有 errors / warnings，把检查结果和当前 YAML 一起发回 AI，让它只修 YAML。</p>
          </div>
          <div class="guide-step">
            <strong>4. 生成失败时这样沟通</strong>
            <p>复制“报错求助”模板，把网页显示的错误、check 结果、当前 YAML 一起发给 AI；要求它保留内容意图，只修字段、长度、图片路径或公式写法。</p>
          </div>
        </div>
      </aside>
      </div>
    </section>
  </main>
  <script>
    const tokenKey = "bit_ppt_session_token";
    const expiresKey = "bit_ppt_session_expires";
    const $ = (id) => document.getElementById(id);
    const authStatus = $("authStatus");
    const runStatus = $("runStatus");
    const copyStatus = $("copyStatus");
    let authRequired = false;
    let signedAuthRequired = false;

    function token() { return localStorage.getItem(tokenKey) || ""; }
    function setStatus(el, text, cls = "") { el.className = "status " + cls; el.textContent = text; }
    function setBusy(busy) { ["loginBtn", "checkBtn", "generateBtn"].forEach((id) => { if ($(id)) $(id).disabled = busy; }); }
    function authHeaders(contentType) {
      const headers = { "content-type": contentType };
      const t = token();
      if (t) headers.authorization = "Bearer " + t;
      return headers;
    }
    function formatDuration(seconds) {
      const days = Math.floor(seconds / 86400);
      if (days >= 1) return days + " 天";
      const hours = Math.floor(seconds / 3600);
      if (hours >= 1) return hours + " 小时";
      return Math.ceil(seconds / 60) + " 分钟";
    }
    function refreshAuthStatus() {
      if (!authRequired) return;
      const t = token();
      const exp = Number(localStorage.getItem(expiresKey) || 0);
      if (!t || !exp) return setStatus(authStatus, "未登录");
      const left = Math.max(0, exp - Math.floor(Date.now() / 1000));
      setStatus(authStatus, left > 0 ? "已登录，剩余 " + formatDuration(left) : "登录已过期", left > 0 ? "ok" : "err");
    }

    async function loadServerConfig() {
      try {
        const res = await fetch("/health");
        const data = await res.json();
        authRequired = Boolean(data.authRequired);
        signedAuthRequired = Boolean(data.signedAuthRequired);
        $("authSection").classList.toggle("hidden", !signedAuthRequired);
        if (!authRequired) {
          localStorage.removeItem(tokenKey);
          localStorage.removeItem(expiresKey);
        } else if (signedAuthRequired) {
          refreshAuthStatus();
        } else {
          setStatus(runStatus, "此服务启用了固定 token 鉴权，网页登录不可用。", "err");
        }
      } catch {
        $("authSection").classList.remove("hidden");
        authRequired = true;
        signedAuthRequired = true;
        refreshAuthStatus();
      }
    }

    async function login() {
      setBusy(true);
      try {
        const res = await fetch("/auth/bit-login", {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({ username: $("username").value.trim(), password: $("password").value })
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.message || "登录失败");
        localStorage.setItem(tokenKey, data.token);
        localStorage.setItem(expiresKey, String(data.expiresAt));
        $("password").value = "";
        refreshAuthStatus();
      } catch (error) {
        setStatus(authStatus, error.message, "err");
      } finally {
        setBusy(false);
      }
    }

    async function checkDeck() {
      setBusy(true);
      try {
        const res = await fetch("/check", { method: "POST", headers: authHeaders("text/yaml"), body: $("yaml").value });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error || "检查失败");
        setStatus(runStatus, JSON.stringify({ errors: data.validation.errors.length, warnings: data.validation.warnings.length, actions: data.actions.length }, null, 2), data.validation.errors.length ? "err" : "ok");
      } catch (error) {
        setStatus(runStatus, error.message, "err");
      } finally {
        setBusy(false);
      }
    }

    async function generateDeck() {
      setBusy(true);
      try {
        const name = encodeURIComponent($("outputName").value || "bit-ppt");
        const res = await fetch("/generate?outputName=" + name, { method: "POST", headers: authHeaders("text/yaml"), body: $("yaml").value });
        if (!res.ok) {
          const data = await res.json().catch(() => ({}));
          throw new Error(data.error || "生成失败");
        }
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = ($("outputName").value || "bit-ppt") + ".pptx";
        a.click();
        URL.revokeObjectURL(url);
        setStatus(runStatus, "生成完成", "ok");
      } catch (error) {
        setStatus(runStatus, error.message, "err");
      } finally {
        setBusy(false);
      }
    }

    async function copyTextFrom(id, label) {
      const value = $(id).value;
      try {
        if (navigator.clipboard && window.isSecureContext) {
          await navigator.clipboard.writeText(value);
        } else {
          const field = document.createElement("textarea");
          field.value = value;
          field.setAttribute("readonly", "");
          field.style.position = "fixed";
          field.style.left = "-9999px";
          field.style.top = "0";
          document.body.appendChild(field);
          field.select();
          const copied = document.execCommand("copy");
          document.body.removeChild(field);
          if (!copied) throw new Error("copy command failed");
        }
        setStatus(copyStatus, label + "已复制", "ok");
      } catch {
        const field = $(id);
        field.focus();
        field.select();
        setStatus(copyStatus, "浏览器阻止自动复制，文本已选中，请按 Ctrl+C", "err");
      }
    }

    function insertExample() {
      $("yaml").value = $("exampleYaml").value;
      setStatus(runStatus, "已插入示例 YAML", "ok");
    }

    function insertFullExample() {
      $("yaml").value = $("fullExampleYaml").value;
      setStatus(runStatus, "已插入完整功能示例 YAML", "ok");
    }

    $("loginBtn").addEventListener("click", login);
    $("logoutBtn").addEventListener("click", () => { localStorage.removeItem(tokenKey); localStorage.removeItem(expiresKey); refreshAuthStatus(); });
    $("checkBtn").addEventListener("click", checkDeck);
    $("generateBtn").addEventListener("click", generateDeck);
    $("copyPromptBtn").addEventListener("click", () => copyTextFrom("aiPrompt", "提示词"));
    $("copyRulesBtn").addEventListener("click", () => copyTextFrom("syntaxRules", "语法规则"));
    $("copyWorkflowBtn").addEventListener("click", () => copyTextFrom("workflowGuide", "使用教程"));
    $("copyErrorHelpBtn").addEventListener("click", () => copyTextFrom("errorHelpPrompt", "报错求助模板"));
    $("insertExampleBtn").addEventListener("click", insertExample);
    $("insertFullExampleBtn").addEventListener("click", insertFullExample);
    loadServerConfig();
    setInterval(refreshAuthStatus, 30000);
  </script>
</body>
</html>`;

function parsePositiveInt(value, fallback) {
  const number = Number(value);
  return Number.isInteger(number) && number > 0 ? number : fallback;
}

function jsonResponse(res, status, payload) {
  const body = JSON.stringify(payload, null, 2);
  res.writeHead(status, {
    "content-type": "application/json; charset=utf-8",
    "content-length": Buffer.byteLength(body),
    "access-control-allow-origin": "*",
    "access-control-allow-methods": "GET,POST,OPTIONS",
    "access-control-allow-headers": "content-type, authorization",
  });
  res.end(body);
}

function textResponse(res, status, text) {
  res.writeHead(status, {
    "content-type": "text/plain; charset=utf-8",
    "content-length": Buffer.byteLength(text),
    "access-control-allow-origin": "*",
  });
  res.end(text);
}

function htmlResponse(res, status, html) {
  res.writeHead(status, {
    "content-type": "text/html; charset=utf-8",
    "content-length": Buffer.byteLength(html),
    "access-control-allow-origin": "*",
  });
  res.end(html);
}

function pptxResponse(res, buffer, fileName) {
  res.writeHead(200, {
    "content-type": PPTX_MIME,
    "content-length": buffer.length,
    "content-disposition": `attachment; filename="${fileName}"`,
    "access-control-allow-origin": "*",
  });
  res.end(buffer);
}

function getAuthToken(options = {}) {
  return options.authToken || process.env.BIT_PPT_TOKEN || "";
}

function getAuthSigningSecret(options = {}) {
  return options.authSigningSecret || process.env.BIT_PPT_AUTH_SECRET || "";
}

function base64UrlDecode(value) {
  const normalized = String(value || "").replaceAll("-", "+").replaceAll("_", "/");
  const padded = normalized + "=".repeat((4 - normalized.length % 4) % 4);
  return Buffer.from(padded, "base64");
}

function base64UrlEncode(buffer) {
  return Buffer.from(buffer).toString("base64").replaceAll("+", "-").replaceAll("/", "_").replaceAll("=", "");
}

function verifySignedAuthToken(token, secret) {
  if (!secret) return false;
  const parts = String(token || "").split(".");
  if (parts.length !== 3) return false;
  const [headerPart, payloadPart, signature] = parts;
  const expected = base64UrlEncode(crypto.createHmac("sha256", secret).update(`${headerPart}.${payloadPart}`).digest());
  const expectedBuffer = Buffer.from(expected);
  const signatureBuffer = Buffer.from(signature);
  if (expectedBuffer.length !== signatureBuffer.length || !crypto.timingSafeEqual(expectedBuffer, signatureBuffer)) return false;

  let payload;
  try {
    payload = JSON.parse(base64UrlDecode(payloadPart).toString("utf8"));
  } catch {
    return false;
  }
  const now = Math.floor(Date.now() / 1000);
  return payload?.aud === "bit-ppt" && typeof payload.sub === "string" && Number(payload.exp) > now;
}

function createSignedAuthToken(username, secret, ttlSeconds = DEFAULT_SESSION_TTL_SECONDS) {
  if (!secret) throw new Error("BIT_PPT_AUTH_SECRET is not configured.");
  const now = Math.floor(Date.now() / 1000);
  const header = base64UrlEncode(Buffer.from(JSON.stringify({ alg: "HS256", typ: "JWT" }), "utf8"));
  const payload = base64UrlEncode(Buffer.from(JSON.stringify({
    sub: String(username),
    aud: "bit-ppt",
    iat: now,
    exp: now + ttlSeconds,
  }), "utf8"));
  const input = `${header}.${payload}`;
  const signature = base64UrlEncode(crypto.createHmac("sha256", secret).update(input).digest());
  return { token: `${input}.${signature}`, expiresAt: now + ttlSeconds };
}

function requireAuth(req, options = {}) {
  const token = getAuthToken(options);
  const signingSecret = getAuthSigningSecret(options);
  if (!token && !signingSecret) return true;
  const authorization = req.headers.authorization || "";
  if (token && authorization === `Bearer ${token}`) return true;
  const bearer = authorization.match(/^Bearer\s+(.+)$/i)?.[1];
  if (bearer && verifySignedAuthToken(bearer, signingSecret)) return true;
  return false;
}

async function parseJsonBody(req, options = {}) {
  const rawBody = await readRequestBody(req, options.bodyLimit || parsePositiveInt(process.env.BIT_PPT_BODY_LIMIT, DEFAULT_BODY_LIMIT));
  return rawBody ? JSON.parse(rawBody) : {};
}

async function verifyBitPassword(username, password, options = {}) {
  if (options.bitAuthVerifier) return options.bitAuthVerifier(username, password);
  const response = await fetch(BIT_CAS_TICKETS_URL, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: `username=${encodeURIComponent(username)}&password=${encodeURIComponent(password)}`,
  });
  if (response.status === 201) return true;
  if (response.status === 400 || response.status === 401) return false;
  throw new Error(`BIT CAS unavailable: ${response.status}`);
}

function parseBool(value) {
  return value === true || value === "true" || value === "1" || value === "yes";
}

function sanitizeFileBase(value) {
  const text = String(value || "deck").replace(/[^\w.-]+/g, "-").replace(/^-+|-+$/g, "");
  return text || "deck";
}

async function readRequestBody(req, limit = DEFAULT_BODY_LIMIT) {
  const chunks = [];
  let total = 0;
  for await (const chunk of req) {
    total += chunk.length;
    if (total > limit) {
      const error = new Error(`Request body exceeds ${limit} bytes.`);
      error.statusCode = 413;
      throw error;
    }
    chunks.push(chunk);
  }
  return Buffer.concat(chunks).toString("utf8");
}

async function parseDeckRequest(req, url, options = {}) {
  const rawBody = await readRequestBody(req, options.bodyLimit || parsePositiveInt(process.env.BIT_PPT_BODY_LIMIT, DEFAULT_BODY_LIMIT));
  const contentType = req.headers["content-type"] || "";
  if (contentType.includes("application/json")) {
    const payload = rawBody ? JSON.parse(rawBody) : {};
    if (payload.deck && typeof payload.deck === "object") {
      return {
        deck: payload.deck,
        deckYaml: YAML.stringify(payload.deck),
        outputName: sanitizeFileBase(payload.outputName || payload.fileName || payload.deck?.meta?.title),
        options: {
          strict: parseBool(payload.strict) || parseBool(url.searchParams.get("strict")),
          fontCn: payload.fontCn,
          fontCnLight: payload.fontCnLight,
          fontEn: payload.fontEn,
          fontSerif: payload.fontSerif,
          fontCode: payload.fontCode,
        },
      };
    }
    const deckYaml = payload.deckYaml || payload.yaml;
    if (typeof deckYaml === "string") {
      return {
        deck: parseDeckYaml(deckYaml, "deckYaml"),
        deckYaml,
        outputName: sanitizeFileBase(payload.outputName || payload.fileName),
        options: {
          strict: parseBool(payload.strict) || parseBool(url.searchParams.get("strict")),
          fontCn: payload.fontCn,
          fontCnLight: payload.fontCnLight,
          fontEn: payload.fontEn,
          fontSerif: payload.fontSerif,
          fontCode: payload.fontCode,
        },
      };
    }
    const error = new Error("JSON body must include deck, deckYaml, or yaml.");
    error.statusCode = 400;
    throw error;
  }
  return {
    deck: parseDeckYaml(rawBody, "request body"),
    deckYaml: rawBody,
    outputName: sanitizeFileBase(url.searchParams.get("outputName") || url.searchParams.get("fileName")),
    options: {
      strict: parseBool(url.searchParams.get("strict")),
      fontCn: url.searchParams.get("fontCn") || undefined,
      fontCnLight: url.searchParams.get("fontCnLight") || undefined,
      fontEn: url.searchParams.get("fontEn") || undefined,
      fontSerif: url.searchParams.get("fontSerif") || undefined,
      fontCode: url.searchParams.get("fontCode") || undefined,
    },
  };
}

function healthPayload(options = {}) {
  const tokenAuthEnabled = Boolean(getAuthToken(options));
  const signedAuthEnabled = Boolean(getAuthSigningSecret(options));
  return {
    ok: true,
    service: "bit-ppt-http",
    capabilities: {
      check: true,
      generate: true,
      bitLogin: true,
      omml: true,
    },
    authRequired: tokenAuthEnabled || signedAuthEnabled,
    signedAuthRequired: signedAuthEnabled,
    maxGenerateConcurrency: parsePositiveInt(options.maxGenerateConcurrency || process.env.BIT_PPT_MAX_GENERATE_CONCURRENCY, DEFAULT_GENERATE_CONCURRENCY),
  };
}

async function handleCheck(req, res, url, options) {
  const { deck } = await parseDeckRequest(req, url, options);
  jsonResponse(res, 200, checkDeck(deck));
}

async function handleBitLogin(req, res, options) {
  const { username, password } = await parseJsonBody(req, options);
  if (!username || !password) {
    jsonResponse(res, 400, { code: 400, message: "Missing username or password." });
    return;
  }
  const ok = await verifyBitPassword(String(username), String(password), options);
  if (!ok) {
    jsonResponse(res, 401, { code: 401, message: "Invalid BIT username or password." });
    return;
  }
  const ttl = parsePositiveInt(options.sessionTtlSeconds || process.env.BIT_PPT_SESSION_TTL_SECONDS, DEFAULT_SESSION_TTL_SECONDS);
  const session = createSignedAuthToken(String(username), getAuthSigningSecret(options), ttl);
  jsonResponse(res, 200, {
    code: 200,
    token: session.token,
    tokenType: "Bearer",
    expiresAt: session.expiresAt,
  });
}

async function handleAuthVerify(req, res, options) {
  const authorization = req.headers.authorization || "";
  const bearer = authorization.match(/^Bearer\s+(.+)$/i)?.[1];
  const valid = Boolean(bearer && verifySignedAuthToken(bearer, getAuthSigningSecret(options)));
  jsonResponse(res, valid ? 200 : 401, { code: valid ? 200 : 401, valid });
}

async function handleGenerate(req, res, url, options) {
  const maxConcurrency = parsePositiveInt(options.maxGenerateConcurrency || process.env.BIT_PPT_MAX_GENERATE_CONCURRENCY, DEFAULT_GENERATE_CONCURRENCY);
  if (!Number.isInteger(options.activeGenerations)) options.activeGenerations = 0;
  if (options.activeGenerations >= maxConcurrency) {
    jsonResponse(res, 429, {
      generated: false,
      error: `Too many concurrent generation requests. Limit is ${maxConcurrency}.`,
    });
    return;
  }
  const parsed = await parseDeckRequest(req, url, options);
  const precheck = checkDeck(parsed.deck);
  if (precheck.validation.errors.length || (parsed.options.strict && precheck.validation.warnings.length)) {
    jsonResponse(res, 422, {
      generated: false,
      error: parsed.options.strict && precheck.validation.warnings.length ? "Deck validation failed strict mode." : "Deck validation failed.",
      validation: precheck.validation,
      repairPrompt: precheck.repairPrompt,
    });
    return;
  }

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "bit-ppt-http-"));
  options.activeGenerations += 1;
  try {
    const inputPath = path.join(tempDir, "deck.yaml");
    const outputPath = path.join(tempDir, `${parsed.outputName}.pptx`);
    fs.writeFileSync(inputPath, parsed.deckYaml, "utf8");
    const result = await generateDeckFile(inputPath, outputPath, parsed.options);
    const buffer = fs.readFileSync(result.output);
    pptxResponse(res, buffer, path.basename(result.output));
  } finally {
    options.activeGenerations -= 1;
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
}

async function handleRequest(req, res, options = {}) {
  const url = new URL(req.url || "/", "http://localhost");
  try {
    if (req.method === "OPTIONS") {
      jsonResponse(res, 200, { ok: true });
      return;
    }
    if (req.method === "GET" && url.pathname === "/health") {
      jsonResponse(res, 200, healthPayload(options));
      return;
    }
    if (req.method === "GET" && url.pathname === "/") {
      htmlResponse(res, 200, WEB_APP_HTML);
      return;
    }
    if (req.method === "POST" && url.pathname === "/auth/bit-login") {
      await handleBitLogin(req, res, options);
      return;
    }
    if (req.method === "POST" && url.pathname === "/auth/verify") {
      await handleAuthVerify(req, res, options);
      return;
    }
    if (["/check", "/generate"].includes(url.pathname) && !requireAuth(req, options)) {
      jsonResponse(res, 401, { error: "Unauthorized." });
      return;
    }
    if (req.method === "POST" && url.pathname === "/check") {
      await handleCheck(req, res, url, options);
      return;
    }
    if (req.method === "POST" && url.pathname === "/generate") {
      await handleGenerate(req, res, url, options);
      return;
    }
    if (["/check", "/generate"].includes(url.pathname)) {
      jsonResponse(res, 405, { error: "Method not allowed." });
      return;
    }
    textResponse(res, 404, "Not found. Use GET /, GET /health, POST /auth/bit-login, POST /check, or POST /generate.");
  } catch (error) {
    if (error.kind === "yaml_syntax") {
      jsonResponse(res, error.statusCode || 400, {
        error: error.message || "YAML syntax error.",
        syntax: error.syntax,
        repairPrompt: error.repairPrompt,
      });
      return;
    }
    if (error.validation) {
      jsonResponse(res, 422, {
        generated: false,
        error: error.message || "Deck validation failed.",
        validation: error.validation,
        repairPrompt: error.repairPrompt,
      });
      return;
    }
    jsonResponse(res, error.statusCode || 400, { error: error.message || String(error) });
  }
}

function createBitPptHttpServer(options = {}) {
  const state = {
    ...options,
    activeGenerations: 0,
  };
  return http.createServer((req, res) => {
    handleRequest(req, res, state);
  });
}

async function startHttpServer(options = {}) {
  const port = Number(options.port || process.env.PORT || 3000);
  const host = options.host || process.env.HOST || "127.0.0.1";
  const server = createBitPptHttpServer(options);
  await new Promise((resolve) => server.listen(port, host, resolve));
  return server;
}

async function main() {
  const portArgIndex = process.argv.findIndex((arg) => arg === "--port" || arg === "-p");
  const hostArgIndex = process.argv.findIndex((arg) => arg === "--host");
  const port = portArgIndex >= 0 ? process.argv[portArgIndex + 1] : process.env.PORT || 3000;
  const host = hostArgIndex >= 0 ? process.argv[hostArgIndex + 1] : process.env.HOST || "127.0.0.1";
  await startHttpServer({ port, host });
  console.log(`bit-ppt-http listening at http://${host}:${port}`);
  console.log("Endpoints: GET /health, POST /check, POST /generate");
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
}

export {
  PPTX_MIME,
  createBitPptHttpServer,
  handleRequest,
  startHttpServer,
};
