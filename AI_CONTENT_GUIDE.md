# AI Content Guide

Use this guide when asking an AI model to create a BIT-style PPT deck. The
model should produce YAML only. It should not write PowerPoint code.

## Required Shape

```yaml
meta:
  title: Presentation title
  subtitle: Optional subtitle
  author: Author or team
  advisor: Organization or advisor
  date: 2026.05
  fonts:
    cn: 微软雅黑
    en: Arial
    serif: SimSun
    code: Consolas

slides:
  - layout: title
  - layout: agenda
    title: 目录
    items: [Section 1, Section 2, Section 3]
```

`meta.fonts` is optional. On Windows, `微软雅黑` is the default Chinese font.
For cross-platform decks, use a commonly installed CJK font such as
`Noto Sans CJK SC` or `Source Han Sans SC`, or pass `--font-cn` in the CLI.

## Supported Layouts

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

## Layout Schemas

### `title`

Uses `meta` only.

```yaml
- layout: title
```

### `agenda`

```yaml
- layout: agenda
  title: 目录
  items:
    - 研究背景
    - 方法设计
    - 实验结果
```

### `section`

```yaml
- layout: section
  sectionNo: "01"
  title: 研究背景
  subtitle: 一句话说明本章节关注点
```

### `bullets`

Use for several peer-level findings. Keep bullets concise. A good bullet is
usually under 32 Chinese characters. The generator can split long bullet lists.

```yaml
- layout: bullets
  title: 核心发现
  lead: 一句话概括本页结论。
  bullets:
    - 第一个要点。
    - 第二个要点。
    - 第三个要点。
```

### `claim`

Use when the slide should make one strong conclusion and list supporting
evidence. Prefer one sentence in `claim`.

```yaml
- layout: claim
  title: 单页核心结论
  claim: 让模型生成结构化内容，让模板负责排版。
  evidence:
    - 证据一。
    - 证据二。
```

### `twoColumn`

Use for contrast, pros/cons, old/new, before/after, or two parallel arguments.

```yaml
- layout: twoColumn
  title: 两栏观点页
  left:
    title: 传统做法
    bullets:
      - 问题一。
      - 问题二。
  right:
    title: 推荐做法
    bullets:
      - 优势一。
      - 优势二。
```

### `cards`

Use for 3 to 6 independent ideas. Keep each card text short.

```yaml
- layout: cards
  title: 多观点卡片
  cards:
    - title: 观点一
      text: 一句话解释。
    - title: 观点二
      text: 一句话解释。
```

### `table`

Use no more than 5 columns. Long tables are automatically split by rows.

```yaml
- layout: table
  title: 对比表
  columns: [维度, 方案 A, 方案 B]
  rows:
    - [成本, 低, 中]
    - [风险, 中, 低]
```

### `comparison`

```yaml
- layout: comparison
  title: 方案对比
  left:
    label: 方案 A
    title: 传统路线
    bullets:
      - 优点或问题。
  right:
    label: 方案 B
    title: 推荐路线
    bullets:
      - 优点或问题。
```

### `timeline`

Use for roadmap, schedule, milestones, or historical development. Keep to 6
items or fewer.

```yaml
- layout: timeline
  title: 路线图
  items:
    - date: 阶段 1
      title: 复刻模板
      text: 固化视觉规范。
```

### `process`

Use for a linear workflow. Keep to 5 steps or fewer.

```yaml
- layout: process
  title: 工作流程
  steps:
    - title: 规划
      text: 选择页面类型。
    - title: 生成
      text: 输出 YAML。
```

### `architecture`

Use for method architecture, system architecture, model pipeline, or module
stack. Keep to 4 layers, each with no more than 5 components.

```yaml
- layout: architecture
  title: 方法架构
  layers:
    - title: 输入层
      components: [文本, 图像, 表格]
      note: 对输入进行结构化编码。
    - title: 模型层
      components: [Encoder, Fusion, Decoder]
      note: 支持 $p_{\theta}(x_i)$ 等行内公式。
```

### `ablation`

Use for ablation studies. Prefer short entries because the layout is table-like.

```yaml
- layout: ablation
  title: 消融实验
  baseline: 完整模型作为对照。
  items:
    - factor: 模块 A
      setting: 移除该模块
      delta: "-2.3%"
      conclusion: 性能下降，说明该模块有效。
```

### `caseStudy`

Use for one example, one sample, one qualitative analysis, or one system usage
case. Image paths are relative to this package directory.

```yaml
- layout: caseStudy
  title: 案例分析
  image: assets/bit-campus-photo.png
  caption: 案例图片说明。
  context:
    - 案例背景。
  method:
    - 处理方法。
  result:
    - 观察结果。
```

### `imageGrid`

Use for qualitative results, visual comparisons, or multiple screenshots. Keep
to 6 images or fewer.

```yaml
- layout: imageGrid
  title: 多图结果
  images:
    - path: assets/bit-campus-photo.png
      caption: 输入
    - path: assets/bit-campus-photo.png
      caption: 输出
```

### `code`

Use for pseudocode, algorithm steps, or a compact code fragment. Keep code under
about 12 short lines.

```yaml
- layout: code
  title: 算法伪代码
  language: Algorithm
  code: |
    Input: data D
    1. encode D
    2. optimize objective L(theta)
  notes:
    - 说明关键约束。
```

### `appendix`

Use for backup material indexes or appendix navigation. Keep to 8 items or
fewer.

```yaml
- layout: appendix
  title: 附录索引
  items:
    - key: A1
      title: 数据集细节
      text: 样本来源和统计信息。
```

### `flowchart`

Use for method flows, data pipelines, system workflows, or dependency diagrams.
The generator draws native PowerPoint shapes and arrow lines, so the result
remains editable. Keep to 10 nodes or fewer.

If `edges` is omitted, nodes are connected in order. Custom `x` and `y` are
allowed and are relative to the internal chart area, but most AI-generated decks
should rely on the automatic layout.

```yaml
- layout: flowchart
  title: 方法流程图
  nodes:
    - id: input
      text: 输入数据
      note: 文本 / 图像 / 表格
    - id: encode
      text: 编码器
      note: 得到 $h_i$
    - id: output
      text: 预测输出
      emphasis: true
  edges:
    - from: input
      to: encode
    - from: encode
      to: output
  note: 流程图使用 PPT 原生形状和箭头。
```

### `chart`

Use for editable statistical charts. Supported `type` values are `bar`, `line`,
`pie`, `doughnut`, `scatter`, and `area`; `bar` and `line` are the safest for
academic reports. Chart labels do not support OMML formulas, so keep math in the
caption or nearby text instead.

```yaml
- layout: chart
  title: 实验结果对比
  type: bar
  categories: [Dataset A, Dataset B, Dataset C]
  valueAxisTitle: Accuracy
  series:
    - name: Baseline
      values: [81.2, 83.5, 84.1]
    - name: Ours
      values: [85.6, 87.2, 88.0]
  caption: 原生 PowerPoint 图表，可继续编辑数据。
```

### Semantic Research Layouts

Use these when the slide has a clear academic role and you want the model to
write into tighter fields instead of a generic bullet page.

```yaml
- layout: problemSolution
  title: 问题与方案
  problem:
    title: 现有问题
    bullets: [问题一, 问题二]
  solution:
    title: 本文方案
    bullets: [方法一, 方法二]
  impact:
    title: 预期收益
    bullets: [收益一, 收益二]

- layout: painOpportunity
  title: 现状痛点与机会
  status:
    title: 现状
    bullets: [背景一, 背景二]
  pain:
    title: 痛点
    bullets: [痛点一, 痛点二]
  opportunity:
    title: 机会
    bullets: [机会一, 机会二]

- layout: experimentDesign
  title: 实验设计
  dataset: [数据集 A, 数据集 B]
  variables: [变量一, 变量二]
  metrics: [Accuracy, F1]
  baselines: [Baseline A, Baseline B]
  procedure: [训练, 验证, 测试]

- layout: resultAnalysis
  title: 结果分析
  finding: 本文方法在关键指标上取得稳定提升。
  metrics:
    - value: "+3.2%"
      label: F1
      note: 相比 baseline
  analysis:
    - 提升主要来自结构化规划模块。

- layout: riskMitigation
  title: 风险与对策
  items:
    - risk: 公式转换失败
      impact: 影响生成
      mitigation: 回传修复提示并保留原始 LaTeX。

- layout: contribution
  title: 主要贡献
  items:
    - title: 贡献一
      text: 简短描述。

- layout: summary
  title: 章节小结
  takeaway: 一句话总结本章。
  points:
    - 支撑要点一。
```

### `metrics`

Use for dashboard-like key numbers. Keep to 4 metrics or fewer.

```yaml
- layout: metrics
  title: 关键指标
  metrics:
    - value: 12+
      label: 正文版式
      note: 覆盖常见汇报场景
```

### `matrix`

Use for 2x2 analysis. Provide exactly 4 cells when possible.

```yaml
- layout: matrix
  title: 方案判断矩阵
  cells:
    - title: 低成本 / 低控制
      text: 简要说明。
    - title: 高成本 / 低控制
      text: 简要说明。
    - title: 低成本 / 高控制
      text: 简要说明。
    - title: 高成本 / 高控制
      text: 简要说明。
```

### `quote`

Use for a concise quote, principle, or memorable thesis.

```yaml
- layout: quote
  title: 引用页
  quote: 让模型生成结构化内容，让模板承担视觉和排版责任。
  source: BIT PPT Template Generator
```

### `formula`

Use for one display formula plus short explanation bullets. The formula is
converted to Office Math / OMML and written into the PPTX OpenXML. Use `$...$`
inside explanation text for small inline formulas.

Prefer wrapping subscript and superscript bodies in braces: write `p_{\theta}`
and `x^{(i)}`. The generator has a small tolerance pass that rewrites common
forms such as `p_\theta`, `x_i`, and `x^2`, but explicit braces are still safer.

```yaml
- layout: formula
  title: 目标函数
  formula:
    latex: "\\mathcal{L}(\\theta) = -\\sum_i y_i \\log p_{\\theta}(x_i)"
  caption: 交叉熵损失函数
  explanation:
    - "$\\theta$ 表示模型参数。"
    - "$p_{\\theta}(x_i)$ 表示预测概率。"
```

### `references`

Use for reference lists. Long lists are automatically split into multiple
slides. Quote strings containing a colon to avoid YAML parsing them as objects.

```yaml
- layout: references
  title: 参考文献
  items:
    - "Author. Title: subtitle. Publisher, 2026."
```

### `imageText`

Image paths are relative to this package directory.

```yaml
- layout: imageText
  title: 图文说明
  image: assets/bit-campus-photo.png
  text:
    - 说明一。
    - 说明二。
```

### `closing`

```yaml
- layout: closing
  title: 谢谢
  subtitle: 敬请各位老师批评指正
```

## Writing Rules

- Prefer one clear claim per slide.
- Do not put paragraph-length text into bullet slides.
- Use `claim` for a decisive conclusion and `bullets` for peer-level findings.
- Use `twoColumn` for contrast, `cards` for independent ideas, and `process`
  for ordered steps.
- Use `architecture` for system/module structure, `ablation` for ablation
  results, `caseStudy` for a single example, and `imageGrid` for qualitative
  comparisons.
- Use `flowchart` for editable process diagrams and `chart` for editable
  statistical charts.
- Split wide tables by columns manually; the generator only splits rows.
- Use `section` slides to separate major parts.
- Do not invent unsupported chart layouts; use `chart`, `table`, `metrics`, or
  `comparison` instead.
- Use `formula` for display equations. Keep one major equation per slide.

## Built-In Overflow Handling

The generator validates structure and estimates text wrapping before writing the
PPTX. Run:

```powershell
node src/generate.mjs input.yaml output.pptx --check
```

The check output includes `validation.errors`, `validation.warnings`, and
`repairPrompt`. Errors stop generation; warnings mean the deck can generate but
should be rewritten for a more stable layout.

Typical repair prompts look like:

```text
Slide 2 flowchart.nodes[3].text: Shorten flowchart.nodes[3].text to 12 characters or split the slide.
Slide 4 chart.series[0].values: Make every chart series have the same number of values as categories.
```

- `bullets`: long lists can be split into continuation slides.
- `table`: long tables can be split by row.
- `references`: long reference lists can be split into continuation slides.

The estimator is conservative because PowerPoint's final font rendering depends
on the local machine. Keep AI-generated text concise even when auto-splitting is
available.

## Inline Math

Inline `$...$` and `\(...\)` math is supported in normal text boxes for:

- `bullets.lead`
- `bullets[]`
- `claim.claim`
- `claim.evidence[]`
- `twoColumn.left/right.text`
- `twoColumn.left/right.bullets[]`
- `cards[].text`
- `comparison.left/right.bullets[]`
- `timeline.items[].text`
- `process.steps[].text`
- `architecture.layers[].components[]`
- `architecture.layers[].note`
- `ablation.baseline`
- `ablation.items[]`
- `caseStudy.context[]`
- `caseStudy.method[]`
- `caseStudy.result[]`
- `caseStudy.caption`
- `imageGrid.images[].caption`
- `code.notes[]`
- `appendix.items[].text`
- `flowchart.nodes[].text`
- `flowchart.nodes[].note`
- `flowchart.note`
- semantic research layouts such as `experimentDesign`, `resultAnalysis`,
  `riskMitigation`, `contribution`, and `summary`
- `metrics[].note`
- `matrix.cells[].text`
- `quote.quote`
- `imageText.text[]`
- `formula.explanation[]`
- `table.columns[]` and `table.rows[][]` are experimentally supported.

References currently treat math syntax as plain text.
