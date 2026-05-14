import { listLayouts as listSupportedLayouts } from "./core/layouts.mjs";
import { ICON_COLOR_TOKENS, ICON_SIZE_PRESETS, listIconsByCategory } from "./icons.mjs";

const writingRules = [
  "Keep slide text short and concrete; use validation warnings as rewrite hints.",
  "Prefer editable PPTX objects: text boxes, shapes, tables, charts, OMML formulas, and native images.",
  "Use `speakerNotes` for presenter scripts; these notes are written to the PPTX notes pane, not the slide canvas.",
  "When an image is unavailable, use `image.mode: placeholder` with a concrete prompt instead of inventing a local file path.",
  "Use local image paths relative to the project root unless an integration provides assets separately.",
  "Run `bit-ppt check <deck.yaml> --json` before generation and use `repairPrompt` to revise content.",
];

const speakerNotesGuide = {
  topic: "speaker-notes",
  purpose: "Add PowerPoint/WPS speaker notes to any slide without rendering them on the slide canvas.",
  fields: {
    speakerNotes: { type: "string | string[]", required: false },
  },
  aliases: ["speaker_notes", "speakerScript", "speaker_script"],
  limits: {
    speakerNotes: { recommendedChars: 1200, recommendedLines: 30 },
  },
  notes: [
    "Use a YAML block scalar for paragraph-style presenter scripts.",
    "Use a string list when the script is naturally segmented into short beats.",
    "Formula text in speaker notes is kept as plain text for now, for example `$L(\\theta)$`.",
    "`notes` remains layout-specific slide content in some layouts; use `speakerNotes` for the actual notes pane.",
  ],
  example: {
    layout: "bullets",
    title: "核心观点",
    bullets: ["输出是可编辑 PPTX。"],
    speakerNotes: "这一页先解释为什么选择 YAML 到 PPTX 的路线。\n公式暂时按普通文本保留，例如 $L(\\theta)$。",
  },
};

const imagePlaceholderGuide = {
  topic: "image-placeholder",
  purpose: "Describe missing images with editable placeholders that users can replace later.",
  whenToUse: "Use when the user has no image yet, cannot upload images, or wants AI to draft an image prompt inside the deck.",
  fields: {
    image: { type: "{ mode: placeholder, prompt, aspectRatio?, placement?, variants? }", required: true },
    prompt: { type: "string", required: true },
    aspectRatio: { type: "16:9 | 4:3 | 3:2 | 1:1 | 3:4 | 9:16 | auto", required: false },
    placement: { type: "top | side | auto", required: false },
    variants: { type: "boolean", required: false },
  },
  limits: {
    prompt: { recommendedChars: 160 },
  },
  notes: [
    "This is a common image object mode, not a standalone slide layout.",
    "Supported by imageText, caseStudy, and imageGrid.",
    "Use `aspectRatio` when the likely image shape is known.",
    "For imageText with unknown aspect ratio, leave it as auto so preflight emits top and side variants.",
    "The generated placeholder is editable PowerPoint text and shapes, not a raster image.",
  ],
  example: {
    layout: "imageText",
    title: "系统架构示意",
    image: {
      mode: "placeholder",
      aspectRatio: "auto",
      prompt: "A clean architecture diagram showing YAML input, validation, PPTX generation, and OMML post-processing.",
    },
    text: ["用户后续可替换占位图。"],
  },
};

const layoutGuides = {
  imageText: {
    layout: "imageText",
    purpose: "Show one image with a short explanatory bullet list.",
    whenToUse: "Use for visual evidence, screenshots, diagrams, and result images that need concise interpretation.",
    fields: {
      layout: { type: "literal", required: true, value: "imageText" },
      title: { type: "string", required: true },
      image: { type: "string | { path, placement, fit } | { mode: placeholder, prompt, aspectRatio?, placement? }", required: true },
      caption: { type: "string", required: false },
      text: { type: "string[]", required: true },
    },
    limits: {
      text: { maxItems: 5, recommendedChars: 38 },
      imagePrompt: { recommendedChars: 160 },
    },
    notes: [
      "Very wide images are automatically placed above the text.",
      "Ordinary landscape, portrait, and square images stay side-by-side so text space remains usable.",
      "`placement: top` or `placement: side` overrides automatic placement.",
      "`fit: contain` avoids cropping; `fit: cover` fills the image box.",
      "If no image is available, set `image.mode: placeholder` and provide a prompt for the future image.",
      "When placeholder aspectRatio is omitted or `auto`, imageText preflight creates top and side layout variants for user selection.",
    ],
    example: {
      layout: "imageText",
      title: "视觉结果说明",
      image: {
        path: "assets/bit-campus-photo.png",
        placement: "auto",
        fit: "contain",
      },
      caption: "图片说明控制在一行内。",
      text: ["图像展示关键现象。", "右侧或下方保留少量解释。"],
    },
  },
  chart: {
    layout: "chart",
    purpose: "Create an editable native PowerPoint chart.",
    whenToUse: "Use for quantitative comparisons, trends, composition, and metric summaries.",
    fields: {
      layout: { type: "literal", required: true, value: "chart" },
      title: { type: "string", required: true },
      type: { type: "bar | line | pie | doughnut | scatter | area", required: false },
      categories: { type: "string[]", required: true },
      series: { type: "{ name, values }[]", required: false },
      values: { type: "number[]", required: false },
      caption: { type: "string", required: false },
    },
    limits: {
      categories: { maxItems: 8, pieMaxItems: 6, recommendedChars: 18 },
      seriesName: { recommendedChars: 16 },
      caption: { recommendedChars: 58 },
    },
    notes: [
      "Every series values list must match the categories length.",
      "Use `values` as a shortcut for a single unnamed series.",
      "Pie and doughnut charts work best with six or fewer categories.",
    ],
    example: {
      layout: "chart",
      title: "指标对比",
      type: "bar",
      categories: ["Baseline", "Ours"],
      series: [{ name: "Accuracy", values: [82.1, 87.4] }],
      caption: "Ours improves accuracy under the same setting.",
    },
  },
  flowchart: {
    layout: "flowchart",
    purpose: "Draw an editable flowchart with PowerPoint shapes and arrow lines.",
    whenToUse: "Use for pipelines, decision paths, model stages, and process diagrams.",
    fields: {
      layout: { type: "literal", required: true, value: "flowchart" },
      title: { type: "string", required: true },
      nodes: { type: "{ id, text, note?, x?, y?, w?, h? }[]", required: true },
      edges: { type: "{ from, to }[]", required: false },
      note: { type: "string", required: false },
    },
    limits: {
      nodes: { maxItems: 10, recommendedTextChars: 12, recommendedNoteChars: 22 },
      note: { recommendedChars: 56 },
    },
    notes: [
      "Node ids must be unique.",
      "Edges must reference existing node ids.",
      "If edges are omitted, nodes are connected in order.",
      "Optional x/y coordinates are relative to the flowchart area.",
    ],
    example: {
      layout: "flowchart",
      title: "生成流程",
      nodes: [
        { id: "yaml", text: "YAML" },
        { id: "check", text: "校验" },
        { id: "pptx", text: "PPTX" },
      ],
      edges: [
        { from: "yaml", to: "check" },
        { from: "check", to: "pptx" },
      ],
    },
  },
  table: {
    layout: "table",
    purpose: "Create an editable comparison or data table.",
    whenToUse: "Use for compact structured comparisons with a small number of columns.",
    fields: {
      layout: { type: "literal", required: true, value: "table" },
      title: { type: "string", required: true },
      columns: { type: "string[]", required: true },
      rows: { type: "array[]", required: true },
    },
    limits: {
      columns: { maxItems: 5 },
      cell: { recommendedChars: 42 },
    },
    notes: [
      "Rows may be split across multiple slides during preflight if the table is too tall.",
      "Keep wide tables under five columns, or split the content by topic.",
      "Inline formulas in table cells are converted to native OMML.",
    ],
    example: {
      layout: "table",
      title: "方法对比",
      columns: ["方法", "输入", "输出"],
      rows: [
        ["Baseline", "文本", "普通 PPTX"],
        ["Ours", "YAML", "可编辑 BIT 风格 PPTX"],
      ],
    },
  },
  formula: {
    layout: "formula",
    purpose: "Show a display formula with short explanatory notes.",
    whenToUse: "Use for core equations, objective functions, and derivations that need native Office Math.",
    fields: {
      layout: { type: "literal", required: true, value: "formula" },
      title: { type: "string", required: true },
      formula: { type: "string | { latex, caption? }", required: true },
      caption: { type: "string", required: false },
      explanation: { type: "string[]", required: false },
    },
    limits: {
      explanation: { maxItems: 4, recommendedChars: 48 },
    },
    notes: [
      "LaTeX is converted to native OMML during PPTX post-processing.",
      "Inline `$...$` and `\\(...\\)` formulas are supported in text-heavy layouts.",
      "Avoid image formulas unless explicitly using a fallback outside this generator.",
    ],
    example: {
      layout: "formula",
      title: "优化目标",
      formula: {
        latex: "\\mathcal{L}=\\sum_i (y_i-\\hat{y}_i)^2",
      },
      explanation: ["目标函数衡量预测误差。"],
    },
  },
};

const additionalLayoutGuides = {
  title: {
    layout: "title",
    purpose: "Create the deck cover slide from top-level metadata.",
    whenToUse: "Use as the first slide; title content usually comes from the deck-level `meta` object.",
    fields: {
      layout: { type: "literal", required: true, value: "title" },
      "meta.title": { type: "string", required: false },
      "meta.subtitle": { type: "string", required: false },
      "meta.author": { type: "string", required: false },
      "meta.advisor": { type: "string", required: false },
      "meta.date": { type: "string", required: false },
    },
    limits: {
      "meta.title": { recommendedChars: 28 },
      "meta.subtitle": { recommendedChars: 42 },
    },
    notes: [
      "The slide object normally only needs `layout: title`.",
      "Cover text is read from deck metadata.",
    ],
    example: { layout: "title" },
  },
  agenda: {
    layout: "agenda",
    purpose: "Show the main agenda or chapter list.",
    whenToUse: "Use near the beginning of a deck to preview sections.",
    fields: {
      layout: { type: "literal", required: true, value: "agenda" },
      title: { type: "string", required: true },
      items: { type: "string[]", required: true },
    },
    limits: { items: { maxItems: 7, recommendedChars: 24 } },
    notes: ["Keep agenda labels short and parallel.", "Use `section` slides for major chapter breaks."],
    example: { layout: "agenda", title: "目录", items: ["研究背景", "方法设计", "实验结果", "总结展望"] },
  },
  section: {
    layout: "section",
    purpose: "Create a visual chapter divider.",
    whenToUse: "Use before a major topic shift or report section.",
    fields: {
      layout: { type: "literal", required: true, value: "section" },
      sectionNo: { type: "string", required: false },
      title: { type: "string", required: true },
      subtitle: { type: "string", required: false },
    },
    limits: { title: { recommendedChars: 24 }, subtitle: { recommendedChars: 42 } },
    notes: ["`sectionNo` can be supplied explicitly, for example `01`.", "Keep the subtitle as context, not body content."],
    example: { layout: "section", sectionNo: "01", title: "研究背景", subtitle: "从问题定义与现有瓶颈切入" },
  },
  bullets: {
    layout: "bullets",
    purpose: "Show a lead sentence and a concise bullet list.",
    whenToUse: "Use for findings, observations, requirements, and short argument chains.",
    fields: {
      layout: { type: "literal", required: true, value: "bullets" },
      title: { type: "string", required: true },
      lead: { type: "string", required: false },
      bullets: { type: "(string | { icon, iconColor?, text })[]", required: true },
      fontSize: { type: "number", required: false },
    },
    limits: { lead: { recommendedChars: 58 }, bullets: { maxItems: 8, recommendedChars: 38 } },
    notes: [
      "Inline `$...$` and `\\(...\\)` formulas are supported.",
      "Long bullet lists may be split during preflight.",
      "A bullet item can be a string, or `{ icon, iconColor?, text }` to replace the dot with an icon (see `bit-ppt guide icons`).",
    ],
    example: { layout: "bullets", title: "核心发现", lead: "结构化输入能降低排版漂移。", bullets: [{ icon: "checkmark", text: "模型只填写短字段。" }, { icon: "settings-gear", text: "模板负责视觉规范。" }, { icon: "magnifier", text: "预检负责溢出修复。" }] },
  },
  claim: {
    layout: "claim",
    purpose: "Emphasize one central claim with supporting evidence.",
    whenToUse: "Use when a slide needs one clear takeaway rather than several equal points.",
    fields: {
      layout: { type: "literal", required: true, value: "claim" },
      title: { type: "string", required: true },
      claim: { type: "string", required: true },
      evidence: { type: "string[]", required: false },
    },
    limits: { claim: { recommendedChars: 54 }, evidence: { maxItems: 5, recommendedChars: 38 } },
    notes: ["Make `claim` a complete sentence.", "Use evidence bullets for proof, not extra claims."],
    example: { layout: "claim", title: "核心结论", claim: "AI 生成 PPT 的关键是稳定内容结构。", evidence: ["YAML 限定字段。", "PPTX 对象保持可编辑。"] },
  },
  twoColumn: {
    layout: "twoColumn",
    purpose: "Compare or separate two parallel content blocks.",
    whenToUse: "Use for before/after, baseline/ours, or two-topic explanation.",
    fields: {
      layout: { type: "literal", required: true, value: "twoColumn" },
      title: { type: "string", required: true },
      left: { type: "{ title, text? | bullets? }", required: true },
      right: { type: "{ title, text? | bullets? }", required: true },
    },
    limits: { columnTitle: { recommendedChars: 16 }, columnText: { recommendedChars: 105 }, columnBullets: { maxItems: 5, recommendedChars: 34 } },
    notes: ["Each column can use either `text` or `bullets`.", "Use `comparison` when the point is explicitly adversarial or directional."],
    example: { layout: "twoColumn", title: "两种路线", left: { title: "传统做法", bullets: ["手工排版", "一致性依赖人工"] }, right: { title: "模板化做法", bullets: ["结构化输入", "生成器固定样式"] } },
  },
  cards: {
    layout: "cards",
    purpose: "Show several peer points as compact cards.",
    whenToUse: "Use for features, modules, contributions, or categories with equal weight.",
    fields: {
      layout: { type: "literal", required: true, value: "cards" },
      title: { type: "string", required: true },
      cards: { type: "{ title, text }[]", required: true },
    },
    limits: { cards: { maxItems: 6 }, cardTitle: { recommendedChars: 14 }, cardText: { recommendedChars: 52 } },
    notes: ["Cards work best when each item has similar importance.", "Use no more than six cards on one slide."],
    example: { layout: "cards", title: "能力模块", cards: [{ title: "校验", text: "检查字段与长度风险。" }, { title: "生成", text: "写入可编辑 PPTX 对象。" }] },
  },
  comparison: {
    layout: "comparison",
    purpose: "Show a direct comparison between two options.",
    whenToUse: "Use when one side is preferred, rejected, or contrasted against another.",
    fields: {
      layout: { type: "literal", required: true, value: "comparison" },
      title: { type: "string", required: true },
      left: { type: "{ label?, title, bullets }", required: true },
      right: { type: "{ label?, title, bullets }", required: true },
    },
    limits: { comparisonTitle: { recommendedChars: 18 }, bullets: { maxItems: 5, recommendedChars: 32 } },
    notes: ["`left.label` and `right.label` are short tags.", "Keep the two sides structurally parallel."],
    example: { layout: "comparison", title: "技术路线取舍", left: { label: "不优先", title: "端到端 Agent", bullets: ["流程封闭", "复核成本高"] }, right: { label: "优先", title: "MCP 工具链", bullets: ["跨客户端", "人类可介入"] } },
  },
  timeline: {
    layout: "timeline",
    purpose: "Show milestones along a horizontal timeline.",
    whenToUse: "Use for roadmaps, phases, history, or scheduled work.",
    fields: {
      layout: { type: "literal", required: true, value: "timeline" },
      title: { type: "string", required: true },
      items: { type: "{ date? | phase?, title, text }[]", required: true },
    },
    limits: { items: { maxItems: 6 }, itemTitle: { recommendedChars: 10 }, itemText: { recommendedChars: 24 } },
    notes: ["Use `date` for calendar time and `phase` for abstract stages.", "Avoid paragraph text in timeline nodes."],
    example: { layout: "timeline", title: "路线图", items: [{ date: "阶段 1", title: "模板", text: "固化视觉规范。" }, { date: "阶段 2", title: "校验", text: "加入预检修复。" }] },
  },
  process: {
    layout: "process",
    purpose: "Show a linear process as connected steps.",
    whenToUse: "Use for workflows, generation pipelines, or execution procedures.",
    fields: {
      layout: { type: "literal", required: true, value: "process" },
      title: { type: "string", required: true },
      steps: { type: "{ title, text }[]", required: true },
    },
    limits: { steps: { maxItems: 5 }, stepTitle: { recommendedChars: 8 }, stepText: { recommendedChars: 24 } },
    notes: ["Steps are rendered left-to-right.", "Use `flowchart` for branching or non-linear processes."],
    example: { layout: "process", title: "生成流程", steps: [{ title: "解析", text: "读取 YAML。" }, { title: "校验", text: "检查字段。" }, { title: "导出", text: "生成 PPTX。" }] },
  },
  problemSolution: {
    layout: "problemSolution",
    purpose: "Present problem, solution, and impact in three panels.",
    whenToUse: "Use for proposal framing or product/research motivation.",
    fields: {
      layout: { type: "literal", required: true, value: "problemSolution" },
      title: { type: "string", required: true },
      problem: { type: "{ label?, title, bullets }", required: true },
      solution: { type: "{ label?, title, bullets }", required: true },
      impact: { type: "{ label?, title, bullets }", required: true },
    },
    limits: { panelTitle: { recommendedChars: 12 }, bullets: { maxItems: 4, recommendedChars: 26 } },
    notes: ["Use this when the three blocks form one argument.", "Each panel supports a custom short `label`."],
    example: { layout: "problemSolution", title: "问题与方案", problem: { title: "排版不稳", bullets: ["模型自由发挥导致溢出。"] }, solution: { title: "结构约束", bullets: ["YAML 固定字段。"] }, impact: { title: "可复核", bullets: ["输出可编辑 PPTX。"] } },
  },
  painOpportunity: {
    layout: "painOpportunity",
    purpose: "Frame current status, pain points, and opportunity.",
    whenToUse: "Use for background analysis before introducing a solution.",
    fields: {
      layout: { type: "literal", required: true, value: "painOpportunity" },
      title: { type: "string", required: true },
      status: { type: "{ title, text? | bullets? }", required: true },
      pain: { type: "{ title, text? | bullets? }", required: true },
      opportunity: { type: "{ title, text? | bullets? }", required: true },
    },
    limits: { panelTitle: { recommendedChars: 12 }, bullets: { maxItems: 4, recommendedChars: 30 } },
    notes: ["Panels use the same structure as `twoColumn` column blocks.", "Best for moving from observation to opportunity."],
    example: { layout: "painOpportunity", title: "现状痛点与机会", status: { title: "现状", bullets: ["PPT 生成需求高。"] }, pain: { title: "痛点", bullets: ["直接生成不稳定。"] }, opportunity: { title: "机会", bullets: ["模板化生成可控。"] } },
  },
  experimentDesign: {
    layout: "experimentDesign",
    purpose: "Summarize an experiment setup.",
    whenToUse: "Use before result slides to define data, variables, metrics, and baselines.",
    fields: {
      layout: { type: "literal", required: true, value: "experimentDesign" },
      title: { type: "string", required: true },
      dataset: { type: "string | string[]", required: false },
      variables: { type: "string | string[]", required: false },
      metrics: { type: "string | string[]", required: false },
      baselines: { type: "string | string[]", required: false },
      procedure: { type: "string[]", required: false },
    },
    limits: { eachBlock: { maxItems: 4, recommendedChars: 26 }, procedure: { maxItems: 5 } },
    notes: ["Scalar strings are accepted and rendered as one bullet.", "Keep procedure items short because they render on one line."],
    example: { layout: "experimentDesign", title: "实验设计", dataset: ["公开数据集 A"], variables: ["是否启用预检"], metrics: ["溢出页数"], baselines: ["直接生成 PPT"], procedure: ["准备输入", "生成", "人工复核"] },
  },
  resultAnalysis: {
    layout: "resultAnalysis",
    purpose: "State one finding, key metrics, and analysis bullets.",
    whenToUse: "Use after experiments to connect numbers to interpretation.",
    fields: {
      layout: { type: "literal", required: true, value: "resultAnalysis" },
      title: { type: "string", required: true },
      finding: { type: "string", required: true },
      metrics: { type: "{ value, label, note? }[]", required: false },
      analysis: { type: "string[]", required: false },
    },
    limits: { finding: { recommendedChars: 52 }, metrics: { maxItems: 3 }, analysis: { maxItems: 4, recommendedChars: 38 } },
    notes: ["Put the strongest result in `finding`.", "Metrics are compact; use `chart` for richer data."],
    example: { layout: "resultAnalysis", title: "结果分析", finding: "预检显著减少长列表造成的页面溢出。", metrics: [{ value: "-80%", label: "溢出页", note: "相对基线" }], analysis: ["主要收益来自 bullets 和 references 拆页。"] },
  },
  riskMitigation: {
    layout: "riskMitigation",
    purpose: "Show risks, impacts, and mitigations in a table.",
    whenToUse: "Use for project planning, deployment risk, or limitation handling.",
    fields: {
      layout: { type: "literal", required: true, value: "riskMitigation" },
      title: { type: "string", required: true },
      items: { type: "{ risk, impact, mitigation }[]", required: true },
    },
    limits: { items: { maxItems: 5 }, cell: { recommendedChars: 28 } },
    notes: ["Keep mitigation actionable.", "Use this for risks; use `table` for arbitrary structured data."],
    example: { layout: "riskMitigation", title: "风险与对策", items: [{ risk: "图片缺失", impact: "生成失败", mitigation: "使用 placeholder 并补 prompt" }] },
  },
  contribution: {
    layout: "contribution",
    purpose: "List the main contributions with numbered emphasis.",
    whenToUse: "Use near the end of a research or project deck.",
    fields: {
      layout: { type: "literal", required: true, value: "contribution" },
      title: { type: "string", required: true },
      items: { type: "{ title, text }[]", required: true },
    },
    limits: { items: { maxItems: 4 }, itemTitle: { recommendedChars: 14 }, itemText: { recommendedChars: 44 } },
    notes: ["Use one contribution per item.", "Avoid turning this into a general summary slide."],
    example: { layout: "contribution", title: "主要贡献", items: [{ title: "可编辑输出", text: "文本、表格、图表均保留为 PPTX 对象。" }, { title: "公式支持", text: "LaTeX 公式转换为原生 OMML。" }] },
  },
  summary: {
    layout: "summary",
    purpose: "Close a section with one takeaway and supporting points.",
    whenToUse: "Use at section endings or before the closing slide.",
    fields: {
      layout: { type: "literal", required: true, value: "summary" },
      title: { type: "string", required: true },
      takeaway: { type: "string", required: true },
      points: { type: "string[]", required: false },
    },
    limits: { takeaway: { recommendedChars: 52 }, points: { maxItems: 5, recommendedChars: 36 } },
    notes: ["`takeaway` should be a sentence, not a heading.", "Use `closing` only for the final thank-you page."],
    example: { layout: "summary", title: "章节小结", takeaway: "结构化内容让 PPT 生成更稳定也更容易修复。", points: ["页面类型明确。", "校验信息可回传模型。"] },
  },
  architecture: {
    layout: "architecture",
    purpose: "Show a layered architecture with components and notes.",
    whenToUse: "Use for systems, model stacks, or pipeline architecture.",
    fields: {
      layout: { type: "literal", required: true, value: "architecture" },
      title: { type: "string", required: true },
      layers: { type: "{ title, components, note? }[]", required: true },
    },
    limits: { layers: { maxItems: 4 }, layerTitle: { recommendedChars: 8 }, components: { maxItems: 5, recommendedChars: 10 }, note: { recommendedChars: 44 } },
    notes: ["Layer order is top-to-bottom.", "Use short component labels so chips stay readable."],
    example: { layout: "architecture", title: "系统架构", layers: [{ title: "输入层", components: ["YAML", "图片"], note: "收集结构化材料。" }, { title: "生成层", components: ["校验", "PPTX", "OMML"], note: "写入可编辑对象。" }] },
  },
  ablation: {
    layout: "ablation",
    purpose: "Summarize ablation factors, settings, deltas, and conclusions.",
    whenToUse: "Use for research experiments that remove or vary components.",
    fields: {
      layout: { type: "literal", required: true, value: "ablation" },
      title: { type: "string", required: true },
      baseline: { type: "string", required: false },
      items: { type: "{ factor, setting, delta, conclusion }[]", required: true },
    },
    limits: { baseline: { recommendedChars: 58 }, items: { maxItems: 6 }, factor: { recommendedChars: 18 }, setting: { recommendedChars: 18 }, delta: { recommendedChars: 18 }, conclusion: { recommendedChars: 28 } },
    notes: ["Use `delta` for numeric or qualitative change.", "Keep conclusions short enough for table-like rendering."],
    example: { layout: "ablation", title: "消融实验", baseline: "基线启用所有模块。", items: [{ factor: "预检", setting: "关闭", delta: "+3 overflow", conclusion: "长列表风险上升" }] },
  },
  caseStudy: {
    layout: "caseStudy",
    purpose: "Show one visual case with context, method, and result.",
    whenToUse: "Use for qualitative examples, screenshots, or representative samples.",
    fields: {
      layout: { type: "literal", required: true, value: "caseStudy" },
      title: { type: "string", required: true },
      image: { type: "string | { path, fit? } | { mode: placeholder, prompt, aspectRatio? }", required: true },
      caption: { type: "string", required: false },
      context: { type: "string[]", required: false },
      method: { type: "string[]", required: false },
      result: { type: "string[]", required: false },
    },
    limits: { caption: { recommendedChars: 40 }, context: { maxItems: 2, recommendedChars: 34 }, method: { maxItems: 2, recommendedChars: 34 }, result: { maxItems: 2, recommendedChars: 34 } },
    notes: ["Image placeholders are supported.", "Use `imageGrid` when comparing multiple images."],
    example: { layout: "caseStudy", title: "案例分析", image: "assets/bit-campus-photo.png", caption: "示例图片，可替换为实验结果图。", context: ["输入是一段论文草稿。"], method: ["生成器按 layout 写入 PPTX。"], result: ["输出可继续编辑。"] },
  },
  imageGrid: {
    layout: "imageGrid",
    purpose: "Show multiple images in a regular grid.",
    whenToUse: "Use for visual result sets, comparisons, or qualitative examples.",
    fields: {
      layout: { type: "literal", required: true, value: "imageGrid" },
      title: { type: "string", required: true },
      images: { type: "(string | { path, caption? } | { mode: placeholder, prompt, caption?, aspectRatio? })[]", required: true },
    },
    limits: { images: { maxItems: 6 }, caption: { recommendedChars: 16 } },
    notes: ["Each image can be a local path or an image object.", "Use short captions to avoid crowding the grid."],
    example: { layout: "imageGrid", title: "多图结果", images: [{ path: "assets/bit-campus-photo.png", caption: "输入" }, { path: "assets/bit-campus-photo.png", caption: "输出" }] },
  },
  code: {
    layout: "code",
    purpose: "Show pseudocode or a compact code block with notes.",
    whenToUse: "Use for algorithms, command snippets, or implementation sketches.",
    fields: {
      layout: { type: "literal", required: true, value: "code" },
      title: { type: "string", required: true },
      language: { type: "string", required: false },
      code: { type: "string", required: false },
      algorithm: { type: "string", required: false },
      noteTitle: { type: "string", required: false },
      notes: { type: "string[]", required: false },
    },
    limits: { code: { recommendedLines: 12, recommendedCharsPerLine: 72 }, notes: { maxItems: 5, recommendedChars: 32 } },
    notes: ["Use either `code` or `algorithm`; `code` takes precedence.", "A YAML block scalar is recommended for multi-line code."],
    example: { layout: "code", title: "算法伪代码", language: "Algorithm", code: "Input: deck D\n1. validate D\n2. generate PPTX", noteTitle: "关键约束", notes: ["代码块保持短行。"] },
  },
  appendix: {
    layout: "appendix",
    purpose: "Create an appendix index or supplemental topic list.",
    whenToUse: "Use for backup slides, extra details, or appendix navigation.",
    fields: {
      layout: { type: "literal", required: true, value: "appendix" },
      title: { type: "string", required: true },
      items: { type: "{ key?, title, text }[]", required: true },
    },
    limits: { items: { maxItems: 8 }, itemTitle: { recommendedChars: 14 }, itemText: { recommendedChars: 42 } },
    notes: ["`key` is optional and defaults to a numbered label.", "Use this as an index, not as dense body text."],
    example: { layout: "appendix", title: "附录索引", items: [{ key: "A1", title: "数据细节", text: "样本来源与筛选规则。" }] },
  },
  metrics: {
    layout: "metrics",
    purpose: "Highlight up to four key metrics.",
    whenToUse: "Use when numeric or compact KPI-style facts should dominate the slide.",
    fields: {
      layout: { type: "literal", required: true, value: "metrics" },
      title: { type: "string", required: true },
      metrics: { type: "{ value, label, note? }[]", required: true },
    },
    limits: { metrics: { maxItems: 4 }, value: { recommendedChars: 8 }, label: { recommendedChars: 12 }, note: { recommendedChars: 30 } },
    notes: ["Use compact values such as `12+`, `-8%`, or `0`.", "Use `chart` when the reader needs to inspect a trend or distribution."],
    example: { layout: "metrics", title: "关键指标", metrics: [{ value: "12+", label: "正文页型", note: "覆盖常见汇报页面" }, { value: "100%", label: "可编辑", note: "原生 PPTX 对象" }] },
  },
  matrix: {
    layout: "matrix",
    purpose: "Show four quadrants or a compact 2x2 decision matrix.",
    whenToUse: "Use for strategy tradeoffs, option classification, or two-axis judgment.",
    fields: {
      layout: { type: "literal", required: true, value: "matrix" },
      title: { type: "string", required: true },
      cells: { type: "{ title, text }[]", required: true },
    },
    limits: { cells: { maxItems: 4 }, cellTitle: { recommendedChars: 18 }, cellText: { recommendedChars: 52 } },
    notes: ["Cells render row-major in a 2x2 grid.", "Provide exactly four cells for a complete matrix."],
    example: { layout: "matrix", title: "方案判断矩阵", cells: [{ title: "低成本 / 高控制", text: "YAML 内容加模板生成。" }, { title: "高成本 / 高控制", text: "完整设计系统与截图校验。" }] },
  },
  quote: {
    layout: "quote",
    purpose: "Show a short quote or memorable statement.",
    whenToUse: "Use for thesis statements, external quotes, or section openers.",
    fields: {
      layout: { type: "literal", required: true, value: "quote" },
      title: { type: "string", required: false },
      quote: { type: "string", required: true },
      source: { type: "string", required: false },
    },
    limits: { quote: { recommendedChars: 70 } },
    notes: ["Keep the quote short enough to remain visually dominant.", "`source` is optional and rendered below the quote."],
    example: { layout: "quote", title: "引用页", quote: "让模型生成结构化内容，让模板承担视觉和排版责任。", source: "BIT PPT Template Generator" },
  },
  references: {
    layout: "references",
    purpose: "List references or citations.",
    whenToUse: "Use near the end of academic or research decks.",
    fields: {
      layout: { type: "literal", required: true, value: "references" },
      title: { type: "string", required: false },
      items: { type: "string[]", required: true },
      fontSize: { type: "number", required: false },
    },
    limits: { items: { recommendedChars: 140 } },
    notes: ["Long reference lists may be split during preflight.", "Keep each reference as one string item."],
    example: { layout: "references", title: "参考文献", items: ["PptxGenJS project documentation. Generate editable PowerPoint presentations with JavaScript."] },
  },
  closing: {
    layout: "closing",
    purpose: "Create the final thank-you slide.",
    whenToUse: "Use as the final slide of a deck.",
    fields: {
      layout: { type: "literal", required: true, value: "closing" },
      title: { type: "string", required: false },
      subtitle: { type: "string", required: false },
    },
    limits: { title: { recommendedChars: 12 }, subtitle: { recommendedChars: 28 } },
    notes: ["Defaults to a BIT green closing slide.", "Use `subtitle` for defense/Q&A wording."],
    example: { layout: "closing", title: "谢谢", subtitle: "敬请各位老师批评指正" },
  },
};

Object.assign(layoutGuides, additionalLayoutGuides);

function listGuideLayouts() {
  return listSupportedLayouts();
}

function getLayoutGuide(layout) {
  return layoutGuides[layout] || null;
}

function getLayoutSchema(layout) {
  const guide = getLayoutGuide(layout);
  if (!guide) return null;
  return {
    layout: guide.layout,
    commonFields: speakerNotesGuide.fields,
    fields: guide.fields,
    limits: guide.limits,
  };
}

function getLayoutExample(layout) {
  const guide = getLayoutGuide(layout);
  return guide?.example || null;
}

function getGuideOverview() {
  return {
    name: "BIT PPT Generator",
    purpose: "Generate editable Beijing Institute of Technology style PPTX files from YAML.",
    workflow: [
      "Choose layouts with `bit-ppt list-layouts`.",
      "Inspect one layout with `bit-ppt guide layout <name>`.",
      "Draft YAML and run `bit-ppt check <deck.yaml> --json`.",
      "Fix validation errors using `repairPrompt`.",
      "Generate with `bit-ppt generate <deck.yaml> <output.pptx>`.",
    ],
    commands: [
      "bit-ppt guide",
      "bit-ppt guide layouts",
      "bit-ppt guide layout <name>",
      "bit-ppt guide schema <name> --json",
      "bit-ppt guide example <name>",
      "bit-ppt guide speaker-notes",
      "bit-ppt guide image-placeholder",
      "bit-ppt guide icons",
      "bit-ppt guide writing-rules",
    ],
    guidedLayouts: listGuideLayouts(),
  };
}

function getSpeakerNotesGuide() {
  return { ...speakerNotesGuide };
}

function getImagePlaceholderGuide() {
  return { ...imagePlaceholderGuide };
}

function getWritingRules() {
  return [...writingRules];
}

function getGuideWorkflow() {
  return getGuideOverview().workflow;
}

function getIconsGuide() {
  const byCategory = listIconsByCategory();
  const totalIcons = Object.values(byCategory).reduce((sum, list) => sum + list.length, 0);
  return {
    topic: "icons",
    purpose: "Decorate cards (and future layouts) with editable BIT-styled vector icons.",
    whenToUse: "Use when a card title benefits from a quick visual cue. Keep one icon per card.",
    fields: {
      icon: { type: "string", required: false, value: "icon name from the catalog below" },
      iconColor: { type: "string", required: false, value: "color token (accent1, primary, green, red, ...) or 6-digit hex; defaults to BIT green" },
    },
    layouts: ["cards", "bullets", "comparison"],
    sizePresets: { ...ICON_SIZE_PRESETS },
    colorTokens: { ...ICON_COLOR_TOKENS },
    totalIcons,
    catalog: byCategory,
    notes: [
      "Icons are injected as native PowerPoint vector shapes — editable after generation.",
      "Default color resolves to BIT green; pass `iconColor: red` for emphasis or any 6-digit hex for custom colors.",
      "Card icon size is fixed by the layout (0.62 inch for single row, 0.46 inch for multi-row); custom sizing is not needed in YAML.",
      "Inside `bullets` (and `comparison` columns), pass `{ icon, iconColor?, text }` to replace the bullet dot with an icon. Plain strings still work.",
      "Icon names are stable; treat them as identifiers, not display labels.",
    ],
    example: {
      layout: "cards",
      title: "三大能力",
      cards: [
        { icon: "magnifier", title: "文献调研", text: "自动搜集、分类、摘要论文。" },
        { icon: "microscope", title: "实验跟踪", text: "记录每次实验的配置、参数与结果。" },
        { icon: "settings-gear", title: "算力调度", text: "通过 AutoDL API 管理 GPU 资源。" },
      ],
    },
  };
}

function getAllGuides() {
  return {
    overview: getGuideOverview(),
    speakerNotes: getSpeakerNotesGuide(),
    imagePlaceholder: getImagePlaceholderGuide(),
    icons: getIconsGuide(),
    writingRules: getWritingRules(),
    layouts: Object.fromEntries(listGuideLayouts().map((layout) => [layout, getLayoutGuide(layout)])),
  };
}

export {
  getAllGuides,
  getGuideOverview,
  getGuideWorkflow,
  getIconsGuide,
  getImagePlaceholderGuide,
  getLayoutExample,
  getLayoutGuide,
  getLayoutSchema,
  getSpeakerNotesGuide,
  getWritingRules,
  listGuideLayouts,
};
