const writingRules = [
  "Keep slide text short and concrete; use validation warnings as rewrite hints.",
  "Prefer editable PPTX objects: text boxes, shapes, tables, charts, OMML formulas, and native images.",
  "Use `speakerNotes` for presenter scripts; these notes are written to the PPTX notes pane, not the slide canvas.",
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

const layoutGuides = {
  imageText: {
    layout: "imageText",
    purpose: "Show one image with a short explanatory bullet list.",
    whenToUse: "Use for visual evidence, screenshots, diagrams, and result images that need concise interpretation.",
    fields: {
      layout: { type: "literal", required: true, value: "imageText" },
      title: { type: "string", required: true },
      image: { type: "string | { path, placement, fit }", required: true },
      caption: { type: "string", required: false },
      text: { type: "string[]", required: true },
    },
    limits: {
      text: { maxItems: 5, recommendedChars: 38 },
    },
    notes: [
      "Very wide images are automatically placed above the text.",
      "Ordinary landscape, portrait, and square images stay side-by-side so text space remains usable.",
      "`placement: top` or `placement: side` overrides automatic placement.",
      "`fit: contain` avoids cropping; `fit: cover` fills the image box.",
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

function listGuideLayouts() {
  return Object.keys(layoutGuides);
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
      "bit-ppt guide writing-rules",
    ],
    guidedLayouts: listGuideLayouts(),
  };
}

function getSpeakerNotesGuide() {
  return { ...speakerNotesGuide };
}

function getWritingRules() {
  return [...writingRules];
}

function getGuideWorkflow() {
  return getGuideOverview().workflow;
}

function getAllGuides() {
  return {
    overview: getGuideOverview(),
    speakerNotes: getSpeakerNotesGuide(),
    writingRules: getWritingRules(),
    layouts: Object.fromEntries(listGuideLayouts().map((layout) => [layout, getLayoutGuide(layout)])),
  };
}

export {
  getAllGuides,
  getGuideOverview,
  getGuideWorkflow,
  getLayoutExample,
  getLayoutGuide,
  getLayoutSchema,
  getSpeakerNotesGuide,
  getWritingRules,
  listGuideLayouts,
};
