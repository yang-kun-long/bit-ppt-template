const SUPPORTED_LAYOUTS = new Set([
  "title",
  "agenda",
  "section",
  "bullets",
  "claim",
  "twoColumn",
  "cards",
  "table",
  "comparison",
  "timeline",
  "process",
  "architecture",
  "ablation",
  "caseStudy",
  "imageGrid",
  "code",
  "appendix",
  "flowchart",
  "chart",
  "problemSolution",
  "painOpportunity",
  "experimentDesign",
  "resultAnalysis",
  "riskMitigation",
  "contribution",
  "summary",
  "metrics",
  "matrix",
  "quote",
  "formula",
  "references",
  "imageText",
  "closing",
]);

const CHART_TYPES = new Set(["bar", "line", "pie", "doughnut", "scatter", "area"]);

const PLACEHOLDER_ASPECT_RATIOS = {
  "16:9": 16 / 9,
  "4:3": 4 / 3,
  "3:2": 3 / 2,
  "1:1": 1,
  "4:5": 4 / 5,
  "3:4": 3 / 4,
  "9:16": 9 / 16,
};

function listLayouts() {
  return [...SUPPORTED_LAYOUTS];
}

export {
  CHART_TYPES,
  PLACEHOLDER_ASPECT_RATIOS,
  SUPPORTED_LAYOUTS,
  listLayouts,
};
