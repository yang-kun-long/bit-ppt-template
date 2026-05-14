import { CHART_TYPES, SUPPORTED_LAYOUTS } from "./layouts.mjs";
import {
  bulletText,
  expandSlidesWithReport,
  getSpeakerNotesValue,
  hasKnownPlaceholderRatio,
  imagePlaceholderNeedsVariants,
  isPlaceholderImage,
  normalizeSpeakerNotes,
  normalizeText,
  resolvePlaceholderRatio,
} from "./preflight.mjs";

function classifyImageRatio(ratio) {
  if (!Number.isFinite(ratio) || ratio <= 0) return "unknown";
  if (ratio >= 2.15) return "panoramic";
  if (ratio >= 1.18) return "landscape";
  if (ratio <= 0.52) return "tall";
  if (ratio <= 0.84) return "portrait";
  return "square";
}

function pureResolveImageInfo(value) {
  if (isPlaceholderImage(value)) {
    const aspectRatio = normalizeText(value.aspectRatio || value.ratio || value.size).toLowerCase();
    const ratio = resolvePlaceholderRatio(aspectRatio);
    return {
      placeholder: true,
      prompt: normalizeText(value.prompt || value.description || value.imagePrompt || value.alt),
      aspectRatio,
      exists: true,
      format: "placeholder",
      ratio,
      orientation: classifyImageRatio(ratio),
      uncertainRatio: !hasKnownPlaceholderRatio(aspectRatio),
    };
  }
  return {
    placeholder: false,
    exists: true,
    path: normalizeText(value),
    ratio: 1,
    orientation: "square",
    unchecked: true,
  };
}

function charCount(value) {
  return [...normalizeText(value)].length;
}

function lineCount(value) {
  const text = normalizeText(value);
  if (!text) return 0;
  return text.split(/\r?\n/).length;
}

function makeIssue(level, slideIndex, pathName, message, repair) {
  return { level, slideIndex, path: pathName, message, repair };
}

function validateDeck(deck, options = {}) {
  const resolveImageInfo = options.resolveImageInfo || pureResolveImageInfo;
  const issues = [];
  const slides = Array.isArray(deck?.slides) ? deck.slides : [];
  const add = (level, slideIndex, pathName, message, repair) => issues.push(makeIssue(level, slideIndex, pathName, message, repair));
  if (!deck || typeof deck !== "object") {
    return {
      errors: [makeIssue("error", 0, "deck", "Input must be a YAML object.", "Rewrite the deck as an object with meta and slides.")],
      warnings: [],
      repairPrompt: "deck: Rewrite the deck as an object with meta and slides.",
    };
  }
  if (!Array.isArray(deck.slides)) {
    add("error", 0, "slides", "`slides` must be an array.", "Return a top-level `slides` array.");
  }

  const maxText = (slideIndex, pathName, value, max, label = "text") => {
    if (value === undefined || value === null || value === "") return;
    const length = charCount(value);
    if (length > max) add("warning", slideIndex, pathName, `${label} is ${length} chars; recommended max is ${max}.`, `Shorten ${pathName} to ${max} characters or split the slide.`);
  };
  const maxItems = (slideIndex, pathName, value, max) => {
    if (!Array.isArray(value)) return;
    if (value.length > max) add("warning", slideIndex, pathName, `${pathName} has ${value.length} items; recommended max is ${max}.`, `Keep only the strongest ${max} items or split into multiple slides.`);
  };
  const requireArray = (slideIndex, pathName, value) => {
    if (value !== undefined && !Array.isArray(value)) add("error", slideIndex, pathName, `${pathName} must be an array.`, `Rewrite ${pathName} as a YAML list.`);
  };
  const checkBullets = (slideIndex, pathName, value, max = 5, textMax = 34) => {
    requireArray(slideIndex, pathName, value);
    maxItems(slideIndex, pathName, value, max);
    if (Array.isArray(value)) value.forEach((item, idx) => maxText(slideIndex, `${pathName}[${idx}]`, bulletText(item), textMax, "bullet"));
  };
  const checkImage = (slideIndex, pathName, value, imageOptions = {}) => {
    if (value === undefined || value === null || value === "") return;
    const image = resolveImageInfo(value);
    if (image.placeholder) {
      if (!image.prompt) add("warning", slideIndex, pathName, "Image placeholder has no prompt or description.", `Add ${pathName}.prompt so the placeholder tells the user what image to add later.`);
      else maxText(slideIndex, `${pathName}.prompt`, image.prompt, 160, "image placeholder prompt");
      if (image.uncertainRatio && !imageOptions.variants) add("warning", slideIndex, pathName, "Image placeholder aspect ratio is unknown.", "Set aspectRatio to 16:9, 4:3, 1:1, or 3:4.");
      return;
    }
    if (!image.exists) {
      add("error", slideIndex, pathName, `Image file does not exist: ${image.path}.`, `Fix ${pathName} to point to an existing image file.`);
      return;
    }
    if (image.path && !image.unchecked && (!image.width || !image.height)) {
      add("warning", slideIndex, pathName, `Image dimensions could not be read: ${image.path}.`, "Use PNG, JPEG, GIF, or WebP when automatic image layout is needed.");
    }
  };

  slides.forEach((slide, idx) => {
    const slideIndex = idx + 1;
    if (!slide || typeof slide !== "object") {
      add("error", slideIndex, "slide", "Each slide must be an object.", "Rewrite this slide as a YAML object with a `layout` field.");
      return;
    }
    if (!SUPPORTED_LAYOUTS.has(slide.layout)) {
      add("error", slideIndex, "layout", `Unknown layout: ${slide.layout || "(missing)"}.`, `Use one of: ${[...SUPPORTED_LAYOUTS].join(", ")}.`);
      return;
    }
    maxText(slideIndex, "title", slide.title, 24, "title");
    const speakerNotesRaw = getSpeakerNotesValue(slide);
    if (speakerNotesRaw !== undefined) {
      if (typeof speakerNotesRaw !== "string" && !Array.isArray(speakerNotesRaw)) {
        add("warning", slideIndex, "speakerNotes", "speakerNotes should be a string block or an array of short strings.", "Rewrite speakerNotes as a YAML block scalar or string list.");
      }
      const speakerNotes = normalizeSpeakerNotes(speakerNotesRaw);
      maxText(slideIndex, "speakerNotes", speakerNotes, 1200, "speaker notes");
      if (lineCount(speakerNotes) > 30) add("warning", slideIndex, "speakerNotes", `speakerNotes has ${lineCount(speakerNotes)} lines; recommended max is 30.`, "Shorten the speaker script or split it across slides.");
    }

    switch (slide.layout) {
      case "agenda":
        checkBullets(slideIndex, "items", slide.items, 7, 24);
        break;
      case "section":
        maxText(slideIndex, "subtitle", slide.subtitle, 42, "subtitle");
        break;
      case "bullets":
        maxText(slideIndex, "lead", slide.lead, 58, "lead");
        checkBullets(slideIndex, "bullets", slide.bullets, 8, 38);
        break;
      case "claim":
        maxText(slideIndex, "claim", slide.claim, 54, "claim");
        checkBullets(slideIndex, "evidence", slide.evidence, 5, 38);
        break;
      case "twoColumn":
        ["left", "right"].forEach((side) => {
          maxText(slideIndex, `${side}.title`, slide[side]?.title, 16, "column title");
          maxText(slideIndex, `${side}.text`, slide[side]?.text, 105, "column text");
          checkBullets(slideIndex, `${side}.bullets`, slide[side]?.bullets, 5, 34);
        });
        break;
      case "cards":
        maxItems(slideIndex, "cards", slide.cards, 6);
        if (Array.isArray(slide.cards)) slide.cards.forEach((card, cardIdx) => {
          maxText(slideIndex, `cards[${cardIdx}].title`, card.title, 14, "card title");
          maxText(slideIndex, `cards[${cardIdx}].text`, card.text, 52, "card text");
        });
        break;
      case "table":
        requireArray(slideIndex, "columns", slide.columns);
        requireArray(slideIndex, "rows", slide.rows);
        if (Array.isArray(slide.columns) && slide.columns.length > 5) add("warning", slideIndex, "columns", `Table has ${slide.columns.length} columns; recommended max is 5.`, "Split wide tables by columns or use multiple slides.");
        if (Array.isArray(slide.rows)) slide.rows.forEach((row, rowIdx) => {
          if (!Array.isArray(row)) add("error", slideIndex, `rows[${rowIdx}]`, "Each table row must be an array.", `Rewrite rows[${rowIdx}] as a YAML list.`);
          else row.forEach((cell, cellIdx) => maxText(slideIndex, `rows[${rowIdx}][${cellIdx}]`, cell, 42, "table cell"));
        });
        break;
      case "comparison":
        ["left", "right"].forEach((side) => {
          maxText(slideIndex, `${side}.title`, slide[side]?.title, 18, "comparison title");
          checkBullets(slideIndex, `${side}.bullets`, slide[side]?.bullets, 5, 32);
        });
        break;
      case "timeline":
        maxItems(slideIndex, "items", slide.items, 6);
        if (Array.isArray(slide.items)) slide.items.forEach((item, itemIdx) => {
          maxText(slideIndex, `items[${itemIdx}].title`, item.title, 10, "timeline title");
          maxText(slideIndex, `items[${itemIdx}].text`, item.text, 24, "timeline text");
        });
        break;
      case "process":
        maxItems(slideIndex, "steps", slide.steps, 5);
        if (Array.isArray(slide.steps)) slide.steps.forEach((step, stepIdx) => {
          maxText(slideIndex, `steps[${stepIdx}].title`, step.title, 8, "step title");
          maxText(slideIndex, `steps[${stepIdx}].text`, step.text, 24, "step text");
        });
        break;
      case "architecture":
        maxItems(slideIndex, "layers", slide.layers, 4);
        if (Array.isArray(slide.layers)) slide.layers.forEach((layer, layerIdx) => {
          maxText(slideIndex, `layers[${layerIdx}].title`, layer.title, 8, "layer title");
          maxItems(slideIndex, `layers[${layerIdx}].components`, layer.components, 5);
          if (Array.isArray(layer.components)) layer.components.forEach((component, compIdx) => maxText(slideIndex, `layers[${layerIdx}].components[${compIdx}]`, component, 10, "component label"));
          maxText(slideIndex, `layers[${layerIdx}].note`, layer.note, 44, "layer note");
        });
        break;
      case "ablation":
        maxText(slideIndex, "baseline", slide.baseline, 58, "baseline");
        maxItems(slideIndex, "items", slide.items, 6);
        if (Array.isArray(slide.items)) slide.items.forEach((item, itemIdx) => ["factor", "setting", "delta", "conclusion"].forEach((key) => maxText(slideIndex, `items[${itemIdx}].${key}`, item[key], key === "conclusion" ? 28 : 18, key)));
        break;
      case "caseStudy":
        checkImage(slideIndex, "image", slide.image);
        checkBullets(slideIndex, "context", slide.context, 2, 34);
        checkBullets(slideIndex, "method", slide.method, 2, 34);
        checkBullets(slideIndex, "result", slide.result, 2, 34);
        maxText(slideIndex, "caption", slide.caption, 40, "caption");
        break;
      case "imageGrid":
        maxItems(slideIndex, "images", slide.images, 6);
        if (Array.isArray(slide.images)) slide.images.forEach((image, imageIdx) => {
          checkImage(slideIndex, `images[${imageIdx}]`, image);
          const caption = image && typeof image === "object" ? image.caption : "";
          maxText(slideIndex, `images[${imageIdx}].caption`, caption, 16, "image caption");
        });
        break;
      case "code":
        if (lineCount(slide.code || slide.algorithm) > 12) add("warning", slideIndex, "code", `Code block has ${lineCount(slide.code || slide.algorithm)} lines; recommended max is 12.`, "Summarize the code as pseudocode under 12 short lines.");
        normalizeText(slide.code || slide.algorithm).split(/\r?\n/).forEach((line, lineIdx) => maxText(slideIndex, `code line ${lineIdx + 1}`, line, 72, "code line"));
        checkBullets(slideIndex, "notes", slide.notes, 5, 32);
        break;
      case "appendix":
        maxItems(slideIndex, "items", slide.items, 8);
        if (Array.isArray(slide.items)) slide.items.forEach((item, itemIdx) => {
          maxText(slideIndex, `items[${itemIdx}].title`, item.title, 14, "appendix title");
          maxText(slideIndex, `items[${itemIdx}].text`, item.text, 42, "appendix text");
        });
        break;
      case "flowchart": {
        requireArray(slideIndex, "nodes", slide.nodes);
        if (!Array.isArray(slide.nodes) || !slide.nodes.length) add("error", slideIndex, "nodes", "Flowchart requires at least one node.", "Add a `nodes` list with id and text fields.");
        maxItems(slideIndex, "nodes", slide.nodes, 10);
        const ids = new Set();
        if (Array.isArray(slide.nodes)) slide.nodes.forEach((node, nodeIdx) => {
          const id = node.id || String(nodeIdx + 1);
          if (ids.has(id)) add("error", slideIndex, `nodes[${nodeIdx}].id`, `Duplicate flowchart node id: ${id}.`, "Use unique ids for every flowchart node.");
          ids.add(id);
          maxText(slideIndex, `nodes[${nodeIdx}].text`, node.text || node.label, 12, "flowchart node text");
          maxText(slideIndex, `nodes[${nodeIdx}].note`, node.note, 22, "flowchart node note");
        });
        if (Array.isArray(slide.edges)) slide.edges.forEach((edge, edgeIdx) => {
          if (!ids.has(edge.from)) add("error", slideIndex, `edges[${edgeIdx}].from`, `Unknown flowchart edge source: ${edge.from}.`, "Point edge.from to an existing node id.");
          if (!ids.has(edge.to)) add("error", slideIndex, `edges[${edgeIdx}].to`, `Unknown flowchart edge target: ${edge.to}.`, "Point edge.to to an existing node id.");
        });
        maxText(slideIndex, "note", slide.note, 56, "flowchart note");
        break;
      }
      case "chart": {
        const type = normalizeText(slide.type || "bar").toLowerCase();
        if (!CHART_TYPES.has(type)) add("error", slideIndex, "type", `Unsupported chart type: ${slide.type}.`, "Use chart type bar, line, pie, doughnut, scatter, or area.");
        const categories = slide.categories || slide.labels || [];
        if (!Array.isArray(categories) || !categories.length) add("error", slideIndex, "categories", "Chart requires categories or labels.", "Add a `categories` list whose length matches every series values list.");
        maxItems(slideIndex, "categories", categories, type === "pie" || type === "doughnut" ? 6 : 8);
        if (Array.isArray(categories)) categories.forEach((category, catIdx) => maxText(slideIndex, `categories[${catIdx}]`, category, 18, "category label"));
        if (slide.series !== undefined && !Array.isArray(slide.series)) add("error", slideIndex, "series", "`series` must be an array.", "Rewrite chart series as a YAML list.");
        if (!Array.isArray(slide.series) && slide.values === undefined) add("error", slideIndex, "series", "Chart requires `series` or shortcut `values`.", "Add `series` with numeric values, or add a single `values` list.");
        const seriesList = Array.isArray(slide.series) ? slide.series : slide.values !== undefined ? [{ name: slide.name || "Series", values: slide.values }] : [];
        seriesList.forEach((series, seriesIdx) => {
          maxText(slideIndex, `series[${seriesIdx}].name`, series.name, 16, "series name");
          if (!Array.isArray(series.values)) add("error", slideIndex, `series[${seriesIdx}].values`, "Chart series values must be an array.", `Rewrite series[${seriesIdx}].values as a numeric list.`);
          else {
            if (Array.isArray(categories) && categories.length && series.values.length !== categories.length) add("error", slideIndex, `series[${seriesIdx}].values`, `Series has ${series.values.length} values but categories has ${categories.length}.`, "Make every chart series have the same number of values as categories.");
            series.values.forEach((value, valueIdx) => {
              if (!Number.isFinite(Number(value))) add("error", slideIndex, `series[${seriesIdx}].values[${valueIdx}]`, `Chart value is not numeric: ${value}.`, "Use numeric chart values only.");
            });
          }
        });
        maxText(slideIndex, "caption", slide.caption, 58, "chart caption");
        break;
      }
      case "problemSolution":
        ["problem", "solution", "impact"].forEach((key) => {
          maxText(slideIndex, `${key}.title`, slide[key]?.title, 12, `${key} title`);
          checkBullets(slideIndex, `${key}.bullets`, slide[key]?.bullets, 4, 26);
        });
        break;
      case "painOpportunity":
        ["status", "pain", "opportunity"].forEach((key) => {
          maxText(slideIndex, `${key}.title`, slide[key]?.title, 12, `${key} title`);
          checkBullets(slideIndex, `${key}.bullets`, slide[key]?.bullets, 4, 30);
        });
        break;
      case "experimentDesign":
        ["dataset", "variables", "metrics", "baselines"].forEach((key) => checkBullets(slideIndex, key, Array.isArray(slide[key]) ? slide[key] : slide[key] ? [slide[key]] : [], 4, 26));
        maxItems(slideIndex, "procedure", slide.procedure, 5);
        break;
      case "resultAnalysis":
        maxText(slideIndex, "finding", slide.finding, 52, "finding");
        maxItems(slideIndex, "metrics", slide.metrics, 3);
        checkBullets(slideIndex, "analysis", slide.analysis, 4, 38);
        break;
      case "riskMitigation":
        maxItems(slideIndex, "items", slide.items, 5);
        if (Array.isArray(slide.items)) slide.items.forEach((item, itemIdx) => ["risk", "impact", "mitigation"].forEach((key) => maxText(slideIndex, `items[${itemIdx}].${key}`, item[key], 28, key)));
        break;
      case "contribution":
        maxItems(slideIndex, "items", slide.items, 4);
        if (Array.isArray(slide.items)) slide.items.forEach((item, itemIdx) => {
          maxText(slideIndex, `items[${itemIdx}].title`, item.title, 14, "contribution title");
          maxText(slideIndex, `items[${itemIdx}].text`, item.text, 44, "contribution text");
        });
        break;
      case "summary":
        maxText(slideIndex, "takeaway", slide.takeaway, 52, "takeaway");
        checkBullets(slideIndex, "points", slide.points, 5, 36);
        break;
      case "metrics":
        maxItems(slideIndex, "metrics", slide.metrics, 4);
        if (Array.isArray(slide.metrics)) slide.metrics.forEach((metric, metricIdx) => {
          maxText(slideIndex, `metrics[${metricIdx}].value`, metric.value, 8, "metric value");
          maxText(slideIndex, `metrics[${metricIdx}].label`, metric.label, 12, "metric label");
          maxText(slideIndex, `metrics[${metricIdx}].note`, metric.note, 30, "metric note");
        });
        break;
      case "matrix":
        maxItems(slideIndex, "cells", slide.cells, 4);
        if (Array.isArray(slide.cells)) slide.cells.forEach((cell, cellIdx) => {
          maxText(slideIndex, `cells[${cellIdx}].title`, cell.title, 18, "matrix title");
          maxText(slideIndex, `cells[${cellIdx}].text`, cell.text, 52, "matrix text");
        });
        break;
      case "quote":
        maxText(slideIndex, "quote", slide.quote, 70, "quote");
        break;
      case "formula":
        if (!slide.formula || (typeof slide.formula === "object" && !slide.formula.latex)) add("warning", slideIndex, "formula", "Formula slide has no formula.", "Add formula.latex or change the slide layout.");
        checkBullets(slideIndex, "explanation", slide.explanation || slide.notes, 4, 48);
        break;
      case "references":
        requireArray(slideIndex, "items", slide.items);
        if (Array.isArray(slide.items)) slide.items.forEach((item, itemIdx) => maxText(slideIndex, `items[${itemIdx}]`, item, 140, "reference"));
        break;
      case "imageText":
        checkImage(slideIndex, "image", slide.image, { variants: imagePlaceholderNeedsVariants(slide) });
        checkBullets(slideIndex, "text", slide.text, 5, 38);
        break;
      default:
        break;
    }
  });

  const errors = issues.filter((item) => item.level === "error");
  const warnings = issues.filter((item) => item.level === "warning");
  const repairPrompt = issues.length
    ? issues.map((item) => `Slide ${item.slideIndex} ${item.path}: ${item.repair || item.message}`).join("\n")
    : "";
  return { errors, warnings, repairPrompt };
}

function checkDeck(deck, options = {}) {
  const validation = validateDeck(deck, options);
  const preflightOnly = expandSlidesWithReport(Array.isArray(deck?.slides) ? deck.slides : []);
  return {
    inputSlides: Array.isArray(deck?.slides) ? deck.slides.length : 0,
    outputSlides: preflightOnly.slides.length,
    actions: preflightOnly.report,
    validation: {
      errors: validation.errors,
      warnings: validation.warnings,
    },
    repairPrompt: validation.repairPrompt,
  };
}

export {
  checkDeck,
  pureResolveImageInfo,
  validateDeck,
};
