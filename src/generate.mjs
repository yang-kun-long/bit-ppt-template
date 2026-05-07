import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath, pathToFileURL } from "node:url";
import JSZip from "jszip";
import pptxgen from "pptxgenjs";
import YAML from "yaml";
const require = createRequire(import.meta.url);
const { latexToOMML } = require("latex-to-omml");

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const W = 13.333;
const H = 7.5;

const theme = {
  green: "006C39",
  darkGreen: "004B28",
  red: "A13F3D",
  ink: "262626",
  muted: "666666",
  light: "F6F8F6",
  line: "DDE7E2",
  white: "FFFFFF",
};

const font = {
  cn: "微软雅黑",
  cnLight: "微软雅黑 Light",
  serif: "SimSun",
  en: "Arial",
  code: "Consolas",
};

const defaultFont = { ...font };

const INCH_PT = 72;
const OVERFLOW_GUARD = 0.9;
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

const assets = {
  campusLine: asset("bit-campus-line.png"),
  seal: asset("bit-seal-small.png"),
  wordmarkWhite: asset("bit-wordmark-white.png"),
  campusPhoto: asset("bit-campus-photo.png"),
  emblemGray: asset("bit-emblem-gray.png"),
};

const assetRatio = {
  [assets.campusLine]: 449 / 166,
  [assets.seal]: 1,
  [assets.wordmarkWhite]: 901 / 252,
  [assets.campusPhoto]: 640 / 426,
  [assets.emblemGray]: 620 / 621,
};

const imageInfoCache = new Map();

function asset(name) {
  return path.join(ROOT, "assets", name);
}

function resolveProjectPath(value, fallback) {
  const source = normalizeText(value || fallback);
  return path.isAbsolute(source) ? source : path.resolve(ROOT, source);
}

function normalizeImageSpec(value, fallback = "assets/bit-campus-photo.png") {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    return {
      path: resolveProjectPath(value.path || value.image || value.src, fallback),
      fit: normalizeText(value.fit || value.sizing).toLowerCase(),
      placement: normalizeText(value.placement || value.position).toLowerCase(),
      alt: normalizeText(value.alt),
    };
  }
  return {
    path: resolveProjectPath(value, fallback),
    fit: "",
    placement: "",
    alt: "",
  };
}

function readPngDimensions(buffer) {
  if (buffer.length < 24) return null;
  if (buffer[0] !== 0x89 || buffer.toString("ascii", 1, 4) !== "PNG") return null;
  return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20), format: "png" };
}

function readGifDimensions(buffer) {
  if (buffer.length < 10) return null;
  const signature = buffer.toString("ascii", 0, 6);
  if (signature !== "GIF87a" && signature !== "GIF89a") return null;
  return { width: buffer.readUInt16LE(6), height: buffer.readUInt16LE(8), format: "gif" };
}

function readJpegDimensions(buffer) {
  if (buffer.length < 4 || buffer[0] !== 0xff || buffer[1] !== 0xd8) return null;
  let offset = 2;
  while (offset + 9 < buffer.length) {
    while (buffer[offset] === 0xff) offset += 1;
    const marker = buffer[offset];
    offset += 1;
    if (marker === 0xd9 || marker === 0xda) break;
    const size = buffer.readUInt16BE(offset);
    if (size < 2 || offset + size > buffer.length) break;
    if ((marker >= 0xc0 && marker <= 0xc3) || (marker >= 0xc5 && marker <= 0xc7) || (marker >= 0xc9 && marker <= 0xcb) || (marker >= 0xcd && marker <= 0xcf)) {
      return {
        width: buffer.readUInt16BE(offset + 5),
        height: buffer.readUInt16BE(offset + 3),
        format: "jpeg",
      };
    }
    offset += size;
  }
  return null;
}

function readWebpDimensions(buffer) {
  if (buffer.length < 30 || buffer.toString("ascii", 0, 4) !== "RIFF" || buffer.toString("ascii", 8, 12) !== "WEBP") return null;
  const chunk = buffer.toString("ascii", 12, 16);
  if (chunk === "VP8X" && buffer.length >= 30) {
    return {
      width: 1 + buffer.readUIntLE(24, 3),
      height: 1 + buffer.readUIntLE(27, 3),
      format: "webp",
    };
  }
  if (chunk === "VP8 " && buffer.length >= 30) {
    return {
      width: buffer.readUInt16LE(26) & 0x3fff,
      height: buffer.readUInt16LE(28) & 0x3fff,
      format: "webp",
    };
  }
  if (chunk === "VP8L" && buffer.length >= 25) {
    const bits = buffer.readUInt32LE(21);
    return {
      width: 1 + (bits & 0x3fff),
      height: 1 + ((bits >> 14) & 0x3fff),
      format: "webp",
    };
  }
  return null;
}

function readImageDimensions(imagePath) {
  if (imageInfoCache.has(imagePath)) return imageInfoCache.get(imagePath);
  let result = null;
  try {
    const buffer = fs.readFileSync(imagePath);
    result = readPngDimensions(buffer) || readJpegDimensions(buffer) || readGifDimensions(buffer) || readWebpDimensions(buffer);
  } catch {
    result = null;
  }
  imageInfoCache.set(imagePath, result);
  return result;
}

function classifyImageRatio(ratio) {
  if (!Number.isFinite(ratio) || ratio <= 0) return "unknown";
  if (ratio >= 2.15) return "panoramic";
  if (ratio >= 1.18) return "landscape";
  if (ratio <= 0.52) return "tall";
  if (ratio <= 0.84) return "portrait";
  return "square";
}

function resolveImageInfo(value, fallback = "assets/bit-campus-photo.png") {
  const spec = normalizeImageSpec(value, fallback);
  const dimensions = readImageDimensions(spec.path);
  const ratio = dimensions?.width && dimensions?.height
    ? dimensions.width / dimensions.height
    : assetRatio[spec.path] || 1;
  return {
    ...spec,
    exists: fs.existsSync(spec.path),
    width: dimensions?.width,
    height: dimensions?.height,
    format: dimensions?.format,
    ratio,
    orientation: classifyImageRatio(ratio),
  };
}

function fitBoxToRatio(box, ratio = 1) {
  let { x, y, w, h } = box;
  const safeRatio = Number.isFinite(ratio) && ratio > 0 ? ratio : 1;
  const boxRatio = w / h;
  if (boxRatio > safeRatio) {
    const fittedW = h * safeRatio;
    x += (w - fittedW) / 2;
    w = fittedW;
  } else {
    const fittedH = w / safeRatio;
    y += (h - fittedH) / 2;
    h = fittedH;
  }
  return { x, y, w, h };
}

function imageFitMode(image, fallback = "cover") {
  const value = normalizeText(image?.fit).toLowerCase();
  if (["contain", "fit", "inside"].includes(value)) return "contain";
  if (["cover", "crop", "fill"].includes(value)) return "cover";
  return fallback;
}

function readDeck(inputFile) {
  const raw = fs.readFileSync(inputFile, "utf8");
  return YAML.parse(raw);
}

function normalizeText(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  if (Array.isArray(value)) return value.map(normalizeText).join(" ");
  if (typeof value === "object") {
    return Object.entries(value)
      .map(([key, val]) => `${key}: ${normalizeText(val)}`)
      .join(" ");
  }
  return String(value);
}

function normalizeLatex(value) {
  return normalizeText(value)
    .replace(/\\\\([A-Za-z]+)/g, "\\$1")
    .replace(/([_^])\\([A-Za-z]+)(?![A-Za-z{])/g, "$1{\\$2}")
    .replace(/([_^])([A-Za-z0-9])(?![A-Za-z0-9{])/g, "$1{$2}");
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function normalizeFonts(fonts = {}) {
  const normalized = {};
  if (fonts.cn) normalized.cn = normalizeText(fonts.cn);
  if (fonts.cjk) normalized.cn = normalizeText(fonts.cjk);
  if (fonts.cnLight) normalized.cnLight = normalizeText(fonts.cnLight);
  if (fonts.cjkLight) normalized.cnLight = normalizeText(fonts.cjkLight);
  if (fonts.en) normalized.en = normalizeText(fonts.en);
  if (fonts.latin) normalized.en = normalizeText(fonts.latin);
  if (fonts.serif) normalized.serif = normalizeText(fonts.serif);
  if (fonts.code) normalized.code = normalizeText(fonts.code);
  return Object.fromEntries(Object.entries(normalized).filter(([, value]) => value));
}

function configureFonts(deck = {}, options = {}) {
  const merged = {
    ...defaultFont,
    ...normalizeFonts(deck.meta?.fonts || {}),
    ...normalizeFonts(options.fonts || {}),
  };
  if (options.fontCn) merged.cn = normalizeText(options.fontCn);
  if (options.fontCjk) merged.cn = normalizeText(options.fontCjk);
  if (options.fontCnLight) merged.cnLight = normalizeText(options.fontCnLight);
  if (options.fontEn) merged.en = normalizeText(options.fontEn);
  if (options.fontSerif) merged.serif = normalizeText(options.fontSerif);
  if (options.fontCode) merged.code = normalizeText(options.fontCode);
  Object.assign(font, merged);
  return { ...font };
}

function isCjk(ch) {
  return /[\u3400-\u9fff\uff00-\uffef]/u.test(ch);
}

function charWidthPt(ch, fontSize, bold = false) {
  const factor = bold ? 1.06 : 1;
  if (ch === "\n") return Infinity;
  if (/\s/.test(ch)) return fontSize * 0.33;
  if (isCjk(ch)) return fontSize * 1.0 * factor;
  if (/[A-Z]/.test(ch)) return fontSize * 0.64 * factor;
  if (/[0-9]/.test(ch)) return fontSize * 0.58 * factor;
  if (/[.,;:!?()[\]{}'"`/\\|+-]/.test(ch)) return fontSize * 0.38 * factor;
  return fontSize * 0.54 * factor;
}

function estimateText(text, boxW, fontSize, opts = {}) {
  const maxPt = Math.max(1, boxW * INCH_PT - (opts.marginPt || 0));
  const lineHeight = opts.lineHeight || 1.18;
  const value = normalizeText(text);
  let lines = 1;
  let current = 0;
  for (const ch of value) {
    const width = charWidthPt(ch, fontSize, opts.bold);
    if (!Number.isFinite(width)) {
      lines += 1;
      current = 0;
      continue;
    }
    if (current > 0 && current + width > maxPt) {
      lines += 1;
      current = width;
    } else {
      current += width;
    }
  }
  return {
    lines,
    height: (lines * fontSize * lineHeight) / INCH_PT,
  };
}

function estimateBulletHeight(text, boxW, fontSize) {
  return estimateText(text, Math.max(0.5, boxW - 0.34), fontSize, { lineHeight: 1.22 }).height + 0.14;
}

function splitBulletsSlide(slide) {
  const bullets = slide.bullets || [];
  const fontSize = slide.fontSize || 16;
  const startY = slide.lead ? 2.92 : 1.72;
  const available = 6.62 - startY;
  const chunks = [];
  let chunk = [];
  let used = 0;
  for (const bullet of bullets) {
    const h = estimateBulletHeight(bullet, 10.4, fontSize);
    if (chunk.length && used + h > available * OVERFLOW_GUARD) {
      chunks.push(chunk);
      chunk = [];
      used = 0;
    }
    chunk.push(bullet);
    used += h;
  }
  if (chunk.length) chunks.push(chunk);
  if (chunks.length <= 1) return [slide];
  return chunks.map((items, idx) => ({
    ...clone(slide),
    title: `${slide.title || "要点"}（${idx + 1}/${chunks.length}）`,
    lead: idx === 0 ? slide.lead : undefined,
    bullets: items,
  }));
}

function estimateTableRowHeight(row, colW, fontSize) {
  const heights = row.map((cell, idx) => estimateText(cell, colW[idx] || colW[0], fontSize, { lineHeight: 1.16, marginPt: 10 }).height + 0.16);
  return Math.max(0.42, ...heights);
}

function splitTableSlide(slide) {
  const columns = slide.columns || [];
  const rows = slide.rows || [];
  const fontSize = rows.length > 5 ? 8.5 : 10;
  const tableW = 11.82;
  const colW = columns.map(() => tableW / Math.max(1, columns.length));
  const available = 4.95 - 0.5;
  const chunks = [];
  let chunk = [];
  let used = 0;
  for (const row of rows) {
    const h = estimateTableRowHeight(row, colW, fontSize);
    if (chunk.length && used + h > available * OVERFLOW_GUARD) {
      chunks.push(chunk);
      chunk = [];
      used = 0;
    }
    chunk.push(row);
    used += h;
  }
  if (chunk.length) chunks.push(chunk);
  if (chunks.length <= 1) return [slide];
  return chunks.map((items, idx) => ({
    ...clone(slide),
    title: `${slide.title || "表格"}（${idx + 1}/${chunks.length}）`,
    rows: items,
  }));
}

function splitReferencesSlide(slide) {
  const items = slide.items || [];
  const fontSize = slide.fontSize || 9.5;
  const available = 4.95;
  const chunks = [];
  let chunk = [];
  let used = 0;
  for (const item of items) {
    const h = estimateText(item, 11.1, fontSize, { lineHeight: 1.18 }).height + 0.11;
    if (chunk.length && used + h > available * OVERFLOW_GUARD) {
      chunks.push(chunk);
      chunk = [];
      used = 0;
    }
    chunk.push(item);
    used += h;
  }
  if (chunk.length) chunks.push(chunk);
  if (chunks.length <= 1) return [slide];
  return chunks.map((itemsChunk, idx) => ({
    ...clone(slide),
    title: `${slide.title || "参考文献"}（${idx + 1}/${chunks.length}）`,
    items: itemsChunk,
  }));
}

function expandSlides(slides = []) {
  return expandSlidesWithReport(slides).slides;
}

function expandSlidesWithReport(slides = []) {
  const report = [];
  const expanded = [];
  slides.forEach((slide, index) => {
    let parts;
    if (slide.layout === "bullets") parts = splitBulletsSlide(slide);
    else if (slide.layout === "table") parts = splitTableSlide(slide);
    else if (slide.layout === "references") parts = splitReferencesSlide(slide);
    else parts = [slide];
    if (parts.length > 1) {
      report.push({
        slideIndex: index + 1,
        layout: slide.layout,
        title: slide.title || slide.layout,
        action: "split",
        parts: parts.length,
      });
    }
    expanded.push(...parts);
  });
  return { slides: expanded, report };
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

function validateDeck(deck) {
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
    if (Array.isArray(value)) value.forEach((item, idx) => maxText(slideIndex, `${pathName}[${idx}]`, item, textMax, "bullet"));
  };
  const checkImage = (slideIndex, pathName, value) => {
    if (value === undefined || value === null || value === "") return;
    const image = resolveImageInfo(value);
    if (!image.exists) {
      add("error", slideIndex, pathName, `Image file does not exist: ${image.path}.`, `Fix ${pathName} to point to an existing image file.`);
      return;
    }
    if (!image.width || !image.height) {
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
        checkImage(slideIndex, "image", slide.image);
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

function addText(slide, text, x, y, w, h, opts = {}) {
  slide.addText(normalizeText(text), {
    x,
    y,
    w,
    h,
    margin: 0,
    fontFace: opts.fontFace || font.cn,
    fontSize: opts.fontSize || 18,
    color: opts.color || theme.ink,
    bold: opts.bold || false,
    breakLine: opts.breakLine || false,
    valign: opts.valign || "mid",
    fit: opts.fit || "shrink",
    ...opts,
  });
}

function addRichText(slide, runs, x, y, w, h, opts = {}) {
  slide.addText(runs, {
    x,
    y,
    w,
    h,
    margin: 0,
    fontFace: opts.fontFace || font.cn,
    fontSize: opts.fontSize || 18,
    color: opts.color || theme.ink,
    valign: opts.valign || "mid",
    fit: opts.fit || "shrink",
    ...opts,
  });
}

function registerEquation(ctx, latex, opts = {}) {
  const marker = `__BIT_OMML_${ctx.equations.length + 1}__`;
  ctx.equations.push({
    marker,
    latex: normalizeLatex(latex),
    display: opts.display !== false,
  });
  return marker;
}

function inlineMathRuns(text, baseOptions, ctx) {
  const source = normalizeText(text);
  const runs = [];
  const pattern = /\$([^$]+)\$|\\\((.+?)\\\)/g;
  let last = 0;
  let match;
  while ((match = pattern.exec(source))) {
    if (match.index > last) {
      runs.push({ text: source.slice(last, match.index), options: baseOptions });
    }
    const latex = match[1] || match[2];
    runs.push({
      text: registerEquation(ctx, latex, { display: false }),
      options: { ...baseOptions, fontFace: font.en },
    });
    last = pattern.lastIndex;
  }
  if (last < source.length) {
    runs.push({ text: source.slice(last), options: baseOptions });
  }
  return runs.length ? runs : [{ text: source, options: baseOptions }];
}

function inlineMathTableCell(text, cellOptions, ctx) {
  if (!ctx) return normalizeText(text);
  const runs = inlineMathRuns(text, cellOptions, ctx);
  return runs.length === 1 && runs[0].text === normalizeText(text) ? normalizeText(text) : runs;
}

function addInlineMathText(slide, text, x, y, w, h, opts, ctx) {
  addRichText(slide, inlineMathRuns(text, opts, ctx), x, y, w, h, opts);
}

function addImageFit(slide, imagePath, box, opts = {}) {
  const dimensions = readImageDimensions(imagePath);
  const ratio = opts.ratio || (dimensions?.width && dimensions?.height ? dimensions.width / dimensions.height : assetRatio[imagePath]) || 1;
  const { x, y, w, h } = fitBoxToRatio(box, ratio);
  slide.addImage({ path: imagePath, x, y, w, h, transparency: opts.transparency || 0 });
}

function addImageCover(slide, imagePath, box, opts = {}) {
  slide.addImage({
    path: imagePath,
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    transparency: opts.transparency || 0,
    sizing: { type: "cover", x: box.x, y: box.y, w: box.w, h: box.h },
  });
}

function addImageInPanel(slide, image, box, opts = {}) {
  const mode = imageFitMode(image, opts.fit || "cover");
  if (mode === "contain") {
    const imageBox = fitBoxToRatio(box, image.ratio);
    const pad = opts.pad ?? 0.12;
    slide.addShape("rect", {
      x: imageBox.x - pad,
      y: imageBox.y - pad,
      w: imageBox.w + pad * 2,
      h: imageBox.h + pad * 2,
      fill: { color: opts.fill || theme.light },
      line: { color: opts.line || theme.line, width: opts.lineWidth || 0.8 },
    });
    addImageFit(slide, image.path, imageBox, { ratio: image.ratio, transparency: opts.transparency || 0 });
    return imageBox;
  }
  slide.addShape("rect", {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    fill: { color: opts.fill || theme.light },
    line: { color: opts.line || theme.line, width: opts.lineWidth || 0.8 },
  });
  addImageCover(slide, image.path, box, { transparency: opts.transparency || 0 });
  return box;
}

function addCoverBrand(slide) {
  slide.background = { color: theme.green };
  addImageCover(slide, assets.campusPhoto, { x: 7.25, y: 0, w: 6.08, h: 7.5 }, { transparency: 12 });
  slide.addShape("rect", { x: 7.25, y: 0, w: 6.08, h: 7.5, fill: { color: theme.darkGreen, transparency: 22 }, line: { transparency: 100 } });
  addImageFit(slide, assets.wordmarkWhite, { x: 0.72, y: 0.5, w: 3.95, h: 1.08 });
  slide.addShape("line", { x: 0.78, y: 6.78, w: 3.8, h: 0, line: { color: theme.white, width: 1, transparency: 25 } });
}

function addPageBrand(slide, pageNo) {
  slide.background = { color: theme.white };
  slide.addShape("rect", { x: 0, y: 0, w: W, h: 0.18, fill: { color: theme.green }, line: { transparency: 100 } });
  addImageFit(slide, assets.campusLine, { x: 10.14, y: 0.36, w: 2.42, h: 0.86 }, { transparency: 4 });
  slide.addImage({ path: assets.seal, x: 12.46, y: 6.72, w: 0.32, h: 0.32, transparency: 4 });
  addText(slide, String(pageNo).padStart(2, "0"), 11.94, 6.77, 0.38, 0.18, {
    fontSize: 8,
    color: theme.muted,
    align: "right",
  });
  slide.addShape("line", { x: 0.72, y: 6.95, w: 10.95, h: 0, line: { color: theme.line, width: 0.75 } });
}

function addTitle(slide, title, subtitle) {
  addText(slide, title, 0.72, 0.64, 8.4, 0.44, {
    fontSize: 18,
    bold: true,
    color: theme.green,
  });
  if (subtitle) {
    addText(slide, subtitle, 0.74, 1.04, 8.8, 0.24, {
      fontSize: 8.5,
      color: theme.muted,
      fontFace: font.en,
    });
  }
  slide.addShape("line", { x: 0.72, y: 1.34, w: 3.8, h: 0, line: { color: theme.green, width: 1.0 } });
}

function addBulletList(slide, bullets, x, y, w, gap = 0.58, options = {}, ctx = null) {
  const size = options.fontSize || 16;
  bullets.forEach((item, idx) => {
    const top = y + idx * gap;
    slide.addShape("ellipse", {
      x,
      y: top + 0.12,
      w: 0.12,
      h: 0.12,
      fill: { color: options.color || theme.green },
      line: { transparency: 100 },
    });
    const textOptions = {
      fontSize: size,
      color: options.textColor || theme.ink,
      valign: "top",
      fit: "shrink",
    };
    if (ctx) addInlineMathText(slide, item, x + 0.28, top, w - 0.28, gap - 0.03, textOptions, ctx);
    else addText(slide, item, x + 0.28, top, w - 0.28, gap - 0.03, textOptions);
  });
}

function addFlowBulletList(slide, bullets, x, y, w, maxH, options = {}, ctx = null) {
  const size = options.fontSize || 16;
  let top = y;
  bullets.forEach((item) => {
    const itemH = Math.min(maxH - (top - y), estimateBulletHeight(item, w, size));
    if (itemH <= 0.12) return;
    slide.addShape("ellipse", {
      x,
      y: top + 0.13,
      w: 0.12,
      h: 0.12,
      fill: { color: options.color || theme.green },
      line: { transparency: 100 },
    });
    const textOptions = {
      fontSize: size,
      color: options.textColor || theme.ink,
      valign: "top",
      fit: "shrink",
    };
    if (ctx) addInlineMathText(slide, item, x + 0.28, top, w - 0.28, itemH, textOptions, ctx);
    else addText(slide, item, x + 0.28, top, w - 0.28, itemH, textOptions);
    top += itemH;
  });
}

function addCompactBulletList(slide, bullets, x, y, w, options = {}, ctx = null) {
  const size = options.fontSize || 10;
  const gap = options.gap || 0.28;
  const dot = options.dotSize || 0.09;
  bullets.forEach((item, idx) => {
    const top = y + idx * gap;
    slide.addShape("ellipse", {
      x,
      y: top + 0.1,
      w: dot,
      h: dot,
      fill: { color: options.color || theme.green },
      line: { transparency: 100 },
    });
    const textOptions = {
      fontSize: size,
      color: options.textColor || theme.ink,
      valign: "mid",
      fit: "shrink",
    };
    if (ctx) addInlineMathText(slide, item, x + 0.26, top, w - 0.26, 0.22, textOptions, ctx);
    else addText(slide, item, x + 0.26, top, w - 0.26, 0.22, textOptions);
  });
}

function layoutTitle(pptx, deck) {
  const slide = pptx.addSlide();
  addCoverBrand(slide);
  addText(slide, deck.meta?.title, 0.78, 2.08, 7.75, 1.16, {
    fontSize: 30,
    bold: true,
    color: theme.white,
    valign: "top",
  });
  addText(slide, deck.meta?.subtitle, 0.82, 3.48, 7.2, 0.54, {
    fontSize: 15,
    color: theme.white,
    transparency: 7,
  });
  const info = [deck.meta?.author, deck.meta?.advisor, deck.meta?.date].filter(Boolean).join("  |  ");
  addText(slide, info, 0.82, 6.42, 7.3, 0.28, {
    fontSize: 11,
    color: theme.white,
  });
}

function layoutAgenda(pptx, slideData, pageNo) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title || "目录", "CONTENTS");
  const items = slideData.items || [];
  items.forEach((item, idx) => {
    const y = 1.78 + idx * 0.78;
    addText(slide, String(idx + 1).padStart(2, "0"), 1.08, y, 0.64, 0.4, {
      fontSize: 18,
      bold: true,
      color: theme.green,
      fontFace: font.en,
    });
    slide.addShape("line", { x: 1.82, y: y + 0.22, w: 0.75, h: 0, line: { color: theme.green, width: 1.3 } });
    addText(slide, item, 2.78, y - 0.03, 7.4, 0.48, {
      fontSize: 19,
      bold: true,
      color: theme.ink,
    });
  });
}

function layoutSection(pptx, slideData, pageNo) {
  const slide = pptx.addSlide();
  slide.background = { color: theme.white };
  slide.addShape("rect", { x: 0, y: 0, w: W, h: H, fill: { color: theme.green }, line: { transparency: 100 } });
  slide.addShape("rect", { x: 0.52, y: 0.44, w: 12.28, h: 6.62, fill: { color: theme.white, transparency: 1 }, line: { color: "DCE9E2", width: 0.8 } });
  addImageFit(slide, assets.emblemGray, { x: 8.75, y: 1.08, w: 3.72, h: 3.72 }, { transparency: 74 });
  addText(slide, slideData.sectionNo || String(pageNo).padStart(2, "0"), 1.1, 1.7, 1.9, 0.72, {
    fontSize: 34,
    bold: true,
    color: theme.green,
    fontFace: font.en,
  });
  addText(slide, slideData.title, 1.08, 2.62, 8.2, 0.64, {
    fontSize: 26,
    bold: true,
    color: theme.ink,
  });
  addText(slide, slideData.subtitle || "", 1.12, 3.38, 7.6, 0.42, {
    fontSize: 13,
    color: theme.muted,
  });
  slide.addShape("line", { x: 1.1, y: 4.18, w: 3.1, h: 0, line: { color: theme.green, width: 2.3 } });
}

function layoutBullets(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title, "KEY FINDINGS");
  const startY = slideData.lead ? 2.92 : 1.72;
  if (slideData.lead) {
    slide.addShape("rect", { x: 0.92, y: 1.72, w: 11.35, h: 0.72, fill: { color: "FAFCFB" }, line: { color: theme.line, width: 0.6 } });
    addInlineMathText(slide, slideData.lead, 1.18, 1.86, 10.8, 0.44, { fontSize: 15, color: theme.ink }, ctx);
  }
  addFlowBulletList(slide, slideData.bullets || [], 1.28, startY, 10.4, 6.62 - startY, { fontSize: slideData.fontSize || 16 }, ctx);
}

function layoutTable(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title, "DATA TABLE");
  const columns = slideData.columns || [];
  const rows = slideData.rows || [];
  const table = [
    columns.map((text) => {
      const options = {
        bold: true,
        color: theme.white,
        fill: { color: theme.green },
        fontFace: font.cn,
      };
      return { text: inlineMathTableCell(text, options, ctx), options };
    }),
    ...rows.map((row, rIdx) =>
      row.map((text) => {
        const options = {
          color: theme.ink,
          fill: { color: rIdx % 2 === 0 ? "FFFFFF" : "F7FAF8" },
          fontFace: font.cn,
        };
        return { text: inlineMathTableCell(text, options, ctx), options };
      }),
    ),
  ];
  slide.addTable(table, {
    x: 0.76,
    y: 1.68,
    w: 11.82,
    h: Math.min(4.95, 0.52 + rows.length * 0.68),
    border: { type: "solid", color: "C9D8D0", pt: 0.6 },
    margin: 0.08,
    fontSize: rows.length > 5 ? 8.5 : 10,
    valign: "mid",
    fit: "shrink",
    rowH: 0.52,
  });
}

function layoutClaim(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title, "MAIN CLAIM");
  addInlineMathText(slide, slideData.claim, 1.02, 1.78, 10.6, 1.16, {
    fontSize: 26,
    bold: true,
    color: theme.green,
    valign: "mid",
  }, ctx);
  slide.addShape("line", { x: 1.04, y: 3.12, w: 4.1, h: 0, line: { color: theme.green, width: 1.3 } });
  addFlowBulletList(slide, slideData.evidence || [], 1.3, 3.58, 10.2, 2.58, { fontSize: 15 }, ctx);
}

function layoutTwoColumn(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title, "TWO COLUMNS");
  addColumnBlock(slide, slideData.left || {}, 0.9, 1.72, 5.35, 4.55, theme.green, ctx);
  addColumnBlock(slide, slideData.right || {}, 7.06, 1.72, 5.35, 4.55, theme.red, ctx);
  slide.addShape("line", { x: 6.67, y: 1.88, w: 0, h: 4.2, line: { color: theme.line, width: 1.1 } });
}

function addColumnBlock(slide, data, x, y, w, h, accent, ctx) {
  addText(slide, data.title || "", x, y, w, 0.36, { fontSize: 17, bold: true, color: accent });
  slide.addShape("line", { x, y: y + 0.5, w: 1.7, h: 0, line: { color: accent, width: 1.2 } });
  if (data.text) {
    addInlineMathText(slide, data.text, x, y + 0.82, w, h - 0.82, { fontSize: 13.5, color: theme.ink, valign: "top" }, ctx);
  } else {
    addFlowBulletList(slide, data.bullets || [], x + 0.12, y + 0.82, w - 0.18, h - 0.82, { fontSize: 13.2, color: accent }, ctx);
  }
}

function layoutCards(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title, "KEY POINTS");
  const cards = (slideData.cards || []).slice(0, 6);
  const cols = cards.length <= 3 ? cards.length || 1 : 3;
  const rows = Math.ceil(cards.length / cols);
  const gap = 0.25;
  const totalW = 11.7;
  const cardW = (totalW - gap * (cols - 1)) / cols;
  const cardH = rows === 1 ? 3.6 : 2.05;
  cards.forEach((card, idx) => {
    const col = idx % cols;
    const row = Math.floor(idx / cols);
    const x = 0.82 + col * (cardW + gap);
    const y = 1.78 + row * (cardH + 0.32);
    slide.addShape("rect", { x, y, w: cardW, h: cardH, fill: { color: idx % 2 ? "F9FBFA" : "FFFFFF" }, line: { color: theme.line, width: 0.8 } });
    addText(slide, card.title || `观点 ${idx + 1}`, x + 0.22, y + 0.2, cardW - 0.44, 0.36, { fontSize: 15, bold: true, color: theme.green });
    addInlineMathText(slide, card.text || "", x + 0.22, y + 0.76, cardW - 0.44, cardH - 0.98, { fontSize: rows === 1 ? 12.5 : 10.8, color: theme.ink, valign: "top" }, ctx);
  });
}

function layoutTimeline(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title, "TIMELINE");
  const items = (slideData.items || []).slice(0, 6);
  const x0 = 1.05;
  const y = 3.2;
  const step = items.length > 1 ? 10.8 / (items.length - 1) : 0;
  slide.addShape("line", { x: x0, y, w: 10.8, h: 0, line: { color: theme.green, width: 2.0 } });
  items.forEach((item, idx) => {
    const x = x0 + idx * step;
    slide.addShape("ellipse", { x: x - 0.12, y: y - 0.12, w: 0.24, h: 0.24, fill: { color: theme.green }, line: { color: theme.white, width: 1 } });
    addText(slide, item.date || item.phase || String(idx + 1), x - 0.68, y - 0.72, 1.36, 0.26, { fontSize: 9, color: theme.green, bold: true, align: "center" });
    addText(slide, item.title || "", x - 0.78, y + 0.32, 1.56, 0.36, { fontSize: 11.5, color: theme.ink, bold: true, align: "center" });
    addInlineMathText(slide, item.text || "", x - 0.86, y + 0.82, 1.72, 1.05, { fontSize: 8.5, color: theme.muted, align: "center", valign: "top" }, ctx);
  });
}

function layoutMetrics(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title, "METRICS");
  const metrics = (slideData.metrics || []).slice(0, 4);
  const cardW = 2.72;
  metrics.forEach((metric, idx) => {
    const x = 1.02 + idx * 3.02;
    slide.addShape("rect", { x, y: 2.05, w: cardW, h: 2.35, fill: { color: idx % 2 ? "F9FBFA" : "FFFFFF" }, line: { color: theme.line, width: 0.8 } });
    addText(slide, metric.value || "", x + 0.18, 2.42, cardW - 0.36, 0.72, { fontSize: 28, bold: true, color: theme.green, align: "center", fontFace: font.en });
    addText(slide, metric.label || "", x + 0.22, 3.24, cardW - 0.44, 0.34, { fontSize: 13.5, bold: true, color: theme.ink, align: "center" });
    addInlineMathText(slide, metric.note || "", x + 0.22, 3.75, cardW - 0.44, 0.38, { fontSize: 9.5, color: theme.muted, align: "center" }, ctx);
  });
}

function layoutQuote(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title || "观点引用", "QUOTE");
  addText(slide, "“", 1.02, 1.72, 0.7, 0.65, { fontSize: 38, color: theme.green, fontFace: font.serif });
  addInlineMathText(slide, slideData.quote, 1.52, 2.08, 9.95, 2.15, { fontSize: 22, color: theme.ink, fontFace: font.serif, valign: "mid" }, ctx);
  slide.addShape("line", { x: 1.54, y: 4.56, w: 2.8, h: 0, line: { color: theme.green, width: 1.2 } });
  addText(slide, slideData.source || "", 1.54, 4.82, 7.6, 0.38, { fontSize: 12, color: theme.muted });
}

function layoutFormula(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title || "公式", "FORMULA");
  const latex = slideData.formula?.latex || slideData.formula || "";
  slide.addShape("rect", { x: 1.0, y: 1.82, w: 11.35, h: 1.85, fill: { color: "FAFCFB" }, line: { color: theme.line, width: 0.7 } });
  if (latex) {
    addText(slide, registerEquation(ctx, latex, { display: true }), 1.28, 2.26, 10.8, 0.9, {
      fontSize: 24,
      color: theme.ink,
      align: "center",
      fontFace: font.en,
    });
  }
  if (slideData.caption) {
    addText(slide, slideData.caption, 1.08, 3.82, 11.05, 0.32, { fontSize: 11, color: theme.muted, align: "center" });
  }
  const explanation = slideData.explanation || slideData.notes || [];
  explanation.forEach((item, idx) => {
    const y = 4.55 + idx * 0.48;
    slide.addShape("ellipse", { x: 1.28, y: y + 0.13, w: 0.12, h: 0.12, fill: { color: theme.green }, line: { transparency: 100 } });
    addInlineMathText(slide, item, 1.56, y, 10.08, 0.38, { fontSize: 13.2, color: theme.ink, valign: "top" }, ctx);
  });
}

function layoutReferences(pptx, slideData, pageNo) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title || "参考文献", "REFERENCES");
  const fontSize = slideData.fontSize || 9.5;
  let y = 1.72;
  (slideData.items || []).forEach((item, idx) => {
    const h = estimateText(item, 10.75, fontSize, { lineHeight: 1.18 }).height + 0.11;
    addText(slide, `[${idx + 1}]`, 0.9, y, 0.52, 0.18, { fontSize, color: theme.green, fontFace: font.en, bold: true });
    addText(slide, item, 1.48, y, 10.8, h, { fontSize, color: theme.ink, valign: "top" });
    y += h;
  });
}

function layoutMatrix(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title, "MATRIX");
  const cells = slideData.cells || [];
  const x0 = 1.28;
  const y0 = 1.78;
  const cellW = 5.05;
  const cellH = 2.1;
  for (let row = 0; row < 2; row++) {
    for (let col = 0; col < 2; col++) {
      const cell = cells[row * 2 + col] || {};
      const x = x0 + col * (cellW + 0.48);
      const y = y0 + row * (cellH + 0.34);
      slide.addShape("rect", { x, y, w: cellW, h: cellH, fill: { color: row === col ? "F1F8F4" : "FFFFFF" }, line: { color: theme.line, width: 0.8 } });
      addText(slide, cell.title || "", x + 0.22, y + 0.22, cellW - 0.44, 0.32, { fontSize: 14.5, bold: true, color: theme.green });
      addInlineMathText(slide, cell.text || "", x + 0.22, y + 0.72, cellW - 0.44, cellH - 0.94, { fontSize: 11, color: theme.ink, valign: "top" }, ctx);
    }
  }
}

function layoutProcess(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title, "PROCESS");
  const steps = (slideData.steps || []).slice(0, 5);
  const stepW = 2.12;
  steps.forEach((step, idx) => {
    const x = 0.95 + idx * 2.38;
    slide.addShape("rect", { x, y: 2.28, w: stepW, h: 1.92, fill: { color: "FFFFFF" }, line: { color: theme.green, width: 1.1 } });
    addText(slide, String(idx + 1).padStart(2, "0"), x + 0.16, 2.48, 0.46, 0.22, { fontSize: 10, color: theme.green, bold: true, fontFace: font.en });
    addText(slide, step.title || "", x + 0.24, 2.88, stepW - 0.48, 0.32, { fontSize: 13, bold: true, color: theme.ink, align: "center" });
    addInlineMathText(slide, step.text || "", x + 0.18, 3.35, stepW - 0.36, 0.54, { fontSize: 8.8, color: theme.muted, align: "center", valign: "top" }, ctx);
    if (idx < steps.length - 1) {
      slide.addShape("chevron", { x: x + stepW + 0.14, y: 3.05, w: 0.35, h: 0.35, fill: { color: theme.green }, line: { transparency: 100 } });
    }
  });
}

function layoutProblemSolution(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title || "问题与方案", "PROBLEM / SOLUTION");
  const blocks = [
    { key: "problem", label: "问题", color: theme.red, x: 0.92 },
    { key: "solution", label: "方案", color: theme.green, x: 4.78 },
    { key: "impact", label: "收益", color: theme.darkGreen, x: 8.64 },
  ];
  blocks.forEach((block) => {
    const data = slideData[block.key] || {};
    slide.addShape("rect", { x: block.x, y: 1.78, w: 3.44, h: 4.48, fill: { color: "FFFFFF" }, line: { color: block.color, width: 1.0 } });
    slide.addShape("rect", { x: block.x, y: 1.78, w: 3.44, h: 0.48, fill: { color: block.color }, line: { transparency: 100 } });
    addText(slide, data.label || block.label, block.x + 0.22, 1.92, 2.9, 0.16, { fontSize: 9.5, bold: true, color: theme.white, fontFace: font.en });
    addText(slide, data.title || "", block.x + 0.25, 2.58, 2.94, 0.46, { fontSize: 16, bold: true, color: block.color });
    addFlowBulletList(slide, data.bullets || [], block.x + 0.28, 3.28, 2.88, 2.42, { fontSize: 10.8, color: block.color }, ctx);
  });
}

function layoutPainOpportunity(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title || "现状痛点与机会", "PAIN / OPPORTUNITY");
  addColumnBlock(slide, slideData.status || { title: "现状" }, 0.92, 1.72, 3.45, 4.45, theme.muted, ctx);
  addColumnBlock(slide, slideData.pain || { title: "痛点" }, 4.92, 1.72, 3.45, 4.45, theme.red, ctx);
  addColumnBlock(slide, slideData.opportunity || { title: "机会" }, 8.92, 1.72, 3.45, 4.45, theme.green, ctx);
}

function layoutExperimentDesign(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title || "实验设计", "EXPERIMENT DESIGN");
  const fields = [
    ["数据集", slideData.dataset],
    ["变量", slideData.variables],
    ["指标", slideData.metrics],
    ["对照", slideData.baselines],
  ];
  fields.forEach(([label, value], idx) => {
    const col = idx % 2;
    const row = Math.floor(idx / 2);
    const x = 0.92 + col * 5.95;
    const y = 1.78 + row * 2.18;
    slide.addShape("rect", { x, y, w: 5.35, h: 1.74, fill: { color: idx % 2 ? "F9FBFA" : "FFFFFF" }, line: { color: theme.line, width: 0.8 } });
    addText(slide, label, x + 0.24, y + 0.18, 1.18, 0.28, { fontSize: 13.5, bold: true, color: theme.green });
    const items = Array.isArray(value) ? value : value ? [value] : [];
    addFlowBulletList(slide, items, x + 0.28, y + 0.62, 4.78, 0.92, { fontSize: 10.8 }, ctx);
  });
  if (slideData.procedure?.length) {
    addText(slide, "流程", 0.94, 6.0, 0.8, 0.24, { fontSize: 12.5, bold: true, color: theme.green });
    addText(slide, slideData.procedure.join("  ->  "), 1.72, 6.0, 10.3, 0.24, { fontSize: 10.5, color: theme.muted });
  }
}

function layoutResultAnalysis(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title || "结果分析", "RESULT ANALYSIS");
  addInlineMathText(slide, slideData.finding || "", 0.98, 1.72, 10.9, 0.78, { fontSize: 22, bold: true, color: theme.green, valign: "mid" }, ctx);
  const metrics = (slideData.metrics || []).slice(0, 3);
  metrics.forEach((metric, idx) => {
    const x = 1.0 + idx * 3.9;
    slide.addShape("rect", { x, y: 2.88, w: 3.35, h: 1.22, fill: { color: "FAFCFB" }, line: { color: theme.line, width: 0.7 } });
    addText(slide, metric.value || "", x + 0.18, 3.06, 1.2, 0.42, { fontSize: 22, bold: true, color: theme.green, fontFace: font.en });
    addText(slide, metric.label || "", x + 1.48, 3.08, 1.64, 0.24, { fontSize: 11.5, bold: true, color: theme.ink });
    addInlineMathText(slide, metric.note || "", x + 1.48, 3.45, 1.65, 0.28, { fontSize: 8.8, color: theme.muted }, ctx);
  });
  addFlowBulletList(slide, slideData.analysis || [], 1.18, 4.55, 10.4, 1.35, { fontSize: 13.2 }, ctx);
}

function layoutRiskMitigation(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title || "风险与对策", "RISK MITIGATION");
  const rows = (slideData.items || []).slice(0, 5);
  const table = [
    ["风险", "影响", "对策"].map((text) => ({ text, options: { bold: true, color: theme.white, fill: { color: theme.green }, fontFace: font.cn } })),
    ...rows.map((item, idx) => ["risk", "impact", "mitigation"].map((key) => {
      const options = { color: theme.ink, fill: { color: idx % 2 === 0 ? "FFFFFF" : "F7FAF8" }, fontFace: font.cn };
      return { text: inlineMathTableCell(item[key] || "", options, ctx), options };
    })),
  ];
  slide.addTable(table, { x: 0.86, y: 1.78, w: 11.62, h: 4.45, border: { type: "solid", color: "C9D8D0", pt: 0.6 }, margin: 0.08, fontSize: 10, valign: "mid", fit: "shrink" });
}

function layoutContribution(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title || "主要贡献", "CONTRIBUTIONS");
  const items = (slideData.items || []).slice(0, 4);
  items.forEach((item, idx) => {
    const y = 1.8 + idx * 1.08;
    addText(slide, String(idx + 1).padStart(2, "0"), 1.04, y, 0.62, 0.36, { fontSize: 18, bold: true, color: theme.green, fontFace: font.en });
    slide.addShape("line", { x: 1.84, y: y + 0.18, w: 0.62, h: 0, line: { color: theme.green, width: 1.1 } });
    addText(slide, item.title || "", 2.68, y - 0.03, 3.0, 0.32, { fontSize: 15, bold: true, color: theme.ink });
    addInlineMathText(slide, item.text || "", 5.86, y - 0.03, 5.9, 0.42, { fontSize: 12, color: theme.muted, valign: "top" }, ctx);
  });
}

function layoutSummary(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title || "章节小结", "SUMMARY");
  addInlineMathText(slide, slideData.takeaway || "", 1.02, 1.78, 10.9, 0.84, { fontSize: 23, bold: true, color: theme.green, valign: "mid" }, ctx);
  addFlowBulletList(slide, slideData.points || [], 1.28, 3.08, 10.4, 2.55, { fontSize: 15 }, ctx);
}

function layoutArchitecture(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title || "方法架构", "ARCHITECTURE");
  const layers = (slideData.layers || []).slice(0, 4);
  const layerH = layers.length <= 3 ? 1.28 : 1.04;
  const gap = layers.length <= 3 ? 0.26 : 0.18;
  layers.forEach((layer, idx) => {
    const y = 1.74 + idx * (layerH + gap);
    const accent = idx % 2 === 0 ? theme.green : theme.darkGreen;
    slide.addShape("rect", { x: 0.92, y, w: 11.45, h: layerH, fill: { color: idx % 2 === 0 ? "F8FBF9" : "FFFFFF" }, line: { color: theme.line, width: 0.8 } });
    slide.addShape("rect", { x: 0.92, y, w: 1.7, h: layerH, fill: { color: accent }, line: { transparency: 100 } });
    addText(slide, layer.title || `Layer ${idx + 1}`, 1.08, y + 0.18, 1.38, 0.42, { fontSize: 12.2, bold: true, color: theme.white, align: "center" });
    const components = (layer.components || []).slice(0, 5);
    const chipW = Math.min(1.62, 8.65 / Math.max(1, components.length));
    components.forEach((component, cIdx) => {
      const x = 2.94 + cIdx * (chipW + 0.18);
      slide.addShape("roundRect", { x, y: y + 0.25, w: chipW, h: 0.44, rectRadius: 0.04, fill: { color: "FFFFFF" }, line: { color: accent, width: 0.75 } });
      addInlineMathText(slide, component, x + 0.1, y + 0.36, chipW - 0.2, 0.12, { fontSize: 8.7, color: accent, bold: true, align: "center" }, ctx);
      if (cIdx < components.length - 1) {
        slide.addShape("chevron", { x: x + chipW + 0.04, y: y + 0.35, w: 0.16, h: 0.16, fill: { color: theme.line }, line: { transparency: 100 } });
      }
    });
    if (layer.note) {
      addInlineMathText(slide, layer.note, 2.94, y + 0.74, 8.75, 0.38, { fontSize: 10.2, color: theme.muted, valign: "mid" }, ctx);
    }
  });
}

function layoutAblation(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title || "消融实验", "ABLATION STUDY");
  if (slideData.baseline) {
    addInlineMathText(slide, slideData.baseline, 0.98, 1.64, 10.95, 0.34, { fontSize: 12.5, color: theme.muted }, ctx);
  }
  const rows = (slideData.items || []).slice(0, 6);
  const headers = ["模块", "设置", "结果变化", "结论"];
  const table = [
    headers.map((text) => ({ text, options: { bold: true, color: theme.white, fill: { color: theme.green }, fontFace: font.cn } })),
    ...rows.map((item, idx) => ["factor", "setting", "delta", "conclusion"].map((key) => {
      const options = { color: theme.ink, fill: { color: idx % 2 === 0 ? "FFFFFF" : "F7FAF8" }, fontFace: font.cn };
      return { text: inlineMathTableCell(item[key] || "", options, ctx), options };
    })),
  ];
  slide.addTable(table, {
    x: 0.82,
    y: 2.12,
    w: 11.74,
    h: Math.min(4.16, 0.52 + rows.length * 0.58),
    border: { type: "solid", color: "C9D8D0", pt: 0.6 },
    margin: 0.08,
    fontSize: rows.length > 4 ? 8.8 : 9.6,
    valign: "mid",
    fit: "shrink",
    colW: [2.0, 3.05, 2.05, 4.64],
  });
}

function layoutCaseStudy(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title || "案例分析", "CASE STUDY");
  const image = resolveImageInfo(slideData.image || "assets/bit-campus-photo.png");
  image.fit = normalizeText(slideData.imageFit || slideData.fit || image.fit).toLowerCase();
  addImageInPanel(slide, image, { x: 1.02, y: 1.84, w: 4.22, h: 3.16 }, { fit: "cover" });
  addInlineMathText(slide, slideData.caption || "", 1.02, 5.16, 4.22, 0.36, { fontSize: 9.5, color: theme.muted, align: "center" }, ctx);
  const blocks = [
    ["背景", slideData.context],
    ["做法", slideData.method],
    ["结果", slideData.result],
  ];
  blocks.forEach(([label, value], idx) => {
    const y = 1.76 + idx * 1.46;
    slide.addShape("rect", { x: 6.02, y, w: 5.85, h: 1.1, fill: { color: idx === 2 ? "F1F8F4" : "FFFFFF" }, line: { color: theme.line, width: 0.75 } });
    addText(slide, label, 6.26, y + 0.18, 0.74, 0.22, { fontSize: 12, bold: true, color: theme.green });
    const items = Array.isArray(value) ? value : value ? [value] : [];
    addCompactBulletList(slide, items, 7.16, y + 0.2, 4.34, { fontSize: 9.8, gap: 0.3, dotSize: 0.08 }, ctx);
  });
}

function layoutImageGrid(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title || "图像结果", "IMAGE GRID");
  const images = (slideData.images || []).slice(0, 6);
  const cols = images.length <= 2 ? images.length || 1 : 3;
  const rows = Math.ceil(images.length / cols);
  const gap = 0.24;
  const gridW = 11.55;
  const cellW = (gridW - gap * (cols - 1)) / cols;
  const cellH = rows === 1 ? 3.9 : 2.02;
  images.forEach((item, idx) => {
    const col = idx % cols;
    const row = Math.floor(idx / cols);
    const x = 0.9 + col * (cellW + gap);
    const y = 1.74 + row * (cellH + 0.36);
    const image = resolveImageInfo(item || "assets/bit-campus-photo.png");
    const caption = item && typeof item === "object" ? item.caption : "";
    addImageInPanel(slide, image, { x: x + 0.1, y: y + 0.1, w: cellW - 0.2, h: cellH - 0.48 }, { fit: "cover", lineWidth: 0.75, pad: 0.1 });
    addInlineMathText(slide, caption, x + 0.12, y + cellH - 0.32, cellW - 0.24, 0.18, { fontSize: 8.4, color: theme.muted, align: "center" }, ctx);
  });
}

function layoutCode(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title || "算法与代码", "CODE / ALGORITHM");
  const code = normalizeText(slideData.code || slideData.algorithm || "");
  slide.addShape("rect", { x: 0.88, y: 1.72, w: 7.15, h: 4.62, fill: { color: "F6F8F6" }, line: { color: theme.line, width: 0.8 } });
  slide.addShape("rect", { x: 0.88, y: 1.72, w: 7.15, h: 0.34, fill: { color: theme.green }, line: { transparency: 100 } });
  addText(slide, slideData.language || "pseudo", 1.08, 1.82, 1.8, 0.1, { fontSize: 7.5, color: theme.white, fontFace: font.en, bold: true });
  addText(slide, code, 1.08, 2.22, 6.75, 3.78, { fontSize: slideData.fontSize || 9.2, color: theme.ink, fontFace: font.code, valign: "top", fit: "shrink" });
  addColumnBlock(slide, { title: slideData.noteTitle || "解释", bullets: slideData.notes || [] }, 8.64, 1.78, 3.56, 4.42, theme.green, ctx);
}

function layoutAppendix(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title || "附录", "APPENDIX");
  const items = (slideData.items || []).slice(0, 8);
  const cols = items.length > 4 ? 2 : 1;
  const colW = cols === 2 ? 5.1 : 10.6;
  items.forEach((item, idx) => {
    const col = cols === 2 ? idx % 2 : 0;
    const row = cols === 2 ? Math.floor(idx / 2) : idx;
    const x = 1.0 + col * 5.82;
    const y = 1.82 + row * 0.88;
    addText(slide, item.key || String(idx + 1).padStart(2, "0"), x, y, 0.62, 0.24, { fontSize: 12, bold: true, color: theme.green, fontFace: font.en, align: "right" });
    slide.addShape("line", { x: x + 0.78, y: y + 0.12, w: 0.5, h: 0, line: { color: theme.green, width: 1.0 } });
    addText(slide, item.title || "", x + 1.48, y - 0.02, colW - 1.48, 0.24, { fontSize: 12.5, bold: true, color: theme.ink });
    addInlineMathText(slide, item.text || "", x + 1.48, y + 0.32, colW - 1.48, 0.24, { fontSize: 8.8, color: theme.muted }, ctx);
  });
}

function resolveFlowNodePositions(nodes, box) {
  const count = nodes.length;
  const cols = count <= 6 ? count || 1 : 4;
  const rows = Math.ceil(count / cols);
  const nodeW = count <= 3 ? 2.28 : count <= 6 ? 1.62 : 1.76;
  const nodeH = 0.74;
  const gapX = cols > 1 ? (box.w - nodeW * cols) / (cols - 1) : 0;
  const gapY = rows > 1 ? Math.min(1.1, (box.h - nodeH * rows) / (rows - 1)) : 0;
  const startY = box.y + (box.h - rows * nodeH - (rows - 1) * gapY) / 2;
  return nodes.map((node, idx) => {
    if (Number.isFinite(node.x) && Number.isFinite(node.y)) {
      return { ...node, x: box.x + node.x, y: box.y + node.y, w: node.w || nodeW, h: node.h || nodeH };
    }
    const row = Math.floor(idx / cols);
    const logicalCol = idx % cols;
    const itemsInRow = row === rows - 1 ? count - row * cols : cols;
    const col = row % 2 === 1 ? itemsInRow - 1 - logicalCol : logicalCol;
    const rowW = itemsInRow * nodeW + Math.max(0, itemsInRow - 1) * gapX;
    const x0 = box.x + (box.w - rowW) / 2;
    return {
      ...node,
      x: x0 + col * (nodeW + gapX),
      y: startY + row * (nodeH + gapY),
      w: node.w || nodeW,
      h: node.h || nodeH,
    };
  });
}

function drawFlowEdge(slide, from, to, options = {}) {
  const x1 = from.x + from.w / 2;
  const y1 = from.y + from.h / 2;
  const x2 = to.x + to.w / 2;
  const y2 = to.y + to.h / 2;
  const dx = x2 - x1;
  const dy = y2 - y1;
  const fromHalfW = from.w / 2;
  const fromHalfH = from.h / 2;
  const toHalfW = to.w / 2;
  const toHalfH = to.h / 2;
  const scaleFrom = Math.min(Math.abs(dx) > 0 ? fromHalfW / Math.abs(dx) : Infinity, Math.abs(dy) > 0 ? fromHalfH / Math.abs(dy) : Infinity);
  const scaleTo = Math.min(Math.abs(dx) > 0 ? toHalfW / Math.abs(dx) : Infinity, Math.abs(dy) > 0 ? toHalfH / Math.abs(dy) : Infinity);
  const sx = x1 + dx * Math.min(0.48, scaleFrom || 0);
  const sy = y1 + dy * Math.min(0.48, scaleFrom || 0);
  const ex = x2 - dx * Math.min(0.48, scaleTo || 0);
  const ey = y2 - dy * Math.min(0.48, scaleTo || 0);
  slide.addShape("line", {
    x: sx,
    y: sy,
    w: ex - sx,
    h: ey - sy,
    line: {
      color: options.color || theme.green,
      width: options.width || 1.15,
      endArrowType: options.arrow === false ? "none" : "triangle",
      transparency: options.transparency || 8,
    },
  });
}

function layoutFlowchart(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title || "流程图", "FLOWCHART");
  const nodes = resolveFlowNodePositions((slideData.nodes || []).slice(0, 10), { x: 0.98, y: 1.7, w: 11.35, h: 4.48 });
  const byId = new Map(nodes.map((node, idx) => [node.id || String(idx + 1), node]));
  const edges = slideData.edges?.length
    ? slideData.edges
    : nodes.slice(0, -1).map((node, idx) => ({ from: node.id || String(idx + 1), to: nodes[idx + 1].id || String(idx + 2) }));
  edges.forEach((edge) => {
    const from = byId.get(edge.from);
    const to = byId.get(edge.to);
    if (from && to) drawFlowEdge(slide, from, to, edge);
  });
  nodes.forEach((node, idx) => {
    const accent = node.color || (idx % 2 === 0 ? theme.green : theme.darkGreen);
    const fill = node.fill || (node.emphasis ? "F1F8F4" : "FFFFFF");
    slide.addShape(node.shape || "roundRect", {
      x: node.x,
      y: node.y,
      w: node.w,
      h: node.h,
      rectRadius: 0.05,
      fill: { color: fill },
      line: { color: accent, width: node.emphasis ? 1.35 : 0.95 },
    });
    addInlineMathText(slide, node.text || node.label || node.id || "", node.x + 0.16, node.y + 0.15, node.w - 0.32, 0.28, {
      fontSize: node.fontSize || 11.2,
      color: accent,
      bold: true,
      align: "center",
    }, ctx);
    if (node.note) {
      addInlineMathText(slide, node.note, node.x + 0.18, node.y + 0.46, node.w - 0.36, 0.18, {
        fontSize: 7.5,
        color: theme.muted,
        align: "center",
      }, ctx);
    }
  });
  if (slideData.note) {
    addInlineMathText(slide, slideData.note, 1.02, 6.02, 10.9, 0.28, { fontSize: 10, color: theme.muted, align: "center" }, ctx);
  }
}

function chartTypeName(type) {
  const value = normalizeText(type || "bar").toLowerCase();
  if (["line", "pie", "doughnut", "scatter", "area"].includes(value)) return value;
  return "bar";
}

function normalizeChartSeries(slideData) {
  const categories = slideData.categories || slideData.labels || [];
  const series = slideData.series?.length ? slideData.series : [{ name: slideData.name || "Series", values: slideData.values || [] }];
  return series.map((item) => ({
    name: item.name || item.label || "Series",
    labels: item.labels || categories,
    values: (item.values || []).map((value) => Number(value)),
  }));
}

function layoutChart(pptx, slideData, pageNo) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  const type = chartTypeName(slideData.type);
  addTitle(slide, slideData.title || "统计图", type.toUpperCase());
  const data = normalizeChartSeries(slideData);
  const showLegend = slideData.showLegend ?? (data.length > 1 || ["pie", "doughnut"].includes(type));
  const chartOptions = {
    x: 0.92,
    y: 1.72,
    w: 10.95,
    h: 4.45,
    chartColors: slideData.colors || [theme.green, theme.red, "577A68", "9DAA57", "4B647A", "B56B45"],
    showLegend,
    legendPos: slideData.legendPos || "b",
    showValue: slideData.showValue ?? ["pie", "doughnut"].includes(type),
    showCategoryName: slideData.showCategoryName ?? false,
    showTitle: false,
    showCatName: false,
    valAxisLabelFontFace: font.en,
    valAxisLabelFontSize: 8,
    valAxisLabelColor: theme.muted,
    catAxisLabelFontFace: font.cn,
    catAxisLabelFontSize: 8,
    catAxisLabelColor: theme.muted,
    catAxisLabelRotate: slideData.rotateLabels || 0,
    catAxisTitle: slideData.categoryAxisTitle,
    valAxisTitle: slideData.valueAxisTitle,
    valGridLine: { color: "DDE7E2", size: 0.5 },
    catAxisLineColor: "C9D8D0",
    valAxisLineColor: "C9D8D0",
    chartArea: { border: { color: "FFFFFF", pt: 0 }, roundedCorners: false },
    plotArea: { border: { color: "FFFFFF", pt: 0 }, fill: { color: "FFFFFF", transparency: 100 } },
  };
  if (type === "bar") {
    chartOptions.barDir = slideData.direction === "horizontal" ? "bar" : "col";
    chartOptions.barGrouping = slideData.grouping || "clustered";
    chartOptions.showValue = slideData.showValue ?? false;
  }
  if (type === "line") {
    chartOptions.lineDataSymbol = slideData.symbol || "circle";
    chartOptions.lineDataSymbolSize = 5;
    chartOptions.lineSize = 2.0;
  }
  slide.addChart(type, data, chartOptions);
  if (slideData.caption) {
    addText(slide, slideData.caption, 1.02, 6.22, 10.7, 0.24, { fontSize: 9.2, color: theme.muted, align: "center" });
  }
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function wrapOmmlForPresentation(omml, display = true) {
  const cleaned = normalizeText(omml);
  if (display) {
    return `<a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">${cleaned}</a14:m>`;
  }
  return `<a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">${cleaned}</a14:m>`;
}

async function postprocessOmml(fileName, equations) {
  if (!equations.length) return;
  const replacements = [];
  for (const equation of equations) {
    let omml;
    try {
      omml = await latexToOMML(equation.latex);
    } catch (error) {
      throw new Error(`Formula conversion failed for "${equation.latex}": ${error.message}`);
    }
    replacements.push({
      marker: equation.marker,
      xml: wrapOmmlForPresentation(omml, equation.display),
    });
  }
  const zip = await JSZip.loadAsync(fs.readFileSync(fileName));
  const slideNames = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name));
  for (const slideName of slideNames) {
    let xml = await zip.file(slideName).async("string");
    let changed = false;
    for (const replacement of replacements) {
      if (!xml.includes(replacement.marker)) continue;
      const marker = escapeRegExp(replacement.marker);
      const runPattern = new RegExp(`<a:r>(?:(?!<\\/a:r>)[\\s\\S])*?<a:t[^>]*>${marker}<\\/a:t>(?:(?!<\\/a:r>)[\\s\\S])*?<\\/a:r>`, "g");
      xml = xml.replace(runPattern, replacement.xml);
      changed = true;
    }
    if (changed) zip.file(slideName, xml);
  }
  const out = await zip.generateAsync({ type: "nodebuffer" });
  fs.writeFileSync(fileName, out);
}

function layoutComparison(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title, "COMPARISON");
  addComparePanel(slide, slideData.left, 0.9, 1.74, theme.muted, "F5F5F5", ctx);
  addComparePanel(slide, slideData.right, 6.92, 1.74, theme.green, "F1F8F4", ctx);
  addText(slide, "VS", 6.12, 3.47, 0.56, 0.3, { fontSize: 14, bold: true, color: theme.red, align: "center", fontFace: font.en });
}

function addComparePanel(slide, data = {}, x, y, accent, fill, ctx) {
  slide.addShape("rect", { x, y, w: 5.25, h: 4.45, fill: { color: fill }, line: { color: accent, width: 1.3 } });
  slide.addShape("rect", { x, y, w: 5.25, h: 0.48, fill: { color: accent }, line: { transparency: 100 } });
  addText(slide, data.label || "", x + 0.22, y + 0.12, 4.8, 0.16, { fontSize: 9, color: theme.white, bold: true, fontFace: font.en });
  addText(slide, data.title || "", x + 0.32, y + 0.88, 4.52, 0.54, { fontSize: 20, color: accent, bold: true });
  addBulletList(slide, data.bullets || [], x + 0.42, y + 1.72, 4.34, 0.6, { fontSize: 13.5, color: accent }, ctx);
}

function chooseImageTextPlacement(slideData, image) {
  const requested = normalizeText(slideData.imagePlacement || slideData.placement || image.placement).toLowerCase();
  if (["top", "above", "wide", "bottomText", "bottom-text"].includes(requested)) return "top";
  if (["side", "left", "right"].includes(requested)) return "side";
  if (image.ratio >= 1.9) return "top";
  return "side";
}

function addBulletGrid(slide, bullets, x, y, w, options = {}, ctx = null) {
  const items = (Array.isArray(bullets) ? bullets : bullets ? [bullets] : []).slice(0, options.maxItems || 5);
  if (!items.length) return;
  const cols = Math.min(options.cols || (items.length <= 2 ? items.length : 3), items.length);
  const gapX = options.gapX || 0.36;
  const gapY = options.gapY || 0.48;
  const colW = (w - gapX * (cols - 1)) / cols;
  const size = options.fontSize || 11.4;
  items.forEach((item, idx) => {
    const col = idx % cols;
    const row = Math.floor(idx / cols);
    const left = x + col * (colW + gapX);
    const top = y + row * gapY;
    slide.addShape("ellipse", {
      x: left,
      y: top + 0.1,
      w: 0.1,
      h: 0.1,
      fill: { color: options.color || theme.green },
      line: { transparency: 100 },
    });
    const textOptions = {
      fontSize: size,
      color: options.textColor || theme.ink,
      valign: "top",
      fit: "shrink",
    };
    if (ctx) addInlineMathText(slide, item, left + 0.22, top, colW - 0.22, 0.36, textOptions, ctx);
    else addText(slide, item, left + 0.22, top, colW - 0.22, 0.36, textOptions);
  });
}

function layoutImageText(pptx, slideData, pageNo, ctx) {
  const slide = pptx.addSlide();
  addPageBrand(slide, pageNo);
  addTitle(slide, slideData.title, "VISUAL EXPLANATION");
  const image = resolveImageInfo(slideData.image || "assets/bit-campus-photo.png");
  image.fit = normalizeText(slideData.imageFit || slideData.fit || image.fit).toLowerCase();
  image.placement = normalizeText(slideData.imagePlacement || slideData.placement || image.placement).toLowerCase();
  const placement = chooseImageTextPlacement(slideData, image);

  if (placement === "top") {
    addImageInPanel(slide, image, { x: 0.98, y: 1.58, w: 11.35, h: 3.52 }, { fit: "contain" });
    if (slideData.caption) {
      addInlineMathText(slide, slideData.caption, 1.02, 5.22, 11.18, 0.22, { fontSize: 9.2, color: theme.muted, align: "center" }, ctx);
    }
    addBulletGrid(slide, slideData.text || [], 1.02, slideData.caption ? 5.66 : 5.46, 11.18, { fontSize: 11.2 }, ctx);
    return;
  }

  const sideFit = image.orientation === "portrait" || image.orientation === "tall" ? "contain" : "cover";
  addImageInPanel(slide, image, { x: 1.02, y: 1.8, w: 4.9, h: 4.08 }, { fit: sideFit });
  if (slideData.caption) {
    addInlineMathText(slide, slideData.caption, 1.0, 6.04, 4.94, 0.22, { fontSize: 8.8, color: theme.muted, align: "center" }, ctx);
  }
  addBulletList(slide, slideData.text || [], 6.82, 2.03, 5.15, 0.76, { fontSize: 16 }, ctx);
}

function layoutClosing(pptx, slideData) {
  const slide = pptx.addSlide();
  slide.background = { color: theme.green };
  addImageFit(slide, assets.emblemGray, { x: 8.6, y: 1.12, w: 4.0, h: 4.0 }, { transparency: 62 });
  addImageFit(slide, assets.wordmarkWhite, { x: 0.78, y: 0.54, w: 3.8, h: 1.02 });
  addText(slide, slideData.title || "谢谢", 0.82, 2.62, 6.4, 0.92, {
    fontSize: 36,
    bold: true,
    color: theme.white,
  });
  addText(slide, slideData.subtitle || "", 0.86, 3.58, 6.4, 0.42, {
    fontSize: 16,
    color: theme.white,
  });
}

function createDeck(deck, options = {}) {
  const resolvedFonts = configureFonts(deck, options);
  const pptx = new pptxgen();
  const ctx = { equations: [] };
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = deck.meta?.author || "BIT";
  pptx.company = "Beijing Institute of Technology";
  pptx.subject = deck.meta?.title || "BIT presentation";
  pptx.title = deck.meta?.title || "BIT presentation";
  pptx.lang = "zh-CN";
  pptx.theme = {
    headFontFace: font.cn,
    bodyFontFace: font.cn,
    lang: "zh-CN",
  };
  pptx.defineLayout({ name: "BIT_WIDE", width: W, height: H });
  pptx.layout = "BIT_WIDE";

  const preflight = expandSlidesWithReport(deck.slides || []);
  let pageNo = 1;
  for (const slide of preflight.slides) {
    switch (slide.layout) {
      case "title":
        layoutTitle(pptx, deck);
        break;
      case "agenda":
        layoutAgenda(pptx, slide, pageNo++);
        break;
      case "section":
        layoutSection(pptx, slide, pageNo++);
        break;
      case "bullets":
        layoutBullets(pptx, slide, pageNo++, ctx);
        break;
      case "claim":
        layoutClaim(pptx, slide, pageNo++, ctx);
        break;
      case "twoColumn":
        layoutTwoColumn(pptx, slide, pageNo++, ctx);
        break;
      case "cards":
        layoutCards(pptx, slide, pageNo++, ctx);
        break;
      case "table":
        layoutTable(pptx, slide, pageNo++, ctx);
        break;
      case "comparison":
        layoutComparison(pptx, slide, pageNo++, ctx);
        break;
      case "timeline":
        layoutTimeline(pptx, slide, pageNo++, ctx);
        break;
      case "process":
        layoutProcess(pptx, slide, pageNo++, ctx);
        break;
      case "problemSolution":
        layoutProblemSolution(pptx, slide, pageNo++, ctx);
        break;
      case "painOpportunity":
        layoutPainOpportunity(pptx, slide, pageNo++, ctx);
        break;
      case "experimentDesign":
        layoutExperimentDesign(pptx, slide, pageNo++, ctx);
        break;
      case "resultAnalysis":
        layoutResultAnalysis(pptx, slide, pageNo++, ctx);
        break;
      case "riskMitigation":
        layoutRiskMitigation(pptx, slide, pageNo++, ctx);
        break;
      case "contribution":
        layoutContribution(pptx, slide, pageNo++, ctx);
        break;
      case "summary":
        layoutSummary(pptx, slide, pageNo++, ctx);
        break;
      case "architecture":
        layoutArchitecture(pptx, slide, pageNo++, ctx);
        break;
      case "ablation":
        layoutAblation(pptx, slide, pageNo++, ctx);
        break;
      case "caseStudy":
        layoutCaseStudy(pptx, slide, pageNo++, ctx);
        break;
      case "imageGrid":
        layoutImageGrid(pptx, slide, pageNo++, ctx);
        break;
      case "code":
        layoutCode(pptx, slide, pageNo++, ctx);
        break;
      case "appendix":
        layoutAppendix(pptx, slide, pageNo++, ctx);
        break;
      case "flowchart":
        layoutFlowchart(pptx, slide, pageNo++, ctx);
        break;
      case "chart":
        layoutChart(pptx, slide, pageNo++);
        break;
      case "metrics":
        layoutMetrics(pptx, slide, pageNo++, ctx);
        break;
      case "quote":
        layoutQuote(pptx, slide, pageNo++, ctx);
        break;
      case "formula":
        layoutFormula(pptx, slide, pageNo++, ctx);
        break;
      case "references":
        layoutReferences(pptx, slide, pageNo++);
        break;
      case "matrix":
        layoutMatrix(pptx, slide, pageNo++, ctx);
        break;
      case "imageText":
        layoutImageText(pptx, slide, pageNo++, ctx);
        break;
      case "closing":
        layoutClosing(pptx, slide);
        break;
      default:
        throw new Error(`Unknown slide layout: ${slide.layout}`);
    }
  }
  return { pptx, preflight: preflight.report, equations: ctx.equations, fonts: resolvedFonts };
}

function listLayouts() {
  return [...SUPPORTED_LAYOUTS];
}

function checkDeck(deck) {
  const validation = validateDeck(deck);
  const preflightOnly = expandSlidesWithReport(deck.slides || []);
  return {
    inputSlides: (deck.slides || []).length,
    outputSlides: preflightOnly.slides.length,
    actions: preflightOnly.report,
    validation: {
      errors: validation.errors,
      warnings: validation.warnings,
    },
    repairPrompt: validation.repairPrompt,
  };
}

function checkDeckFile(input) {
  const deck = readDeck(input);
  return checkDeck(deck);
}

async function generateDeckFile(input, output, options = {}) {
  const deck = readDeck(input);
  const check = checkDeck(deck);
  if (check.validation.errors.length) {
    const error = new Error("Deck validation failed.");
    error.validation = check.validation;
    error.repairPrompt = check.repairPrompt;
    throw error;
  }
  fs.mkdirSync(path.dirname(output), { recursive: true });
  const { pptx, preflight, equations, fonts: resolvedFonts } = createDeck(deck, options);
  let fileName = output;
  try {
    await pptx.writeFile({ fileName });
    await postprocessOmml(fileName, equations);
  } catch (error) {
    if (error?.code !== "EBUSY") throw error;
    const parsed = path.parse(output);
    const stamp = new Date().toISOString().replace(/[-:]/g, "").replace(/\..+/, "");
    fileName = path.join(parsed.dir, `${parsed.name}-${stamp}${parsed.ext}`);
    await pptx.writeFile({ fileName });
    await postprocessOmml(fileName, equations);
  }
  return { output: fileName, preflight, validation: check.validation, repairPrompt: check.repairPrompt, fonts: resolvedFonts };
}

function parseLegacyOptions(args) {
  const options = {};
  for (let idx = 0; idx < args.length; idx += 1) {
    const arg = args[idx];
    if (arg === "--font-cn" || arg === "--font-cjk") options.fontCn = args[++idx];
    else if (arg === "--font-cn-light") options.fontCnLight = args[++idx];
    else if (arg === "--font-en" || arg === "--font-latin") options.fontEn = args[++idx];
    else if (arg === "--font-serif") options.fontSerif = args[++idx];
    else if (arg === "--font-code") options.fontCode = args[++idx];
  }
  return options;
}

async function main() {
  const input = process.argv[2] || path.join(ROOT, "content", "example.yaml");
  const output = process.argv[3] || path.join(ROOT, "output", "example.pptx");
  const checkOnly = process.argv.includes("--check");
  if (checkOnly) {
    const result = checkDeckFile(input);
    console.log(JSON.stringify(result, null, 2));
    if (result.validation.errors.length) process.exitCode = 1;
    return;
  }
  try {
    const result = await generateDeckFile(input, output, parseLegacyOptions(process.argv.slice(4)));
    if (result.validation.warnings.length) {
      console.warn(`Validation warning(s): ${result.validation.warnings.length}`);
      console.warn(result.repairPrompt);
    }
    if (result.preflight.length) {
      console.log(`Preflight adjusted ${result.preflight.length} slide(s): ${result.preflight.map((item) => `${item.title}->${item.parts}`).join(", ")}`);
    }
    console.log(`Generated ${result.output}`);
  } catch (error) {
    if (error.validation) {
      console.error(JSON.stringify({ error: error.message, validation: error.validation, repairPrompt: error.repairPrompt }, null, 2));
      process.exitCode = 1;
      return;
    }
    throw error;
  }
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
}

export {
  ROOT,
  checkDeck,
  checkDeckFile,
  configureFonts,
  createDeck,
  defaultFont,
  generateDeckFile,
  listLayouts,
  readDeck,
  validateDeck,
};
