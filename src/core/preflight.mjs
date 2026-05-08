import { PLACEHOLDER_ASPECT_RATIOS } from "./layouts.mjs";

const INCH_PT = 72;
const OVERFLOW_GUARD = 0.9;

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

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function hasKnownPlaceholderRatio(value) {
  const ratio = normalizeText(value).toLowerCase();
  if (!ratio || ratio === "auto" || ratio === "unknown") return false;
  if (PLACEHOLDER_ASPECT_RATIOS[ratio]) return true;
  const colon = ratio.match(/^(\d+(?:\.\d+)?):(\d+(?:\.\d+)?)$/);
  if (colon) return Number(colon[1]) > 0 && Number(colon[2]) > 0;
  const numeric = Number(ratio);
  return Number.isFinite(numeric) && numeric > 0;
}

function resolvePlaceholderRatio(value, fallback = 16 / 9) {
  const ratio = normalizeText(value).toLowerCase();
  if (PLACEHOLDER_ASPECT_RATIOS[ratio]) return PLACEHOLDER_ASPECT_RATIOS[ratio];
  const colon = ratio.match(/^(\d+(?:\.\d+)?):(\d+(?:\.\d+)?)$/);
  if (colon) {
    const w = Number(colon[1]);
    const h = Number(colon[2]);
    if (w > 0 && h > 0) return w / h;
  }
  const numeric = Number(ratio);
  if (Number.isFinite(numeric) && numeric > 0) return numeric;
  return fallback;
}

function getSpeakerNotesValue(slide) {
  if (!slide || typeof slide !== "object") return undefined;
  return slide.speakerNotes ?? slide.speaker_notes ?? slide.speakerScript ?? slide.speaker_script;
}

function normalizeSpeakerNotes(value) {
  if (value === null || value === undefined) return "";
  if (Array.isArray(value)) return value.map(normalizeText).filter(Boolean).join("\n");
  return normalizeText(value).replace(/\r\n?/g, "\n").trim();
}

function isPlaceholderImage(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const mode = normalizeText(value.mode || value.type).toLowerCase();
  return mode === "placeholder" || value.placeholder === true || value.noImage === true;
}

function imagePlaceholderNeedsVariants(slide) {
  if (!slide || slide.layout !== "imageText" || !isPlaceholderImage(slide.image)) return false;
  if (slide.placeholderVariants === false) return false;
  if (slide.image && typeof slide.image === "object" && slide.image.variants === false) return false;
  const ratio = normalizeText(slide.image.aspectRatio || slide.image.ratio || slide.image.size).toLowerCase();
  return !hasKnownPlaceholderRatio(ratio);
}

function splitImagePlaceholderSlide(slide) {
  if (!imagePlaceholderNeedsVariants(slide)) return [slide];
  const baseImage = slide.image && typeof slide.image === "object" && !Array.isArray(slide.image) ? slide.image : { mode: "placeholder" };
  return [
    {
      ...clone(slide),
      title: `${slide.title || "图文说明"}（横图方案）`,
      image: { ...baseImage, mode: "placeholder", aspectRatio: "16:9", placement: "top", variants: false },
      placement: "top",
    },
    {
      ...clone(slide),
      title: `${slide.title || "图文说明"}（侧图方案）`,
      image: { ...baseImage, mode: "placeholder", aspectRatio: "4:3", placement: "side", variants: false },
      placement: "side",
    },
  ];
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

function expandSlidesWithReport(slides = []) {
  const report = [];
  const expanded = [];
  slides.forEach((slide, index) => {
    let parts;
    if (slide.layout === "bullets") parts = splitBulletsSlide(slide);
    else if (slide.layout === "table") parts = splitTableSlide(slide);
    else if (slide.layout === "references") parts = splitReferencesSlide(slide);
    else if (slide.layout === "imageText") parts = splitImagePlaceholderSlide(slide);
    else parts = [slide];
    if (parts.length > 1) {
      report.push({
        slideIndex: index + 1,
        layout: slide.layout,
        title: slide.title || slide.layout,
        action: imagePlaceholderNeedsVariants(slide) ? "placeholderVariants" : "split",
        parts: parts.length,
      });
    }
    expanded.push(...parts);
  });
  return { slides: expanded, report };
}

function expandSlides(slides = []) {
  return expandSlidesWithReport(slides).slides;
}

export {
  clone,
  expandSlides,
  expandSlidesWithReport,
  getSpeakerNotesValue,
  hasKnownPlaceholderRatio,
  imagePlaceholderNeedsVariants,
  isPlaceholderImage,
  normalizeSpeakerNotes,
  normalizeText,
  resolvePlaceholderRatio,
};
