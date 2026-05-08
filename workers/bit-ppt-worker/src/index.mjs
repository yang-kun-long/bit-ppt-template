import YAML from "yaml";

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

function normalizeText(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  if (Array.isArray(value)) return value.map(normalizeText).join(" ");
  if (typeof value === "object") return Object.entries(value).map(([key, val]) => `${key}: ${normalizeText(val)}`).join(" ");
  return String(value);
}

function charCount(value) {
  return [...normalizeText(value)].length;
}

function isPlaceholderImage(value) {
  return Boolean(value && typeof value === "object" && !Array.isArray(value) && (value.mode === "placeholder" || value.placeholder === true || value.noImage === true));
}

function placeholderNeedsVariants(slide) {
  if (!slide || slide.layout !== "imageText" || !isPlaceholderImage(slide.image)) return false;
  if (slide.placeholderVariants === false || slide.image.variants === false) return false;
  const ratio = normalizeText(slide.image.aspectRatio || slide.image.ratio || slide.image.size).toLowerCase();
  return !ratio || ratio === "auto" || ratio === "unknown";
}

function expandSlidesWithReport(slides = []) {
  const report = [];
  const expanded = [];
  slides.forEach((slide, index) => {
    if (placeholderNeedsVariants(slide)) {
      report.push({
        slideIndex: index + 1,
        layout: slide.layout,
        title: slide.title || slide.layout,
        action: "placeholderVariants",
        parts: 2,
      });
      expanded.push(
        { ...slide, title: `${slide.title || "图文说明"}（横图方案）`, image: { ...slide.image, aspectRatio: "16:9", placement: "top", variants: false } },
        { ...slide, title: `${slide.title || "图文说明"}（侧图方案）`, image: { ...slide.image, aspectRatio: "4:3", placement: "side", variants: false } },
      );
      return;
    }
    expanded.push(slide);
  });
  return { slides: expanded, report };
}

function makeIssue(level, slideIndex, path, message, repair) {
  return { level, slideIndex, path, message, repair };
}

function validateDeck(deck) {
  const issues = [];
  const add = (level, slideIndex, path, message, repair) => issues.push(makeIssue(level, slideIndex, path, message, repair));
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
  const slides = Array.isArray(deck.slides) ? deck.slides : [];
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
    if (charCount(slide.title) > 24) add("warning", slideIndex, "title", "title is too long for the BIT template.", "Shorten title to 24 characters or split the slide.");
    if (slide.layout === "imageText") {
      if (!slide.image) add("error", slideIndex, "image", "imageText requires image.", "Add image path or image.mode: placeholder.");
      if (isPlaceholderImage(slide.image) && !normalizeText(slide.image.prompt || slide.image.description || slide.image.alt)) {
        add("warning", slideIndex, "image.prompt", "Image placeholder has no prompt or description.", "Add image.prompt for later replacement.");
      }
      const text = Array.isArray(slide.text) ? slide.text : [];
      if (text.length > 5) add("warning", slideIndex, "text", "imageText has too many text bullets.", "Keep imageText text to 5 bullets or fewer.");
      text.forEach((item, itemIdx) => {
        if (charCount(item) > 38) add("warning", slideIndex, `text[${itemIdx}]`, "bullet is too long for imageText.", "Shorten this bullet to 38 characters or split the slide.");
      });
    }
  });
  const errors = issues.filter((item) => item.level === "error");
  const warnings = issues.filter((item) => item.level === "warning");
  const repairPrompt = issues.length ? issues.map((item) => `Slide ${item.slideIndex} ${item.path}: ${item.repair || item.message}`).join("\n") : "";
  return { errors, warnings, repairPrompt };
}

function checkDeck(deck) {
  const validation = validateDeck(deck);
  const preflight = expandSlidesWithReport(Array.isArray(deck?.slides) ? deck.slides : []);
  return {
    inputSlides: Array.isArray(deck?.slides) ? deck.slides.length : 0,
    outputSlides: preflight.slides.length,
    actions: preflight.report,
    validation: {
      errors: validation.errors,
      warnings: validation.warnings,
    },
    repairPrompt: validation.repairPrompt,
    runtime: "cloudflare-worker-spike",
    capabilities: {
      check: true,
      generate: false,
      omml: false,
    },
  };
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data, null, 2), {
    status,
    headers: {
      "content-type": "application/json; charset=utf-8",
      "access-control-allow-origin": "*",
      "access-control-allow-methods": "GET,POST,OPTIONS",
      "access-control-allow-headers": "content-type",
    },
  });
}

async function parseDeckRequest(request) {
  const contentType = request.headers.get("content-type") || "";
  if (contentType.includes("application/json")) {
    const body = await request.json();
    if (body.deck) return body.deck;
    if (body.deckYaml) return YAML.parse(body.deckYaml);
    if (body.yaml) return YAML.parse(body.yaml);
    throw new Error("JSON body must include deck, deckYaml, or yaml.");
  }
  return YAML.parse(await request.text());
}

export default {
  async fetch(request) {
    const url = new URL(request.url);
    if (request.method === "OPTIONS") return json({ ok: true });
    if (request.method === "GET" && url.pathname === "/health") {
      return json({
        ok: true,
        service: "bit-ppt-worker-spike",
        capabilities: {
          check: true,
          generate: false,
          omml: false,
        },
      });
    }
    if (request.method === "POST" && url.pathname === "/check") {
      try {
        const deck = await parseDeckRequest(request);
        return json(checkDeck(deck));
      } catch (error) {
        return json({ error: error.message || String(error) }, 400);
      }
    }
    return json({ error: "Not found. Use GET /health or POST /check." }, 404);
  },
};
