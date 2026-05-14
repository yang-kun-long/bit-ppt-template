import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const ICONS_DIR = path.join(ROOT, "assets", "icons");
const EMU_PER_INCH = 914400;

const BIT_PALETTE = {
  accent1: "006C39",
  accent2: "004B28",
  accent3: "A13F3D",
  accent4: "006C39",
  primary: "006C39",
  green: "006C39",
  darkgreen: "004B28",
  red: "A13F3D",
  ink: "262626",
  muted: "666666",
  white: "FFFFFF",
};

const SIZE_PRESETS = {
  xs: 0.36,
  sm: 0.5,
  md: 0.65,
  lg: 0.85,
  xl: 1.1,
};

const DEFAULT_ICON_SIZE = SIZE_PRESETS.md;

function resolveIconColorHex(color) {
  if (!color) return BIT_PALETTE.primary;
  const key = String(color).trim().toLowerCase();
  if (BIT_PALETTE[key]) return BIT_PALETTE[key];
  return String(color).replace(/^#/, "").toUpperCase();
}

function resolveIconSize(value) {
  if (value == null) return DEFAULT_ICON_SIZE;
  if (typeof value === "number") return value;
  const key = String(value).trim().toLowerCase();
  if (SIZE_PRESETS[key] != null) return SIZE_PRESETS[key];
  const num = Number(value);
  return Number.isFinite(num) ? num : DEFAULT_ICON_SIZE;
}

let manifestCache = null;

function loadManifest() {
  if (manifestCache) return manifestCache;
  const raw = fs.readFileSync(path.join(ICONS_DIR, "manifest.json"), "utf-8");
  manifestCache = JSON.parse(raw);
  return manifestCache;
}

export function listIconNames() {
  return Object.keys(loadManifest().icons);
}

export const ICON_SIZE_PRESETS = { ...SIZE_PRESETS };

export const ICON_COLOR_TOKENS = { ...BIT_PALETTE };

export function listIconsByCategory() {
  const out = {};
  const icons = loadManifest().icons;
  for (const [name, meta] of Object.entries(icons)) {
    const cat = meta.category || "misc";
    (out[cat] ??= []).push(name);
  }
  for (const list of Object.values(out)) list.sort();
  return out;
}

export function getIconMeta(name) {
  return loadManifest().icons[name];
}

export function hasIcon(name) {
  return Boolean(getIconMeta(name));
}

function loadFragment(name) {
  const meta = getIconMeta(name);
  if (!meta) throw new Error(`Unknown icon: ${name}`);
  const fragmentPath = path.join(ICONS_DIR, meta.fragment);
  return fs.readFileSync(fragmentPath, "utf-8");
}

const inchToEmu = (inches) => Math.round(inches * EMU_PER_INCH);

function rewriteFragment(xml, { x, y, w, h, color, idBase }) {
  let out = xml;
  out = out.replace(/<a:off x="\d+" y="\d+"\/>/, `<a:off x="${x}" y="${y}"/>`);
  out = out.replace(/<a:ext cx="\d+" cy="\d+"\/>/, `<a:ext cx="${w}" cy="${h}"/>`);

  let counter = idBase;
  out = out.replace(/<p:cNvPr id="\d+"/g, () => `<p:cNvPr id="${counter++}"`);

  const hex = resolveIconColorHex(color);
  out = out.replace(/<a:schemeClr val="[^"]+"\/>/g, `<a:srgbClr val="${hex}"/>`);
  if (color) {
    out = out.replace(/<a:srgbClr val="[^"]+"\/>/g, `<a:srgbClr val="${hex}"/>`);
  }

  return out;
}

export function createIconCollector() {
  const requests = [];
  let nextId = 1;
  return {
    request(icon, options = {}) {
      if (!hasIcon(icon)) {
        throw new Error(`Unknown icon: ${icon}. Use listIconNames() to see options.`);
      }
      const id = `${Date.now().toString(36)}-${nextId++}`;
      const req = { id, icon, color: options.color || null };
      requests.push(req);
      return req.id;
    },
    list() {
      return requests;
    },
  };
}

export function addIconMarker(slide, requestId, { x, y, w, h }) {
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: "FFFFFF", transparency: 100 },
    line: { color: "FFFFFF", transparency: 100 },
    objectName: `__BIT_ICON__${requestId}`,
  });
}

export function addIcon(slide, collector, icon, { x, y, w, h, size, color, fit = "contain" } = {}) {
  const meta = getIconMeta(icon);
  if (!meta) throw new Error(`Unknown icon: ${icon}`);

  let targetW = w;
  let targetH = h;

  if (targetW == null && targetH == null) {
    const s = resolveIconSize(size);
    targetW = s;
    targetH = s;
  } else if (targetW == null) {
    targetW = targetH;
  } else if (targetH == null) {
    targetH = targetW;
  }

  if (fit === "contain" && meta.native_size_cm) {
    const [nw, nh] = meta.native_size_cm;
    const ratio = nw / nh;
    const boxRatio = targetW / targetH;
    if (boxRatio > ratio) {
      const newW = targetH * ratio;
      x += (targetW - newW) / 2;
      targetW = newW;
    } else if (boxRatio < ratio) {
      const newH = targetW / ratio;
      y += (targetH - newH) / 2;
      targetH = newH;
    }
  }

  const id = collector.request(icon, { color });
  addIconMarker(slide, id, { x, y, w: targetW, h: targetH });
  return id;
}

export async function postprocessIcons(fileName, iconRequests) {
  if (!iconRequests.length) return;

  const buffer = fs.readFileSync(fileName);
  const zip = await JSZip.loadAsync(buffer);
  const slideNames = Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));

  let idBase = 100000;

  for (const slideName of slideNames) {
    let xml = await zip.file(slideName).async("string");
    let changed = false;

    for (const req of iconRequests) {
      const markerName = `__BIT_ICON__${req.id}`;
      const escaped = markerName.replace(/[-\/\\^$*+?.()|[\]{}]/g, "\\$&");
      const pattern = new RegExp(
        `<p:sp>(?:(?!<\\/p:sp>)[\\s\\S])*?<p:cNvPr id="\\d+" name="${escaped}"(?:\\/>|>[\\s\\S]*?<\\/p:cNvPr>)(?:(?!<\\/p:sp>)[\\s\\S])*?<\\/p:sp>`,
        "g",
      );
      const match = pattern.exec(xml);
      if (!match) continue;

      const markerXml = match[0];
      const offMatch = markerXml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
      const extMatch = markerXml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
      if (!offMatch || !extMatch) continue;

      const target = {
        x: parseInt(offMatch[1], 10),
        y: parseInt(offMatch[2], 10),
        w: parseInt(extMatch[1], 10),
        h: parseInt(extMatch[2], 10),
        color: req.color,
        idBase,
      };
      idBase += 200;

      const fragment = loadFragment(req.icon);
      const replacement = rewriteFragment(fragment, target);

      xml = xml.replace(markerXml, replacement);
      changed = true;
    }

    if (changed) zip.file(slideName, xml);
  }

  const out = await zip.generateAsync({ type: "nodebuffer" });
  fs.writeFileSync(fileName, out);
}
