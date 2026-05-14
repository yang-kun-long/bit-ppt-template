import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import test from "node:test";
import JSZip from "jszip";
import YAML from "yaml";
import {
  ROOT,
  generateDeckFile,
  validateDeck,
} from "../src/generate.mjs";
import {
  ICON_COLOR_TOKENS,
  ICON_SIZE_PRESETS,
  getIconMeta,
  hasIcon,
  listIconNames,
  listIconsByCategory,
} from "../src/icons.mjs";
import { getIconsGuide } from "../src/layout-guides.mjs";

const execFileAsync = promisify(execFile);
const CLI = path.join(ROOT, "bin", "bit-ppt.mjs");

function tempFile(name) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "bit-ppt-icons-test-"));
  return { dir, path: path.join(dir, name) };
}

function writeYamlDeck(deck) {
  const file = tempFile("deck.yaml");
  fs.writeFileSync(file.path, YAML.stringify(deck), "utf8");
  return file;
}

async function readSlideXml(pptxPath, slideIndex = 1) {
  const zip = await JSZip.loadAsync(fs.readFileSync(pptxPath));
  return zip.file(`ppt/slides/slide${slideIndex}.xml`).async("string");
}

async function runCli(args) {
  try {
    const result = await execFileAsync(process.execPath, [CLI, ...args], {
      cwd: ROOT,
      maxBuffer: 20 * 1024 * 1024,
    });
    return { status: 0, stdout: result.stdout, stderr: result.stderr };
  } catch (error) {
    return { status: error.code ?? 1, stdout: error.stdout ?? "", stderr: error.stderr ?? "" };
  }
}

test("icon manifest exposes core helpers", () => {
  const names = listIconNames();
  assert.ok(names.length >= 46, `expected at least 46 icons, got ${names.length}`);
  assert.ok(names.includes("magnifier"));
  assert.ok(names.includes("settings-gear"));
  assert.equal(hasIcon("magnifier"), true);
  assert.equal(hasIcon("not-a-real-icon"), false);

  const meta = getIconMeta("magnifier");
  assert.ok(meta);
  assert.equal(meta.kind, "grpSp");
  assert.equal(Array.isArray(meta.native_size_cm), true);
  assert.equal(meta.native_size_cm.length, 2);
});

test("icons are grouped into the documented categories", () => {
  const byCategory = listIconsByCategory();
  const expected = ["tech", "creative", "academic", "office", "people", "thinking", "sport", "misc"];
  for (const cat of expected) {
    assert.ok(Array.isArray(byCategory[cat]) && byCategory[cat].length > 0, `category ${cat} missing or empty`);
  }
});

test("BIT color tokens resolve to expected hex values", () => {
  assert.equal(ICON_COLOR_TOKENS.primary, "006C39");
  assert.equal(ICON_COLOR_TOKENS.green, "006C39");
  assert.equal(ICON_COLOR_TOKENS.accent1, "006C39");
  assert.equal(ICON_COLOR_TOKENS.red, "A13F3D");
  assert.equal(ICON_COLOR_TOKENS.white, "FFFFFF");
});

test("size presets cover sm/md/lg", () => {
  assert.equal(typeof ICON_SIZE_PRESETS.sm, "number");
  assert.equal(typeof ICON_SIZE_PRESETS.md, "number");
  assert.equal(typeof ICON_SIZE_PRESETS.lg, "number");
  assert.ok(ICON_SIZE_PRESETS.sm < ICON_SIZE_PRESETS.md);
  assert.ok(ICON_SIZE_PRESETS.md < ICON_SIZE_PRESETS.lg);
});

test("generated PPTX replaces icon markers with vector geometry in BIT green by default", async () => {
  const deckFile = writeYamlDeck({
    slides: [
      {
        layout: "cards",
        title: "Default icon color",
        cards: [
          { icon: "magnifier", title: "调研", text: "默认配色应当是 BIT 主绿。" },
        ],
      },
    ],
  });
  const output = path.join(deckFile.dir, "default-color.pptx");
  const result = await generateDeckFile(deckFile.path, output);
  assert.equal(result.validation.errors.length, 0);

  const xml = await readSlideXml(result.output);
  assert.equal(xml.includes("__BIT_ICON__"), false, "marker should be replaced");
  assert.ok(xml.includes("custGeom"), "icon vector geometry should be present");
  assert.ok(xml.includes('srgbClr val="006C39"'), "default color should be BIT green");
  assert.equal(xml.includes('schemeClr val="accent'), false, "should not emit scheme colors anymore");
});

test("iconColor token resolves to BIT red", async () => {
  const deckFile = writeYamlDeck({
    slides: [
      {
        layout: "cards",
        title: "Custom token color",
        cards: [
          { icon: "target-arrow", iconColor: "red", title: "突破", text: "应当渲染为 BIT 红。" },
        ],
      },
    ],
  });
  const output = path.join(deckFile.dir, "red-token.pptx");
  const result = await generateDeckFile(deckFile.path, output);
  assert.equal(result.validation.errors.length, 0);

  const xml = await readSlideXml(result.output);
  assert.ok(xml.includes('srgbClr val="A13F3D"'), "iconColor: red should produce BIT red");
});

test("explicit hex iconColor is honored as-is", async () => {
  const deckFile = writeYamlDeck({
    slides: [
      {
        layout: "cards",
        title: "Custom hex color",
        cards: [
          { icon: "lightbulb", iconColor: "123456", title: "灵感", text: "传入十六进制颜色。" },
        ],
      },
    ],
  });
  const output = path.join(deckFile.dir, "hex-color.pptx");
  const result = await generateDeckFile(deckFile.path, output);
  assert.equal(result.validation.errors.length, 0);

  const xml = await readSlideXml(result.output);
  assert.ok(xml.includes('srgbClr val="123456"'), "explicit hex must appear in slide XML");
});

test("bullets accept object form { icon, iconColor?, text } without validation warnings", () => {
  const deck = {
    slides: [
      {
        layout: "bullets",
        title: "图标 bullet",
        bullets: [
          { icon: "checkmark", text: "对象形式不应触发字符长度警告。" },
          { icon: "magnifier", iconColor: "red", text: "iconColor 也可以指定。" },
          "字符串形式继续支持。",
        ],
      },
    ],
  };
  const validation = validateDeck(deck);
  assert.equal(validation.errors.length, 0);
  assert.equal(validation.warnings.length, 0, JSON.stringify(validation.warnings, null, 2));
});

test("bullets text length warning measures only the .text field on object items", () => {
  const longText = "字符".repeat(80);
  const deck = {
    slides: [
      {
        layout: "bullets",
        title: "Long bullet",
        bullets: [{ icon: "checkmark", text: longText }],
      },
    ],
  };
  const validation = validateDeck(deck);
  assert.equal(validation.errors.length, 0);
  assert.equal(validation.warnings.length, 1);
  assert.match(validation.warnings[0].path, /bullets\[0\]/);
  assert.match(validation.warnings[0].message, /bullet is \d+ chars/);
});

test("generated bullets layout injects icons in BIT green by default", async () => {
  const deckFile = writeYamlDeck({
    slides: [
      {
        layout: "bullets",
        title: "Bullet 图标",
        bullets: [
          { icon: "checkmark", text: "项目一" },
          "字符串项二",
        ],
      },
    ],
  });
  const output = path.join(deckFile.dir, "bullet-icons.pptx");
  const result = await generateDeckFile(deckFile.path, output);
  assert.equal(result.validation.errors.length, 0);
  const xml = await readSlideXml(result.output);
  assert.equal(xml.includes("__BIT_ICON__"), false);
  assert.ok(xml.includes("custGeom"));
  assert.ok(xml.includes('srgbClr val="006C39"'));
});

test("getIconsGuide reports the live catalog", () => {
  const guide = getIconsGuide();
  assert.equal(guide.topic, "icons");
  assert.ok(guide.totalIcons >= 46);
  assert.deepEqual(guide.layouts.sort(), ["bullets", "cards", "comparison"]);
  assert.ok(guide.catalog.academic.includes("microscope"));
  assert.equal(guide.colorTokens.primary, "006C39");
  assert.ok(typeof guide.sizePresets.md === "number");
});

test("CLI guide icons --json exposes catalog and presets", async () => {
  const result = await runCli(["guide", "icons", "--json"]);
  assert.equal(result.status, 0);
  const payload = JSON.parse(result.stdout);
  assert.equal(payload.topic, "icons");
  assert.ok(payload.totalIcons >= 46);
  assert.ok(payload.layouts.includes("cards"));
  assert.ok(payload.layouts.includes("bullets"));
  assert.equal(payload.colorTokens.red, "A13F3D");
  assert.ok(Array.isArray(payload.catalog.thinking));
});

test("CLI help mentions guide icons", async () => {
  const result = await runCli(["--help"]);
  assert.equal(result.status, 0);
  assert.match(result.stdout, /bit-ppt guide icons/);
});
