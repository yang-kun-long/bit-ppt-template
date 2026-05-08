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
  checkDeck,
  checkDeckFile,
  generateDeckFile,
  listLayouts,
  validateDeck,
} from "../src/generate.mjs";

const execFileAsync = promisify(execFile);
const CLI = path.join(ROOT, "bin", "bit-ppt.mjs");

async function runCli(args, options = {}) {
  try {
    const result = await execFileAsync(process.execPath, [CLI, ...args], {
      cwd: ROOT,
      maxBuffer: 20 * 1024 * 1024,
      ...options,
    });
    return { status: 0, stdout: result.stdout, stderr: result.stderr };
  } catch (error) {
    return {
      status: error.code ?? 1,
      stdout: error.stdout ?? "",
      stderr: error.stderr ?? "",
    };
  }
}

function tempFile(name) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "bit-ppt-test-"));
  return {
    dir,
    path: path.join(dir, name),
  };
}

function writeYamlDeck(deck) {
  const file = tempFile("deck.yaml");
  fs.writeFileSync(file.path, YAML.stringify(deck), "utf8");
  return file;
}

test("listLayouts includes core and image layouts", () => {
  const layouts = listLayouts();
  assert.ok(layouts.includes("title"));
  assert.ok(layouts.includes("chart"));
  assert.ok(layouts.includes("imageText"));
  assert.ok(layouts.includes("imageGrid"));
});

test("checkDeckFile accepts the example deck", () => {
  const result = checkDeckFile(path.join(ROOT, "content", "example.yaml"));
  assert.equal(result.inputSlides, 18);
  assert.equal(result.validation.errors.length, 0);
  assert.equal(result.repairPrompt, "");
});

test("checkDeckFile returns repair prompt for invalid fixture", () => {
  const result = checkDeckFile(path.join(ROOT, "content", "invalid-deck-test.yaml"));
  assert.equal(result.validation.errors.length, 2);
  assert.match(result.repairPrompt, /series\[0\]\.values/);
  assert.match(result.repairPrompt, /edges\[0\]\.to/);
});

test("CLI list-layouts --json prints JSON", async () => {
  const result = await runCli(["list-layouts", "--json"]);
  assert.equal(result.status, 0);
  const layouts = JSON.parse(result.stdout);
  assert.ok(layouts.includes("imageText"));
});

test("CLI help points to progressive guide commands", async () => {
  const result = await runCli(["--help"]);
  assert.equal(result.status, 0);
  assert.match(result.stdout, /bit-ppt guide layout imageText/);
  assert.match(result.stdout, /bit-ppt guide speaker-notes/);
  assert.match(result.stdout, /bit-ppt guide image-placeholder/);
  assert.match(result.stdout, /bit-ppt doctor/);
  assert.match(result.stdout, /Progressive guide/);
});

test("CLI guide prints overview", async () => {
  const result = await runCli(["guide"]);
  assert.equal(result.status, 0);
  assert.match(result.stdout, /BIT PPT Generator/);
  assert.match(result.stdout, /Guided layouts/);
});

test("CLI guide layout returns focused text", async () => {
  const result = await runCli(["guide", "layout", "imageText"]);
  assert.equal(result.status, 0);
  assert.match(result.stdout, /imageText/);
  assert.match(result.stdout, /Very wide images/);
});

test("CLI guide schema --json returns layout schema", async () => {
  const result = await runCli(["guide", "schema", "imageText", "--json"]);
  assert.equal(result.status, 0);
  const payload = JSON.parse(result.stdout);
  assert.equal(payload.layout, "imageText");
  assert.equal(payload.fields.image.required, true);
  assert.equal(payload.commonFields.speakerNotes.type, "string | string[]");
});

test("CLI guide speaker-notes returns focused guide", async () => {
  const result = await runCli(["guide", "speaker-notes", "--json"]);
  assert.equal(result.status, 0);
  const payload = JSON.parse(result.stdout);
  assert.equal(payload.topic, "speaker-notes");
  assert.equal(payload.fields.speakerNotes.type, "string | string[]");
  assert.match(payload.notes.join("\n"), /plain text/);
});

test("CLI guide image-placeholder returns focused guide", async () => {
  const result = await runCli(["guide", "image-placeholder", "--json"]);
  assert.equal(result.status, 0);
  const payload = JSON.parse(result.stdout);
  assert.equal(payload.topic, "image-placeholder");
  assert.match(payload.fields.image.type, /placeholder/);
  assert.match(payload.notes.join("\n"), /imageText/);
});

test("CLI guide example --json returns example deck fragment", async () => {
  const result = await runCli(["guide", "example", "chart", "--json"]);
  assert.equal(result.status, 0);
  const payload = JSON.parse(result.stdout);
  assert.equal(payload.layout, "chart");
  assert.ok(Array.isArray(payload.categories));
});

test("CLI guide workflow --json returns agent workflow", async () => {
  const result = await runCli(["guide", "workflow", "--json"]);
  assert.equal(result.status, 0);
  const payload = JSON.parse(result.stdout);
  assert.ok(Array.isArray(payload));
  assert.match(payload.join("\n"), /repairPrompt/);
});

test("CLI guide all --json returns overview and layout guides", async () => {
  const result = await runCli(["guide", "all", "--json"]);
  assert.equal(result.status, 0);
  const payload = JSON.parse(result.stdout);
  assert.equal(payload.overview.name, "BIT PPT Generator");
  assert.ok(payload.layouts.imageText);
  assert.ok(payload.layouts.chart);
});

test("CLI doctor --json reports environment status", async () => {
  const result = await runCli(["doctor", "--json"]);
  assert.equal(result.status, 0);
  const payload = JSON.parse(result.stdout);
  assert.equal(payload.ok, true);
  assert.ok(payload.checks.some((item) => item.name === "fixture:example" && item.ok));
  assert.ok(payload.checks.some((item) => item.name === "output:writable" && item.ok));
});

test("CLI check --json exits successfully for valid deck", async () => {
  const result = await runCli(["check", "content/example.yaml", "--json"]);
  assert.equal(result.status, 0);
  const payload = JSON.parse(result.stdout);
  assert.equal(payload.validation.errors.length, 0);
});

test("CLI check exits non-zero for invalid deck", async () => {
  const result = await runCli(["check", "content/invalid-deck-test.yaml", "--json"]);
  assert.notEqual(result.status, 0);
  const payload = JSON.parse(result.stdout);
  assert.equal(payload.validation.errors.length, 2);
});

test("CLI check --strict exits non-zero on warnings", async () => {
  const result = await runCli(["check", "content/formula-test.yaml", "--json", "--strict"]);
  assert.notEqual(result.status, 0);
  const payload = JSON.parse(result.stdout);
  assert.equal(payload.validation.errors.length, 0);
  assert.ok(payload.validation.warnings.length > 0);
});

test("generateDeckFile writes a PPTX", async () => {
  const output = tempFile("example.pptx");
  const result = await generateDeckFile(path.join(ROOT, "content", "example.yaml"), output.path);
  assert.equal(result.validation.errors.length, 0);
  assert.ok(fs.existsSync(result.output));
  assert.ok(fs.statSync(result.output).size > 1000);
});

test("generateDeckFile writes speaker notes into notes slides", async () => {
  const deckFile = writeYamlDeck({
    meta: {
      title: "Speaker notes test",
    },
    slides: [
      {
        layout: "bullets",
        title: "备注测试",
        bullets: ["页面正文保持可编辑。"],
        speakerNotes: [
          "第一段演讲稿会写入 PowerPoint 备注区。",
          "公式先按粗文字保留：$L(\\theta)$。",
        ],
      },
    ],
  });
  const output = path.join(deckFile.dir, "speaker-notes.pptx");
  const result = await generateDeckFile(deckFile.path, output);
  assert.equal(result.validation.errors.length, 0);

  const zip = await JSZip.loadAsync(fs.readFileSync(result.output));
  const notesXml = await zip.file("ppt/notesSlides/notesSlide1.xml").async("string");
  assert.match(notesXml, /第一段演讲稿/);
  assert.match(notesXml, /\$L\(\\theta\)\$/);
});

test("validateDeck reports missing image paths", () => {
  const deck = {
    slides: [
      {
        layout: "imageText",
        title: "Missing image",
        image: "assets/not-found.png",
        text: ["A valid explanation."],
      },
    ],
  };
  const validation = validateDeck(deck);
  assert.equal(validation.errors.length, 1);
  assert.match(validation.errors[0].message, /Image file does not exist/);
});

test("imageText placeholder expands uncertain ratio into layout variants", () => {
  const deck = {
    slides: [
      {
        layout: "imageText",
        title: "Placeholder image",
        image: {
          mode: "placeholder",
          prompt: "A workflow diagram showing YAML, validation, and PPTX generation.",
        },
        text: ["Use a placeholder first."],
      },
    ],
  };
  const result = checkDeck(deck);
  assert.equal(result.validation.errors.length, 0);
  assert.equal(result.validation.warnings.length, 0);
  assert.equal(result.inputSlides, 1);
  assert.equal(result.outputSlides, 2);
  assert.equal(result.actions[0].action, "placeholderVariants");
});

test("generateDeckFile writes editable image placeholder text", async () => {
  const deckFile = writeYamlDeck({
    slides: [
      {
        layout: "imageText",
        title: "Placeholder image",
        image: {
          mode: "placeholder",
          aspectRatio: "16:9",
          placement: "top",
          prompt: "A clean architecture diagram for the PPT generator.",
        },
        text: ["The user can replace this placeholder later."],
      },
    ],
  });
  const output = path.join(deckFile.dir, "placeholder-image.pptx");
  const result = await generateDeckFile(deckFile.path, output);
  assert.equal(result.validation.errors.length, 0);

  const zip = await JSZip.loadAsync(fs.readFileSync(result.output));
  const slideXml = await zip.file("ppt/slides/slide1.xml").async("string");
  assert.match(slideXml, /待补图片/);
  assert.match(slideXml, /A clean architecture diagram/);
});

test("imageText accepts object image syntax for landscape and side placement", () => {
  const deck = {
    slides: [
      {
        layout: "imageText",
        title: "Wide image",
        image: {
          path: "assets/bit-campus-photo.png",
          placement: "top",
          fit: "contain",
        },
        text: ["Wide image uses top placement."],
      },
      {
        layout: "imageText",
        title: "Side image",
        image: {
          path: "assets/bit-emblem-gray.png",
          placement: "side",
          fit: "contain",
        },
        text: ["Square image can stay beside text."],
      },
    ],
  };
  const result = checkDeck(deck);
  assert.equal(result.validation.errors.length, 0);
  assert.equal(result.inputSlides, 2);
});

test("CLI generate writes output for imageText object syntax", async () => {
  const deckFile = writeYamlDeck({
    slides: [
      {
        layout: "imageText",
        title: "Image layout",
        image: {
          path: "assets/bit-campus-photo.png",
          placement: "top",
          fit: "contain",
        },
        text: ["Top image.", "Bottom text."],
      },
    ],
  });
  const output = path.join(deckFile.dir, "image-layout.pptx");
  const result = await runCli(["generate", deckFile.path, output]);
  assert.equal(result.status, 0);
  assert.ok(fs.existsSync(output));
  assert.ok(fs.statSync(output).size > 1000);
});

test("CLI generate --json prints machine-readable result", async () => {
  const output = tempFile("example-json.pptx");
  const result = await runCli(["generate", "content/example.yaml", output.path, "--json"]);
  assert.equal(result.status, 0);
  assert.equal(result.stderr, "");
  const payload = JSON.parse(result.stdout);
  assert.equal(payload.validation.errors.length, 0);
  assert.equal(payload.output, output.path);
  assert.ok(fs.existsSync(output.path));
});

test("CLI generate --strict exits non-zero on warnings", async () => {
  const output = tempFile("strict-warning.pptx");
  const result = await runCli(["generate", "content/formula-test.yaml", output.path, "--strict"]);
  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /strict mode/);
  assert.equal(fs.existsSync(output.path), false);
});
