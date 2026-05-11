import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import test from "node:test";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
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
import {
  PPTX_MIME,
  createBitPptHttpServer,
} from "../src/http-server.mjs";

const execFileAsync = promisify(execFile);
const CLI = path.join(ROOT, "bin", "bit-ppt.mjs");
const MCP_CLI = path.join(ROOT, "bin", "bit-ppt-mcp.mjs");

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

function stripNodeDeprecationWarnings(stderr) {
  return stderr
    .split(/\r?\n/)
    .filter((line) => !/^\(node:\d+\) \[DEP\d+\] DeprecationWarning:/.test(line))
    .filter((line) => !/^\(Use `node --trace-deprecation/.test(line))
    .join("\n")
    .trim();
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

async function withMcpClient(callback) {
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: [MCP_CLI],
    cwd: ROOT,
    stderr: "pipe",
  });
  const client = new Client({ name: "bit-ppt-test-client", version: "1.0.0" });
  await client.connect(transport);
  try {
    return await callback(client);
  } finally {
    await client.close();
  }
}

async function withHttpServer(callback, options = {}) {
  const server = createBitPptHttpServer(options);
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  try {
    const address = server.address();
    return await callback(`http://127.0.0.1:${address.port}`);
  } finally {
    await new Promise((resolve, reject) => {
      server.close((error) => error ? reject(error) : resolve());
    });
  }
}

function base64UrlEncode(value) {
  return Buffer.from(value).toString("base64").replaceAll("+", "-").replaceAll("/", "_").replaceAll("=", "");
}

function signedAuthToken(secret, payload = {}) {
  const now = Math.floor(Date.now() / 1000);
  const header = base64UrlEncode(JSON.stringify({ alg: "HS256", typ: "JWT" }));
  const body = base64UrlEncode(JSON.stringify({
    sub: "test-user",
    aud: "bit-ppt",
    iat: now,
    exp: now + 300,
    ...payload,
  }));
  const input = `${header}.${body}`;
  const sig = crypto.createHmac("sha256", secret).update(input).digest("base64url");
  return `${input}.${sig}`;
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

test("MCP server exposes core tools", async () => {
  await withMcpClient(async (client) => {
    const result = await client.listTools();
    const names = result.tools.map((tool) => tool.name);
    assert.ok(names.includes("list_layouts"));
    assert.ok(names.includes("validate_deck"));
    assert.ok(names.includes("preflight_deck"));
    assert.ok(names.includes("generate_pptx"));
    assert.ok(names.includes("get_guide"));
  });
});

test("MCP preflight_deck accepts raw YAML", async () => {
  await withMcpClient(async (client) => {
    const deckYaml = YAML.stringify({
      slides: [
        {
          layout: "imageText",
          title: "Placeholder",
          image: {
            mode: "placeholder",
            prompt: "A generated diagram placeholder.",
          },
          text: ["Use placeholder first."],
        },
      ],
    });
    const result = await client.callTool({
      name: "preflight_deck",
      arguments: { deckYaml },
    });
    assert.equal(result.structuredContent.outputSlides, 2);
    assert.equal(result.structuredContent.actions[0].action, "placeholderVariants");
  });
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

test("HTTP /check matches checkDeck for raw YAML", async () => {
  await withHttpServer(async (baseUrl) => {
    const deck = {
      slides: [
        {
          layout: "imageText",
          title: "HTTP check",
          image: {
            mode: "placeholder",
            prompt: "A generated diagram placeholder.",
          },
          text: ["Use shared validation."],
        },
      ],
    };
    const deckYaml = YAML.stringify(deck);
    const response = await fetch(`${baseUrl}/check`, {
      method: "POST",
      headers: { "content-type": "application/yaml" },
      body: deckYaml,
    });
    assert.equal(response.status, 200);
    assert.deepEqual(await response.json(), checkDeck(deck));
  });
});

test("HTTP /check returns line and column for YAML syntax errors", async () => {
  await withHttpServer(async (baseUrl) => {
    const response = await fetch(`${baseUrl}/check`, {
      method: "POST",
      headers: { "content-type": "application/yaml" },
      body: "slides:\n  - layout bullets\n    title: Broken\n",
    });
    assert.equal(response.status, 400);
    const payload = await response.json();
    assert.match(payload.error, /line \d+, column \d+/);
    assert.equal(payload.syntax.errors[0].level, "error");
    assert.ok(Number.isInteger(payload.syntax.errors[0].line));
    assert.ok(Number.isInteger(payload.syntax.errors[0].column));
    assert.match(payload.syntax.errors[0].context, /layout bullets/);
    assert.match(payload.repairPrompt, /Fix the YAML syntax first/);
  });
});

test("HTTP /generate returns a valid PPTX download", async () => {
  await withHttpServer(async (baseUrl) => {
    const deckYaml = YAML.stringify({
      meta: {
        title: "HTTP PPTX",
      },
      slides: [
        {
          layout: "bullets",
          title: "HTTP 生成",
          bullets: ["上传 YAML。", "下载 PPTX。"],
        },
      ],
    });
    const response = await fetch(`${baseUrl}/generate`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ deckYaml, outputName: "http-test" }),
    });
    assert.equal(response.status, 200);
    assert.equal(response.headers.get("content-type"), PPTX_MIME);
    assert.match(response.headers.get("content-disposition"), /http-test\.pptx/);
    const buffer = Buffer.from(await response.arrayBuffer());
    assert.ok(buffer.length > 1000);
    const zip = await JSZip.loadAsync(buffer);
    assert.ok(zip.file("ppt/presentation.xml"));
  });
});

test("HTTP token auth protects check and generate when configured", async () => {
  await withHttpServer(async (baseUrl) => {
    const deckYaml = YAML.stringify({
      slides: [
        {
          layout: "bullets",
          title: "Auth check",
          bullets: ["Protected endpoint."],
        },
      ],
    });
    const health = await fetch(`${baseUrl}/health`);
    assert.equal((await health.json()).authRequired, true);

    const unauthorized = await fetch(`${baseUrl}/check`, {
      method: "POST",
      headers: { "content-type": "application/yaml" },
      body: deckYaml,
    });
    assert.equal(unauthorized.status, 401);

    const authorized = await fetch(`${baseUrl}/check`, {
      method: "POST",
      headers: {
        "authorization": "Bearer test-token",
        "content-type": "application/yaml",
      },
      body: deckYaml,
    });
    assert.equal(authorized.status, 200);
    assert.equal((await authorized.json()).validation.errors.length, 0);
  }, { authToken: "test-token" });
});

test("HTTP signed auth token is accepted when configured", async () => {
  const secret = "signed-test-secret";
  await withHttpServer(async (baseUrl) => {
    const deckYaml = YAML.stringify({
      slides: [
        {
          layout: "bullets",
          title: "Signed auth check",
          bullets: ["Protected endpoint."],
        },
      ],
    });
    const health = await fetch(`${baseUrl}/health`);
    const healthPayload = await health.json();
    assert.equal(healthPayload.authRequired, true);
    assert.equal(healthPayload.signedAuthRequired, true);

    const emptyBearer = await fetch(`${baseUrl}/check`, {
      method: "POST",
      headers: {
        "authorization": "Bearer ",
        "content-type": "application/yaml",
      },
      body: deckYaml,
    });
    assert.equal(emptyBearer.status, 401);

    const authorized = await fetch(`${baseUrl}/check`, {
      method: "POST",
      headers: {
        "authorization": `Bearer ${signedAuthToken(secret)}`,
        "content-type": "application/yaml",
      },
      body: deckYaml,
    });
    assert.equal(authorized.status, 200);
    assert.equal((await authorized.json()).validation.errors.length, 0);
  }, { authSigningSecret: secret });
});

test("HTTP BIT login issues signed token for protected endpoints", async () => {
  const secret = "bit-login-test-secret";
  await withHttpServer(async (baseUrl) => {
    const page = await fetch(`${baseUrl}/`);
    assert.equal(page.status, 200);
    assert.match(await page.text(), /BIT PPT Generator/);

    const badLogin = await fetch(`${baseUrl}/auth/bit-login`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ username: "u", password: "wrong" }),
    });
    assert.equal(badLogin.status, 401);

    const login = await fetch(`${baseUrl}/auth/bit-login`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ username: "u", password: "right" }),
    });
    assert.equal(login.status, 200);
    const session = await login.json();
    assert.equal(session.tokenType, "Bearer");
    assert.ok(session.token);

    const deckYaml = YAML.stringify({
      slides: [{ layout: "bullets", title: "BIT login", bullets: ["ok"] }],
    });
    const check = await fetch(`${baseUrl}/check`, {
      method: "POST",
      headers: {
        "authorization": `Bearer ${session.token}`,
        "content-type": "application/yaml",
      },
      body: deckYaml,
    });
    assert.equal(check.status, 200);
    assert.equal((await check.json()).validation.errors.length, 0);
  }, {
    authSigningSecret: secret,
    bitAuthVerifier: async (username, password) => username === "u" && password === "right",
  });
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
  assert.equal(stripNodeDeprecationWarnings(result.stderr), "");
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
