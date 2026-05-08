import fs from "node:fs";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import YAML from "yaml";
import { parseDeckYaml } from "./core/yaml-parse.mjs";
import {
  checkDeck,
  generateDeckFile,
} from "./generate.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const DEFAULT_BODY_LIMIT = 10 * 1024 * 1024;
const DEFAULT_GENERATE_CONCURRENCY = 1;

function parsePositiveInt(value, fallback) {
  const number = Number(value);
  return Number.isInteger(number) && number > 0 ? number : fallback;
}

function jsonResponse(res, status, payload) {
  const body = JSON.stringify(payload, null, 2);
  res.writeHead(status, {
    "content-type": "application/json; charset=utf-8",
    "content-length": Buffer.byteLength(body),
    "access-control-allow-origin": "*",
    "access-control-allow-methods": "GET,POST,OPTIONS",
    "access-control-allow-headers": "content-type, authorization",
  });
  res.end(body);
}

function textResponse(res, status, text) {
  res.writeHead(status, {
    "content-type": "text/plain; charset=utf-8",
    "content-length": Buffer.byteLength(text),
    "access-control-allow-origin": "*",
  });
  res.end(text);
}

function pptxResponse(res, buffer, fileName) {
  res.writeHead(200, {
    "content-type": PPTX_MIME,
    "content-length": buffer.length,
    "content-disposition": `attachment; filename="${fileName}"`,
    "access-control-allow-origin": "*",
  });
  res.end(buffer);
}

function getAuthToken(options = {}) {
  return options.authToken || process.env.BIT_PPT_TOKEN || "";
}

function requireAuth(req, options = {}) {
  const token = getAuthToken(options);
  if (!token) return true;
  const authorization = req.headers.authorization || "";
  if (authorization === `Bearer ${token}`) return true;
  return false;
}

function parseBool(value) {
  return value === true || value === "true" || value === "1" || value === "yes";
}

function sanitizeFileBase(value) {
  const text = String(value || "deck").replace(/[^\w.-]+/g, "-").replace(/^-+|-+$/g, "");
  return text || "deck";
}

async function readRequestBody(req, limit = DEFAULT_BODY_LIMIT) {
  const chunks = [];
  let total = 0;
  for await (const chunk of req) {
    total += chunk.length;
    if (total > limit) {
      const error = new Error(`Request body exceeds ${limit} bytes.`);
      error.statusCode = 413;
      throw error;
    }
    chunks.push(chunk);
  }
  return Buffer.concat(chunks).toString("utf8");
}

async function parseDeckRequest(req, url, options = {}) {
  const rawBody = await readRequestBody(req, options.bodyLimit || parsePositiveInt(process.env.BIT_PPT_BODY_LIMIT, DEFAULT_BODY_LIMIT));
  const contentType = req.headers["content-type"] || "";
  if (contentType.includes("application/json")) {
    const payload = rawBody ? JSON.parse(rawBody) : {};
    if (payload.deck && typeof payload.deck === "object") {
      return {
        deck: payload.deck,
        deckYaml: YAML.stringify(payload.deck),
        outputName: sanitizeFileBase(payload.outputName || payload.fileName || payload.deck?.meta?.title),
        options: {
          strict: parseBool(payload.strict) || parseBool(url.searchParams.get("strict")),
          fontCn: payload.fontCn,
          fontCnLight: payload.fontCnLight,
          fontEn: payload.fontEn,
          fontSerif: payload.fontSerif,
          fontCode: payload.fontCode,
        },
      };
    }
    const deckYaml = payload.deckYaml || payload.yaml;
    if (typeof deckYaml === "string") {
      return {
        deck: parseDeckYaml(deckYaml, "deckYaml"),
        deckYaml,
        outputName: sanitizeFileBase(payload.outputName || payload.fileName),
        options: {
          strict: parseBool(payload.strict) || parseBool(url.searchParams.get("strict")),
          fontCn: payload.fontCn,
          fontCnLight: payload.fontCnLight,
          fontEn: payload.fontEn,
          fontSerif: payload.fontSerif,
          fontCode: payload.fontCode,
        },
      };
    }
    const error = new Error("JSON body must include deck, deckYaml, or yaml.");
    error.statusCode = 400;
    throw error;
  }
  return {
    deck: parseDeckYaml(rawBody, "request body"),
    deckYaml: rawBody,
    outputName: sanitizeFileBase(url.searchParams.get("outputName") || url.searchParams.get("fileName")),
    options: {
      strict: parseBool(url.searchParams.get("strict")),
      fontCn: url.searchParams.get("fontCn") || undefined,
      fontCnLight: url.searchParams.get("fontCnLight") || undefined,
      fontEn: url.searchParams.get("fontEn") || undefined,
      fontSerif: url.searchParams.get("fontSerif") || undefined,
      fontCode: url.searchParams.get("fontCode") || undefined,
    },
  };
}

function healthPayload(options = {}) {
  return {
    ok: true,
    service: "bit-ppt-http",
    capabilities: {
      check: true,
      generate: true,
      omml: true,
    },
    authRequired: Boolean(getAuthToken(options)),
    maxGenerateConcurrency: parsePositiveInt(options.maxGenerateConcurrency || process.env.BIT_PPT_MAX_GENERATE_CONCURRENCY, DEFAULT_GENERATE_CONCURRENCY),
  };
}

async function handleCheck(req, res, url, options) {
  const { deck } = await parseDeckRequest(req, url, options);
  jsonResponse(res, 200, checkDeck(deck));
}

async function handleGenerate(req, res, url, options) {
  const maxConcurrency = parsePositiveInt(options.maxGenerateConcurrency || process.env.BIT_PPT_MAX_GENERATE_CONCURRENCY, DEFAULT_GENERATE_CONCURRENCY);
  if (!Number.isInteger(options.activeGenerations)) options.activeGenerations = 0;
  if (options.activeGenerations >= maxConcurrency) {
    jsonResponse(res, 429, {
      generated: false,
      error: `Too many concurrent generation requests. Limit is ${maxConcurrency}.`,
    });
    return;
  }
  const parsed = await parseDeckRequest(req, url, options);
  const precheck = checkDeck(parsed.deck);
  if (precheck.validation.errors.length || (parsed.options.strict && precheck.validation.warnings.length)) {
    jsonResponse(res, 422, {
      generated: false,
      error: parsed.options.strict && precheck.validation.warnings.length ? "Deck validation failed strict mode." : "Deck validation failed.",
      validation: precheck.validation,
      repairPrompt: precheck.repairPrompt,
    });
    return;
  }

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "bit-ppt-http-"));
  options.activeGenerations += 1;
  try {
    const inputPath = path.join(tempDir, "deck.yaml");
    const outputPath = path.join(tempDir, `${parsed.outputName}.pptx`);
    fs.writeFileSync(inputPath, parsed.deckYaml, "utf8");
    const result = await generateDeckFile(inputPath, outputPath, parsed.options);
    const buffer = fs.readFileSync(result.output);
    pptxResponse(res, buffer, path.basename(result.output));
  } finally {
    options.activeGenerations -= 1;
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
}

async function handleRequest(req, res, options = {}) {
  const url = new URL(req.url || "/", "http://localhost");
  try {
    if (req.method === "OPTIONS") {
      jsonResponse(res, 200, { ok: true });
      return;
    }
    if (req.method === "GET" && url.pathname === "/health") {
      jsonResponse(res, 200, healthPayload(options));
      return;
    }
    if (["/check", "/generate"].includes(url.pathname) && !requireAuth(req, options)) {
      jsonResponse(res, 401, { error: "Unauthorized." });
      return;
    }
    if (req.method === "POST" && url.pathname === "/check") {
      await handleCheck(req, res, url, options);
      return;
    }
    if (req.method === "POST" && url.pathname === "/generate") {
      await handleGenerate(req, res, url, options);
      return;
    }
    if (["/check", "/generate"].includes(url.pathname)) {
      jsonResponse(res, 405, { error: "Method not allowed." });
      return;
    }
    textResponse(res, 404, "Not found. Use GET /health, POST /check, or POST /generate.");
  } catch (error) {
    if (error.kind === "yaml_syntax") {
      jsonResponse(res, error.statusCode || 400, {
        error: error.message || "YAML syntax error.",
        syntax: error.syntax,
        repairPrompt: error.repairPrompt,
      });
      return;
    }
    if (error.validation) {
      jsonResponse(res, 422, {
        generated: false,
        error: error.message || "Deck validation failed.",
        validation: error.validation,
        repairPrompt: error.repairPrompt,
      });
      return;
    }
    jsonResponse(res, error.statusCode || 400, { error: error.message || String(error) });
  }
}

function createBitPptHttpServer(options = {}) {
  const state = {
    ...options,
    activeGenerations: 0,
  };
  return http.createServer((req, res) => {
    handleRequest(req, res, state);
  });
}

async function startHttpServer(options = {}) {
  const port = Number(options.port || process.env.PORT || 3000);
  const host = options.host || process.env.HOST || "127.0.0.1";
  const server = createBitPptHttpServer(options);
  await new Promise((resolve) => server.listen(port, host, resolve));
  return server;
}

async function main() {
  const portArgIndex = process.argv.findIndex((arg) => arg === "--port" || arg === "-p");
  const hostArgIndex = process.argv.findIndex((arg) => arg === "--host");
  const port = portArgIndex >= 0 ? process.argv[portArgIndex + 1] : process.env.PORT || 3000;
  const host = hostArgIndex >= 0 ? process.argv[hostArgIndex + 1] : process.env.HOST || "127.0.0.1";
  await startHttpServer({ port, host });
  console.log(`bit-ppt-http listening at http://${host}:${port}`);
  console.log("Endpoints: GET /health, POST /check, POST /generate");
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
}

export {
  PPTX_MIME,
  createBitPptHttpServer,
  handleRequest,
  startHttpServer,
};
