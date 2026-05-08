import fs from "node:fs";
import path from "node:path";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import * as z from "zod/v4";
import { parseDeckYaml } from "./core/yaml-parse.mjs";
import {
  ROOT,
  checkDeck,
  checkDeckFile,
  generateDeckFile,
  listLayouts,
  validateDeck,
} from "./generate.mjs";
import {
  getAllGuides,
  getGuideOverview,
  getGuideWorkflow,
  getImagePlaceholderGuide,
  getLayoutExample,
  getLayoutGuide,
  getLayoutSchema,
  getSpeakerNotesGuide,
  getWritingRules,
  listGuideLayouts,
} from "./layout-guides.mjs";

function packageVersion() {
  try {
    const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, "package.json"), "utf8"));
    return pkg.version || "0.0.0";
  } catch {
    return "0.0.0";
  }
}

function resolveLocalPath(value) {
  const source = String(value || "");
  return path.isAbsolute(source) ? source : path.resolve(ROOT, source);
}

function loadDeckArgs(args) {
  if (args.deckYaml) return parseDeckYaml(args.deckYaml, "deckYaml");
  if (args.inputPath) return parseDeckYaml(fs.readFileSync(resolveLocalPath(args.inputPath), "utf8"), args.inputPath);
  throw new Error("Provide either inputPath or deckYaml.");
}

function textResult(data) {
  return {
    structuredContent: data,
    content: [{ type: "text", text: JSON.stringify(data, null, 2) }],
  };
}

function guideResult(topic = "overview", name) {
  switch (topic) {
    case "overview":
      return getGuideOverview();
    case "all":
      return getAllGuides();
    case "workflow":
      return getGuideWorkflow();
    case "layouts":
      return listGuideLayouts();
    case "layout": {
      if (!name) throw new Error("get_guide topic=layout requires name.");
      const guide = getLayoutGuide(name);
      if (!guide) throw new Error(`No guide available for layout: ${name}`);
      return guide;
    }
    case "schema": {
      if (!name) throw new Error("get_guide topic=schema requires name.");
      const schema = getLayoutSchema(name);
      if (!schema) throw new Error(`No schema available for layout: ${name}`);
      return schema;
    }
    case "example": {
      if (!name) throw new Error("get_guide topic=example requires name.");
      const example = getLayoutExample(name);
      if (!example) throw new Error(`No example available for layout: ${name}`);
      return example;
    }
    case "speaker-notes":
      return getSpeakerNotesGuide();
    case "image-placeholder":
      return getImagePlaceholderGuide();
    case "writing-rules":
      return getWritingRules();
    default:
      throw new Error(`Unknown guide topic: ${topic}`);
  }
}

const deckInputSchema = {
  inputPath: z.string().optional().describe("Path to a YAML deck file. Relative paths resolve from the project root."),
  deckYaml: z.string().optional().describe("Raw YAML deck content. Use this instead of inputPath for generated drafts."),
};

const fontOptionsSchema = {
  fontCn: z.string().optional().describe("Chinese/CJK font override."),
  fontCnLight: z.string().optional().describe("Light Chinese/CJK font override."),
  fontEn: z.string().optional().describe("Latin font override."),
  fontSerif: z.string().optional().describe("Serif font override."),
  fontCode: z.string().optional().describe("Code font override."),
};

export function createBitPptMcpServer() {
  const server = new McpServer({
    name: "bit-ppt-template",
    version: packageVersion(),
  });

  server.registerTool("list_layouts", {
    title: "List Layouts",
    description: "List supported BIT PPT slide layouts.",
    inputSchema: {},
  }, async () => textResult({ layouts: listLayouts() }));

  server.registerTool("get_guide", {
    title: "Get Progressive Guide",
    description: "Return focused guide content for layouts, schemas, examples, speaker notes, image placeholders, or writing rules.",
    inputSchema: {
      topic: z.enum(["overview", "all", "workflow", "layouts", "layout", "schema", "example", "speaker-notes", "image-placeholder", "writing-rules"]).optional().default("overview"),
      name: z.string().optional().describe("Layout name for topic layout, schema, or example."),
    },
  }, async ({ topic, name }) => textResult(guideResult(topic, name)));

  server.registerTool("validate_deck", {
    title: "Validate Deck",
    description: "Validate a YAML deck and return validation errors, warnings, and repairPrompt.",
    inputSchema: deckInputSchema,
  }, async (args) => {
    const deck = loadDeckArgs(args);
    const validation = validateDeck(deck);
    return textResult({
      validation: {
        errors: validation.errors,
        warnings: validation.warnings,
      },
      repairPrompt: validation.repairPrompt,
    });
  });

  server.registerTool("preflight_deck", {
    title: "Preflight Deck",
    description: "Run validation plus preflight planning such as automatic splits and placeholder layout variants.",
    inputSchema: deckInputSchema,
  }, async (args) => textResult(checkDeck(loadDeckArgs(args))));

  server.registerTool("get_repair_prompt", {
    title: "Get Repair Prompt",
    description: "Return only the repairPrompt for a YAML deck, plus validation counts.",
    inputSchema: deckInputSchema,
  }, async (args) => {
    const result = checkDeck(loadDeckArgs(args));
    return textResult({
      repairPrompt: result.repairPrompt,
      errors: result.validation.errors.length,
      warnings: result.validation.warnings.length,
    });
  });

  server.registerTool("generate_pptx", {
    title: "Generate PPTX",
    description: "Generate an editable PPTX from a YAML file. Generation stops on validation errors; strict mode also stops on warnings.",
    inputSchema: {
      inputPath: z.string().describe("Path to the input YAML deck file. Relative paths resolve from the project root."),
      outputPath: z.string().describe("Path to the output PPTX file. Relative paths resolve from the project root."),
      strict: z.boolean().optional().describe("Treat validation warnings as failures before generation."),
      ...fontOptionsSchema,
    },
  }, async (args) => {
    const input = resolveLocalPath(args.inputPath);
    const output = resolveLocalPath(args.outputPath);
    if (args.strict) {
      const check = checkDeckFile(input);
      if (check.validation.errors.length || check.validation.warnings.length) {
        return textResult({
          generated: false,
          error: "Deck validation failed strict mode.",
          validation: check.validation,
          repairPrompt: check.repairPrompt,
        });
      }
    }
    const result = await generateDeckFile(input, output, args);
    return textResult({ generated: true, ...result });
  });

  return server;
}

export async function startStdioMcpServer() {
  const server = createBitPptMcpServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
}
