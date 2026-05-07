#!/usr/bin/env node
import path from "node:path";
import {
  checkDeckFile,
  defaultFont,
  generateDeckFile,
  listLayouts,
} from "../src/generate.mjs";
import {
  getGuideOverview,
  getLayoutExample,
  getLayoutGuide,
  getLayoutSchema,
  getWritingRules,
  listGuideLayouts,
} from "../src/layout-guides.mjs";

function printHelp() {
  console.log(`bit-ppt

Usage:
  bit-ppt generate <input.yaml> <output.pptx> [font options]
  bit-ppt check <input.yaml> [--json]
  bit-ppt list-layouts [--json]
  bit-ppt guide [topic] [name] [--json]

Progressive guide:
  bit-ppt guide
  bit-ppt guide layouts
  bit-ppt guide layout imageText
  bit-ppt guide schema imageText --json
  bit-ppt guide example imageText
  bit-ppt guide writing-rules

Font options:
  --font-cn <name>        Chinese/CJK font, default: ${defaultFont.cn}
  --font-cn-light <name>  Light Chinese/CJK font, default: ${defaultFont.cnLight}
  --font-en <name>        Latin font, default: ${defaultFont.en}
  --font-serif <name>     Serif font, default: ${defaultFont.serif}
  --font-code <name>      Code font, default: ${defaultFont.code}
`);
}

function printJson(value) {
  console.log(JSON.stringify(value, null, 2));
}

function printGuideOverview(overview) {
  console.log(`${overview.name}

${overview.purpose}

Workflow:
${overview.workflow.map((item, idx) => `  ${idx + 1}. ${item}`).join("\n")}

Guide commands:
${overview.commands.map((item) => `  ${item}`).join("\n")}

Guided layouts:
  ${overview.guidedLayouts.join(", ")}
`);
}

function printLayoutGuide(guide) {
  console.log(`${guide.layout}

Purpose:
  ${guide.purpose}

Use when:
  ${guide.whenToUse}

Fields:
${Object.entries(guide.fields).map(([name, spec]) => {
    const required = spec.required ? "required" : "optional";
    const value = spec.value ? ` = ${spec.value}` : "";
    return `  ${name}: ${spec.type}${value} (${required})`;
  }).join("\n")}

Limits:
${Object.entries(guide.limits || {}).map(([name, spec]) => `  ${name}: ${Object.entries(spec).map(([key, value]) => `${key}=${value}`).join(", ")}`).join("\n") || "  none"}

Notes:
${guide.notes.map((item) => `  - ${item}`).join("\n")}
`);
}

function printExample(example) {
  console.log(JSON.stringify(example, null, 2));
}

function printWritingRules(rules) {
  console.log(`Writing rules:
${rules.map((item) => `  - ${item}`).join("\n")}
`);
}

function printAvailableGuideLayouts() {
  console.log(listGuideLayouts().join("\n"));
}

function printGuideHelp() {
  console.log(`bit-ppt guide

Usage:
  bit-ppt guide
  bit-ppt guide layouts
  bit-ppt guide layout <name> [--json]
  bit-ppt guide schema <name> --json
  bit-ppt guide example <name> [--json]
  bit-ppt guide writing-rules [--json]
`);
}

function parseOptions(args) {
  const options = {};
  const positional = [];
  for (let idx = 0; idx < args.length; idx += 1) {
    const arg = args[idx];
    if (arg === "--json") options.json = true;
    else if (arg === "--font-cn" || arg === "--font-cjk") options.fontCn = args[++idx];
    else if (arg === "--font-cn-light") options.fontCnLight = args[++idx];
    else if (arg === "--font-en" || arg === "--font-latin") options.fontEn = args[++idx];
    else if (arg === "--font-serif") options.fontSerif = args[++idx];
    else if (arg === "--font-code") options.fontCode = args[++idx];
    else if (arg === "-h" || arg === "--help") options.help = true;
    else positional.push(arg);
  }
  return { options, positional };
}

function printCheck(result, asJson) {
  if (asJson) {
    console.log(JSON.stringify(result, null, 2));
    return;
  }
  const errors = result.validation.errors.length;
  const warnings = result.validation.warnings.length;
  console.log(`slides: ${result.inputSlides} -> ${result.outputSlides}`);
  console.log(`errors: ${errors}`);
  console.log(`warnings: ${warnings}`);
  if (result.actions.length) {
    console.log(`actions: ${result.actions.map((item) => `${item.title}->${item.parts}`).join(", ")}`);
  }
  if (result.repairPrompt) {
    console.log("\nrepair prompt:");
    console.log(result.repairPrompt);
  }
}

function printGuide(command, name, asJson) {
  if (!command) {
    const overview = getGuideOverview();
    if (asJson) printJson(overview);
    else printGuideOverview(overview);
    return;
  }

  if (command === "help" || command === "-h" || command === "--help") {
    printGuideHelp();
    return;
  }

  if (command === "layouts") {
    const layouts = listGuideLayouts();
    if (asJson) printJson(layouts);
    else printAvailableGuideLayouts();
    return;
  }

  if (command === "layout") {
    if (!name) throw new Error("Usage: bit-ppt guide layout <name>");
    const guide = getLayoutGuide(name);
    if (!guide) throw new Error(`No guide available for layout: ${name}`);
    if (asJson) printJson(guide);
    else printLayoutGuide(guide);
    return;
  }

  if (command === "schema") {
    if (!name) throw new Error("Usage: bit-ppt guide schema <name>");
    const schema = getLayoutSchema(name);
    if (!schema) throw new Error(`No schema available for layout: ${name}`);
    printJson(schema);
    return;
  }

  if (command === "example") {
    if (!name) throw new Error("Usage: bit-ppt guide example <name>");
    const example = getLayoutExample(name);
    if (!example) throw new Error(`No example available for layout: ${name}`);
    if (asJson) printJson(example);
    else printExample(example);
    return;
  }

  if (command === "writing-rules") {
    const rules = getWritingRules();
    if (asJson) printJson(rules);
    else printWritingRules(rules);
    return;
  }

  throw new Error(`Unknown guide topic: ${command}`);
}

async function main() {
  const [command, ...rest] = process.argv.slice(2);
  const { options, positional } = parseOptions(rest);
  if (!command || command === "help" || command === "-h" || command === "--help" || options.help) {
    printHelp();
    return;
  }

  if (command === "list-layouts") {
    const layouts = listLayouts();
    if (options.json) console.log(JSON.stringify(layouts, null, 2));
    else console.log(layouts.join("\n"));
    return;
  }

  if (command === "guide") {
    printGuide(positional[0], positional[1], options.json);
    return;
  }

  if (command === "check") {
    const input = positional[0];
    if (!input) throw new Error("Missing input YAML path.");
    const result = checkDeckFile(path.resolve(input));
    printCheck(result, options.json);
    if (result.validation.errors.length) process.exitCode = 1;
    return;
  }

  if (command === "generate") {
    const [input, output] = positional;
    if (!input || !output) throw new Error("Usage: bit-ppt generate <input.yaml> <output.pptx>");
    const result = await generateDeckFile(path.resolve(input), path.resolve(output), options);
    if (result.validation.warnings.length) {
      console.warn(`Validation warning(s): ${result.validation.warnings.length}`);
      console.warn(result.repairPrompt);
    }
    if (result.preflight.length) {
      console.log(`Preflight adjusted ${result.preflight.length} slide(s): ${result.preflight.map((item) => `${item.title}->${item.parts}`).join(", ")}`);
    }
    console.log(`Generated ${result.output}`);
    console.log(`Fonts: cn=${result.fonts.cn}, en=${result.fonts.en}, serif=${result.fonts.serif}, code=${result.fonts.code}`);
    return;
  }

  throw new Error(`Unknown command: ${command}`);
}

main().catch((error) => {
  if (error.validation) {
    console.error(JSON.stringify({ error: error.message, validation: error.validation, repairPrompt: error.repairPrompt }, null, 2));
  } else {
    console.error(error.message || error);
  }
  process.exitCode = 1;
});
