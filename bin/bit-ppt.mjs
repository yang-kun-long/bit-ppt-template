#!/usr/bin/env node
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import {
  ROOT,
  checkDeckFile,
  defaultFont,
  generateDeckFile,
  listLayouts,
} from "../src/generate.mjs";
import {
  getAllGuides,
  getGuideOverview,
  getGuideWorkflow,
  getLayoutExample,
  getLayoutGuide,
  getLayoutSchema,
  getWritingRules,
  listGuideLayouts,
} from "../src/layout-guides.mjs";

function printHelp() {
  console.log(`bit-ppt

Usage:
  bit-ppt generate <input.yaml> <output.pptx> [--json] [--strict] [font options]
  bit-ppt check <input.yaml> [--json] [--strict]
  bit-ppt list-layouts [--json]
  bit-ppt guide [topic] [name] [--json]
  bit-ppt doctor [--json]

Progressive guide:
  bit-ppt guide
  bit-ppt guide all --json
  bit-ppt guide workflow --json
  bit-ppt guide layouts
  bit-ppt guide layout imageText
  bit-ppt guide schema imageText --json
  bit-ppt guide example imageText
  bit-ppt guide writing-rules

Quality options:
  --json                  Print machine-readable JSON where supported
  --strict                Treat validation warnings as failures

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

function printWorkflow(workflow) {
  console.log(`Workflow:
${workflow.map((item, idx) => `  ${idx + 1}. ${item}`).join("\n")}
`);
}

function printGuideHelp() {
  console.log(`bit-ppt guide

Usage:
  bit-ppt guide
  bit-ppt guide all --json
  bit-ppt guide workflow [--json]
  bit-ppt guide layouts
  bit-ppt guide layout <name> [--json]
  bit-ppt guide schema <name> --json
  bit-ppt guide example <name> [--json]
  bit-ppt guide writing-rules [--json]
`);
}

function doctorCheck(name, ok, message, details = {}) {
  return { name, ok, message, ...details };
}

function canImportPackage(packageName) {
  try {
    import.meta.resolve(packageName);
    return true;
  } catch {
    return false;
  }
}

function runDoctor() {
  const checks = [];
  const version = process.versions.node;
  const major = Number(version.split(".")[0]);
  checks.push(doctorCheck("node", major >= 18, `Node.js ${version}`, { version }));

  ["pptxgenjs", "yaml", "jszip", "latex-to-omml"].forEach((packageName) => {
    const available = canImportPackage(packageName);
    checks.push(doctorCheck(`dependency:${packageName}`, available, available ? `${packageName} is available` : `${packageName} is missing`));
  });

  ["assets/bit-campus-photo.png", "assets/bit-campus-line.png", "content/example.yaml", "content/body-layout-test.yaml", "content/chart-flow-test.yaml"].forEach((relativePath) => {
    const fileName = path.join(ROOT, relativePath);
    checks.push(doctorCheck(`file:${relativePath}`, fs.existsSync(fileName), fs.existsSync(fileName) ? "found" : "missing", { path: fileName }));
  });

  [
    ["fixture:example", "content/example.yaml"],
    ["fixture:body-layouts", "content/body-layout-test.yaml"],
    ["fixture:charts", "content/chart-flow-test.yaml"],
  ].forEach(([name, relativePath]) => {
    try {
      const result = checkDeckFile(path.join(ROOT, relativePath));
      checks.push(doctorCheck(name, result.validation.errors.length === 0, result.validation.errors.length ? "validation errors found" : "check passed", {
        warnings: result.validation.warnings.length,
        errors: result.validation.errors.length,
      }));
    } catch (error) {
      checks.push(doctorCheck(name, false, error.message || String(error)));
    }
  });

  try {
    const outputDir = path.join(ROOT, "output");
    fs.mkdirSync(outputDir, { recursive: true });
    const probe = path.join(outputDir, `.bit-ppt-doctor-${process.pid}.tmp`);
    fs.writeFileSync(probe, "ok");
    fs.unlinkSync(probe);
    checks.push(doctorCheck("output:writable", true, "output directory is writable", { path: outputDir }));
  } catch (error) {
    checks.push(doctorCheck("output:writable", false, error.message || String(error)));
  }

  const ok = checks.every((item) => item.ok);
  return {
    ok,
    platform: process.platform,
    arch: process.arch,
    cwd: process.cwd(),
    root: ROOT,
    checks,
  };
}

function printDoctor(result) {
  console.log(`BIT PPT doctor: ${result.ok ? "ok" : "failed"}

Root:
  ${result.root}

Checks:
${result.checks.map((item) => `  ${item.ok ? "OK" : "FAIL"} ${item.name}: ${item.message}`).join("\n")}
`);
}

function parseOptions(args) {
  const options = {};
  const positional = [];
  for (let idx = 0; idx < args.length; idx += 1) {
    const arg = args[idx];
    if (arg === "--json") options.json = true;
    else if (arg === "--strict") options.strict = true;
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

function hasStrictFailure(result, options) {
  return Boolean(options.strict && result.validation.warnings.length);
}

function applyCheckExitCode(result, options) {
  if (result.validation.errors.length || hasStrictFailure(result, options)) {
    process.exitCode = 1;
  }
}

function makeStrictWarningError(result) {
  const error = new Error("Deck validation warning(s) failed strict mode.");
  error.validation = result.validation;
  error.repairPrompt = result.repairPrompt;
  return error;
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

  if (command === "all") {
    const guides = getAllGuides();
    if (asJson) printJson(guides);
    else printGuideOverview(guides.overview);
    return;
  }

  if (command === "workflow") {
    const workflow = getGuideWorkflow();
    if (asJson) printJson(workflow);
    else printWorkflow(workflow);
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

  if (command === "doctor") {
    const result = runDoctor();
    if (options.json) printJson(result);
    else printDoctor(result);
    if (!result.ok) process.exitCode = 1;
    return;
  }

  if (command === "check") {
    const input = positional[0];
    if (!input) throw new Error("Missing input YAML path.");
    const result = checkDeckFile(path.resolve(input));
    printCheck(result, options.json);
    applyCheckExitCode(result, options);
    return;
  }

  if (command === "generate") {
    const [input, output] = positional;
    if (!input || !output) throw new Error("Usage: bit-ppt generate <input.yaml> <output.pptx>");
    if (options.strict) {
      const strictCheck = checkDeckFile(path.resolve(input));
      if (hasStrictFailure(strictCheck, options)) throw makeStrictWarningError(strictCheck);
    }
    const result = await generateDeckFile(path.resolve(input), path.resolve(output), options);
    if (options.json) {
      printJson(result);
      return;
    }
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
