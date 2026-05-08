import YAML, { LineCounter } from "yaml";

function lineContext(source, line, radius = 2) {
  const lines = String(source || "").replace(/\r\n?/g, "\n").split("\n");
  const start = Math.max(1, line - radius);
  const end = Math.min(lines.length, line + radius);
  const width = String(end).length;
  const context = [];
  for (let current = start; current <= end; current += 1) {
    context.push(`${String(current).padStart(width, " ")} | ${lines[current - 1] ?? ""}`);
  }
  return context.join("\n");
}

function pointerLine(column) {
  return `${" ".repeat(Math.max(0, column - 1))}^`;
}

function makeYamlSyntaxError(source, errors, label = "YAML") {
  const diagnostics = errors.map((error) => {
    const start = error.linePos?.[0] || { line: 1, col: 1 };
    const end = error.linePos?.[1] || start;
    return {
      level: "error",
      code: error.code || error.name || "YAML_PARSE_ERROR",
      message: error.message,
      line: start.line,
      column: start.col,
      endLine: end.line,
      endColumn: end.col,
      context: lineContext(source, start.line),
      pointer: pointerLine(start.col),
    };
  });
  const first = diagnostics[0] || {
    code: "YAML_PARSE_ERROR",
    message: "YAML could not be parsed.",
    line: 1,
    column: 1,
    context: "",
    pointer: "^",
  };
  const repairPrompt = diagnostics
    .map((item) => [
      `${label} syntax error at line ${item.line}, column ${item.column}: ${item.message}`,
      item.context,
      item.pointer,
      "Fix the YAML syntax first, then keep the deck schema unchanged.",
    ].filter(Boolean).join("\n"))
    .join("\n\n");
  const wrapped = new Error(`${label} syntax error at line ${first.line}, column ${first.column}: ${first.message}`);
  wrapped.statusCode = 400;
  wrapped.kind = "yaml_syntax";
  wrapped.syntax = {
    errors: diagnostics,
  };
  wrapped.repairPrompt = repairPrompt;
  return wrapped;
}

function parseDeckYaml(source, label = "YAML") {
  const lineCounter = new LineCounter();
  const doc = YAML.parseDocument(source, {
    lineCounter,
    prettyErrors: false,
  });
  if (doc.errors.length) {
    for (const error of doc.errors) {
      if (Array.isArray(error.pos)) {
        error.linePos = error.pos.map((pos) => lineCounter.linePos(pos));
      }
    }
    throw makeYamlSyntaxError(source, doc.errors, label);
  }
  return doc.toJSON();
}

export {
  parseDeckYaml,
};
