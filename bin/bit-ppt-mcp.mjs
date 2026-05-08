#!/usr/bin/env node
import { startStdioMcpServer } from "../src/mcp-server.mjs";

if (process.argv.includes("--help") || process.argv.includes("-h")) {
  console.log(`bit-ppt-mcp

Usage:
  bit-ppt-mcp

Starts the BIT PPT MCP server over stdio. Configure your MCP client to launch
this command from the project root or from the installed npm package.

Tools:
  list_layouts
  get_guide
  validate_deck
  preflight_deck
  get_repair_prompt
  generate_pptx
`);
  process.exit(0);
}

startStdioMcpServer().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
