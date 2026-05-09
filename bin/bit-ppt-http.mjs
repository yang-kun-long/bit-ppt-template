#!/usr/bin/env node
import { startHttpServer } from "../src/http-server.mjs";

const args = process.argv.slice(2);
if (args.includes("--help") || args.includes("-h")) {
  console.log(`BIT PPT Generator Web UI

Usage:
  bit-ppt-generator [--host <host>] [--port <port>]
  bit-ppt-http [--host <host>] [--port <port>]

Examples:
  bit-ppt-generator
  bit-ppt-generator --port 3001
  bit-ppt-http --host 0.0.0.0 --port 3001

Environment:
  HOST                         Default: 127.0.0.1
  PORT                         Default: 3000
  BIT_PPT_AUTH_SECRET          Enables BIT login and signed web sessions
  BIT_PPT_SESSION_TTL_SECONDS  Signed web session lifetime
  BIT_PPT_TOKEN                Enables fixed Bearer-token API auth
`);
  process.exit(0);
}

const portArgIndex = process.argv.findIndex((arg) => arg === "--port" || arg === "-p");
const hostArgIndex = process.argv.findIndex((arg) => arg === "--host");
const explicitPort = portArgIndex >= 0 || Boolean(process.env.PORT);
const port = portArgIndex >= 0 ? process.argv[portArgIndex + 1] : process.env.PORT || 3000;
const host = hostArgIndex >= 0 ? process.argv[hostArgIndex + 1] : process.env.HOST || "127.0.0.1";

async function startWithFallback() {
  let currentPort = Number(port);
  const maxPort = explicitPort ? currentPort : currentPort + 10;
  while (currentPort <= maxPort) {
    try {
      await startHttpServer({ port: currentPort, host });
      return currentPort;
    } catch (error) {
      if (error.code !== "EADDRINUSE" || explicitPort || currentPort >= maxPort) throw error;
      currentPort += 1;
    }
  }
  return currentPort;
}

startWithFallback()
  .then((actualPort) => {
    const url = `http://${host}:${actualPort}/`;
    console.log(`BIT PPT Generator running at ${url}`);
    console.log("Open the URL in your browser. Local mode does not require login unless auth environment variables are set.");
    console.log("Endpoints: GET /health, POST /check, POST /generate");
  })
  .catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
