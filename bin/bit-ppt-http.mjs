#!/usr/bin/env node
import { startHttpServer } from "../src/http-server.mjs";

const portArgIndex = process.argv.findIndex((arg) => arg === "--port" || arg === "-p");
const hostArgIndex = process.argv.findIndex((arg) => arg === "--host");
const port = portArgIndex >= 0 ? process.argv[portArgIndex + 1] : process.env.PORT || 3000;
const host = hostArgIndex >= 0 ? process.argv[hostArgIndex + 1] : process.env.HOST || "127.0.0.1";

startHttpServer({ port, host })
  .then(() => {
    console.log(`bit-ppt-http listening at http://${host}:${port}`);
    console.log("Endpoints: GET /health, POST /check, POST /generate");
  })
  .catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
