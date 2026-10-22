#!/usr/bin/env node
const fs = require("node:fs");
const path = require("node:path");
const { spawnSync } = require("node:child_process");
const { prepareApp } = require("./prepare-app");

try {
  const [script, ...args] = process.argv.slice(2);
  if (!script) throw new Error("Usage: node scripts/run-app.js NPM-SCRIPT [ARGUMENTS...]");
  const directory = prepareApp();
  if (!fs.existsSync(path.join(directory, "node_modules"))) throw new Error("Install the app dependencies first: npm run install:app");
  const result = spawnSync("npm", ["run", script, "--", ...args], { cwd: directory, stdio: "inherit" });
  if (result.error) throw result.error;
  process.exitCode = result.status ?? 1;
} catch (error) {
  console.error(error.message);
  process.exitCode = 1;
}
