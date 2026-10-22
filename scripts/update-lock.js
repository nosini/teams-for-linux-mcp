#!/usr/bin/env node
const fs = require("node:fs");
const path = require("node:path");
const { spawnSync } = require("node:child_process");
const { createHash } = require("node:crypto");
const { prepareApp } = require("./prepare-app");

try {
  const directory = prepareApp({ validateLock: false });
  // Start from upstream's new resolutions, including security fixes to
  // transitive dependencies that do not change package.json.
  const upstreamLock = fs.readFileSync(path.resolve(__dirname, "../upstream/teams-for-linux/package-lock.json"));
  fs.writeFileSync(path.join(directory, "package-lock.json"), upstreamLock);
  const result = spawnSync("npm", ["install", "--package-lock-only", "--ignore-scripts"], { cwd: directory, stdio: "inherit" });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error("npm could not update the app lockfile");
  fs.copyFileSync(path.join(directory, "package-lock.json"), path.resolve(__dirname, "../dependencies/package-lock.json"));
  fs.writeFileSync(path.resolve(__dirname, "../dependencies/upstream-lock.sha256"), `${createHash("sha256").update(upstreamLock).digest("hex")}\n`);
  prepareApp();
} catch (error) {
  console.error(error.message);
  process.exitCode = 1;
}
