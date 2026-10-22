#!/usr/bin/env node
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { spawnSync } = require("node:child_process");
const { createHash } = require("node:crypto");

const projectRoot = path.resolve(__dirname, "..");

function merge(base, overrides) {
  const result = structuredClone(base);
  for (const [key, value] of Object.entries(overrides)) {
    if (value === null) delete result[key];
    else if (typeof value === "object" && !Array.isArray(value)) result[key] = merge(result[key] || {}, value);
    else result[key] = value;
  }
  return result;
}

function run(command, args, cwd, options = {}) {
  const result = spawnSync(command, args, { cwd, encoding: "utf8", ...options });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error(`${command} ${args.join(" ")} failed:\n${result.stderr || result.stdout}`);
  return result.stdout;
}

function overlay(source, destination) {
  fs.mkdirSync(destination, { recursive: true });
  for (const entry of fs.readdirSync(source, { withFileTypes: true })) {
    const from = path.join(source, entry.name), to = path.join(destination, entry.name);
    if (entry.isDirectory()) {
      fs.mkdirSync(to, { recursive: true });
      overlay(from, to);
    } else {
      if (fs.existsSync(to)) throw new Error(`Our file now exists upstream: ${to}. Review the overlap before building.`);
      fs.copyFileSync(from, to);
      fs.chmodSync(to, fs.statSync(from).mode);
    }
  }
}

function checkLock(manifest, lock) {
  const root = lock.packages?.[""];
  if (!root || lock.lockfileVersion !== 3) throw new Error("Expected an npm v3 dependency lockfile");
  for (const field of ["name", "version", "dependencies", "devDependencies", "optionalDependencies"]) {
    const actual = root[field] ?? (field.endsWith("Dependencies") || field === "dependencies" ? {} : null);
    const expected = manifest[field] ?? (field.endsWith("Dependencies") || field === "dependencies" ? {} : null);
    const sorted = (value) => typeof value === "object" && value !== null
      ? JSON.stringify(Object.entries(value).sort(([a], [b]) => a.localeCompare(b))) : JSON.stringify(value);
    if (sorted(actual) !== sorted(expected)) throw new Error(`Dependency lock is out of date (${field}). Run npm run update:lock.`);
  }
}

function copyUpstream(source, destination) {
  if (!fs.existsSync(path.join(source, "app/index.js"))) {
    throw new Error("Upstream source is missing. Run git submodule update --init --recursive.");
  }
  if (fs.existsSync(path.join(source, ".git"))) {
    // Copy committed files only. Local changes to upstream must be reviewed,
    // rather than silently included in a build pinned to a different source.
    run("git", ["diff", "--exit-code", "HEAD", "--"], source);
    const files = run("git", ["ls-files", "-z"], source).split("\0").filter(Boolean);
    for (const filename of files) {
      const to = path.join(destination, filename);
      fs.mkdirSync(path.dirname(to), { recursive: true });
      fs.cpSync(path.join(source, filename), to, { preserveTimestamps: true, verbatimSymlinks: true });
    }
  } else {
    // flatpak-builder can strip Git metadata after checking out submodules.
    fs.cpSync(source, destination, { recursive: true, verbatimSymlinks: true, filter: (file) => ![".git", "node_modules"].includes(path.basename(file)) });
  }
}

function prepareApp({ root = projectRoot, validateLock = true } = {}) {
  const parent = path.join(root, "build"), output = path.join(parent, "teams-office");
  fs.mkdirSync(parent, { recursive: true });
  const temporary = fs.mkdtempSync(path.join(parent, ".teams-office-"));
  try {
    copyUpstream(path.join(root, "upstream/teams-for-linux"), temporary);
    const patchDirectory = path.join(root, "patches");
    const patches = fs.readdirSync(patchDirectory).filter((file) => file.endsWith(".patch")).sort();
    for (const patch of patches) {
      // Stop Git at the build directory: applying a patch must never consult
      // or modify the parent repository's index.
      const options = { env: { ...process.env, GIT_CEILING_DIRECTORIES: parent } };
      const filename = path.join(patchDirectory, patch);
      run("git", ["apply", "--no-index", "--check", filename], temporary, options);
      run("git", ["apply", "--no-index", filename], temporary, options);
    }
    overlay(path.join(root, "src"), path.join(temporary, "app"));
    overlay(path.join(root, "tools"), path.join(temporary, "scripts"));
    overlay(path.join(root, "tests"), path.join(temporary, "tests"));
    const manifestPath = path.join(temporary, "package.json");
    const manifest = merge(JSON.parse(fs.readFileSync(manifestPath)), JSON.parse(fs.readFileSync(path.join(root, "package.overrides.json"))));
    const lockPath = path.join(root, "dependencies/package-lock.json");
    if (validateLock) {
      const recorded = fs.readFileSync(path.join(root, "dependencies/upstream-lock.sha256"), "utf8").trim();
      const actual = createHash("sha256").update(fs.readFileSync(path.join(temporary, "package-lock.json"))).digest("hex");
      if (recorded !== actual) throw new Error("Upstream's dependency lock changed. Run npm run update:lock.");
      checkLock(manifest, JSON.parse(fs.readFileSync(lockPath)));
    }
    fs.writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);
    fs.copyFileSync(lockPath, path.join(temporary, "package-lock.json"));
    // Keep installed dependencies and local build/test output across prepares.
    // Everything else is rebuilt, so removed source files cannot linger.
    for (const entry of ["node_modules", "dist", "test-results", "playwright-report"]) {
      const from = path.join(output, entry);
      if (fs.existsSync(from)) fs.renameSync(from, path.join(temporary, entry));
    }
    fs.rmSync(output, { recursive: true, force: true });
    fs.renameSync(temporary, output);
    return output;
  } finally {
    fs.rmSync(temporary, { recursive: true, force: true });
  }
}

if (require.main === module) {
  try { console.log(`Prepared ${path.relative(projectRoot, prepareApp())}${os.EOL}`); }
  catch (error) { console.error(error.message); process.exitCode = 1; }
}

module.exports = { prepareApp, merge, checkLock };
