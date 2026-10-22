const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const { createHash } = require("node:crypto");
const { prepareApp } = require("../scripts/prepare-app");

function fixture(t) {
  const parent = path.resolve(__dirname, "../build/test-fixtures");
  fs.mkdirSync(parent, { recursive: true });
  const root = fs.mkdtempSync(path.join(parent, "prepare-"));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const write = (filename, value) => {
    const target = path.join(root, filename);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, typeof value === "string" ? value : `${JSON.stringify(value)}\n`);
  };
  const manifest = { name: "teams-for-linux", version: "1.0.0", scripts: { start: "electron app", prestart: "npm ci" }, dependencies: { upstream: "1.0.0" } };
  write("upstream/teams-for-linux/package.json", manifest);
  write("upstream/teams-for-linux/package-lock.json", { lockfileVersion: 3, packages: { "": manifest } });
  write("dependencies/upstream-lock.sha256", createHash("sha256").update(fs.readFileSync(path.join(root, "upstream/teams-for-linux/package-lock.json"))).digest("hex") + "\n");
  write("upstream/teams-for-linux/app/index.js", "// upstream\n");
  write("upstream/teams-for-linux/tests/unit/upstream.test.js", "// upstream test\n");
  write("upstream/teams-for-linux/README.md", "Upstream README\n");
  write("README.md", "Our README\n");
  write("package.overrides.json", { scripts: { prestart: null }, dependencies: { bridge: "2.0.0" } });
  write("dependencies/package-lock.json", { name: manifest.name, version: manifest.version, lockfileVersion: 3,
    packages: { "": { name: manifest.name, version: manifest.version, dependencies: { upstream: "1.0.0", bridge: "2.0.0" } } } });
  write("src/mcp/index.js", "// bridge\n");
  write("tools/call.js", "// relay tool\n");
  write("tests/unit/bridge.test.js", "// bridge test\n");
  write("patches/0001-integration.patch", "diff --git a/app/index.js b/app/index.js\n--- a/app/index.js\n+++ b/app/index.js\n@@ -1 +1,2 @@\n // upstream\n+// integration\n");
  return { root, write, read: (filename) => fs.readFileSync(path.join(root, filename), "utf8") };
}

test("assembles exported upstream sources, patches, dependencies and both test suites without changing source", (t) => {
  const { root, read } = fixture(t);
  prepareApp({ root });
  const base = "build/teams-office/";
  assert.equal(read(base + "app/index.js"), "// upstream\n// integration\n");
  assert.equal(read(base + "app/mcp/index.js"), "// bridge\n");
  assert.equal(read(base + "scripts/call.js"), "// relay tool\n");
  assert.equal(read(base + "tests/unit/upstream.test.js"), "// upstream test\n");
  assert.equal(read(base + "tests/unit/bridge.test.js"), "// bridge test\n");
  const manifest = JSON.parse(read(base + "package.json"));
  assert.equal(manifest.scripts.prestart, undefined);
  assert.deepEqual(manifest.dependencies, { upstream: "1.0.0", bridge: "2.0.0" });
  assert.equal(read("upstream/teams-for-linux/app/index.js"), "// upstream\n");
  assert.equal(read("upstream/teams-for-linux/README.md"), "Upstream README\n");
  assert.equal(read("README.md"), "Our README\n");
});

test("repreparing removes deleted overlay files while preserving installed dependencies and artifacts", (t) => {
  const { root, write, read } = fixture(t);
  prepareApp({ root });
  write("build/teams-office/node_modules/dependency.txt", "installed\n");
  write("build/teams-office/dist/package.txt", "artifact\n");
  fs.unlinkSync(path.join(root, "src/mcp/index.js"));
  prepareApp({ root });
  assert.equal(fs.existsSync(path.join(root, "build/teams-office/app/mcp/index.js")), false);
  assert.equal(read("build/teams-office/node_modules/dependency.txt"), "installed\n");
  assert.equal(read("build/teams-office/dist/package.txt"), "artifact\n");
});

test("an incompatible upstream update fails without replacing the last working assembly", (t) => {
  const { root, write, read } = fixture(t);
  prepareApp({ root });
  write("upstream/teams-for-linux/app/index.js", "// incompatible upstream\n");
  assert.throws(() => prepareApp({ root }), /patch does not apply/);
  assert.equal(read("build/teams-office/app/index.js"), "// upstream\n// integration\n");
  assert.equal(fs.readdirSync(path.join(root, "build")).some((file) => file.startsWith(".teams-office-")), false);
});

test("an upstream file conflicting with our overlay needs explicit review", (t) => {
  const { root, write } = fixture(t);
  write("upstream/teams-for-linux/app/mcp/index.js", "// newly added upstream feature\n");
  assert.throws(() => prepareApp({ root }), /Our file now exists upstream/);
});

test("changed upstream dependencies cannot silently use the previous lockfile", (t) => {
  const { root, write } = fixture(t);
  write("upstream/teams-for-linux/package.json", { name: "teams-for-linux", version: "1.0.0", dependencies: { upstream: "1.1.0" } });
  assert.throws(() => prepareApp({ root }), /Dependency lock is out of date \(dependencies\)/);
  assert.doesNotThrow(() => prepareApp({ root, validateLock: false }));
});

test("a clone without initialized submodules gives an actionable error", (t) => {
  const { root } = fixture(t);
  fs.rmSync(path.join(root, "upstream/teams-for-linux"), { recursive: true });
  assert.throws(() => prepareApp({ root }), /git submodule update --init --recursive/);
});

test("upstream transitive dependency changes require a lock refresh even with unchanged manifest specs", (t) => {
  const { root, write } = fixture(t);
  write("upstream/teams-for-linux/package-lock.json", { lockfileVersion: 3, packages: { "node_modules/upstream": { version: "1.0.1" } } });
  assert.throws(() => prepareApp({ root }), /Upstream's dependency lock changed/);
});
