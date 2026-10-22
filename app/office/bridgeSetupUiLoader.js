const fs = require("node:fs");

// Only the fixed, locally installed UI module can reload. No MCP request can
// supply a path or source. Packaged builds keep the normal immutable module.
function createBridgeSetupUiLoader({ development = process.defaultApp === true, filename = require.resolve("./bridgeSetupUi") } = {}) {
  const stamp = () => {
    const stat = fs.statSync(filename, { bigint: true });
    return `${stat.dev}:${stat.ino}:${stat.mtimeNs}:${stat.size}`;
  };
  let current = require(filename).bridgeSetupUi;
  let revision = development ? stamp() : null;
  return (request) => {
    if (development) {
      const nextRevision = stamp();
      if (nextRevision !== revision) {
        const previous = require.cache[filename];
        delete require.cache[filename];
        try {
          const next = require(filename).bridgeSetupUi;
          if (typeof next !== "function") throw new Error("Invalid UI module");
          current = next;
          revision = nextRevision;
          module.children = module.children.filter((child) => child.id !== filename || child === require.cache[filename]);
        } catch {
          if (previous) require.cache[filename] = previous;
          module.children = module.children.filter((child) => child.id !== filename || child === previous);
          throw new Error("Office setup UI could not reload");
        }
      }
    }
    return `(${current.toString()})(${JSON.stringify(request)})`;
  };
}

const bridgeSetupUiScript = createBridgeSetupUiLoader();
module.exports = { bridgeSetupUiScript, createBridgeSetupUiLoader };
