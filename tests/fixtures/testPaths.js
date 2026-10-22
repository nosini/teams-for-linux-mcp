const path = require("node:path");

// Tests execute in build/teams-office. Keep socket fixtures at the project
// root so the extra assembly directory does not exceed Unix sun_path limits.
function testPath(...parts) {
  const root = process.env.TEAMS_MCP_TEST_TMP || path.resolve(__dirname, "../../../../build/t");
  return path.join(root, ...parts);
}

module.exports = { testPath };
