import globals from "./build/teams-office/node_modules/globals/index.js";
import pluginJs from "./build/teams-office/node_modules/@eslint/js/src/index.js";

export default [
  { files: ["**/*.js"], languageOptions: { sourceType: "commonjs", globals: { ...globals.browser, ...globals.node } } },
  pluginJs.configs.recommended,
  { rules: { "no-var": "error", eqeqeq: ["error", "always", { null: "ignore" }], "no-unused-vars": ["error", { argsIgnorePattern: "^_" }] } },
];
