const js = require("@eslint/js");
const globals = require("globals");

module.exports = [
  js.configs.recommended,
  {
    ignores: ["dist/", "node_modules/"],
  },
  {
    // Browser-side add-in code, bundled by webpack as ES modules.
    // Office/OfficeRuntime are declared per-file via `/* global */` comments.
    files: ["src/**/*.js"],
    languageOptions: {
      ecmaVersion: "latest",
      sourceType: "module",
      globals: globals.browser,
    },
  },
  {
    // Node-side tooling: webpack config, dev server, build scripts.
    files: ["eslint.config.js", "webpack.config.js", "server.js", "scripts/**/*.js"],
    languageOptions: {
      ecmaVersion: "latest",
      sourceType: "commonjs",
      globals: globals.node,
    },
  },
];
