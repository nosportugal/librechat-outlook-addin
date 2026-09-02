const path = require("path");
const fs = require("fs");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const {LABELS, makeLabeledIconBuffer} = require("./scripts/label-icons.js");

// Load .env so the dev server can inject the same env-config.js that server.js
// serves in production. Without this, window.__ENV is empty and API calls
// resolve against the dev-server origin (causing 404s).
require("dotenv").config();

const isProduction = process.env.NODE_ENV === "production";

// Environment badge for the main app icon during `npm start` / `npm run dev`.
// Mirrors server.js: empty => PROD => no label. Source icons live in assets/.
const ICON_LABEL = (process.env.ICON_LABEL || "").trim().toLowerCase();
const MAIN_ICON_RE = /^\/assets\/icon-(16|32|64|80|128)\.png$/;

function envConfigJS() {
  const env = {
    LIBRECHAT_API_URL: process.env.LIBRECHAT_API_URL || "",
    LIBRECHAT_AGENT_ID: process.env.LIBRECHAT_AGENT_ID || "",
    ENTRA_CLIENT_ID: process.env.ENTRA_CLIENT_ID || "",
    ENTRA_TENANT_ID: process.env.ENTRA_TENANT_ID || "",
    ENTRA_API_SCOPE: process.env.ENTRA_API_SCOPE || "",
    APP_NAME: process.env.APP_NAME || "",
    APP_LOGO_URL: process.env.APP_LOGO_URL || "",
  };
  return `window.__ENV = ${JSON.stringify(env)};`;
}

function getDevServerHttpsOptions() {
  const certDir = path.join(
    process.env.HOME || process.env.USERPROFILE,
    ".office-addin-dev-certs",
  );
  const certFile = path.join(certDir, "localhost.crt");
  const keyFile = path.join(certDir, "localhost.key");
  const caFile = path.join(certDir, "ca.crt");

  if (fs.existsSync(certFile) && fs.existsSync(keyFile)) {
    const opts = {cert: certFile, key: keyFile};
    if (fs.existsSync(caFile)) opts.ca = caFile;
    return opts;
  }
  // Fall back to webpack-dev-server self-signed cert
  return true;
}

module.exports = {
  entry: {
    taskpane: "./src/taskpane/taskpane.js",
    commands: "./src/commands/commands.js",
    "prompt-dialog": "./src/dialogs/prompt-dialog.js",
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].js",
    clean: true,
  },
  optimization: {
    // @azure/msal-browser (~200 KB) is needed by both the taskpane and the
    // commands entry points; split it (and any other shared code) into one
    // chunk so it is not bundled twice. HtmlWebpackPlugin picks the chunk up
    // automatically for its listed entries.
    splitChunks: {chunks: "all"},
  },
  devServer: {
    static: {
      directory: path.resolve(__dirname, "dist"),
    },
    host: "0.0.0.0",
    port: 3000,
    ...(isProduction
      ? {}
      : {
          server: {
            type: "https",
            options: getDevServerHttpsOptions(),
          },
        }),
    allowedHosts: "all",
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
    setupMiddlewares: (middlewares, devServer) => {
      devServer.app.get("/env-config.js", (_req, res) => {
        res.setHeader("Content-Type", "application/javascript");
        res.setHeader("Cache-Control", "no-cache");
        res.end(envConfigJS());
      });

      // Serve the environment-labeled main app icon when ICON_LABEL is set.
      // Generated on the fly from the source icon in assets/ (dev server does
      // not run the build-time label-icons step).
      if (ICON_LABEL && LABELS[ICON_LABEL]) {
        devServer.app.get(MAIN_ICON_RE, async (req, res) => {
          const srcPath = path.resolve(__dirname, "." + req.path);
          try {
            const buffer = await makeLabeledIconBuffer(srcPath, ICON_LABEL);
            res.setHeader("Content-Type", "image/png");
            res.setHeader("Cache-Control", "no-cache");
            res.end(buffer);
          } catch (err) {
            console.warn(
              `[label-icons] dev-server failed for ${req.path}:`,
              err.message,
            );
            res.status(404).end();
          }
        });
      }
      return middlewares;
    },
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: "./src/taskpane/taskpane.html",
      chunks: ["taskpane"],
    }),
    new HtmlWebpackPlugin({
      filename: "commands.html",
      template: "./src/commands/commands.html",
      chunks: ["commands"],
    }),
    new HtmlWebpackPlugin({
      filename: "prompt-dialog.html",
      template: "./src/dialogs/prompt-dialog.html",
      chunks: ["prompt-dialog"],
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: "assets",
          to: "assets",
          noErrorOnMissing: true,
        },
      ],
    }),
  ],
  module: {
    rules: [
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"],
      },
    ],
  },
  resolve: {
    extensions: [".js"],
  },
  devtool: isProduction ? false : "source-map",
};
