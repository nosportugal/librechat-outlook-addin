const http = require("http");
const fs = require("fs");
const path = require("path");

const PORT = process.env.PORT || 3000;
const DIST = path.join(__dirname, "dist");

// Environment badge for the main app icon (DEV/STB). Empty => PROD => no label.
// Build-time generates icon-<size>-dev.png / -stb.png; here we pick the variant.
const ICON_LABEL = (process.env.ICON_LABEL || "").trim().toLowerCase();
const MAIN_ICON_RE = /^\/assets\/icon-(16|32|64|80|128)\.png$/;

const MIME = {
  ".html": "text/html",
  ".js": "application/javascript",
  ".css": "text/css",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".json": "application/json",
  ".map": "application/json",
  ".txt": "text/plain",
};

function envConfigJS() {
  const env = {
    LIBRECHAT_API_URL: process.env.LIBRECHAT_API_URL || "",
    LIBRECHAT_AGENT_ID: process.env.LIBRECHAT_AGENT_ID || "",
    LIBRECHAT_API_KEY_HELP: process.env.LIBRECHAT_API_KEY_HELP || "",
    APP_NAME: process.env.APP_NAME || "",
    APP_LOGO_URL: process.env.APP_LOGO_URL || "",
  };
  return `window.__ENV = ${JSON.stringify(env)};`;
}

const server = http.createServer((req, res) => {
  const url = new URL(req.url, `http://${req.headers.host}`);
  let pathname = url.pathname;

  if (pathname === "/env-config.js") {
    res.writeHead(200, {
      "Content-Type": "application/javascript",
      "Cache-Control": "no-cache",
    });
    res.end(envConfigJS());
    return;
  }

  if (pathname === "/") pathname = "/taskpane.html";

  // Swap the main app icon for its environment-labeled variant when present.
  if (ICON_LABEL) {
    const match = pathname.match(MAIN_ICON_RE);
    if (match) {
      const labeled = `/assets/icon-${match[1]}-${ICON_LABEL}.png`;
      if (fs.existsSync(path.join(DIST, labeled))) {
        pathname = labeled;
      }
    }
  }

  const filePath = path.join(DIST, pathname);
  const resolved = path.resolve(filePath);
  if (!resolved.startsWith(DIST)) {
    res.writeHead(403);
    res.end();
    return;
  }

  fs.readFile(resolved, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end("Not Found");
      return;
    }
    const ext = path.extname(resolved);
    res.writeHead(200, {
      "Content-Type": MIME[ext] || "application/octet-stream",
      "Access-Control-Allow-Origin": "*",
    });
    res.end(data);
  });
});

server.listen(PORT, "0.0.0.0", () => {
  console.log(`Listening on http://0.0.0.0:${PORT}`);
});
