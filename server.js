const http = require("http");
const fs = require("fs");
const path = require("path");
const {recordTelemetryEvent, shutdownTelemetry} = require("./otel");

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
    ENTRA_CLIENT_ID: process.env.ENTRA_CLIENT_ID || "",
    ENTRA_TENANT_ID: process.env.ENTRA_TENANT_ID || "",
    ENTRA_API_SCOPE: process.env.ENTRA_API_SCOPE || "",
    APP_NAME: process.env.APP_NAME || "",
    APP_LOGO_URL: process.env.APP_LOGO_URL || "",
  };
  return `window.__ENV = ${JSON.stringify(env)};`;
}

function bounded(value, allowed, fallback = "unknown") {
  return allowed.includes(value) ? value : fallback;
}

function boundedAction(action) {
  return bounded(action, ["summarize", "reply", "tone"]);
}

function boundedOutcome(outcome) {
  return bounded(outcome, ["completed", "failed", "cancelled"]);
}

function boundedErrorKind(kind) {
  return bounded(
    kind,
    [
      "auth",
      "network",
      "timeout",
      "http_4xx",
      "http_5xx",
      "office",
      "invalid_response",
      "unknown",
    ],
    "unknown",
  );
}

function normalizeTelemetryEvent(event) {
  if (!event || typeof event !== "object") return null;
  const action = boundedAction(event.action);
  const outcome = boundedOutcome(event.outcome);
  const errorKind = boundedErrorKind(event.errorKind);
  const phaseDurations = event.phaseDurations || {};

  return {
    type: bounded(event.type, [
      "addin_installed",
      "addin_active_day",
      "outlook_action_started",
      "outlook_action_finished",
    ]),
    action,
    outcome,
    errorKind,
    durationMs:
      typeof event.durationMs === "number" ? event.durationMs : undefined,
    phaseDurations: {
      reading_email:
        typeof phaseDurations.reading_email === "number"
          ? phaseDurations.reading_email
          : undefined,
      thinking:
        typeof phaseDurations.thinking === "number"
          ? phaseDurations.thinking
          : undefined,
      inserting_result:
        typeof phaseDurations.inserting_result === "number"
          ? phaseDurations.inserting_result
          : undefined,
    },
  };
}

async function handleTelemetry(req, res) {
  if (req.method !== "POST") {
    res.writeHead(405, {"Content-Type": "text/plain"});
    res.end("Method Not Allowed");
    return;
  }

  const auth = req.headers.authorization || "";
  if (!auth.startsWith("Bearer ")) {
    res.writeHead(401, {"Content-Type": "text/plain"});
    res.end("Unauthorized");
    return;
  }

  const chunks = [];
  for await (const chunk of req) {
    chunks.push(chunk);
  }
  const raw = Buffer.concat(chunks).toString("utf8");
  if (raw.length > 32 * 1024) {
    res.writeHead(413, {"Content-Type": "text/plain"});
    res.end("Payload Too Large");
    return;
  }

  let event;
  try {
    event = JSON.parse(raw);
  } catch {
    res.writeHead(400, {"Content-Type": "text/plain"});
    res.end("Bad Request");
    return;
  }

  const normalized = normalizeTelemetryEvent(event);
  if (!normalized || !normalized.type) {
    res.writeHead(422, {"Content-Type": "text/plain"});
    res.end("Unprocessable Entity");
    return;
  }

  recordTelemetryEvent(event);

  res.writeHead(202, {"Content-Type": "application/json"});
  res.end(JSON.stringify({ok: true}));
}

const server = http.createServer(async (req, res) => {
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

  if (pathname === "/__telemetry") {
    await handleTelemetry(req, res);
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

function shutdown() {
  void shutdownTelemetry().finally(() => server.close());
}

process.once("SIGTERM", shutdown);
process.once("SIGINT", shutdown);
