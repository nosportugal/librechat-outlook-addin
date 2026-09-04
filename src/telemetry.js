import {getAccessToken} from "./auth.js";

const INSTALL_KEY = "telemetry_install_recorded";
const ACTIVE_DAY_KEY = "telemetry_active_day";

function nowDateKey() {
  return new Date().toISOString().slice(0, 10);
}

function safeSessionId() {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function getSessionId() {
  if (typeof localStorage === "undefined") return null;
  let sessionId = localStorage.getItem("telemetry_session_id");
  if (!sessionId) {
    sessionId = safeSessionId();
    try {
      localStorage.setItem("telemetry_session_id", sessionId);
    } catch {
      return null;
    }
  }
  return sessionId;
}

function normalizeErrorKind(err) {
  const message = err && err.message ? String(err.message) : String(err || "");
  const lower = message.toLowerCase();
  if (lower.includes("interactionrequiredautherror")) return "auth";
  if (lower.includes("acquiretoken") || lower.includes("auth")) return "auth";
  if (lower.includes("timeout")) return "timeout";
  if (lower.includes("network") || lower.includes("fetch")) return "network";
  if (/http[_ ]?4\d\d/.test(lower)) return "http_4xx";
  if (/http[_ ]?5\d\d/.test(lower)) return "http_5xx";
  if (lower.includes("office") || lower.includes("no email")) return "office";
  if (lower.includes("response") || lower.includes("invalid"))
    return "invalid_response";
  return "unknown";
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

function boundedPhase(phase) {
  return bounded(phase, ["reading_email", "thinking", "inserting_result"]);
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

async function postTelemetry(event) {
  // This function is always called fire-and-forget. Keep every operation,
  // including auth and storage access, inside the catch boundary so a rejected
  // promise can never become an unhandled rejection in the add-in host.
  try {
    const token = await getAccessToken({allowInteractive: false});
    const sessionId = getSessionId();
    if (!sessionId) return;

    const body = JSON.stringify({...event, sessionId});
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 5000);
    try {
      const response = await fetch("/__telemetry", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${token}`,
        },
        body,
        signal: controller.signal,
        keepalive: true,
      });
      if (!response.ok) {
        // Silent failure: telemetry must never break the add-in.
      }
    } finally {
      clearTimeout(timeout);
    }
  } catch {
    // Silent failure: auth, storage, serialization, network, timeout, and
    // server errors must never affect the user's action.
  }
}

function recordAdoption() {
  if (typeof localStorage === "undefined") return;
  try {
    if (!localStorage.getItem(INSTALL_KEY)) {
      localStorage.setItem(INSTALL_KEY, "1");
      void postTelemetry({type: "addin_installed"});
    }
    const today = nowDateKey();
    if (localStorage.getItem(ACTIVE_DAY_KEY) !== today) {
      localStorage.setItem(ACTIVE_DAY_KEY, today);
      void postTelemetry({type: "addin_active_day"});
    }
  } catch {
    // No telemetry if storage is unavailable.
  }
}

function recordActionStarted(action) {
  void postTelemetry({type: "outlook_action_started", action});
}

function recordActionFinished(action, outcome, details = {}) {
  void postTelemetry({
    type: "outlook_action_finished",
    action,
    outcome,
    errorKind: details.errorKind || null,
    durationMs: details.durationMs,
    phaseDurations: details.phaseDurations || {},
  });
}

function recordError(action, err, durationMs, phaseDurations = {}) {
  recordActionFinished(action, "failed", {
    errorKind: normalizeErrorKind(err),
    durationMs,
    phaseDurations,
  });
}

export {
  recordAdoption,
  recordActionStarted,
  recordActionFinished,
  recordError,
  normalizeErrorKind,
  boundedAction,
  boundedOutcome,
  boundedPhase,
  boundedErrorKind,
};
