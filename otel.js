/* global process, require, module */

const endpoint = (process.env.OTEL_EXPORTER_OTLP_ENDPOINT || "").replace(
  /\/+$/,
  "",
);
const SERVICE_NAME = "librechat-outlook-addin";
const {version: SERVICE_VERSION} = require("./package.json");

const METRIC_DEFINITIONS = {
  installsTotal: {
    kind: "counter",
    name: "addin_installs_total",
    options: {
      description: "Authenticated add-in installs observed once per profile",
    },
  },
  activeDaysTotal: {
    kind: "counter",
    name: "addin_active_days_total",
    options: {
      description: "Authenticated browser profiles active in a UTC day",
    },
  },
  actionsTotal: {
    kind: "counter",
    name: "outlook_actions_total",
    options: {description: "Outlook AI actions started, by action"},
  },
  outcomesTotal: {
    kind: "counter",
    name: "outlook_action_outcomes_total",
    options: {description: "Outlook AI action outcomes, by action and outcome"},
  },
  duration: {
    kind: "histogram",
    name: "outlook_action_duration",
    options: {
      description: "Outlook AI action duration in milliseconds",
      unit: "ms",
      advice: {
        explicitBucketBoundaries: [
          100, 250, 500, 1000, 2500, 5000, 10000, 30000, 60000, 120000,
        ],
      },
    },
  },
  phaseDuration: {
    kind: "histogram",
    name: "outlook_action_phase_duration",
    options: {
      description: "Outlook AI action phase duration in milliseconds",
      unit: "ms",
      advice: {
        explicitBucketBoundaries: [
          10, 25, 50, 100, 250, 500, 1000, 2500, 5000, 10000, 30000, 60000,
        ],
      },
    },
  },
};

const ACTIONS = new Set(["summarize", "reply", "tone"]);
const OUTCOMES = new Set(["completed", "failed", "cancelled"]);
const PHASES = ["reading_email", "thinking", "inserting_result"];
const ERROR_KINDS = new Set([
  "auth",
  "network",
  "timeout",
  "http_4xx",
  "http_5xx",
  "office",
  "invalid_response",
  "unknown",
]);
const EVENT_TYPES = new Set([
  "addin_installed",
  "addin_active_day",
  "outlook_action_started",
  "outlook_action_finished",
]);

function normalizeTelemetryEvent(event) {
  if (!event || typeof event !== "object") return null;
  if (!EVENT_TYPES.has(event.type)) return null;

  const action = ACTIONS.has(event.action) ? event.action : "unknown";
  const outcome = OUTCOMES.has(event.outcome) ? event.outcome : "unknown";
  const errorKind = ERROR_KINDS.has(event.errorKind)
    ? event.errorKind
    : "unknown";
  const phaseDurations = {};
  for (const phase of PHASES) {
    const value = event.phaseDurations?.[phase];
    if (Number.isFinite(value) && value >= 0) phaseDurations[phase] = value;
  }

  return {
    type: event.type,
    action,
    outcome,
    errorKind,
    durationMs:
      Number.isFinite(event.durationMs) && event.durationMs >= 0
        ? event.durationMs
        : undefined,
    phaseDurations,
  };
}

function createNoopSink() {
  return {
    recordTelemetryEvent() {},
    shutdown: async () => {},
  };
}

function createSink() {
  if (!endpoint) return createNoopSink();

  try {
    const {
      PeriodicExportingMetricReader,
    } = require("@opentelemetry/sdk-metrics");
    const {
      OTLPMetricExporter,
    } = require("@opentelemetry/exporter-metrics-otlp-http");
    const {MeterProvider} = require("@opentelemetry/sdk-metrics");
    const {resourceFromAttributes} = require("@opentelemetry/resources");
    const {
      ATTR_SERVICE_NAME,
      ATTR_SERVICE_VERSION,
    } = require("@opentelemetry/semantic-conventions");

    const exporter = new OTLPMetricExporter({url: `${endpoint}/v1/metrics`});
    const reader = new PeriodicExportingMetricReader({
      exporter,
      exportIntervalMillis: 60000,
    });
    const meterProvider = new MeterProvider({
      resource: resourceFromAttributes({
        [ATTR_SERVICE_NAME]: SERVICE_NAME,
        [ATTR_SERVICE_VERSION]: SERVICE_VERSION,
      }),
      readers: [reader],
    });
    const meter = meterProvider.getMeter(SERVICE_NAME);
    const create = (definition) =>
      definition.kind === "histogram"
        ? meter.createHistogram(definition.name, definition.options)
        : meter.createCounter(definition.name, definition.options);

    const installsTotal = create(METRIC_DEFINITIONS.installsTotal);
    const activeDaysTotal = create(METRIC_DEFINITIONS.activeDaysTotal);
    const actionsTotal = create(METRIC_DEFINITIONS.actionsTotal);
    const outcomesTotal = create(METRIC_DEFINITIONS.outcomesTotal);
    const duration = create(METRIC_DEFINITIONS.duration);
    const phaseDuration = create(METRIC_DEFINITIONS.phaseDuration);

    return {
      recordTelemetryEvent(event) {
        switch (event.type) {
          case "addin_installed":
            installsTotal.add(1);
            break;
          case "addin_active_day":
            activeDaysTotal.add(1);
            break;
          case "outlook_action_started":
            actionsTotal.add(1, {action: event.action});
            break;
          case "outlook_action_finished":
            outcomesTotal.add(1, {
              action: event.action,
              outcome: event.outcome,
              ...(event.outcome === "failed"
                ? {error_kind: event.errorKind}
                : {}),
            });
            if (event.durationMs !== undefined) {
              duration.record(event.durationMs, {action: event.action});
            }
            for (const phase of PHASES) {
              if (event.phaseDurations[phase] !== undefined) {
                phaseDuration.record(event.phaseDurations[phase], {
                  action: event.action,
                  phase,
                });
              }
            }
            break;
          default:
            break;
        }
      },
      shutdown: () => meterProvider.shutdown(),
    };
  } catch {
    // A telemetry dependency or collector misconfiguration must not prevent
    // the static server from starting.
    return createNoopSink();
  }
}

const sink = createSink();

function recordTelemetryEvent(event) {
  try {
    const normalized = normalizeTelemetryEvent(event);
    if (normalized) sink.recordTelemetryEvent(normalized);
  } catch {
    // Telemetry is best-effort and must never affect request handling.
  }
}

module.exports = {
  METRIC_DEFINITIONS,
  normalizeTelemetryEvent,
  recordTelemetryEvent,
  shutdownTelemetry: sink.shutdown,
};
