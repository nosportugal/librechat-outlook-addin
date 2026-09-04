import {describe, expect, it, vi} from "vitest";

vi.stubEnv("OTEL_EXPORTER_OTLP_ENDPOINT", "");

const {METRIC_DEFINITIONS, normalizeTelemetryEvent, recordTelemetryEvent} =
  await import("../otel.js");

describe("Outlook OpenTelemetry sink", () => {
  it("defines the adoption, action, outcome, and timing metrics", () => {
    expect(METRIC_DEFINITIONS.installsTotal.name).toBe("addin_installs_total");
    expect(METRIC_DEFINITIONS.activeDaysTotal.name).toBe(
      "addin_active_days_total",
    );
    expect(METRIC_DEFINITIONS.actionsTotal.name).toBe("outlook_actions_total");
    expect(METRIC_DEFINITIONS.outcomesTotal.name).toBe(
      "outlook_action_outcomes_total",
    );
    expect(METRIC_DEFINITIONS.duration.name).toBe("outlook_action_duration");
    expect(METRIC_DEFINITIONS.phaseDuration.name).toBe(
      "outlook_action_phase_duration",
    );
  });

  it("normalizes untrusted event labels to bounded values", () => {
    const event = normalizeTelemetryEvent({
      type: "outlook_action_finished",
      action: "invented-action",
      outcome: "raw-outcome",
      errorKind: "email contents",
      durationMs: 42,
      phaseDurations: {invented_phase: 99, thinking: 12},
    });

    expect(event.action).toBe("unknown");
    expect(event.outcome).toBe("unknown");
    expect(event.errorKind).toBe("unknown");
    expect(event.phaseDurations).toEqual({thinking: 12});
  });

  it("is a no-op without a collector endpoint", () => {
    expect(() => recordTelemetryEvent({type: "addin_installed"})).not.toThrow();
  });
});
