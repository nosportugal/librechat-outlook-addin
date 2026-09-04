import {describe, it, expect, beforeEach, vi} from "vitest";

vi.mock("../src/auth.js", () => ({
  getAccessToken: vi.fn(async () => "test-token"),
}));

function makeStorage() {
  const data = new Map();
  return {
    getItem: (key) => (data.has(key) ? data.get(key) : null),
    setItem: (key, value) => data.set(key, value),
    removeItem: (key) => data.delete(key),
    clear: () => data.clear(),
  };
}

describe("telemetry helpers", () => {
  beforeEach(() => {
    vi.resetModules();
    vi.restoreAllMocks();
    vi.stubGlobal("localStorage", makeStorage());
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => ({ok: true})),
    );
  });

  it("deduplicates install and active-day events", async () => {
    const telemetry = await import("../src/telemetry.js");
    const {recordAdoption} = telemetry;
    const fetchMock = fetch;

    recordAdoption();
    await Promise.resolve();
    await Promise.resolve();

    expect(fetchMock).toHaveBeenCalledTimes(2);

    fetchMock.mockClear();
    recordAdoption();
    await Promise.resolve();
    await Promise.resolve();

    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("normalizes errors into bounded kinds", async () => {
    const telemetry = await import("../src/telemetry.js");
    const {normalizeErrorKind} = telemetry;

    expect(normalizeErrorKind(new Error("InteractionRequiredAuthError"))).toBe(
      "auth",
    );
    expect(normalizeErrorKind(new Error("timeout"))).toBe("timeout");
    expect(normalizeErrorKind(new Error("HTTP 500"))).toBe("http_5xx");
  });

  it("bounded helpers preserve known labels only", async () => {
    const telemetry = await import("../src/telemetry.js");
    const {boundedAction, boundedOutcome, boundedPhase, boundedErrorKind} =
      telemetry;

    expect(boundedAction("summarize")).toBe("summarize");
    expect(boundedAction("nope")).toBe("unknown");
    expect(boundedOutcome("completed")).toBe("completed");
    expect(boundedOutcome("boom")).toBe("unknown");
    expect(boundedPhase("thinking")).toBe("thinking");
    expect(boundedPhase("other")).toBe("unknown");
    expect(boundedErrorKind("auth")).toBe("auth");
    expect(boundedErrorKind("raw text")).toBe("unknown");
  });
});
