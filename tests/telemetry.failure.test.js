import {describe, it, expect, beforeEach, vi} from "vitest";

const mockGetAccessToken = vi.fn();

vi.mock("../src/auth.js", () => ({
  getAccessToken: mockGetAccessToken,
}));

function makeStorage() {
  const data = new Map();
  return {
    getItem: (key) => (data.has(key) ? data.get(key) : null),
    setItem: (key, value) => data.set(key, value),
  };
}

describe("telemetry failure isolation", () => {
  beforeEach(() => {
    vi.resetModules();
    vi.clearAllMocks();
    vi.stubGlobal("localStorage", makeStorage());
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => ({ok: true})),
    );
    mockGetAccessToken.mockResolvedValue("test-token");
  });

  it("does not reject or throw when token acquisition fails", async () => {
    mockGetAccessToken.mockRejectedValue(new Error("auth unavailable"));
    const telemetry = await import("../src/telemetry.js");

    expect(() => telemetry.recordActionStarted("summarize")).not.toThrow();
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(fetch).not.toHaveBeenCalled();
  });

  it("does not reject or throw when telemetry fetch fails", async () => {
    fetch.mockRejectedValue(new Error("collector unavailable"));
    const telemetry = await import("../src/telemetry.js");

    expect(() => telemetry.recordActionStarted("reply")).not.toThrow();
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(fetch).toHaveBeenCalledTimes(1);
  });

  it("does not wait indefinitely for a telemetry server", async () => {
    vi.useFakeTimers();
    fetch.mockImplementation(
      (_url, options) =>
        new Promise((_resolve, reject) => {
          options.signal.addEventListener("abort", () =>
            reject(new DOMException("aborted", "AbortError")),
          );
        }),
    );
    const telemetry = await import("../src/telemetry.js");

    telemetry.recordActionStarted("tone");
    await vi.advanceTimersByTimeAsync(5000);

    expect(fetch).toHaveBeenCalledTimes(1);
    vi.useRealTimers();
  });

  it("does not throw when browser storage is unavailable", async () => {
    vi.stubGlobal("localStorage", {
      getItem: () => {
        throw new Error("storage unavailable");
      },
      setItem: () => {
        throw new Error("storage unavailable");
      },
    });
    const telemetry = await import("../src/telemetry.js");

    expect(() => telemetry.recordAdoption()).not.toThrow();
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(fetch).not.toHaveBeenCalled();
  });
});
