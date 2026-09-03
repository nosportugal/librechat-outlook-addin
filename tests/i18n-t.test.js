import {describe, it, expect, beforeEach} from "vitest";

// Import the real module (no mocks) and drive it via initI18n + URL/Office.
import {initI18n, t} from "../src/i18n.js";

describe("t() substitution", () => {
  beforeEach(() => {
    // Force a deterministic locale.
    window.history.replaceState({}, "", "?lang=en-GB");
    initI18n("AI Assistant");
  });

  it("substitutes {app} with the configured app name", () => {
    expect(t("settings.title")).toBe("Account");
  });

  it("substitutes positional {0} with the first extra argument", () => {
    expect(t("auth.signedInAs", "user@nos.pt")).toBe(
      "Signed in as user@nos.pt",
    );
  });

  it("leaves {0} untouched when no argument is provided", () => {
    expect(t("auth.signedInAs")).toBe("Signed in as {0}");
  });

  it("returns the key itself for unknown keys", () => {
    expect(t("does.not.exist")).toBe("does.not.exist");
  });
});
