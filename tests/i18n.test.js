import {describe, it, expect} from "vitest";
import fs from "node:fs";
import path from "node:path";
import {fileURLToPath} from "node:url";

const REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

function loadTranslations() {
  const src = fs.readFileSync(path.join(REPO, "src/i18n.js"), "utf8");
  const m = src.match(/const translations = (\{[\s\S]*?\n\});/);
  if (!m) throw new Error("translations object not found in i18n.js");
  // Evaluate the object literal in isolation. It is pure data (no functions).
  return new Function(`return ${m[1]};`)();
}

const translations = loadTranslations();
const en = translations["en-GB"];
const pt = translations["pt-PT"];

// Keys the Entra epic removes (API-key UI is gone).
const REMOVED_KEYS = [
  "settings.apiKeyLabel",
  "settings.apiKeyDescription",
  "settings.apiKeyLinkLabel",
  "settings.apiKeyPlaceholder",
  "settings.testConnection",
  "settings.testing",
  "settings.connectionSuccess",
  "settings.connectionFailed",
  "error.noApiKey",
  "notify.noApiKey",
];

// Keys the Entra epic adds.
const ADDED_KEYS = [
  "auth.signedInAs",
  "auth.signingIn",
  "auth.signInRequired",
  "auth.signInButton",
  "auth.naaUnsupported",
  "auth.failed",
  "auth.identityMismatch",
  "notify.signInRequired",
  "notify.authFailed",
];

describe("i18n locale parity", () => {
  it("en-GB and pt-PT expose the exact same set of keys", () => {
    expect(Object.keys(en).sort()).toEqual(Object.keys(pt).sort());
  });

  it("no value is empty or missing in either locale", () => {
    for (const [key, value] of Object.entries(en)) {
      expect(value, `en-GB ${key}`).toBeTruthy();
    }
    for (const [key, value] of Object.entries(pt)) {
      expect(value, `pt-PT ${key}`).toBeTruthy();
    }
  });
});

describe("i18n — Entra SSO epic key changes", () => {
  it.each(ADDED_KEYS)("adds %s in both locales", (key) => {
    expect(en).toHaveProperty(key);
    expect(pt).toHaveProperty(key);
  });

  it.each(REMOVED_KEYS)("removes %s from both locales", (key) => {
    expect(en).not.toHaveProperty(key);
    expect(pt).not.toHaveProperty(key);
  });
});
