import {describe, it, expect} from "vitest";
import fs from "node:fs";
import path from "node:path";
import {fileURLToPath} from "node:url";

// The duplicated envConfigJS() in server.js and webpack.config.js must stay in
// sync. We extract the function body from each file and evaluate it in a
// sandbox where `process.env` is controlled, then parse the emitted
// `window.__ENV = {...}` payload.
const REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

function loadEnvConfigFn(relFile) {
  const src = fs.readFileSync(path.join(REPO, relFile), "utf8");
  const m = src.match(/function envConfigJS\(\) \{[\s\S]*?\n\}/);
  if (!m) throw new Error(`envConfigJS not found in ${relFile}`);
  return m[0];
}

function envVarsFrom(relFile, env) {
  const decl = loadEnvConfigFn(relFile);
  const fakeProcess = {env};
  const fn = new Function("process", `${decl}; return envConfigJS();`);
  const output = fn(fakeProcess);
  return JSON.parse(
    output.trim().replace(/^window\.__ENV = /, "").replace(/;$/, ""),
  );
}

const ENTRA = {
  ENTRA_CLIENT_ID: "client-id-x",
  ENTRA_TENANT_ID: "tenant-id-y",
  ENTRA_API_SCOPE: "api://librechat/access_as_user",
};

describe.each([["server.js"], ["webpack.config.js"]])(
  "env plumbing in %s",
  (file) => {
    it("exposes ENTRA_CLIENT_ID / ENTRA_TENANT_ID / ENTRA_API_SCOPE", () => {
      const vars = envVarsFrom(file, ENTRA);
      expect(vars.ENTRA_CLIENT_ID).toBe(ENTRA.ENTRA_CLIENT_ID);
      expect(vars.ENTRA_TENANT_ID).toBe(ENTRA.ENTRA_TENANT_ID);
      expect(vars.ENTRA_API_SCOPE).toBe(ENTRA.ENTRA_API_SCOPE);
    });

    it("drops LIBRECHAT_API_KEY_HELP", () => {
      const vars = envVarsFrom(file, {
        ...ENTRA,
        LIBRECHAT_API_KEY_HELP: "should-be-dropped",
      });
      expect(vars).not.toHaveProperty("LIBRECHAT_API_KEY_HELP");
    });

    it("still exposes the existing vars (API URL, agent id, branding)", () => {
      const vars = envVarsFrom(file, {
        ...ENTRA,
        LIBRECHAT_API_URL: "https://chat.example.com",
        LIBRECHAT_AGENT_ID: "agent_1",
        APP_NAME: "Acme AI",
        APP_LOGO_URL: "https://x/logo.png",
      });
      expect(vars.LIBRECHAT_API_URL).toBe("https://chat.example.com");
      expect(vars.LIBRECHAT_AGENT_ID).toBe("agent_1");
      expect(vars.APP_NAME).toBe("Acme AI");
      expect(vars.APP_LOGO_URL).toBe("https://x/logo.png");
    });
  },
);
