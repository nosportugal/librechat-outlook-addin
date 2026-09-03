import {describe, expect, it} from "vitest";

import {parseManifestVersion} from "../src/taskpane/manifestVersion.js";

describe("parseManifestVersion", () => {
  it("reads the manifest version from the taskpane URL", () => {
    expect(parseManifestVersion("?manifestVersion=0.6.0.0")).toBe("0.6.0.0");
  });

  it("returns unknown when the URL is not stamped", () => {
    expect(parseManifestVersion("?action=settings")).toBe("unknown");
    expect(parseManifestVersion("?manifestVersion=")).toBe("unknown");
  });
});