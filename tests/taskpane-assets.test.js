import {readFileSync} from "node:fs";
import {resolve} from "node:path";
import {describe, expect, it} from "vitest";

const repoRoot = resolve(import.meta.dirname, "..");

function readRepoFile(relativePath) {
  return readFileSync(resolve(repoRoot, relativePath), "utf8");
}

describe("taskpane stylesheet loading", () => {
  it("loads the stylesheet as an external resource for CSP-compatible hosts", () => {
    const html = readRepoFile("src/taskpane/taskpane.html");
    const javascript = readRepoFile("src/taskpane/taskpane.js");

    expect(html).toContain('<link rel="stylesheet" href="/taskpane.css" />');
    expect(javascript).not.toContain('import "./taskpane.css";');
  });
});
