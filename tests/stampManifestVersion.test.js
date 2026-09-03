import {describe, expect, it} from "vitest";

import {stampManifestVersion} from "../scripts/stampManifestVersion.js";

const SAMPLE_MANIFEST = `
  <Version>0.6.0.0</Version>
  <SourceLocation DefaultValue="https://localhost:3000/taskpane.html" />
  <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html" />
  <bt:Url id="SettingsTaskpane.Url" DefaultValue="https://localhost:3000/taskpane.html?action=settings" />
  <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html" />
`;

describe("stampManifestVersion", () => {
  it("stamps both taskpane URLs with the manifest version", () => {
    const stamped = stampManifestVersion(SAMPLE_MANIFEST);

    expect(stamped).toContain('taskpane.html?manifestVersion=0.6.0.0"');
    expect(stamped.match(/manifestVersion=0\.6\.0\.0/g)).toHaveLength(3);
    expect(stamped).toContain(
      "taskpane.html?action=settings&manifestVersion=0.6.0.0",
    );
  });

  it("leaves non-taskpane URLs untouched", () => {
    expect(stampManifestVersion(SAMPLE_MANIFEST)).toContain(
      'id="Commands.Url" DefaultValue="https://localhost:3000/commands.html"',
    );
  });
});
