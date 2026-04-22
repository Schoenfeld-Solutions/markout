import {
  DEFAULT_MAX_JAVASCRIPT_ASSET_BYTES,
  findOversizedJavaScriptAssets,
} from "../scripts/check-bundle-budgets";

describe("bundle budget checks", () => {
  it("returns only assets larger than the configured JavaScript budget", () => {
    const oversizedAssets = findOversizedJavaScriptAssets([
      {
        filePath: "/repo/dist/taskpane.js",
        sizeBytes: DEFAULT_MAX_JAVASCRIPT_ASSET_BYTES + 1,
      },
      {
        filePath: "/repo/dist/taskpane-runtime.chunk.js",
        sizeBytes: DEFAULT_MAX_JAVASCRIPT_ASSET_BYTES,
      },
      {
        filePath: "/repo/dist/commands.js",
        sizeBytes: 113,
      },
    ]);

    expect(oversizedAssets).toEqual([
      {
        filePath: "/repo/dist/taskpane.js",
        sizeBytes: DEFAULT_MAX_JAVASCRIPT_ASSET_BYTES + 1,
      },
    ]);
  });
});
