import {
  MarkOutError,
  getAllRuntimeChannelConfigs,
  getChannelScopedKey,
  getRuntimeChannelConfig,
  isMarkOutErrorCode,
  resolveRuntimeChannelConfig,
} from "../src/lib/runtime";

describe("runtime channel config", () => {
  it("defines three unique runtime channels", () => {
    const configs = getAllRuntimeChannelConfigs();
    const addInIds = new Set(configs.map((config) => config.addInId));
    const storageNamespaces = new Set(
      configs.map((config) => config.storageNamespace)
    );

    expect(configs).toHaveLength(3);
    expect(addInIds.size).toBe(3);
    expect(storageNamespaces.size).toBe(3);
  });

  it("resolves the explicit channel query before path heuristics", () => {
    expect(
      resolveRuntimeChannelConfig(
        "https://localhost:3000/taskpane.html?channel=production"
      ).channelId
    ).toBe("production");

    expect(
      resolveRuntimeChannelConfig(
        "https://example.invalid/outlook/taskpane.html?channel=beta"
      ).channelId
    ).toBe("beta");
  });

  it("falls back to host and pathname heuristics when the channel query is absent", () => {
    expect(
      resolveRuntimeChannelConfig("https://localhost:3000/taskpane.html")
        .channelId
    ).toBe("local");
    expect(
      resolveRuntimeChannelConfig(
        "https://schoenfeld-solutions.github.io/markout/outlook-beta/taskpane.html"
      ).channelId
    ).toBe("beta");
    expect(
      resolveRuntimeChannelConfig(
        "https://schoenfeld-solutions.github.io/markout/outlook/taskpane.html"
      ).channelId
    ).toBe("production");
  });

  it("generates channel-scoped storage keys", () => {
    expect(
      getChannelScopedKey(getRuntimeChannelConfig("beta"), "notification")
    ).toBe("markout.beta.notification");
  });

  it("recognizes typed MarkOut errors", () => {
    const error = new MarkOutError(
      "restore-state-too-large",
      "Restore-state storage is full."
    );

    expect(isMarkOutErrorCode(error, "restore-state-too-large")).toBe(true);
    expect(isMarkOutErrorCode(error, "unsupported-body-type")).toBe(false);
    expect(
      isMarkOutErrorCode(new Error("boom"), "restore-state-too-large")
    ).toBe(false);
  });
});
