import {
  MarkOutError,
  createInMemoryDiagnosticSink,
  getErrorDiagnosticMetadata,
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

  it("keeps hosted runtime URLs queryless and local URLs channel-explicit", () => {
    const beta = getRuntimeChannelConfig("beta");
    const production = getRuntimeChannelConfig("production");
    const local = getRuntimeChannelConfig("local");

    expect(beta.commandsUrl).toBe(
      "https://schoenfeld-solutions.github.io/markout/outlook-beta/commands.html"
    );
    expect(beta.launcheventUrl).toBe(
      "https://schoenfeld-solutions.github.io/markout/outlook-beta/launchevent.js"
    );
    expect(beta.taskpaneUrl).toBe(
      "https://schoenfeld-solutions.github.io/markout/outlook-beta/taskpane.html"
    );
    expect(production.commandsUrl).toBe(
      "https://schoenfeld-solutions.github.io/markout/outlook/commands.html"
    );
    expect(production.launcheventUrl).toBe(
      "https://schoenfeld-solutions.github.io/markout/outlook/launchevent.js"
    );
    expect(production.taskpaneUrl).toBe(
      "https://schoenfeld-solutions.github.io/markout/outlook/taskpane.html"
    );
    expect(local.commandsUrl).toBe(
      "https://localhost:3000/commands.html?channel=local"
    );
    expect(local.launcheventUrl).toBe(
      "https://localhost:3000/launchevent.js?channel=local"
    );
    expect(local.taskpaneUrl).toBe(
      "https://localhost:3000/taskpane.html?channel=local"
    );
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
    expect(resolveRuntimeChannelConfig("not a url").channelId).toBe(
      "production"
    );
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

  it("keeps diagnostics bounded and immutable", () => {
    const sink = createInMemoryDiagnosticSink(
      2,
      () => new Date("2026-04-25T10:00:00.000Z")
    );

    sink.record({
      area: "preview",
      code: "preview.render.started",
      level: "debug",
    });
    sink.record({
      area: "render",
      code: "draft.render.started",
      level: "debug",
    });
    sink.record({
      area: "restore",
      code: "draft.restore.succeeded",
      level: "info",
    });

    const snapshot = sink.snapshot();
    expect(snapshot.map((event) => event.code)).toEqual([
      "draft.render.started",
      "draft.restore.succeeded",
    ]);
    expect(snapshot[0]?.id).toBe(2);
    expect(snapshot[1]?.timestamp).toBe("2026-04-25T10:00:00.000Z");

    snapshot[0]!.metadata.mutated = true;
    expect(sink.snapshot()[0]?.metadata).toEqual({});
  });

  it("redacts sensitive diagnostic metadata and omits error messages", () => {
    const sink = createInMemoryDiagnosticSink(
      4,
      () => new Date("2026-04-25T10:00:00.000Z")
    );
    const error = new MarkOutError(
      "unsupported-body-type",
      "Message body contains private draft content."
    );

    sink.record({
      area: "body-io",
      code: "body.write.failed",
      level: "error",
      message: "  write failed\nwith whitespace  ",
      metadata: {
        bodyHtml: "<p>secret</p>",
        errorMessage: error.message,
        safeCount: 4,
        tokenValue: "secret-token",
        ...getErrorDiagnosticMetadata(error),
      },
    });

    expect(sink.snapshot()).toEqual([
      {
        area: "body-io",
        code: "body.write.failed",
        id: 1,
        level: "error",
        message: "write failed with whitespace",
        metadata: {
          bodyHtml: "[redacted]",
          errorCode: "unsupported-body-type",
          errorMessage: "[redacted]",
          errorName: "MarkOutError",
          safeCount: 4,
          tokenValue: "[redacted]",
        },
        timestamp: "2026-04-25T10:00:00.000Z",
      },
    ]);
  });

  it("normalizes diagnostic capacity, truncates long strings, and clears events", () => {
    const sink = createInMemoryDiagnosticSink(
      0,
      () => new Date("2026-04-25T10:00:00.000Z")
    );

    sink.record({
      area: "notification",
      code: "notification.transient.shown",
      level: "info",
      message: "x".repeat(200),
      metadata: {
        detail: "y".repeat(160),
        omitted: undefined,
      },
    });
    expect(sink.snapshot()[0]?.message).toMatch(/\.\.\.$/);
    expect(String(sink.snapshot()[0]?.metadata.detail)).toMatch(/\.\.\.$/);

    sink.record({
      area: "render",
      code: "draft.render.succeeded",
      level: "info",
    });

    expect(sink.snapshot()).toEqual([
      {
        area: "render",
        code: "draft.render.succeeded",
        id: 2,
        level: "info",
        metadata: {},
        timestamp: "2026-04-25T10:00:00.000Z",
      },
    ]);

    sink.clear();
    expect(sink.snapshot()).toEqual([]);
  });

  it("reports non-MarkOut error diagnostics without leaking messages", () => {
    expect(getErrorDiagnosticMetadata(new TypeError("private value"))).toEqual({
      errorName: "TypeError",
    });
    expect(getErrorDiagnosticMetadata("private value")).toEqual({
      errorName: "UnknownError",
    });
  });
});
