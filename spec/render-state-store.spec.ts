import { createOfficeRenderStateStore } from "../src/lib/render-state-store";
import { getRuntimeChannelConfig } from "../src/lib/runtime";
import { FakeMailboxItem, installOfficeEnvironment } from "./helpers";

describe("render state store", () => {
  const runtimeChannelConfig = getRuntimeChannelConfig("production");

  beforeEach(() => {
    installOfficeEnvironment();
  });

  it("persists rendered and pending render states", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    const renderStateStore = createOfficeRenderStateStore(
      mailboxItem,
      runtimeChannelConfig
    );

    await renderStateStore.setPendingRenderState("<div>Original</div>");
    expect(await renderStateStore.getRenderState()).toEqual(
      expect.objectContaining({
        channelId: "production",
        originalHtml: "<div>Original</div>",
        phase: "pending",
        storedAt: expect.any(String),
      })
    );

    await renderStateStore.setRenderedRenderState("<div>Original</div>");
    expect(await renderStateStore.getRenderState()).toEqual(
      expect.objectContaining({
        channelId: "production",
        originalHtml: "<div>Original</div>",
        phase: "rendered",
        storedAt: expect.any(String),
      })
    );

    await renderStateStore.clearRenderState();
    expect(await renderStateStore.getRenderState()).toBeNull();
  });

  it("supports legacy raw-html and false sentinel values", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.customProperties.set(
      "markout.originalHtml",
      "<div>Legacy</div>"
    );

    const renderStateStore = createOfficeRenderStateStore(
      mailboxItem,
      runtimeChannelConfig
    );
    expect(await renderStateStore.getRenderState()).toEqual(
      expect.objectContaining({
        channelId: "production",
        originalHtml: "<div>Legacy</div>",
        phase: "rendered",
      })
    );

    mailboxItem.customProperties.set("markout.originalHtml", "false");
    expect(await renderStateStore.getRenderState()).toBeNull();
  });

  it("surfaces custom property load failures", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.failNextLoadCustomProperties = true;
    const renderStateStore = createOfficeRenderStateStore(
      mailboxItem,
      runtimeChannelConfig
    );

    await expect(renderStateStore.getRenderState()).rejects.toMatchObject({
      message: "Loading custom properties failed.",
      name: "CustomPropertiesLoadError",
    });
  });

  it("surfaces custom property save failures", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.customProperties.failNextSave = true;
    const renderStateStore = createOfficeRenderStateStore(
      mailboxItem,
      runtimeChannelConfig
    );

    await expect(
      renderStateStore.setPendingRenderState("<div>Original</div>")
    ).rejects.toMatchObject({
      message: "Custom property save failed.",
      name: "CustomPropertiesSaveError",
    });
  });

  it("falls back to session data when the render state is too large for custom properties", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.customProperties.nextSaveError = {
      message:
        "Specified argument was out of the range of valid values. Parameter name: customProperties",
      name: "Sys.ArgumentOutOfRangeException",
    };
    const renderStateStore = createOfficeRenderStateStore(
      mailboxItem,
      runtimeChannelConfig
    );
    const largeHtml = "<div>" + "A".repeat(4000) + "</div>";

    await renderStateStore.setPendingRenderState(largeHtml);

    expect(await renderStateStore.getRenderState()).toEqual(
      expect.objectContaining({
        channelId: "production",
        originalHtml: largeHtml,
        phase: "pending",
      })
    );
    expect(
      mailboxItem.customProperties.get("markout.production.originalHtml")
    ).toBeUndefined();
    expect(
      mailboxItem.sessionData.get("markout.production.originalHtml")
    ).toContain('"phase":"pending"');

    await renderStateStore.clearRenderState();
    expect(
      mailboxItem.sessionData.get("markout.production.originalHtml")
    ).toBeUndefined();
  });

  it("fails closed when large render state cannot be stored in session data", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.customProperties.nextSaveError = {
      message:
        "Specified argument was out of the range of valid values. Parameter name: customProperties",
      name: "Sys.ArgumentOutOfRangeException",
    };
    mailboxItem.sessionData.nextSetError = {
      message: "Quota exceeded.",
      name: "QuotaExceededError",
    };
    const renderStateStore = createOfficeRenderStateStore(
      mailboxItem,
      runtimeChannelConfig
    );
    const largeHtml = "<div>" + "A".repeat(80000) + "</div>";

    await expect(
      renderStateStore.setPendingRenderState(largeHtml)
    ).rejects.toMatchObject({
      code: "restore-state-too-large",
      message:
        "MarkOut couldn't persist the original draft HTML because Outlook's restore-state storage is full for this channel.",
      name: "MarkOutError",
    });
  });

  it("treats missing session-data keys as an empty render state", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.sessionData.nextGetError = {
      message: "The specified key was not found.",
      name: "KeyNotFound",
    };
    const renderStateStore = createOfficeRenderStateStore(
      mailboxItem,
      runtimeChannelConfig
    );

    await expect(renderStateStore.getRenderState()).resolves.toBeNull();
  });

  it("ignores missing session-data keys during cleanup", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    const renderStateStore = createOfficeRenderStateStore(
      mailboxItem,
      runtimeChannelConfig
    );

    await renderStateStore.setPendingRenderState("<div>Original</div>");
    mailboxItem.sessionData.nextRemoveError = {
      message: "The specified key was not found.",
      name: "KeyNotFound",
    };

    await expect(renderStateStore.clearRenderState()).resolves.toBeUndefined();
  });

  it("fails when no active compose item is available", () => {
    installOfficeEnvironment({ mailboxItem: undefined });

    try {
      createOfficeRenderStateStore();
      throw new Error("Expected createOfficeRenderStateStore to throw");
    } catch (error) {
      expect(error).toMatchObject({
        code: "office-compose-item-missing",
        message: "MarkOut requires an active Outlook compose item.",
        name: "MarkOutError",
      });
    }
  });
});
