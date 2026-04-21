import { createOfficeRenderStateStore } from "../src/lib/render-state-store";
import { FakeMailboxItem, installOfficeEnvironment } from "./helpers";

describe("render state store", () => {
  beforeEach(() => {
    installOfficeEnvironment();
  });

  it("persists rendered and pending render states", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    const renderStateStore = createOfficeRenderStateStore(mailboxItem);

    await renderStateStore.setPendingRenderState("<div>Original</div>");
    expect(await renderStateStore.getRenderState()).toEqual({
      originalHtml: "<div>Original</div>",
      phase: "pending",
    });

    await renderStateStore.setRenderedRenderState("<div>Original</div>");
    expect(await renderStateStore.getRenderState()).toEqual({
      originalHtml: "<div>Original</div>",
      phase: "rendered",
    });

    await renderStateStore.clearRenderState();
    expect(await renderStateStore.getRenderState()).toBeNull();
  });

  it("supports legacy raw-html and false sentinel values", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.customProperties.set(
      "markout.originalHtml",
      "<div>Legacy</div>"
    );

    const renderStateStore = createOfficeRenderStateStore(mailboxItem);
    expect(await renderStateStore.getRenderState()).toEqual({
      originalHtml: "<div>Legacy</div>",
      phase: "rendered",
    });

    mailboxItem.customProperties.set("markout.originalHtml", "false");
    expect(await renderStateStore.getRenderState()).toBeNull();
  });

  it("surfaces custom property load failures", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.failNextLoadCustomProperties = true;
    const renderStateStore = createOfficeRenderStateStore(mailboxItem);

    await expect(renderStateStore.getRenderState()).rejects.toMatchObject({
      message: "Loading custom properties failed.",
      name: "CustomPropertiesLoadError",
    });
  });

  it("surfaces custom property save failures", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.customProperties.failNextSave = true;
    const renderStateStore = createOfficeRenderStateStore(mailboxItem);

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
    const renderStateStore = createOfficeRenderStateStore(mailboxItem);
    const largeHtml = "<div>" + "A".repeat(4000) + "</div>";

    await renderStateStore.setPendingRenderState(largeHtml);

    expect(await renderStateStore.getRenderState()).toEqual({
      originalHtml: largeHtml,
      phase: "pending",
    });
    expect(
      mailboxItem.customProperties.get("markout.originalHtml")
    ).toBeUndefined();
    expect(mailboxItem.sessionData.get("markout.originalHtml")).toContain(
      '"phase":"pending"'
    );

    await renderStateStore.clearRenderState();
    expect(mailboxItem.sessionData.get("markout.originalHtml")).toBeUndefined();
  });

  it("fails when no active compose item is available", () => {
    installOfficeEnvironment({ mailboxItem: undefined });

    expect(() => createOfficeRenderStateStore()).toThrow(
      "MarkOut requires an active Outlook compose item."
    );
  });
});
