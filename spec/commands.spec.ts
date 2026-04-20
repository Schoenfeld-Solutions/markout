import {
  FakeMailboxItem,
  createCommandEvent,
  installOfficeEnvironment,
} from "./helpers";

describe("commands", () => {
  beforeEach(() => {
    jest.resetModules();
    jest.spyOn(console, "error").mockImplementation(() => undefined);
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  it("registers the command handler on Office ready", async () => {
    const environment = installOfficeEnvironment({
      mailboxItem: new FakeMailboxItem("<div>Original</div>"),
    });
    const commandsModule = await import("../src/commands/commands");

    await environment.triggerReady();

    expect(Office.actions.associate).toHaveBeenCalledWith(
      "renderCurrentItem",
      commandsModule.renderCurrentItem
    );
  });

  it("completes the command event when rendering succeeds", async () => {
    installOfficeEnvironment({
      mailboxItem: new FakeMailboxItem("<div>Original</div>"),
    });
    const renderItemMock = jest.fn().mockResolvedValue("rendered");
    jest.doMock("../src/lib/item", () => ({
      renderItem: renderItemMock,
    }));

    const { renderCurrentItem } = await import("../src/commands/commands");
    const event = createCommandEvent();

    await renderCurrentItem(event);

    expect(renderItemMock).toHaveBeenCalledTimes(1);
    expect(event.completed).toHaveBeenCalledTimes(1);
  });

  it("shows a notification when rendering fails and an active item exists", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    installOfficeEnvironment({ mailboxItem });
    jest.doMock("../src/lib/item", () => ({
      renderItem: jest.fn().mockRejectedValue(new Error("boom")),
    }));

    const { renderCurrentItem } = await import("../src/commands/commands");
    const event = createCommandEvent();

    await renderCurrentItem(event);

    expect(mailboxItem.notificationMessages.replaceAsync).toHaveBeenCalledWith(
      "markout.render",
      expect.objectContaining({
        message:
          "MarkOut could not render this draft. Open the task pane to inspect the content and try again.",
      })
    );
    expect(event.completed).toHaveBeenCalledTimes(1);
  });

  it("completes cleanly when rendering fails without an active item", async () => {
    installOfficeEnvironment({ mailboxItem: undefined });
    jest.doMock("../src/lib/item", () => ({
      renderItem: jest.fn().mockRejectedValue(new Error("boom")),
    }));

    const { renderCurrentItem } = await import("../src/commands/commands");
    const event = createCommandEvent();

    await renderCurrentItem(event);

    expect(event.completed).toHaveBeenCalledTimes(1);
  });

  it("keeps completion stable when showing the notification throws", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.throwOnNotificationReplace = true;
    installOfficeEnvironment({ mailboxItem });
    jest.doMock("../src/lib/item", () => ({
      renderItem: jest.fn().mockRejectedValue(new Error("boom")),
    }));

    const { renderCurrentItem } = await import("../src/commands/commands");
    const event = createCommandEvent();

    await expect(renderCurrentItem(event)).resolves.toBeUndefined();
    expect(event.completed).toHaveBeenCalledTimes(1);
  });
});
