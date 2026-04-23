import {
  FakeMailboxItem,
  createCommandEvent,
  installOfficeEnvironment,
} from "./helpers";

describe("launch events", () => {
  beforeEach(() => {
    jest.resetModules();
    jest.spyOn(console, "error").mockImplementation(() => undefined);
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  it("registers Smart Alert handlers on Office ready", async () => {
    const environment = installOfficeEnvironment({
      mailboxItem: new FakeMailboxItem("<div>Original</div>"),
    });
    const launchEventModule = await import("../src/launchevent/launchevent");

    await environment.triggerReady();

    expect(Office.actions.associate).toHaveBeenCalledWith(
      "onMessageSendHandler",
      launchEventModule.onMessageSendHandler
    );
    expect(Office.actions.associate).toHaveBeenCalledWith(
      "onAppointmentSendHandler",
      launchEventModule.onAppointmentSendHandler
    );
  });

  it("allows send without rendering when auto-render is disabled", async () => {
    installOfficeEnvironment({
      mailboxItem: new FakeMailboxItem("<div>Original</div>"),
    });
    const ensureRenderedMock = jest.fn();
    jest.doMock("../src/lib/item", () => ({
      createItemRenderer: () => ({
        ensureRendered: ensureRenderedMock,
      }),
    }));

    const { onMessageSendHandler } =
      await import("../src/launchevent/launchevent");
    const event = createCommandEvent();

    await onMessageSendHandler(event);

    expect(ensureRenderedMock).not.toHaveBeenCalled();
    expect(event.completed).toHaveBeenCalledWith({ allowEvent: true });
  });

  it("allows send after a successful auto-render", async () => {
    const environment = installOfficeEnvironment({
      mailboxItem: new FakeMailboxItem("<div>Original</div>"),
    });
    environment.roamingSettings.set("markout.autorender", true);
    const ensureRenderedMock = jest.fn().mockResolvedValue(true);
    jest.doMock("../src/lib/item", () => ({
      createItemRenderer: () => ({
        ensureRendered: ensureRenderedMock,
      }),
    }));

    const { onMessageSendHandler } =
      await import("../src/launchevent/launchevent");
    const event = createCommandEvent();

    await onMessageSendHandler(event);

    expect(ensureRenderedMock).toHaveBeenCalledTimes(1);
    expect(event.completed).toHaveBeenCalledWith({ allowEvent: true });
  });

  it("soft-blocks send when auto-render fails", async () => {
    const environment = installOfficeEnvironment({
      mailboxItem: new FakeMailboxItem("<div>Original</div>"),
    });
    environment.roamingSettings.set("markout.autorender", true);
    jest.doMock("../src/lib/item", () => ({
      createItemRenderer: () => ({
        ensureRendered: jest.fn().mockRejectedValue(new Error("boom")),
      }),
    }));

    const { onAppointmentSendHandler } =
      await import("../src/launchevent/launchevent");
    const event = createCommandEvent();

    await onAppointmentSendHandler(event);

    expect(event.completed).toHaveBeenCalledWith({
      allowEvent: false,
      errorMessage:
        "MarkOut could not render this draft before send. Open the MarkOut task pane, review the content, then try again.",
    });
  });

  it("soft-blocks send when the lazy render runtime cannot be loaded", async () => {
    const environment = installOfficeEnvironment({
      mailboxItem: new FakeMailboxItem("<div>Original</div>"),
    });
    environment.roamingSettings.set("markout.autorender", true);
    jest.doMock("../src/lib/item", () => {
      throw new Error("chunk load failed");
    });

    const { onMessageSendHandler } =
      await import("../src/launchevent/launchevent");
    const event = createCommandEvent();

    await onMessageSendHandler(event);

    expect(event.completed).toHaveBeenCalledWith({
      allowEvent: false,
      errorMessage:
        "MarkOut could not render this draft before send. Open the MarkOut task pane, review the content, then try again.",
    });
  });

  it("soft-blocks send when the compose item context is missing", async () => {
    const environment = installOfficeEnvironment({ mailboxItem: undefined });
    environment.roamingSettings.set("markout.autorender", true);

    const { onMessageSendHandler } =
      await import("../src/launchevent/launchevent");
    const event = createCommandEvent();

    await onMessageSendHandler(event);

    expect(event.completed).toHaveBeenCalledWith({
      allowEvent: false,
      errorMessage:
        "MarkOut could not render this draft before send. Open the MarkOut task pane, review the content, then try again.",
    });
  });
});
