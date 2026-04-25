import { createComposeNotificationService } from "../src/lib/compose-notifications";
import { getRuntimeChannelConfig } from "../src/lib/runtime";
import { FakeMailboxItem, installOfficeEnvironment } from "./helpers";

describe("compose notification service", () => {
  beforeEach(() => {
    installOfficeEnvironment();
    jest.useFakeTimers();
  });

  afterEach(() => {
    jest.useRealTimers();
  });

  it("shows a persistent informational notification in Outlook when supported", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Draft</div>");
    const notificationService = createComposeNotificationService(mailboxItem);

    await expect(
      notificationService.showAutoRenderNotification({
        message: "MarkOut auto-render is enabled",
      })
    ).resolves.toBe("outlook");

    expect(mailboxItem.lastNotificationDetails).toMatchObject({
      message: "MarkOut auto-render is enabled",
      persistent: true,
      type: Office.MailboxEnums.ItemNotificationMessageType
        .InformationalMessage,
    });
  });

  it("falls back to a pane-local notification when Outlook replaceAsync fails", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Draft</div>");
    mailboxItem.failNextNotificationReplace = true;
    const notificationService = createComposeNotificationService(mailboxItem);

    await expect(
      notificationService.showAutoRenderNotification({
        message: "MarkOut auto-render is enabled",
      })
    ).resolves.toBe("pane");
  });

  it("shows transient compose notifications and clears them after a short delay", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Draft</div>");
    const notificationService = createComposeNotificationService(mailboxItem);

    await expect(
      notificationService.showTransientNotification({
        intent: "success",
        message: "Rendered Markdown was inserted at the current body cursor.",
      })
    ).resolves.toBe("outlook");

    expect(mailboxItem.lastNotificationDetails).toMatchObject({
      icon: "Icon.16x16",
      message: "Rendered Markdown was inserted at the current body cursor.",
      type: Office.MailboxEnums.ItemNotificationMessageType
        .InformationalMessage,
    });

    jest.runOnlyPendingTimers();
    await Promise.resolve();

    expect(mailboxItem.notificationMessages.removeAsync).toHaveBeenCalled();
    expect(mailboxItem.lastNotificationDetails).toBeNull();
  });

  it("keeps only the newest transient notification timer active", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Draft</div>");
    const notificationService = createComposeNotificationService(mailboxItem);

    await notificationService.showTransientNotification({
      intent: "info",
      message: "First transient message.",
    });
    await notificationService.showTransientNotification({
      intent: "success",
      message: "Second transient message.",
    });

    expect(mailboxItem.lastNotificationDetails).toMatchObject({
      message: "Second transient message.",
    });

    jest.runOnlyPendingTimers();
    await Promise.resolve();

    expect(mailboxItem.notificationMessages.removeAsync).toHaveBeenCalledTimes(
      1
    );
    expect(mailboxItem.lastNotificationDetails).toBeNull();
  });

  it("cancels a pending transient timer when the notification is cleared manually", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Draft</div>");
    const notificationService = createComposeNotificationService(mailboxItem);

    await notificationService.showTransientNotification({
      intent: "warning",
      message: "Drop a Markdown or text file to load content into MarkOut.",
    });
    await notificationService.clearTransientNotification();

    expect(mailboxItem.notificationMessages.removeAsync).toHaveBeenCalledTimes(
      1
    );

    jest.runOnlyPendingTimers();
    await Promise.resolve();

    expect(mailboxItem.notificationMessages.removeAsync).toHaveBeenCalledTimes(
      1
    );
  });

  it("uses the error infobar type for transient errors", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Draft</div>");
    const notificationService = createComposeNotificationService(mailboxItem);

    await notificationService.showTransientNotification({
      intent: "error",
      message: "Selection state could not be read from Outlook.",
    });

    expect(mailboxItem.lastNotificationDetails).toMatchObject({
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
    });
  });

  it("retries transient errors as informational infobars when the host rejects error message details", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Draft</div>");
    mailboxItem.notificationReplaceInterceptor = (details) =>
      String(details.type) ===
      String(Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage)
        ? {
            message: "Error infobar rejected.",
            name: "NotificationReplaceError",
          }
        : null;
    const notificationService = createComposeNotificationService(mailboxItem);

    await expect(
      notificationService.showTransientNotification({
        intent: "error",
        message: "Selection state could not be read from Outlook.",
      })
    ).resolves.toBe("outlook");

    expect(mailboxItem.notificationMessages.replaceAsync).toHaveBeenCalledTimes(
      2
    );
    expect(mailboxItem.lastNotificationDetails).toMatchObject({
      icon: "Icon.16x16",
      message: "Selection state could not be read from Outlook.",
      type: Office.MailboxEnums.ItemNotificationMessageType
        .InformationalMessage,
    });
  });

  it("does not treat a transient infobar dismiss as an auto-render dismissal", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Draft</div>");
    const notificationService = createComposeNotificationService(mailboxItem);
    const dismissHandler = jest.fn();

    notificationService.onAutoRenderDismiss(dismissHandler);
    await notificationService.showTransientNotification({
      intent: "info",
      message: "The current draft was rendered successfully.",
    });
    await mailboxItem.triggerInfobarDismiss();
    await Promise.resolve();

    expect(dismissHandler).not.toHaveBeenCalled();
    await expect(
      notificationService.hasAutoRenderBeenDismissed()
    ).resolves.toBe(false);
  });

  it("ignores non-dismiss informational infobar clicks", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Draft</div>");
    const notificationService = createComposeNotificationService(mailboxItem);
    const dismissHandler = jest.fn();

    notificationService.onAutoRenderDismiss(dismissHandler);

    for (const handler of mailboxItem.infobarHandlers) {
      await handler({
        infobarDetails: {
          actionType: "OpenTaskPane",
          infobarType: Office.MailboxEnums.InfobarType.Informational,
        } as unknown as Office.InfobarDetails,
        type: Office.EventType.InfobarClicked,
      } as unknown as Office.InfobarClickedEventArgs);
    }

    expect(dismissHandler).not.toHaveBeenCalled();
    await expect(
      notificationService.hasAutoRenderBeenDismissed()
    ).resolves.toBe(false);
  });

  it("ignores dismiss clicks for non-informational infobars", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Draft</div>");
    const notificationService = createComposeNotificationService(mailboxItem);
    const dismissHandler = jest.fn();

    notificationService.onAutoRenderDismiss(dismissHandler);

    for (const handler of mailboxItem.infobarHandlers) {
      await handler({
        infobarDetails: {
          actionType: Office.MailboxEnums.InfobarActionType.Dismiss,
          infobarType: Office.MailboxEnums.InfobarType.Error,
        } as Office.InfobarDetails,
        type: Office.EventType.InfobarClicked,
      } as unknown as Office.InfobarClickedEventArgs);
    }

    expect(dismissHandler).not.toHaveBeenCalled();
    await expect(
      notificationService.hasAutoRenderBeenDismissed()
    ).resolves.toBe(false);
  });

  it("persists infobar dismiss state per item and clears it again", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Draft</div>");
    const notificationService = createComposeNotificationService(mailboxItem);
    const dismissHandler = jest.fn();

    notificationService.onAutoRenderDismiss(dismissHandler);
    await mailboxItem.triggerInfobarDismiss();
    await Promise.resolve();

    expect(dismissHandler).toHaveBeenCalledTimes(1);
    await expect(
      notificationService.hasAutoRenderBeenDismissed()
    ).resolves.toBe(true);

    await notificationService.clearAutoRenderDismissed();

    await expect(
      notificationService.hasAutoRenderBeenDismissed()
    ).resolves.toBe(false);
  });

  it("scopes persisted auto-render dismissals by runtime channel", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Draft</div>");
    const productionNotificationService = createComposeNotificationService(
      mailboxItem,
      getRuntimeChannelConfig("production")
    );
    const betaNotificationService = createComposeNotificationService(
      mailboxItem,
      getRuntimeChannelConfig("beta")
    );

    await productionNotificationService.markAutoRenderDismissed();

    await expect(
      productionNotificationService.hasAutoRenderBeenDismissed()
    ).resolves.toBe(true);
    await expect(
      betaNotificationService.hasAutoRenderBeenDismissed()
    ).resolves.toBe(false);
  });

  it("falls back to in-memory dismiss tracking when sessionData writes fail", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Draft</div>");
    mailboxItem.sessionData.nextSetError = {
      message: "Session set failed.",
      name: "SessionSetError",
    };
    const notificationService = createComposeNotificationService(mailboxItem);

    await notificationService.markAutoRenderDismissed();

    await expect(
      notificationService.hasAutoRenderBeenDismissed()
    ).resolves.toBe(true);
  });

  it("keeps in-memory dismiss tracking when sessionData reads fail", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Draft</div>");
    const notificationService = createComposeNotificationService(mailboxItem);

    await notificationService.markAutoRenderDismissed();
    mailboxItem.sessionData.nextGetError = {
      message: "Session get failed.",
      name: "SessionGetError",
    };

    await expect(
      notificationService.hasAutoRenderBeenDismissed()
    ).resolves.toBe(true);
  });

  it("removes the current notification when clearAutoRenderNotification is called", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Draft</div>");
    const notificationService = createComposeNotificationService(mailboxItem);

    await notificationService.showAutoRenderNotification({
      message: "MarkOut auto-render is enabled",
    });

    await notificationService.clearAutoRenderNotification();

    expect(mailboxItem.lastNotificationDetails).toBeNull();
  });

  it("falls back to a pane-local transient notification when Outlook replaceAsync fails", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Draft</div>");
    mailboxItem.notificationReplaceInterceptor = () => ({
      message: "Notification replace failed.",
      name: "NotificationReplaceError",
    });
    const notificationService = createComposeNotificationService(mailboxItem);

    await expect(
      notificationService.showTransientNotification({
        intent: "warning",
        message: "Drop a Markdown or text file to load content into MarkOut.",
      })
    ).resolves.toBe("pane");
  });

  it("falls back to pane-local notifications when no Outlook item is available", async () => {
    const notificationService = createComposeNotificationService(null);

    await expect(
      notificationService.showAutoRenderNotification({
        message: "MarkOut auto-render is enabled",
      })
    ).resolves.toBe("pane");
    await expect(
      notificationService.showTransientNotification({
        intent: "info",
        message: "The current draft was rendered successfully.",
      })
    ).resolves.toBe("pane");
  });
});
