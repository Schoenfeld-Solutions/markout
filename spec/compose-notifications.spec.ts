import { createComposeNotificationService } from "../src/lib/compose-notifications";
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
});
