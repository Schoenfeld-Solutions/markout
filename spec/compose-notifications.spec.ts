import { createComposeNotificationService } from "../src/lib/compose-notifications";
import { FakeMailboxItem, installOfficeEnvironment } from "./helpers";

describe("compose notification service", () => {
  beforeEach(() => {
    installOfficeEnvironment();
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
});
