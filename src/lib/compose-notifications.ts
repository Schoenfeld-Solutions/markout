const AUTO_RENDER_DISMISSED_KEY = "markout.autorender.notificationDismissed";
const AUTO_RENDER_NOTIFICATION_KEY = "markout.autorender.notification";

export type NotificationSurface = "outlook" | "pane";

export interface AutoRenderNotificationCopy {
  message: string;
}

export interface ComposeNotificationService {
  clearAutoRenderDismissed(): Promise<void>;
  clearAutoRenderNotification(): Promise<void>;
  hasAutoRenderBeenDismissed(): Promise<boolean>;
  markAutoRenderDismissed(): Promise<void>;
  onAutoRenderDismiss(handler: () => void): void;
  showAutoRenderNotification(
    copy: AutoRenderNotificationCopy
  ): Promise<NotificationSurface>;
}

interface AsyncSessionDataLike {
  getAsync(
    name: string,
    callback: (result: Office.AsyncResult<string | undefined>) => void
  ): void;
  removeAsync(
    name: string,
    callback: (result: Office.AsyncResult<void>) => void
  ): void;
  setAsync(
    name: string,
    value: string,
    callback: (result: Office.AsyncResult<void>) => void
  ): void;
}

interface NotificationMessagesLike {
  removeAsync?(
    key: string,
    callback?: (result: Office.AsyncResult<void>) => void
  ): void;
  replaceAsync(
    key: string,
    details: Office.NotificationMessageDetails,
    callback?: (result: Office.AsyncResult<void>) => void
  ): void;
}

interface NotificationAwareItemLike {
  addHandlerAsync?(
    eventType: Office.EventType,
    handler: (event: Office.InfobarClickedEventArgs) => void,
    callback?: (result: Office.AsyncResult<void>) => void
  ): void;
  notificationMessages?: NotificationMessagesLike;
  sessionData?: AsyncSessionDataLike;
}

const inMemoryDismissals = new WeakMap<NotificationAwareItemLike, boolean>();

class OutlookComposeNotificationService implements ComposeNotificationService {
  public constructor(private readonly item: NotificationAwareItemLike | null) {}

  public async clearAutoRenderDismissed(): Promise<void> {
    const currentItem = this.item;

    if (currentItem === null) {
      return;
    }

    const sessionData = currentItem.sessionData;

    if (sessionData === undefined) {
      inMemoryDismissals.delete(currentItem);
      return;
    }

    await new Promise<void>((resolve, reject) => {
      sessionData.removeAsync(AUTO_RENDER_DISMISSED_KEY, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(toOfficeError(result.error));
          return;
        }

        resolve();
      });
    })
      .then(() => {
        inMemoryDismissals.delete(currentItem);
      })
      .catch(() => {
        inMemoryDismissals.delete(currentItem);
      });
  }

  public async clearAutoRenderNotification(): Promise<void> {
    const notificationMessages = this.item?.notificationMessages;
    const removeNotification =
      notificationMessages?.removeAsync?.bind(notificationMessages);

    if (
      notificationMessages === undefined ||
      typeof removeNotification !== "function"
    ) {
      return;
    }

    await new Promise<void>((resolve, reject) => {
      removeNotification.call(
        notificationMessages,
        AUTO_RENDER_NOTIFICATION_KEY,
        (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            reject(toOfficeError(result.error));
            return;
          }

          resolve();
        }
      );
    }).catch(() => undefined);
  }

  public async hasAutoRenderBeenDismissed(): Promise<boolean> {
    const currentItem = this.item;

    if (currentItem === null) {
      return false;
    }

    const sessionData = currentItem.sessionData;

    if (sessionData === undefined) {
      return inMemoryDismissals.get(currentItem) === true;
    }

    try {
      const dismissedValue = await new Promise<string | undefined>(
        (resolve, reject) => {
          sessionData.getAsync(AUTO_RENDER_DISMISSED_KEY, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
              reject(toOfficeError(result.error));
              return;
            }

            resolve(result.value);
          });
        }
      );

      if (dismissedValue === undefined) {
        return inMemoryDismissals.get(currentItem) === true;
      }

      return dismissedValue === "true";
    } catch {
      return inMemoryDismissals.get(currentItem) === true;
    }
  }

  public async markAutoRenderDismissed(): Promise<void> {
    const currentItem = this.item;

    if (currentItem === null) {
      return;
    }

    inMemoryDismissals.set(currentItem, true);
    const sessionData = currentItem.sessionData;

    if (sessionData === undefined) {
      return;
    }

    await new Promise<void>((resolve, reject) => {
      sessionData.setAsync(AUTO_RENDER_DISMISSED_KEY, "true", (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(toOfficeError(result.error));
          return;
        }

        resolve();
      });
    }).catch(() => undefined);
  }

  public onAutoRenderDismiss(handler: () => void): void {
    if (
      this.item === null ||
      typeof this.item.addHandlerAsync !== "function" ||
      typeof Office === "undefined"
    ) {
      return;
    }

    this.item.addHandlerAsync(Office.EventType.InfobarClicked, (event) => {
      const actionType = String(event.infobarDetails.actionType);
      const informationalType = String(event.infobarDetails.infobarType);

      if (
        actionType !== String(Office.MailboxEnums.InfobarActionType.Dismiss) ||
        informationalType !==
          String(Office.MailboxEnums.InfobarType.Informational)
      ) {
        return;
      }

      void this.markAutoRenderDismissed().finally(() => {
        handler();
      });
    });
  }

  public async showAutoRenderNotification(
    copy: AutoRenderNotificationCopy
  ): Promise<NotificationSurface> {
    const notificationMessages = this.item?.notificationMessages;

    if (notificationMessages === undefined) {
      return "pane";
    }

    try {
      await new Promise<void>((resolve, reject) => {
        notificationMessages.replaceAsync(
          AUTO_RENDER_NOTIFICATION_KEY,
          {
            icon: "Icon.16x16",
            message: copy.message,
            persistent: true,
            type: Office.MailboxEnums.ItemNotificationMessageType
              .InformationalMessage,
          },
          (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
              reject(toOfficeError(result.error));
              return;
            }

            resolve();
          }
        );
      });

      return "outlook";
    } catch {
      return "pane";
    }
  }
}

function getCurrentComposeItem(): NotificationAwareItemLike | null {
  if (typeof Office === "undefined") {
    return null;
  }

  const mailbox = Office.context.mailbox;
  if (mailbox.item === undefined) {
    return null;
  }

  return mailbox.item;
}

function toOfficeError(
  error: { message: string; name: string } | undefined
): Error {
  const normalizedError = new Error(
    error?.message ?? "An unknown Outlook error occurred."
  );
  normalizedError.name = error?.name ?? "OfficeAsyncError";
  return normalizedError;
}

export function createComposeNotificationService(
  item: NotificationAwareItemLike | null = getCurrentComposeItem()
): ComposeNotificationService {
  return new OutlookComposeNotificationService(item);
}
