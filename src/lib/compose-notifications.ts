import {
  getChannelScopedKey,
  resolveRuntimeChannelConfig,
  type RuntimeChannelConfig,
} from "./runtime";

const TRANSIENT_NOTIFICATION_TIMEOUT_MS = 4200;
const NOTIFICATION_MESSAGE_MAX_LENGTH = 145;

export type NotificationSurface = "outlook" | "pane";
export type NotificationIntent = "error" | "info" | "success" | "warning";

export interface AutoRenderNotificationCopy {
  message: string;
}

export interface ComposeTransientNotificationCopy {
  intent: NotificationIntent;
  message: string;
}

export interface ComposeNotificationService {
  clearAutoRenderDismissed(): Promise<void>;
  clearAutoRenderNotification(): Promise<void>;
  clearTransientNotification(): Promise<void>;
  hasAutoRenderBeenDismissed(): Promise<boolean>;
  markAutoRenderDismissed(): Promise<void>;
  onAutoRenderDismiss(handler: () => void): void;
  showAutoRenderNotification(
    copy: AutoRenderNotificationCopy
  ): Promise<NotificationSurface>;
  showTransientNotification(
    copy: ComposeTransientNotificationCopy
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
  addAsync?(
    key: string,
    details: Office.NotificationMessageDetails,
    callback?: (result: Office.AsyncResult<void>) => void
  ): void;
  removeAsync?(
    key: string,
    callback?: (result: Office.AsyncResult<void>) => void
  ): void;
  replaceAsync?(
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
  private transientGeneration = 0;
  private transientTimeoutId: ReturnType<typeof setTimeout> | null = null;

  public constructor(
    private readonly item: NotificationAwareItemLike | null,
    private readonly runtimeChannelConfig: RuntimeChannelConfig
  ) {}

  private getAutoRenderDismissedKey(): string {
    return getChannelScopedKey(
      this.runtimeChannelConfig,
      "autorender.notificationDismissed"
    );
  }

  private getAutoRenderNotificationKey(): string {
    return getChannelScopedKey(
      this.runtimeChannelConfig,
      "autorender.notification"
    );
  }

  private getTransientNotificationKey(): string {
    return getChannelScopedKey(
      this.runtimeChannelConfig,
      "compose.notification"
    );
  }

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
      sessionData.removeAsync(this.getAutoRenderDismissedKey(), (result) => {
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
    await this.removeNotification(this.getAutoRenderNotificationKey()).catch(
      () => undefined
    );
  }

  public async clearTransientNotification(): Promise<void> {
    this.transientGeneration += 1;
    this.clearTransientTimeout();

    await this.removeNotification(this.getTransientNotificationKey()).catch(
      () => undefined
    );
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
          sessionData.getAsync(this.getAutoRenderDismissedKey(), (result) => {
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
      sessionData.setAsync(
        this.getAutoRenderDismissedKey(),
        "true",
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

      if (this.transientTimeoutId !== null) {
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
    try {
      await this.upsertNotification(this.getAutoRenderNotificationKey(), {
        icon: "Icon.16x16",
        message: normalizeNotificationMessage(copy.message),
        persistent: true,
        type: Office.MailboxEnums.ItemNotificationMessageType
          .InformationalMessage,
      });

      return "outlook";
    } catch {
      return "pane";
    }
  }

  public async showTransientNotification(
    copy: ComposeTransientNotificationCopy
  ): Promise<NotificationSurface> {
    const generation = this.transientGeneration + 1;
    this.transientGeneration = generation;
    this.clearTransientTimeout();

    for (const details of buildTransientNotificationAttempts(copy)) {
      try {
        await this.upsertNotification(
          this.getTransientNotificationKey(),
          details
        );
        this.scheduleTransientRemoval(generation);
        return "outlook";
      } catch {
        continue;
      }
    }

    return "pane";
  }

  private clearTransientTimeout(): void {
    if (this.transientTimeoutId !== null) {
      globalThis.clearTimeout(this.transientTimeoutId);
      this.transientTimeoutId = null;
    }
  }

  private async removeNotification(key: string): Promise<void> {
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
      removeNotification.call(notificationMessages, key, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(toOfficeError(result.error));
          return;
        }

        resolve();
      });
    });
  }

  private async upsertNotification(
    key: string,
    details: Office.NotificationMessageDetails
  ): Promise<void> {
    const notificationMessages = this.item?.notificationMessages;

    if (
      notificationMessages === undefined ||
      (typeof notificationMessages.replaceAsync !== "function" &&
        typeof notificationMessages.addAsync !== "function")
    ) {
      throw new Error("Outlook notification messages are not available.");
    }

    await new Promise<void>((resolve, reject) => {
      const handler = (result: Office.AsyncResult<void>) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(toOfficeError(result.error));
          return;
        }

        resolve();
      };

      if (typeof notificationMessages.replaceAsync === "function") {
        notificationMessages.replaceAsync(key, details, handler);
        return;
      }

      notificationMessages.addAsync?.(key, details, handler);
    });
  }

  private scheduleTransientRemoval(generation: number): void {
    this.transientTimeoutId = globalThis.setTimeout(() => {
      if (this.transientGeneration !== generation) {
        return;
      }

      this.transientTimeoutId = null;
      void this.removeNotification(this.getTransientNotificationKey()).catch(
        () => undefined
      );
    }, TRANSIENT_NOTIFICATION_TIMEOUT_MS);
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

function mapNotificationIntent(
  intent: NotificationIntent
): Office.MailboxEnums.ItemNotificationMessageType {
  return intent === "error"
    ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
    : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage;
}

function buildTransientNotificationAttempts(
  copy: ComposeTransientNotificationCopy
): Office.NotificationMessageDetails[] {
  const message = normalizeNotificationMessage(copy.message);
  const informationalMessage = {
    icon: "Icon.16x16",
    message,
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
  } satisfies Office.NotificationMessageDetails;

  if (copy.intent !== "error") {
    return [informationalMessage];
  }

  return [
    {
      message,
      type: mapNotificationIntent(copy.intent),
    } satisfies Office.NotificationMessageDetails,
    informationalMessage,
  ];
}

function normalizeNotificationMessage(message: string): string {
  const normalizedMessage = message.replaceAll(/\s+/g, " ").trim();

  if (normalizedMessage.length <= NOTIFICATION_MESSAGE_MAX_LENGTH) {
    return normalizedMessage;
  }

  return `${normalizedMessage.slice(0, NOTIFICATION_MESSAGE_MAX_LENGTH - 3).trimEnd()}...`;
}

export function createComposeNotificationService(
  item: NotificationAwareItemLike | null = getCurrentComposeItem(),
  runtimeChannelConfig: RuntimeChannelConfig = resolveRuntimeChannelConfig()
): ComposeNotificationService {
  return new OutlookComposeNotificationService(item, runtimeChannelConfig);
}
