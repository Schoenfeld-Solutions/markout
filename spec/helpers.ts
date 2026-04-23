import { readFileSync } from "fs";
import { join } from "path";
import type { JSDOM as JSDOMType } from "jsdom";
import { TextDecoder, TextEncoder } from "util";

interface OfficeErrorLike {
  message: string;
  name: string;
}

interface ReadyInfo {
  host: Office.HostType;
}

const runtime = globalThis as typeof globalThis & {
  DOMParser?: typeof DOMParser;
  Office?: typeof Office;
  TextDecoder?: typeof TextDecoder;
  TextEncoder?: typeof TextEncoder;
};

runtime.TextDecoder = TextDecoder;
runtime.TextEncoder = TextEncoder;

export function installDomParser(): void {
  // eslint-disable-next-line @typescript-eslint/no-require-imports
  const { JSDOM } = require("jsdom") as { JSDOM: typeof JSDOMType };
  runtime.DOMParser = new JSDOM().window.DOMParser;
}

export function readFile(name: string, stripNewlines = false): string {
  let content = readFileSync(join(__dirname, name), "utf8");

  content = content.replace(/\r/g, "");
  if (stripNewlines) {
    content = content.replace(/\r?\n/g, "").trim();
  }

  return content;
}

export class FakeRoamingSettings {
  public failNextSave = false;
  private readonly values = new Map<string, unknown>();

  public get(name: string): unknown {
    return this.values.get(name);
  }

  public saveAsync(callback: (result: Office.AsyncResult<void>) => void): void {
    if (this.failNextSave) {
      this.failNextSave = false;
      callback(
        failedAsyncResult<void>({
          message: "Roaming settings save failed.",
          name: "RoamingSettingsSaveError",
        })
      );
      return;
    }

    callback(succeededAsyncResult<void>(undefined));
  }

  public set(name: string, value: unknown): void {
    this.values.set(name, value);
  }
}

export class FakeCustomProperties {
  public failNextSave = false;
  public nextSaveError: OfficeErrorLike | null = null;
  private readonly values = new Map<string, string>();

  public get(name: string): string | undefined {
    return this.values.get(name);
  }

  public remove(name: string): void {
    this.values.delete(name);
  }

  public saveAsync(callback: (result: Office.AsyncResult<void>) => void): void {
    if (this.nextSaveError !== null) {
      const saveError = this.nextSaveError;
      this.nextSaveError = null;
      callback(failedAsyncResult<void>(saveError));
      return;
    }

    if (this.failNextSave) {
      this.failNextSave = false;
      callback(
        failedAsyncResult<void>({
          message: "Custom property save failed.",
          name: "CustomPropertiesSaveError",
        })
      );
      return;
    }

    callback(succeededAsyncResult<void>(undefined));
  }

  public set(name: string, value: string): void {
    this.values.set(name, value);
  }
}

export class FakeSessionData {
  public nextGetError: OfficeErrorLike | null = null;
  public nextRemoveError: OfficeErrorLike | null = null;
  public nextSetError: OfficeErrorLike | null = null;
  private readonly values = new Map<string, string>();

  public get(name: string): string | undefined {
    return this.values.get(name);
  }

  public getAsync(
    name: string,
    callback: (result: Office.AsyncResult<string | undefined>) => void
  ): void {
    if (this.nextGetError !== null) {
      const getError = this.nextGetError;
      this.nextGetError = null;
      callback(failedAsyncResult<string | undefined>(getError));
      return;
    }

    callback(succeededAsyncResult<string | undefined>(this.values.get(name)));
  }

  public removeAsync(
    name: string,
    callback: (result: Office.AsyncResult<void>) => void
  ): void {
    if (this.nextRemoveError !== null) {
      const removeError = this.nextRemoveError;
      this.nextRemoveError = null;
      callback(failedAsyncResult<void>(removeError));
      return;
    }

    this.values.delete(name);
    callback(succeededAsyncResult<void>(undefined));
  }

  public set(name: string, value: string): void {
    this.values.set(name, value);
  }

  public setAsync(
    name: string,
    value: string,
    callback: (result: Office.AsyncResult<void>) => void
  ): void {
    if (this.nextSetError !== null) {
      const setError = this.nextSetError;
      this.nextSetError = null;
      callback(failedAsyncResult<void>(setError));
      return;
    }

    this.values.set(name, value);
    callback(succeededAsyncResult<void>(undefined));
  }
}

export class FakeBody {
  public failNextGet = false;
  public failNextGetType = false;
  public failNextSetSelected = false;
  public failNextSet = false;
  public lastSelectedHtml = "";
  public type: Office.CoercionType = "html" as Office.CoercionType;

  public constructor(private html: string) {}

  public get currentHtml(): string {
    return this.html;
  }

  public getAsync(
    _coercionType: Office.CoercionType,
    callback: (result: Office.AsyncResult<string>) => void
  ): void {
    if (this.failNextGet) {
      this.failNextGet = false;
      callback(
        failedAsyncResult<string>({
          message: "Body read failed.",
          name: "BodyGetError",
        })
      );
      return;
    }

    callback(succeededAsyncResult(this.html));
  }

  public getTypeAsync(
    callback: (result: Office.AsyncResult<Office.CoercionType>) => void
  ): void {
    if (this.failNextGetType) {
      this.failNextGetType = false;
      callback(
        failedAsyncResult<Office.CoercionType>({
          message: "Body type read failed.",
          name: "BodyTypeError",
        })
      );
      return;
    }

    callback(succeededAsyncResult(this.type));
  }

  public setAsync(
    value: string,
    _options: { coercionType: Office.CoercionType },
    callback: (result: Office.AsyncResult<void>) => void
  ): void {
    if (this.failNextSet) {
      this.failNextSet = false;
      callback(
        failedAsyncResult<void>({
          message: "Body write failed.",
          name: "BodySetError",
        })
      );
      return;
    }

    this.html = value;
    callback(succeededAsyncResult<void>(undefined));
  }

  public setSelectedDataAsync(
    value: string,
    _options: { coercionType: Office.CoercionType },
    callback: (result: Office.AsyncResult<void>) => void
  ): void {
    if (this.failNextSetSelected) {
      this.failNextSetSelected = false;
      callback(
        failedAsyncResult<void>({
          message: "Selected body write failed.",
          name: "BodySetSelectedError",
        })
      );
      return;
    }

    this.lastSelectedHtml = value;
    this.html = value;
    callback(succeededAsyncResult<void>(undefined));
  }
}

export class FakeMailboxItem {
  public readonly body: FakeBody;
  public readonly customProperties = new FakeCustomProperties();
  public infobarHandlers: ((
    event: Office.InfobarClickedEventArgs
  ) => void | Promise<void>)[] = [];
  public lastNotificationDetails: Office.NotificationMessageDetails | null =
    null;
  public notificationReplaceInterceptor:
    | ((details: Office.NotificationMessageDetails) => OfficeErrorLike | null)
    | null = null;
  public nextHtmlSelectionError: OfficeErrorLike | null = null;
  public nextTextSelectionError: OfficeErrorLike | null = null;
  public selectionHtml = "";
  public selectionSource: "body" | "subject" = "body";
  public selectionText = "";
  public readonly sessionData = new FakeSessionData();
  public failNextLoadCustomProperties = false;
  public failNextNotificationRemove = false;
  public failNextNotificationReplace = false;
  public throwOnNotificationReplace = false;
  public readonly notificationMessages = {
    addAsync: jest.fn(
      (
        _key: string,
        details: Office.NotificationMessageDetails,
        callback?: (result: Office.AsyncResult<void>) => void
      ) => {
        this.lastNotificationDetails = details;
        callback?.(succeededAsyncResult<void>(undefined));
      }
    ),
    removeAsync: jest.fn(
      (_key: string, callback?: (result: Office.AsyncResult<void>) => void) => {
        if (this.failNextNotificationRemove) {
          this.failNextNotificationRemove = false;
          callback?.(
            failedAsyncResult<void>({
              message: "Notification remove failed.",
              name: "NotificationRemoveError",
            })
          );
          return;
        }

        this.lastNotificationDetails = null;
        callback?.(succeededAsyncResult<void>(undefined));
      }
    ),
    replaceAsync: jest.fn(
      (
        _key: string,
        details: Office.NotificationMessageDetails,
        callback?: (result: Office.AsyncResult<void>) => void
      ) => {
        if (this.throwOnNotificationReplace) {
          throw new Error("Notification replace failed.");
        }

        const replaceError = this.notificationReplaceInterceptor?.(details);
        if (replaceError !== null && replaceError !== undefined) {
          callback?.(failedAsyncResult<void>(replaceError));
          return;
        }

        if (this.failNextNotificationReplace) {
          this.failNextNotificationReplace = false;
          callback?.(
            failedAsyncResult<void>({
              message: "Notification replace failed.",
              name: "NotificationReplaceError",
            })
          );
          return;
        }

        this.lastNotificationDetails = details;
        callback?.(succeededAsyncResult<void>(undefined));
      }
    ),
  };

  public constructor(initialHtml: string) {
    this.body = new FakeBody(initialHtml);
  }

  public loadCustomPropertiesAsync(
    callback: (result: Office.AsyncResult<FakeCustomProperties>) => void
  ): void {
    if (this.failNextLoadCustomProperties) {
      this.failNextLoadCustomProperties = false;
      callback(
        failedAsyncResult<FakeCustomProperties>({
          message: "Loading custom properties failed.",
          name: "CustomPropertiesLoadError",
        })
      );
      return;
    }

    callback(succeededAsyncResult(this.customProperties));
  }

  public getSelectedDataAsync(
    coercionType: Office.CoercionType,
    callback: (
      result: Office.AsyncResult<{ data: string; sourceProperty: string }>
    ) => void
  ): void {
    if (
      coercionType === Office.CoercionType.Html &&
      this.nextHtmlSelectionError !== null
    ) {
      const error = this.nextHtmlSelectionError;
      this.nextHtmlSelectionError = null;
      callback(failedAsyncResult(error));
      return;
    }

    if (
      coercionType === Office.CoercionType.Text &&
      this.nextTextSelectionError !== null
    ) {
      const error = this.nextTextSelectionError;
      this.nextTextSelectionError = null;
      callback(failedAsyncResult(error));
      return;
    }

    callback(
      succeededAsyncResult({
        data:
          coercionType === Office.CoercionType.Html
            ? this.selectionHtml
            : this.selectionText,
        sourceProperty: this.selectionSource,
      })
    );
  }

  public addHandlerAsync(
    eventType: Office.EventType,
    handler: (event: Office.InfobarClickedEventArgs) => void | Promise<void>,
    callback?: (result: Office.AsyncResult<void>) => void
  ): void {
    if (eventType === Office.EventType.InfobarClicked) {
      this.infobarHandlers.push(handler);
    }

    callback?.(succeededAsyncResult<void>(undefined));
  }

  public async triggerInfobarDismiss(): Promise<void> {
    for (const handler of this.infobarHandlers) {
      await handler({
        infobarDetails: {
          actionType: Office.MailboxEnums.InfobarActionType.Dismiss,
          infobarType: Office.MailboxEnums.InfobarType.Informational,
        } as Office.InfobarDetails,
        type: Office.EventType.InfobarClicked,
      } as unknown as Office.InfobarClickedEventArgs);
    }
  }
}

export interface FakeOfficeEnvironment {
  mailboxItem: FakeMailboxItem | undefined;
  roamingSettings: FakeRoamingSettings;
  triggerOfficeThemeChange(
    officeTheme: Partial<Office.OfficeTheme>
  ): Promise<void>;
  triggerReady(host?: Office.HostType): Promise<void>;
}

export function installOfficeEnvironment(options?: {
  displayLanguage?: string;
  mailboxItem?: FakeMailboxItem | undefined;
  roamingSettings?: FakeRoamingSettings;
}): FakeOfficeEnvironment {
  const readyCallbacks: ((info: ReadyInfo) => void | Promise<void>)[] = [];
  const officeThemeChangedHandlers: ((
    args: Office.OfficeThemeChangedEventArgs
  ) => void | Promise<void>)[] = [];
  const mailboxItem = options?.mailboxItem;
  const displayLanguage = options?.displayLanguage ?? "en-US";
  const roamingSettings = options?.roamingSettings ?? new FakeRoamingSettings();
  const outlookHost = "Outlook" as unknown as Office.HostType;
  const officeTheme: Office.OfficeTheme = {
    bodyBackgroundColor: "#ffffff",
    bodyForegroundColor: "#1b1a19",
    controlBackgroundColor: "#ffffff",
    controlForegroundColor: "#1b1a19",
    isDarkTheme: false,
    themeId: 3,
  };

  runtime.Office = {
    actions: {
      associate: jest.fn(),
    },
    AddinCommands: {},
    AsyncResultStatus: {
      Failed: "failed",
      Succeeded: "succeeded",
    },
    CoercionType: {
      Html: "html",
      Text: "text",
    },
    context: {
      displayLanguage,
      mailbox: {
        addHandlerAsync: jest.fn(
          (
            eventType: Office.EventType,
            handler: (
              args: Office.OfficeThemeChangedEventArgs
            ) => void | Promise<void>,
            callback?: (result: Office.AsyncResult<void>) => void
          ) => {
            if (eventType === Office.EventType.OfficeThemeChanged) {
              officeThemeChangedHandlers.push(handler);
            }

            callback?.(succeededAsyncResult<void>(undefined));
          }
        ),
        item: mailboxItem,
        officeTheme,
      },
      roamingSettings,
    },
    EventType: {
      InfobarClicked: "olkInfobarClicked",
      OfficeThemeChanged: "officeThemeChanged",
    },
    HostType: {
      Outlook: "Outlook",
    },
    MailboxEnums: {
      InfobarActionType: {
        Dismiss: "Dismiss",
      },
      InfobarType: {
        Error: 2,
        Informational: 0,
        Insight: 3,
        ProgressIndicator: 1,
      },
      ItemNotificationMessageType: {
        ErrorMessage: "errorMessage",
        InformationalMessage: "informationalMessage",
      },
    },
    onReady(
      callback: (info: ReadyInfo) => void | Promise<void>
    ): Promise<ReadyInfo> {
      readyCallbacks.push(callback);
      return Promise.resolve({ host: outlookHost });
    },
  } as unknown as typeof Office;

  return {
    mailboxItem,
    roamingSettings,
    async triggerOfficeThemeChange(
      nextOfficeTheme: Partial<Office.OfficeTheme>
    ): Promise<void> {
      Object.assign(officeTheme, nextOfficeTheme);

      for (const handler of officeThemeChangedHandlers) {
        await handler({
          officeTheme,
          type: "officeThemeChanged",
        });
      }
    },
    async triggerReady(host = Office.HostType.Outlook): Promise<void> {
      for (const readyCallback of readyCallbacks) {
        await readyCallback({ host });
      }
    },
  };
}

export function createCommandEvent(): Office.AddinCommands.Event {
  return {
    completed: jest.fn(),
  } as unknown as Office.AddinCommands.Event;
}

export function failedAsyncResult<T>(
  error: OfficeErrorLike
): Office.AsyncResult<T> {
  return {
    error,
    status: Office.AsyncResultStatus.Failed,
  } as Office.AsyncResult<T>;
}

export function succeededAsyncResult<T>(value: T): Office.AsyncResult<T> {
  return {
    status: Office.AsyncResultStatus.Succeeded,
    value,
  } as Office.AsyncResult<T>;
}
