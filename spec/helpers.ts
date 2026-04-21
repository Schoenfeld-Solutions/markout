import { readFileSync } from "fs";
import { JSDOM } from "jsdom";
import { join } from "path";

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
};

export function installDomParser(): void {
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
  public failNextSet = false;

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
}

export class FakeMailboxItem {
  public readonly body: FakeBody;
  public readonly customProperties = new FakeCustomProperties();
  public readonly sessionData = new FakeSessionData();
  public failNextLoadCustomProperties = false;
  public throwOnNotificationReplace = false;
  public readonly notificationMessages = {
    replaceAsync: jest.fn(
      (
        _key: string,
        _details: Office.NotificationMessageDetails,
        callback?: (result: Office.AsyncResult<void>) => void
      ) => {
        if (this.throwOnNotificationReplace) {
          throw new Error("Notification replace failed.");
        }

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
}

export interface FakeOfficeEnvironment {
  mailboxItem: FakeMailboxItem | undefined;
  roamingSettings: FakeRoamingSettings;
  triggerReady(host?: Office.HostType): Promise<void>;
}

export function installOfficeEnvironment(options?: {
  mailboxItem?: FakeMailboxItem | undefined;
  roamingSettings?: FakeRoamingSettings;
}): FakeOfficeEnvironment {
  const readyCallbacks: ((info: ReadyInfo) => void | Promise<void>)[] = [];
  const mailboxItem = options?.mailboxItem;
  const roamingSettings = options?.roamingSettings ?? new FakeRoamingSettings();
  const outlookHost = "Outlook" as unknown as Office.HostType;

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
    },
    context: {
      mailbox: {
        item: mailboxItem,
      },
      roamingSettings,
    },
    HostType: {
      Outlook: "Outlook",
    },
    MailboxEnums: {
      ItemNotificationMessageType: {
        ErrorMessage: "errorMessage",
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
