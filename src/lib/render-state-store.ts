import {
  MarkOutError,
  getChannelScopedKey,
  resolveRuntimeChannelConfig,
  type ChannelId,
  type RuntimeChannelConfig,
} from "./runtime";

interface CustomPropertiesLike {
  get(name: string): string | undefined;
  remove(name: string): void;
  saveAsync(callback: (result: Office.AsyncResult<void>) => void): void;
  set(name: string, value: string): void;
}

interface SessionDataLike {
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

interface MailboxItemWithCustomPropertiesLike {
  loadCustomPropertiesAsync(
    callback: (result: Office.AsyncResult<CustomPropertiesLike>) => void
  ): void;
  sessionData?: SessionDataLike;
}

export type RenderStatePhase = "pending" | "rendered";

export interface RenderState {
  channelId: ChannelId;
  originalHtml: string;
  phase: RenderStatePhase;
  storedAt: string;
}

export interface RenderStateStore {
  clearRenderState(): Promise<void>;
  getRenderState(): Promise<RenderState | null>;
  setPendingRenderState(originalHtml: string): Promise<void>;
  setRenderedRenderState(originalHtml: string): Promise<void>;
}

interface LegacyRenderStatePayload {
  channelId?: unknown;
  html?: unknown;
  originalHtml?: unknown;
  phase?: unknown;
  schemaVersion?: unknown;
  storedAt?: unknown;
}

interface StoredRenderStatePayload {
  channelId: ChannelId;
  originalHtml: string;
  phase: RenderStatePhase;
  schemaVersion: 2;
  storedAt: string;
}

class OfficeRenderStateStore implements RenderStateStore {
  public constructor(
    private readonly mailboxItem: MailboxItemWithCustomPropertiesLike,
    private readonly runtimeChannelConfig: RuntimeChannelConfig
  ) {}

  public async clearRenderState(): Promise<void> {
    const customProperties = await this.loadCustomProperties();

    for (const key of this.getStorageKeys()) {
      customProperties.remove(key);
    }

    await this.saveCustomProperties(customProperties);
    await this.clearSessionState();
  }

  public async getRenderState(): Promise<RenderState | null> {
    return this.resolveStoredRenderState(await this.loadStoredRenderState());
  }

  public async setPendingRenderState(originalHtml: string): Promise<void> {
    await this.saveRenderState({
      channelId: this.runtimeChannelConfig.channelId,
      originalHtml,
      phase: "pending",
      storedAt: new Date().toISOString(),
    });
  }

  public async setRenderedRenderState(originalHtml: string): Promise<void> {
    await this.saveRenderState({
      channelId: this.runtimeChannelConfig.channelId,
      originalHtml,
      phase: "rendered",
      storedAt: new Date().toISOString(),
    });
  }

  private getPrimaryStorageKey(): string {
    return getChannelScopedKey(this.runtimeChannelConfig, "originalHtml");
  }

  private getStorageKeys(): string[] {
    return [this.getPrimaryStorageKey(), "markout.originalHtml"];
  }

  private async loadCustomProperties(): Promise<CustomPropertiesLike> {
    return new Promise<CustomPropertiesLike>((resolve, reject) => {
      this.mailboxItem.loadCustomPropertiesAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(toOfficeError(result.error));
          return;
        }

        resolve(result.value);
      });
    });
  }

  private async saveCustomProperties(
    customProperties: CustomPropertiesLike
  ): Promise<void> {
    await new Promise<void>((resolve, reject) => {
      customProperties.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(toOfficeError(result.error));
          return;
        }

        resolve();
      });
    });
  }

  private async saveRenderState(renderState: RenderState): Promise<void> {
    await this.saveStoredRenderState(
      JSON.stringify(toStoredPayload(renderState))
    );
  }

  private async saveStoredRenderState(
    serializedRenderState: string
  ): Promise<void> {
    const customProperties = await this.loadCustomProperties();

    try {
      customProperties.set(this.getPrimaryStorageKey(), serializedRenderState);
      customProperties.remove("markout.originalHtml");
      await this.saveCustomProperties(customProperties);
      await this.clearSessionState();
      return;
    } catch (error) {
      if (!isCapacityError(error) || !supportsSessionData(this.mailboxItem)) {
        throw error;
      }
    }

    await this.saveSessionState(serializedRenderState);
    await this.clearCustomPropertyStateQuietly();
  }

  private async loadStoredRenderState(): Promise<string | undefined> {
    const sessionState = await this.loadSessionState();

    if (sessionState !== undefined) {
      return sessionState;
    }

    const customProperties = await this.loadCustomProperties();

    for (const key of this.getStorageKeys()) {
      const storedValue = customProperties.get(key);

      if (storedValue !== undefined) {
        return storedValue;
      }
    }

    return undefined;
  }

  private async clearCustomPropertyStateQuietly(): Promise<void> {
    try {
      const customProperties = await this.loadCustomProperties();

      for (const key of this.getStorageKeys()) {
        customProperties.remove(key);
      }

      await this.saveCustomProperties(customProperties);
    } catch {
      // Ignore cleanup failures after successfully saving a session-scoped state.
    }
  }

  private async clearSessionState(): Promise<void> {
    const sessionData = this.mailboxItem.sessionData;

    if (sessionData === undefined) {
      return;
    }

    for (const key of this.getStorageKeys()) {
      await new Promise<void>((resolve, reject) => {
        sessionData.removeAsync(key, (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            if (isMissingSessionStateError(result.error)) {
              resolve();
              return;
            }

            reject(toOfficeError(result.error));
            return;
          }

          resolve();
        });
      });
    }
  }

  private async loadSessionState(): Promise<string | undefined> {
    const sessionData = this.mailboxItem.sessionData;

    if (sessionData === undefined) {
      return undefined;
    }

    for (const key of this.getStorageKeys()) {
      const storedValue = await new Promise<string | undefined>(
        (resolve, reject) => {
          sessionData.getAsync(key, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
              if (isMissingSessionStateError(result.error)) {
                resolve(undefined);
                return;
              }

              reject(toOfficeError(result.error));
              return;
            }

            resolve(result.value);
          });
        }
      );

      if (storedValue !== undefined) {
        return storedValue;
      }
    }

    return undefined;
  }

  private async saveSessionState(renderState: string): Promise<void> {
    const sessionData = this.mailboxItem.sessionData;

    if (sessionData === undefined) {
      throw createRestoreStateTooLargeError();
    }

    try {
      await new Promise<void>((resolve, reject) => {
        sessionData.setAsync(
          this.getPrimaryStorageKey(),
          renderState,
          (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
              reject(toOfficeError(result.error));
              return;
            }

            resolve();
          }
        );
      });
    } catch (error) {
      if (isCapacityError(error)) {
        throw createRestoreStateTooLargeError(error);
      }

      throw error;
    }
  }

  private resolveStoredRenderState(
    storedValue: string | undefined
  ): RenderState | null {
    if (storedValue === undefined || storedValue === "false") {
      return null;
    }

    try {
      const parsedValue = JSON.parse(storedValue) as LegacyRenderStatePayload;
      return normalizeParsedRenderState(
        parsedValue,
        this.runtimeChannelConfig.channelId
      );
    } catch {
      return {
        channelId: this.runtimeChannelConfig.channelId,
        originalHtml: storedValue,
        phase: "rendered",
        storedAt: new Date(0).toISOString(),
      };
    }
  }
}

function normalizeParsedRenderState(
  parsedValue: LegacyRenderStatePayload,
  expectedChannelId: ChannelId
): RenderState | null {
  if (typeof parsedValue.html === "string") {
    return {
      channelId: expectedChannelId,
      originalHtml: parsedValue.html,
      phase: "rendered",
      storedAt: new Date(0).toISOString(),
    };
  }

  if (
    typeof parsedValue.originalHtml !== "string" ||
    (parsedValue.phase !== "pending" && parsedValue.phase !== "rendered")
  ) {
    return null;
  }

  const channelId =
    parsedValue.channelId === undefined
      ? expectedChannelId
      : normalizeChannelId(parsedValue.channelId);

  if (channelId === null || channelId !== expectedChannelId) {
    return null;
  }

  return {
    channelId,
    originalHtml: parsedValue.originalHtml,
    phase: parsedValue.phase,
    storedAt:
      typeof parsedValue.storedAt === "string"
        ? parsedValue.storedAt
        : new Date(0).toISOString(),
  };
}

function normalizeChannelId(value: unknown): ChannelId | null {
  return value === "beta" || value === "local" || value === "production"
    ? value
    : null;
}

function toStoredPayload(renderState: RenderState): StoredRenderStatePayload {
  return {
    channelId: renderState.channelId,
    originalHtml: renderState.originalHtml,
    phase: renderState.phase,
    schemaVersion: 2,
    storedAt: renderState.storedAt,
  };
}

function getCurrentMailboxItem(): MailboxItemWithCustomPropertiesLike {
  const mailboxItem = Office.context.mailbox.item;

  if (mailboxItem === undefined) {
    throw new MarkOutError(
      "office-compose-item-missing",
      "MarkOut requires an active Outlook compose item."
    );
  }

  return mailboxItem;
}

function createRestoreStateTooLargeError(cause?: unknown): MarkOutError {
  return new MarkOutError(
    "restore-state-too-large",
    "MarkOut couldn't persist the original draft HTML because Outlook's restore-state storage is full for this channel.",
    { cause }
  );
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

function isCapacityError(error: unknown): boolean {
  if (!(error instanceof Error)) {
    return false;
  }

  return (
    /ArgumentOutOfRange/i.test(error.name) ||
    /customproperties/i.test(error.message) ||
    /out of the range of valid values/i.test(error.message) ||
    /quota/i.test(error.message)
  );
}

function isMissingSessionStateError(
  error: { message: string; name: string } | undefined
): boolean {
  return (
    /KeyNotFound/i.test(error?.name ?? "") ||
    /specified key was not found/i.test(error?.message ?? "")
  );
}

function supportsSessionData(
  mailboxItem: MailboxItemWithCustomPropertiesLike
): mailboxItem is MailboxItemWithCustomPropertiesLike & {
  sessionData: SessionDataLike;
} {
  return mailboxItem.sessionData !== undefined;
}

export function createOfficeRenderStateStore(
  mailboxItem: MailboxItemWithCustomPropertiesLike = getCurrentMailboxItem(),
  runtimeChannelConfig: RuntimeChannelConfig = resolveRuntimeChannelConfig()
): RenderStateStore {
  return new OfficeRenderStateStore(mailboxItem, runtimeChannelConfig);
}
