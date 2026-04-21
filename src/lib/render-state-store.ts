const ORIGINAL_HTML_KEY = "markout.originalHtml";
const PERSISTENT_RENDER_SOURCE_KEY_PREFIX = "markout.renderSource.";

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

interface PersistentStorageLike {
  getItem(name: string): string | null;
  removeItem(name: string): void;
  setItem(name: string, value: string): void;
}

export type RenderStatePhase = "pending" | "rendered";

export interface RenderState {
  originalHtml: string;
  phase: RenderStatePhase;
}

export interface RenderStateStore {
  clearRenderState(): Promise<void>;
  getRenderState(): Promise<RenderState | null>;
  setPendingRenderState(originalHtml: string): Promise<void>;
  setRenderedRenderState(originalHtml: string): Promise<void>;
}

interface LegacyRenderStatePayload {
  html?: unknown;
  originalHtml?: unknown;
  phase?: unknown;
  originalHtmlStorage?: unknown;
  originalHtmlStorageKey?: unknown;
}

class OfficeRenderStateStore implements RenderStateStore {
  public constructor(
    private readonly mailboxItem: MailboxItemWithCustomPropertiesLike,
    private readonly persistentStorage: PersistentStorageLike | undefined
  ) {}

  public async clearRenderState(): Promise<void> {
    const persistentStorageKey = await this.getPersistentStorageKey();
    const customProperties = await this.loadCustomProperties();
    customProperties.remove(ORIGINAL_HTML_KEY);
    await this.saveCustomProperties(customProperties);
    await this.clearSessionState();
    this.clearPersistentRenderSource(persistentStorageKey);
  }

  public async getRenderState(): Promise<RenderState | null> {
    return this.resolveStoredRenderState(await this.loadStoredRenderState());
  }

  public async setPendingRenderState(originalHtml: string): Promise<void> {
    await this.saveRenderState({
      originalHtml,
      phase: "pending",
    });
  }

  public async setRenderedRenderState(originalHtml: string): Promise<void> {
    await this.saveRenderState({
      originalHtml,
      phase: "rendered",
    });
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
    const currentPersistentStorageKey = await this.getPersistentStorageKey();

    if (this.persistentStorage !== undefined) {
      const persistentStorageKey =
        currentPersistentStorageKey ?? createPersistentStorageKey();

      try {
        this.savePersistentRenderSource(
          persistentStorageKey,
          renderState.originalHtml
        );
      } catch {
        await this.saveStoredRenderState(JSON.stringify(renderState));
        return;
      }

      try {
        await this.saveStoredRenderState(
          JSON.stringify({
            originalHtmlStorage: "local",
            originalHtmlStorageKey: persistentStorageKey,
            phase: renderState.phase,
          } satisfies LegacyRenderStatePayload)
        );
      } catch (error) {
        if (currentPersistentStorageKey !== persistentStorageKey) {
          this.clearPersistentRenderSource(persistentStorageKey);
        }

        throw error;
      }

      if (
        currentPersistentStorageKey !== undefined &&
        currentPersistentStorageKey !== persistentStorageKey
      ) {
        this.clearPersistentRenderSource(currentPersistentStorageKey);
      }

      return;
    }

    await this.saveStoredRenderState(JSON.stringify(renderState));
  }

  private async saveStoredRenderState(
    serializedRenderState: string
  ): Promise<void> {
    try {
      const customProperties = await this.loadCustomProperties();
      customProperties.set(ORIGINAL_HTML_KEY, serializedRenderState);
      await this.saveCustomProperties(customProperties);
      await this.clearSessionState();
    } catch (error) {
      if (!supportsSessionData(this.mailboxItem) || !isCapacityError(error)) {
        throw error;
      }

      await this.saveSessionState(serializedRenderState);
      await this.clearCustomPropertyStateQuietly();
    }
  }

  private async loadStoredRenderState(): Promise<string | undefined> {
    const sessionState = await this.loadSessionState();

    if (sessionState !== undefined) {
      return sessionState;
    }

    const customProperties = await this.loadCustomProperties();
    return customProperties.get(ORIGINAL_HTML_KEY);
  }

  private async clearCustomPropertyStateQuietly(): Promise<void> {
    try {
      const customProperties = await this.loadCustomProperties();
      customProperties.remove(ORIGINAL_HTML_KEY);
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

    await new Promise<void>((resolve, reject) => {
      sessionData.removeAsync(ORIGINAL_HTML_KEY, (result) => {
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

  private async loadSessionState(): Promise<string | undefined> {
    const sessionData = this.mailboxItem.sessionData;

    if (sessionData === undefined) {
      return undefined;
    }

    return new Promise<string | undefined>((resolve, reject) => {
      sessionData.getAsync(ORIGINAL_HTML_KEY, (result) => {
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
    });
  }

  private async saveSessionState(renderState: string): Promise<void> {
    const sessionData = this.mailboxItem.sessionData;

    if (sessionData === undefined) {
      throw new Error("MarkOut couldn't persist large render state data.");
    }

    await new Promise<void>((resolve, reject) => {
      sessionData.setAsync(ORIGINAL_HTML_KEY, renderState, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(toOfficeError(result.error));
          return;
        }

        resolve();
      });
    });
  }

  private async getPersistentStorageKey(): Promise<string | undefined> {
    const storedRenderState = await this.loadStoredRenderState();

    if (storedRenderState === undefined || storedRenderState === "false") {
      return undefined;
    }

    try {
      const parsedValue = JSON.parse(
        storedRenderState
      ) as LegacyRenderStatePayload;
      return getPersistentStorageKeyFromPayload(parsedValue);
    } catch {
      return undefined;
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
      return this.normalizeParsedRenderState(parsedValue);
    } catch {
      return {
        originalHtml: storedValue,
        phase: "rendered",
      };
    }
  }

  private normalizeParsedRenderState(
    parsedValue: LegacyRenderStatePayload
  ): RenderState | null {
    if (typeof parsedValue.html === "string") {
      return {
        originalHtml: parsedValue.html,
        phase: "rendered",
      };
    }

    if (
      typeof parsedValue.originalHtml === "string" &&
      (parsedValue.phase === "pending" || parsedValue.phase === "rendered")
    ) {
      return {
        originalHtml: parsedValue.originalHtml,
        phase: parsedValue.phase,
      };
    }

    const persistentStorageKey =
      getPersistentStorageKeyFromPayload(parsedValue);

    if (
      persistentStorageKey !== undefined &&
      (parsedValue.phase === "pending" || parsedValue.phase === "rendered")
    ) {
      const persistentRenderSource =
        this.persistentStorage?.getItem(persistentStorageKey);

      if (
        persistentRenderSource === null ||
        persistentRenderSource === undefined
      ) {
        throw new Error(
          "MarkOut couldn't recover the stored draft source for this compose session."
        );
      }

      return {
        originalHtml: persistentRenderSource,
        phase: parsedValue.phase,
      };
    }

    return null;
  }

  private clearPersistentRenderSource(
    persistentStorageKey: string | undefined
  ): void {
    if (
      this.persistentStorage === undefined ||
      persistentStorageKey === undefined
    ) {
      return;
    }

    this.persistentStorage.removeItem(persistentStorageKey);
  }

  private savePersistentRenderSource(
    persistentStorageKey: string,
    originalHtml: string
  ): void {
    if (this.persistentStorage === undefined) {
      return;
    }

    this.persistentStorage.setItem(persistentStorageKey, originalHtml);
  }
}

function getCurrentMailboxItem(): MailboxItemWithCustomPropertiesLike {
  const mailboxItem = Office.context.mailbox.item;

  if (mailboxItem === undefined) {
    throw new Error("MarkOut requires an active Outlook compose item.");
  }

  return mailboxItem;
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
    /out of the range of valid values/i.test(error.message)
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

function getPersistentRenderStateStorage(): PersistentStorageLike | undefined {
  try {
    const persistentStorage = globalThis.localStorage as
      | Partial<PersistentStorageLike>
      | undefined;

    if (
      persistentStorage === undefined ||
      typeof persistentStorage.getItem !== "function" ||
      typeof persistentStorage.removeItem !== "function" ||
      typeof persistentStorage.setItem !== "function"
    ) {
      return undefined;
    }

    return persistentStorage as PersistentStorageLike;
  } catch {
    return undefined;
  }
}

function createPersistentStorageKey(): string {
  if (typeof globalThis.crypto.randomUUID === "function") {
    return `${PERSISTENT_RENDER_SOURCE_KEY_PREFIX}${globalThis.crypto.randomUUID()}`;
  }

  return `${PERSISTENT_RENDER_SOURCE_KEY_PREFIX}${Date.now().toString(36)}.${Math.random()
    .toString(36)
    .slice(2, 10)}`;
}

function getPersistentStorageKeyFromPayload(
  parsedValue: LegacyRenderStatePayload
): string | undefined {
  if (
    parsedValue.originalHtmlStorage === "local" &&
    typeof parsedValue.originalHtmlStorageKey === "string"
  ) {
    return parsedValue.originalHtmlStorageKey;
  }

  return undefined;
}

export function createOfficeRenderStateStore(
  mailboxItem: MailboxItemWithCustomPropertiesLike = getCurrentMailboxItem(),
  persistentStorage:
    | PersistentStorageLike
    | undefined = getPersistentRenderStateStorage()
): RenderStateStore {
  return new OfficeRenderStateStore(mailboxItem, persistentStorage);
}
