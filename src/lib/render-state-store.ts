const ORIGINAL_HTML_KEY = "markout.originalHtml";

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
}

class OfficeRenderStateStore implements RenderStateStore {
  public constructor(
    private readonly mailboxItem: MailboxItemWithCustomPropertiesLike
  ) {}

  public async clearRenderState(): Promise<void> {
    const customProperties = await this.loadCustomProperties();
    customProperties.remove(ORIGINAL_HTML_KEY);
    await this.saveCustomProperties(customProperties);
    await this.clearSessionState();
  }

  public async getRenderState(): Promise<RenderState | null> {
    const sessionState = await this.loadSessionState();

    if (sessionState !== undefined) {
      return normalizeStoredRenderState(sessionState);
    }

    const customProperties = await this.loadCustomProperties();
    return normalizeStoredRenderState(customProperties.get(ORIGINAL_HTML_KEY));
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
    const serializedRenderState = JSON.stringify(renderState);

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
}

function getCurrentMailboxItem(): MailboxItemWithCustomPropertiesLike {
  const mailboxItem = Office.context.mailbox.item;

  if (mailboxItem === undefined) {
    throw new Error("MarkOut requires an active Outlook compose item.");
  }

  return mailboxItem;
}

function normalizeStoredRenderState(
  storedValue: string | undefined
): RenderState | null {
  if (storedValue === undefined || storedValue === "false") {
    return null;
  }

  try {
    const parsedValue = JSON.parse(storedValue) as LegacyRenderStatePayload;
    return normalizeParsedRenderState(parsedValue);
  } catch {
    return {
      originalHtml: storedValue,
      phase: "rendered",
    };
  }
}

function normalizeParsedRenderState(
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

  return null;
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

function supportsSessionData(
  mailboxItem: MailboxItemWithCustomPropertiesLike
): mailboxItem is MailboxItemWithCustomPropertiesLike & {
  sessionData: SessionDataLike;
} {
  return mailboxItem.sessionData !== undefined;
}

export function createOfficeRenderStateStore(
  mailboxItem: MailboxItemWithCustomPropertiesLike = getCurrentMailboxItem()
): RenderStateStore {
  return new OfficeRenderStateStore(mailboxItem);
}
