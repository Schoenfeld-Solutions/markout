const ORIGINAL_HTML_KEY = "markout.originalHtml";

interface CustomPropertiesLike {
  get(name: string): string | undefined;
  remove(name: string): void;
  saveAsync(callback: (result: Office.AsyncResult<void>) => void): void;
  set(name: string, value: string): void;
}

interface MailboxItemWithCustomPropertiesLike {
  loadCustomPropertiesAsync(
    callback: (result: Office.AsyncResult<CustomPropertiesLike>) => void
  ): void;
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
  }

  public async getRenderState(): Promise<RenderState | null> {
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
    const customProperties = await this.loadCustomProperties();
    customProperties.set(ORIGINAL_HTML_KEY, JSON.stringify(renderState));
    await this.saveCustomProperties(customProperties);
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

export function createOfficeRenderStateStore(
  mailboxItem: MailboxItemWithCustomPropertiesLike = getCurrentMailboxItem()
): RenderStateStore {
  return new OfficeRenderStateStore(mailboxItem);
}
