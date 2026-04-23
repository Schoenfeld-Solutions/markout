import { MarkOutError } from "./runtime";

export type SelectionSource = "body" | "subject";

export interface ComposeSelection {
  hasSelection: boolean;
  html: string | null;
  source: SelectionSource;
  text: string;
}

export interface BodyAccessor {
  getHtml(): Promise<string>;
  getSelection(): Promise<ComposeSelection>;
  setHtml(html: string): Promise<void>;
  replaceSelectionWithHtml(html: string): Promise<void>;
}

interface ItemBodyLike {
  getAsync(
    coercionType: Office.CoercionType,
    callback: (result: Office.AsyncResult<string>) => void
  ): void;
  getTypeAsync(
    callback: (result: Office.AsyncResult<Office.CoercionType>) => void
  ): void;
  setAsync(
    value: string,
    options: { coercionType: Office.CoercionType },
    callback: (result: Office.AsyncResult<void>) => void
  ): void;
  setSelectedDataAsync(
    value: string,
    options: { coercionType: Office.CoercionType },
    callback: (result: Office.AsyncResult<void>) => void
  ): void;
}

interface SelectedDataValue {
  data?: string;
  sourceProperty?: string;
}

interface MailboxItemLike {
  body: ItemBodyLike;
  getSelectedDataAsync(
    coercionType: Office.CoercionType,
    callback: (result: Office.AsyncResult<SelectedDataValue>) => void
  ): void;
}

class OfficeBodyAccessor implements BodyAccessor {
  public constructor(private readonly mailboxItem: MailboxItemLike) {}

  public async getHtml(): Promise<string> {
    return new Promise<string>((resolve, reject) => {
      this.mailboxItem.body.getAsync(Office.CoercionType.Html, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(toOfficeError(result.error));
          return;
        }

        resolve(result.value);
      });
    });
  }

  public async getSelection(): Promise<ComposeSelection> {
    const textSelection = await this.readSelectedData(Office.CoercionType.Text);
    const source = normalizeSelectionSource(textSelection.sourceProperty);

    if (source === "body") {
      const htmlSelection = await this.tryReadSelectedHtml();

      return {
        hasSelection: textSelection.data.trim().length > 0,
        html: htmlSelection,
        source,
        text: textSelection.data,
      };
    }

    return {
      hasSelection: textSelection.data.trim().length > 0,
      html: null,
      source,
      text: textSelection.data,
    };
  }

  public async setHtml(html: string): Promise<void> {
    await new Promise<void>((resolve, reject) => {
      this.mailboxItem.body.setAsync(
        html,
        { coercionType: Office.CoercionType.Html },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            reject(toOfficeError(result.error));
            return;
          }

          resolve();
        }
      );
    });
  }

  public async replaceSelectionWithHtml(html: string): Promise<void> {
    const bodyType = await this.getBodyType();

    if (bodyType !== Office.CoercionType.Html) {
      const error = new MarkOutError(
        "unsupported-body-type",
        "MarkOut can only insert rendered content into an HTML compose body."
      );
      throw error;
    }

    await new Promise<void>((resolve, reject) => {
      this.mailboxItem.body.setSelectedDataAsync(
        html,
        { coercionType: Office.CoercionType.Html },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            reject(toOfficeError(result.error));
            return;
          }

          resolve();
        }
      );
    });
  }

  private async getBodyType(): Promise<Office.CoercionType> {
    return new Promise<Office.CoercionType>((resolve, reject) => {
      this.mailboxItem.body.getTypeAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(toOfficeError(result.error));
          return;
        }

        resolve(result.value);
      });
    });
  }

  private async readSelectedData(
    coercionType: Office.CoercionType
  ): Promise<{ data: string; sourceProperty: string | undefined }> {
    return new Promise((resolve, reject) => {
      this.mailboxItem.getSelectedDataAsync(coercionType, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(toOfficeError(result.error));
          return;
        }

        const value = result.value;

        resolve({
          data: value.data ?? "",
          sourceProperty: value.sourceProperty,
        });
      });
    });
  }

  private async tryReadSelectedHtml(): Promise<string | null> {
    try {
      const htmlSelection = await this.readSelectedData(
        Office.CoercionType.Html
      );
      return htmlSelection.data;
    } catch (error) {
      if (
        error instanceof Error &&
        ["InvalidFormatError", "InvalidSelection"].includes(error.name)
      ) {
        return null;
      }

      throw error;
    }
  }
}

function normalizeSelectionSource(
  sourceProperty: string | undefined
): SelectionSource {
  return sourceProperty === "subject" ? "subject" : "body";
}

function getCurrentMailboxItem(): MailboxItemLike {
  const mailboxItem = Office.context.mailbox.item;

  if (mailboxItem === undefined) {
    throw new MarkOutError(
      "office-compose-item-missing",
      "MarkOut requires an active Outlook compose item."
    );
  }

  if (
    typeof mailboxItem.getSelectedDataAsync !== "function" ||
    typeof mailboxItem.body.setSelectedDataAsync !== "function" ||
    typeof mailboxItem.body.getTypeAsync !== "function"
  ) {
    throw new MarkOutError(
      "office-selection-api-unavailable",
      "MarkOut requires an Outlook compose item with body selection APIs."
    );
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

export function createOfficeBodyAccessor(
  mailboxItem: MailboxItemLike = getCurrentMailboxItem()
): BodyAccessor {
  return new OfficeBodyAccessor(mailboxItem);
}
