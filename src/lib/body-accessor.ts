export interface BodyAccessor {
  getHtml(): Promise<string>;
  setHtml(html: string): Promise<void>;
}

interface ItemBodyLike {
  getAsync(
    coercionType: Office.CoercionType,
    callback: (result: Office.AsyncResult<string>) => void
  ): void;
  setAsync(
    value: string,
    options: { coercionType: Office.CoercionType },
    callback: (result: Office.AsyncResult<void>) => void
  ): void;
}

interface MailboxItemLike {
  body: ItemBodyLike;
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
}

function getCurrentMailboxItem(): MailboxItemLike {
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

export function createOfficeBodyAccessor(
  mailboxItem: MailboxItemLike = getCurrentMailboxItem()
): BodyAccessor {
  return new OfficeBodyAccessor(mailboxItem);
}
