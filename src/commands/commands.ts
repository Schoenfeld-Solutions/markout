import { renderItem } from "../lib/item";

export const COMMAND_ERROR_MESSAGE =
  "MarkOut could not render this draft. Open the task pane to inspect the content and try again.";

export async function renderCurrentItem(
  event: Office.AddinCommands.Event
): Promise<void> {
  try {
    await renderItem();
  } catch (error) {
    console.error("MarkOut failed to render the current draft.", error);
    const currentItem = Office.context.mailbox.item;

    if (currentItem === undefined) {
      return;
    }

    const notification: Office.NotificationMessageDetails = {
      icon: "Icon.80x80",
      message: COMMAND_ERROR_MESSAGE,
      persistent: true,
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
    };

    try {
      currentItem.notificationMessages.replaceAsync(
        "markout.render",
        notification
      );
    } catch (notificationError) {
      console.error(
        "MarkOut could not show the render failure notification.",
        notificationError
      );
    }
  } finally {
    event.completed();
  }
}

export function registerCommandHandlers(): void {
  Office.actions.associate("renderCurrentItem", renderCurrentItem);
}

if (typeof Office !== "undefined") {
  void Office.onReady(() => {
    registerCommandHandlers();
  });
}
