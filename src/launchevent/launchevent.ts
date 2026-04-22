export const SMART_ALERT_ERROR_MESSAGE =
  "MarkOut could not render this draft before send. Open the MarkOut task pane, review the content, then try again.";

async function handleSendEvent(
  event: Office.AddinCommands.Event
): Promise<void> {
  try {
    const { createOfficeSettingsStore } = await import(
      /* webpackChunkName: "launchevent-settings" */ "../lib/config"
    );
    const settingsStore = createOfficeSettingsStore();

    if (!settingsStore.getAutoRender()) {
      event.completed({ allowEvent: true });
      return;
    }

    const { ensureRendered } = await import(
      /* webpackChunkName: "launchevent-render" */ "../lib/item"
    );
    await ensureRendered();
    event.completed({ allowEvent: true });
  } catch (error) {
    console.error("MarkOut auto-render failed before send.", error);
    event.completed({
      allowEvent: false,
      errorMessage: SMART_ALERT_ERROR_MESSAGE,
    } as Office.AddinCommands.EventCompletedOptions);
  }
}

export async function onMessageSendHandler(
  event: Office.AddinCommands.Event
): Promise<void> {
  await handleSendEvent(event);
}

export async function onAppointmentSendHandler(
  event: Office.AddinCommands.Event
): Promise<void> {
  await handleSendEvent(event);
}

export function registerLaunchEventHandlers(): void {
  Office.actions.associate(
    "onAppointmentSendHandler",
    onAppointmentSendHandler
  );
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}

if (typeof Office !== "undefined") {
  void Office.onReady(() => {
    registerLaunchEventHandlers();
  });
}
