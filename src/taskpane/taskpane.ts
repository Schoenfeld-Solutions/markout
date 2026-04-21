// @ts-expect-error Webpack resolves taskpane CSS side-effect imports at build time.
import "./taskpane.css";
import { createOfficeSettingsStore } from "../lib/config";
import { Debounce } from "../lib/debounce";
import { DefaultHtmlSanitizer } from "../lib/html-sanitizer";
import { renderItem } from "../lib/item";
import { renderMarkdown } from "../lib/renderer";

const PREVIEW_MARKDOWN = `
# MarkOut Preview

Compose your draft in Markdown, then render it into Outlook-ready email HTML.

## What this pane controls

- Preview the active email theme
- Edit the inline stylesheet
- Enable or disable auto-render on send

\`\`\`ts
console.log("MarkOut keeps the authoring flow Markdown-first.");
\`\`\`
`.trim();

type StatusTone = "error" | "idle" | "info" | "success";

interface TaskpaneElements {
  appBody: HTMLElement;
  autoRenderButton: HTMLButtonElement;
  preview: HTMLElement;
  refreshPreview: HTMLButtonElement;
  renderButton: HTMLButtonElement;
  sideloadCopy: HTMLElement;
  sideloadDetails: HTMLElement;
  sideloadMessage: HTMLElement;
  sideloadStage: HTMLElement;
  sideloadTitle: HTMLElement;
  statusMessage: HTMLElement;
  themeEditor: HTMLTextAreaElement;
}

const htmlSanitizer = new DefaultHtmlSanitizer();

function setElementVisible(element: HTMLElement, visible: boolean): void {
  element.hidden = !visible;
  element.setAttribute("aria-hidden", String(!visible));

  if (visible) {
    element.style.removeProperty("display");
    return;
  }

  element.style.display = "none";
}

function getRequiredElement<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);

  if (!(element instanceof HTMLElement)) {
    throw new Error(`Required element "${id}" was not found.`);
  }

  return element as T;
}

function getElements(): TaskpaneElements {
  return {
    appBody: getRequiredElement("app-body"),
    autoRenderButton: getRequiredElement("autorender-button"),
    preview: getRequiredElement("mo-preview"),
    refreshPreview: getRequiredElement("refresh-preview"),
    renderButton: getRequiredElement("render-button"),
    sideloadCopy: getRequiredElement("sideload-copy"),
    sideloadDetails: getRequiredElement("sideload-details"),
    sideloadMessage: getRequiredElement("sideload-msg"),
    sideloadStage: getRequiredElement("sideload-stage"),
    sideloadTitle: getRequiredElement("sideload-title"),
    statusMessage: getRequiredElement("status-message"),
    themeEditor: getRequiredElement("theme-editor"),
  };
}

function setStatus(
  elements: TaskpaneElements,
  tone: StatusTone,
  message: string
): void {
  elements.statusMessage.dataset.tone = tone;
  elements.statusMessage.textContent = message;
}

function updateAutoRenderButton(
  elements: TaskpaneElements,
  enabled: boolean
): void {
  elements.autoRenderButton.textContent = `Auto-render on send: ${enabled ? "On" : "Off"}`;
}

function formatError(error: unknown): string {
  if (error instanceof Error) {
    return `${error.name}: ${error.message}`;
  }

  return String(error);
}

function setBootState(
  elements: TaskpaneElements,
  message: string,
  tone: Exclude<StatusTone, "idle"> = "info"
): void {
  setElementVisible(elements.sideloadStage, true);
  elements.sideloadStage.dataset.tone = tone;
  elements.sideloadStage.textContent = message;
}

function clearBootState(elements: TaskpaneElements): void {
  elements.sideloadStage.textContent = "";
  setElementVisible(elements.sideloadStage, false);
  elements.sideloadDetails.textContent = "";
  setElementVisible(elements.sideloadDetails, false);
}

function showFallback(
  elements: TaskpaneElements,
  title: string,
  message: string,
  error?: unknown
): void {
  elements.sideloadTitle.textContent = title;
  elements.sideloadCopy.textContent = message;
  elements.sideloadDetails.textContent =
    error === undefined ? "" : formatError(error);
  setElementVisible(elements.sideloadDetails, error !== undefined);
  setBootState(
    elements,
    error === undefined ? message : "Initialization failed.",
    error === undefined ? "info" : "error"
  );
  setElementVisible(elements.sideloadMessage, true);
  setElementVisible(elements.appBody, false);
}

function showApplication(elements: TaskpaneElements): void {
  clearBootState(elements);
  setElementVisible(elements.sideloadMessage, false);
  setElementVisible(elements.appBody, true);
}

function reportBootFailure(error: unknown): void {
  try {
    const elements = getElements();
    showFallback(
      elements,
      "MarkOut could not initialize",
      getInitializationFailureMessage(),
      error
    );
  } catch (fallbackError) {
    console.error(
      "MarkOut failed to render the fallback state.",
      fallbackError
    );
  }
}

function getInitializationFailureMessage(): string {
  if (window.location.hostname === "localhost") {
    return "Use Outlook's Add from File flow with manifest-localhost.xml to load the local development build.";
  }

  return "Reload Outlook on the web and reopen MarkOut from Apps. If the problem persists, remove and reinstall the beta manifest.";
}

async function updatePreview(
  elements: TaskpaneElements,
  stylesheet: string
): Promise<void> {
  const previewHtml = await renderMarkdown({
    css: stylesheet,
    markdown: PREVIEW_MARKDOWN,
  });

  elements.preview.innerHTML = htmlSanitizer.sanitize(previewHtml);
}

async function initializeTaskpane(): Promise<void> {
  const elements = getElements();
  const settingsStore = createOfficeSettingsStore();

  setBootState(elements, "Task pane script loaded. Initializing settings...");
  setElementVisible(elements.sideloadMessage, true);
  setElementVisible(elements.appBody, false);
  elements.themeEditor.value = settingsStore.getStylesheet();
  updateAutoRenderButton(elements, settingsStore.getAutoRender());

  try {
    setBootState(elements, "Rendering initial preview...");
    await updatePreview(elements, settingsStore.getStylesheet());
  } catch (error) {
    console.error("MarkOut failed to render the initial preview.", error);
    setBootState(elements, "Initial preview failed.", "error");
    setStatus(
      elements,
      "error",
      "Preview could not be rendered, but the task pane is ready."
    );
  }

  showApplication(elements);

  const saveTheme = new Debounce(async () => {
    try {
      await settingsStore.save();
      setStatus(elements, "success", "Theme saved.");
    } catch (error) {
      console.error("MarkOut failed to save theme changes.", error);
      setStatus(elements, "error", "Theme changes could not be saved.");
    }
  }, 900);

  elements.themeEditor.addEventListener("input", async () => {
    settingsStore.setStylesheet(elements.themeEditor.value);
    setStatus(elements, "info", "Saving theme changes...");

    try {
      await updatePreview(elements, settingsStore.getStylesheet());
      saveTheme.trigger();
    } catch (error) {
      console.error("MarkOut failed to refresh the preview.", error);
      setStatus(elements, "error", "Preview could not be refreshed.");
    }
  });

  elements.refreshPreview.addEventListener("click", async () => {
    setStatus(elements, "info", "Refreshing preview...");

    try {
      await updatePreview(elements, settingsStore.getStylesheet());
      setStatus(elements, "success", "Preview refreshed.");
    } catch (error) {
      console.error("MarkOut failed to refresh the preview.", error);
      setStatus(elements, "error", "Preview could not be refreshed.");
    }
  });

  elements.renderButton.addEventListener("click", async () => {
    elements.renderButton.disabled = true;
    setStatus(
      elements,
      "info",
      "Applying Markdown rendering to the current draft..."
    );

    try {
      const result = await renderItem();
      setStatus(
        elements,
        "success",
        result === "rendered"
          ? "Draft rendered successfully."
          : "Original draft HTML restored successfully."
      );
    } catch (error) {
      console.error("MarkOut failed to render the current draft.", error);
      setStatus(
        elements,
        "error",
        "MarkOut could not update the current draft."
      );
    } finally {
      elements.renderButton.disabled = false;
    }
  });

  elements.autoRenderButton.addEventListener("click", async () => {
    const nextValue = !settingsStore.getAutoRender();
    settingsStore.setAutoRender(nextValue);

    try {
      await settingsStore.save();
      updateAutoRenderButton(elements, nextValue);
      setStatus(
        elements,
        "success",
        `Auto-render on send ${nextValue ? "enabled" : "disabled"}.`
      );
    } catch (error) {
      console.error("MarkOut failed to update the auto-render setting.", error);
      settingsStore.setAutoRender(!nextValue);
      setStatus(elements, "error", "Auto-render setting could not be updated.");
    }
  });
}

window.addEventListener("error", (event) => {
  console.error(
    "MarkOut captured a window error.",
    event.error ?? event.message
  );
  reportBootFailure(event.error ?? event.message);
});

window.addEventListener("unhandledrejection", (event) => {
  console.error("MarkOut captured an unhandled rejection.", event.reason);
  reportBootFailure(event.reason);
});

async function bootTaskpane(): Promise<void> {
  const elements = getElements();
  console.info("[MarkOut] taskpane script loaded");
  setElementVisible(elements.sideloadMessage, true);
  setElementVisible(elements.appBody, false);
  setBootState(
    elements,
    "Task pane script loaded. Waiting for Office.onReady..."
  );

  if (typeof Office === "undefined") {
    reportBootFailure(
      new Error("Office.js did not load before taskpane startup.")
    );
    return;
  }

  const info = await Office.onReady();
  console.info("[MarkOut] Office.onReady", { host: info.host });
  setBootState(elements, `Office.onReady resolved for ${String(info.host)}.`);

  if (info.host !== Office.HostType.Outlook) {
    showFallback(
      elements,
      "Open MarkOut from Outlook",
      "This task pane is only available when MarkOut is launched inside Outlook compose."
    );
    return;
  }

  await initializeTaskpane();
}

void bootTaskpane().catch((error: unknown) => {
  console.error("MarkOut failed to initialize the task pane.", error);
  reportBootFailure(error);
});
