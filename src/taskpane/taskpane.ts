// @ts-expect-error Webpack resolves taskpane CSS side-effect imports at build time.
import "./taskpane.css";
import { createOfficeSettingsStore } from "../lib/config";
import { Debounce } from "../lib/debounce";
import type { RenderItemResult } from "../lib/item";
import type { RenderOptions } from "../lib/renderer";

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
  sideloadMessage: HTMLElement;
  sideloadTitle: HTMLElement;
  statusMessage: HTMLElement;
  themeEditor: HTMLTextAreaElement;
}

interface PreviewDependencies {
  renderMarkdown: (options: RenderOptions) => Promise<string>;
  sanitize: (html: string) => string;
}

interface RenderItemModule {
  renderItem: () => Promise<RenderItemResult>;
}

let itemModulePromise: Promise<RenderItemModule> | null = null;
let previewDependenciesPromise: Promise<PreviewDependencies> | null = null;

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
    sideloadMessage: getRequiredElement("sideload-msg"),
    sideloadTitle: getRequiredElement("sideload-title"),
    statusMessage: getRequiredElement("status-message"),
    themeEditor: getRequiredElement("theme-editor"),
  };
}

async function loadItemModule(): Promise<RenderItemModule> {
  itemModulePromise ??= import("../lib/item");
  return itemModulePromise;
}

async function loadPreviewDependencies(): Promise<PreviewDependencies> {
  previewDependenciesPromise ??= Promise.all([
    import("../lib/html-sanitizer"),
    import("../lib/renderer"),
  ]).then(([htmlSanitizerModule, rendererModule]) => {
    const htmlSanitizer = new htmlSanitizerModule.DefaultHtmlSanitizer();
    return {
      renderMarkdown: rendererModule.renderMarkdown,
      sanitize: (html: string) => htmlSanitizer.sanitize(html),
    };
  });

  return previewDependenciesPromise;
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

function showFallback(
  elements: TaskpaneElements,
  title: string,
  message: string
): void {
  elements.sideloadTitle.textContent = title;
  elements.sideloadCopy.textContent = message;
  elements.sideloadMessage.hidden = false;
  elements.appBody.hidden = true;
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
  const previewDependencies = await loadPreviewDependencies();
  const previewHtml = await previewDependencies.renderMarkdown({
    css: stylesheet,
    markdown: PREVIEW_MARKDOWN,
  });

  elements.preview.innerHTML = previewDependencies.sanitize(previewHtml);
}

async function initializeTaskpane(): Promise<void> {
  const elements = getElements();
  const settingsStore = createOfficeSettingsStore();

  elements.sideloadMessage.hidden = true;
  elements.appBody.hidden = false;
  elements.themeEditor.value = settingsStore.getStylesheet();
  updateAutoRenderButton(elements, settingsStore.getAutoRender());
  await updatePreview(elements, settingsStore.getStylesheet());

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
    await updatePreview(elements, settingsStore.getStylesheet());
    saveTheme.trigger();
  });

  elements.refreshPreview.addEventListener("click", async () => {
    setStatus(elements, "info", "Refreshing preview...");
    await updatePreview(elements, settingsStore.getStylesheet());
    setStatus(elements, "success", "Preview refreshed.");
  });

  elements.renderButton.addEventListener("click", async () => {
    elements.renderButton.disabled = true;
    setStatus(
      elements,
      "info",
      "Applying Markdown rendering to the current draft..."
    );

    try {
      const itemModule = await loadItemModule();
      const result = await itemModule.renderItem();
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

void Office.onReady((info) => {
  const elements = getElements();

  if (info.host !== Office.HostType.Outlook) {
    showFallback(
      elements,
      "Open MarkOut from Outlook",
      "This task pane is only available when MarkOut is launched inside Outlook compose."
    );
    return;
  }

  void initializeTaskpane().catch((error: unknown) => {
    console.error("MarkOut failed to initialize the task pane.", error);
    showFallback(
      elements,
      "MarkOut could not initialize",
      getInitializationFailureMessage()
    );
  });
});
