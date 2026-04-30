/** @jest-environment jsdom */

import { act } from "react";
import { defaultStylesheet } from "../src/lib/config";
import { createInMemoryDiagnosticSink } from "../src/lib/runtime";
import {
  createMutableSettingsStore,
  createNotificationService,
  createTaskpaneServices,
  flushTaskpane,
  mountTaskpaneApp,
} from "./taskpane-app-harness";

(
  globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true;

describe("taskpane app flows", () => {
  beforeEach(() => {
    jest.spyOn(console, "error").mockImplementation((message) => {
      const text = String(message);
      if (
        text.includes("not wrapped in act") ||
        text.includes("change in the order of Hooks") ||
        text.startsWith("MarkOut failed")
      ) {
        return;
      }

      throw new Error(text);
    });
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  it("confirms the intro and starts directly on insert after dismissal", async () => {
    const settingsStore = createMutableSettingsStore();
    const mounted = await mountTaskpaneApp({ settingsStore });

    try {
      expect(mounted.container.textContent).toContain("Intro");

      await mounted.click("#intro-confirm-button");
      await waitForCondition(
        () => mounted.container.querySelector("#markdown-input") !== null
      );

      expect(settingsStore.setIntroDismissed).toHaveBeenCalledWith(true);
      expect(settingsStore.save).toHaveBeenCalledTimes(1);
      expect(mounted.container.querySelector("#markdown-input")).not.toBeNull();
    } finally {
      mounted.cleanup();
    }

    const dismissedSettingsStore = createMutableSettingsStore({
      introDismissed: true,
    });
    const dismissed = await mountTaskpaneApp({
      settingsStore: dismissedSettingsStore,
    });

    try {
      expect(dismissed.container.querySelector("#intro-confirm-button")).toBe(
        null
      );
      expect(
        dismissed.container.querySelector("#markdown-input")
      ).not.toBeNull();
    } finally {
      dismissed.cleanup();
    }
  });

  it("persists settings changes and rolls back when roaming settings fail", async () => {
    const settingsStore = createMutableSettingsStore({ introDismissed: true });
    const mounted = await mountTaskpaneApp({ settingsStore });

    try {
      await mounted.click("#panel-button-settings");
      await mounted.click("#autorender-switch");

      expect(settingsStore.setAutoRender).toHaveBeenCalledWith(true);
      expect(settingsStore.state.autoRender).toBe(true);

      settingsStore.save.mockRejectedValueOnce(new Error("roaming down"));

      await mounted.click("#show-help-switch");
      await waitForCondition(() =>
        mounted.container.textContent.includes("Settings could not be updated.")
      );

      expect(settingsStore.setHelpVisible).toHaveBeenCalledWith(false);
      expect(settingsStore.state.helpVisible).toBe(true);
      expect(mounted.container.textContent).toContain(
        "Settings could not be updated."
      );
    } finally {
      mounted.cleanup();
    }
  });

  it("updates language preference from settings controls", async () => {
    const settingsStore = createMutableSettingsStore({ introDismissed: true });
    const mounted = await mountTaskpaneApp({ settingsStore });

    try {
      await mounted.click("#panel-button-settings");
      await changeSelectValue(
        mounted.container.querySelector<HTMLSelectElement>(
          "#language-preference-select"
        ),
        "de-DE"
      );

      expect(settingsStore.setLanguagePreference).toHaveBeenCalledWith("de-DE");
      expect(
        mounted.container
          .querySelector("#taskpane-shell")
          ?.getAttribute("data-locale")
      ).toBe("de-DE");
    } finally {
      mounted.cleanup();
    }
  });

  it("executes insert, selection render, and unchanged draft render flows", async () => {
    const diagnosticSink = createInMemoryDiagnosticSink();
    const services = createTaskpaneServices({
      getSelection: jest.fn().mockResolvedValue({
        hasSelection: true,
        html: "<p>selected</p>",
        source: "body",
        text: "selected",
      }),
      insertRenderedMarkdown: jest.fn().mockResolvedValue("replaced"),
      renderEntireDraft: jest.fn().mockResolvedValue("unchanged"),
    });
    const notificationService = createNotificationService();
    const mounted = await mountTaskpaneApp({
      diagnosticSink,
      initialMarkdownInput: "# Heading",
      notificationService,
      services,
      settingsStore: createMutableSettingsStore({ introDismissed: true }),
    });

    try {
      await flushTaskpane();

      await mounted.click("#insert-rendered-markdown-button");
      expect(
        services.composeMarkdown.insertRenderedMarkdown
      ).toHaveBeenCalledWith("# Heading");
      expect(
        notificationService.showTransientNotification
      ).toHaveBeenCalledWith(
        expect.objectContaining({
          intent: "success",
          message: "Rendered Markdown replaced the current selection.",
        })
      );

      await waitForCondition(() => {
        const button = mounted.container.querySelector<HTMLButtonElement>(
          "#render-selection-button"
        );
        return button !== null && !button.disabled;
      });

      await mounted.click("#render-selection-button");
      expect(services.composeMarkdown.renderSelection).toHaveBeenCalledTimes(1);
      expect(
        notificationService.showTransientNotification
      ).toHaveBeenCalledWith(
        expect.objectContaining({
          intent: "success",
          message: "The current body selection was rendered successfully.",
        })
      );

      await mounted.click("#render-entire-draft-button");
      expect(services.renderEntireDraft).toHaveBeenCalledTimes(1);
      expect(
        notificationService.showTransientNotification
      ).toHaveBeenCalledWith(
        expect.objectContaining({
          intent: "info",
          message:
            "No Markdown-looking draft blocks were found, so the message body was left unchanged.",
        })
      );
      expect(diagnosticSink.snapshot().map((event) => event.code)).toEqual(
        expect.arrayContaining([
          "fragment.insert.started",
          "fragment.insert.replaced-selection",
          "selection.render.started",
          "selection.render.succeeded",
          "draft.render.started",
          "draft.render.unchanged",
        ])
      );
    } finally {
      mounted.cleanup();
    }
  });

  it("surfaces action failures and clears the busy state for retry", async () => {
    const services = createTaskpaneServices({
      renderEntireDraft: jest.fn().mockRejectedValue(new Error("draft failed")),
    });
    const notificationService = createNotificationService();
    const mounted = await mountTaskpaneApp({
      notificationService,
      services,
      settingsStore: createMutableSettingsStore({ introDismissed: true }),
    });

    try {
      const renderDraftButton =
        mounted.container.querySelector<HTMLButtonElement>(
          "#render-entire-draft-button"
        );

      await mounted.click("#render-entire-draft-button");
      await waitForCondition(() => renderDraftButton?.disabled === false);

      expect(
        notificationService.showTransientNotification
      ).toHaveBeenCalledWith(
        expect.objectContaining({
          intent: "error",
          message: "draft failed",
        })
      );
      expect(renderDraftButton?.disabled).toBe(false);
    } finally {
      mounted.cleanup();
    }
  });

  it("handles dropped markdown, unsupported files, missing files, and read failures", async () => {
    const originalFileReader = window.FileReader;
    const notificationService = createNotificationService();
    const mounted = await mountTaskpaneApp({
      notificationService,
      settingsStore: createMutableSettingsStore({ introDismissed: true }),
    });

    try {
      await waitForCondition(
        () =>
          mounted.container.querySelector(
            '[data-testid="taskpane-dropzone"]'
          ) !== null
      );

      Object.defineProperty(window, "FileReader", {
        configurable: true,
        value: createFileReaderClass("## Loaded"),
      });

      await dispatchDrop(
        mounted.container.querySelector('[data-testid="taskpane-dropzone"]'),
        new File(["ignored"], "loaded.md")
      );
      expect(
        mounted.container.querySelector<HTMLTextAreaElement>("#markdown-input")
          ?.value
      ).toBe("## Loaded");
      expect(
        notificationService.showTransientNotification
      ).toHaveBeenCalledWith(
        expect.objectContaining({
          intent: "success",
          message: "loaded.md loaded into the insert pane.",
        })
      );

      await dispatchDrop(
        mounted.container.querySelector('[data-testid="taskpane-dropzone"]'),
        new File(["ignored"], "loaded.html")
      );
      expect(
        notificationService.showTransientNotification
      ).toHaveBeenCalledWith(
        expect.objectContaining({
          intent: "error",
          message:
            "Only .md, .markdown, and .txt files are supported in the insert pane.",
        })
      );

      await dispatchDrop(
        mounted.container.querySelector('[data-testid="taskpane-dropzone"]')
      );
      expect(
        notificationService.showTransientNotification
      ).toHaveBeenCalledWith(
        expect.objectContaining({
          intent: "warning",
          message: "Drop a Markdown or text file to load content into MarkOut.",
        })
      );

      Object.defineProperty(window, "FileReader", {
        configurable: true,
        value: createFileReaderClass(null),
      });

      await dispatchDrop(
        mounted.container.querySelector('[data-testid="taskpane-dropzone"]'),
        new File(["ignored"], "broken.md")
      );
      expect(
        notificationService.showTransientNotification
      ).toHaveBeenCalledWith(
        expect.objectContaining({
          intent: "error",
          message: "broken.md could not be read.",
        })
      );
    } finally {
      Object.defineProperty(window, "FileReader", {
        configurable: true,
        value: originalFileReader,
      });
      mounted.cleanup();
    }
  });

  it("lints, resets, persists, and reports stylesheet save failures", async () => {
    const settingsStore = createMutableSettingsStore({
      introDismissed: true,
      stylesheet: ".mo { color: inherit; }",
    });
    const mounted = await mountTaskpaneApp({ settingsStore });

    try {
      await mounted.click("#panel-button-settings");
      await mounted.click("#lint-stylesheet-button");

      expect(mounted.container.textContent).toContain("No lint findings.");

      settingsStore.save.mockRejectedValueOnce(new Error("save failed"));
      const resetButton = findButtonByText(
        mounted.container,
        "Reset default stylesheet"
      );
      act(() => {
        resetButton.click();
      });
      await waitForStylesheetDebounce();

      expect(settingsStore.setStylesheet).toHaveBeenCalledWith(
        defaultStylesheet
      );
      await waitForCondition(() =>
        mounted.container.textContent.includes(
          "Stylesheet changes could not be persisted."
        )
      );

      expect(mounted.container.textContent).toContain(
        "Stylesheet changes could not be persisted."
      );
    } finally {
      mounted.cleanup();
    }
  });
});

async function changeSelectValue(
  select: HTMLSelectElement | null,
  value: string
): Promise<void> {
  if (select === null) {
    throw new Error("Expected select to exist.");
  }

  act(() => {
    select.value = value;
    select.dispatchEvent(new Event("change", { bubbles: true }));
  });
  await flushTaskpane();
}

async function waitForCondition(
  predicate: () => boolean | undefined
): Promise<void> {
  for (let attempt = 0; attempt < 20; attempt += 1) {
    if (predicate()) {
      return;
    }

    await flushTaskpane();
  }

  throw new Error("Timed out waiting for taskpane condition.");
}

async function waitForStylesheetDebounce(): Promise<void> {
  await new Promise<void>((resolve) => {
    window.setTimeout(resolve, 760);
  });
  await flushTaskpane();
}

async function dispatchDrop(
  dropzone: Element | null,
  file?: File
): Promise<void> {
  if (dropzone === null) {
    throw new Error("Expected dropzone to exist.");
  }

  const event = new Event("drop", {
    bubbles: true,
    cancelable: true,
  });

  Object.defineProperty(event, "dataTransfer", {
    value: {
      files: {
        item: () => file ?? null,
      },
    },
  });

  act(() => {
    dropzone.dispatchEvent(event);
  });
  await flushTaskpane();
}

function createFileReaderClass(result: string | null): typeof FileReader {
  return class TestFileReader {
    public onerror: (() => void) | null = null;
    public onload: (() => void) | null = null;
    public result: string | ArrayBuffer | null = result;

    public readAsText(): void {
      if (result === null) {
        this.onerror?.();
        return;
      }

      this.onload?.();
    }
  } as unknown as typeof FileReader;
}

function findButtonByText(
  container: HTMLElement,
  text: string
): HTMLButtonElement {
  const button = Array.from(container.querySelectorAll("button")).find(
    (candidate) => candidate.textContent.trim() === text
  );

  if (button === undefined) {
    throw new Error(`Expected button "${text}" to exist.`);
  }

  return button;
}
