/** @jest-environment jsdom */

import { act } from "react";
import { createRoot, type Root } from "react-dom/client";
import {
  TaskpaneApp,
  readDroppedMarkdownFile,
  supportsMarkdownFile,
} from "../src/taskpane/app";
import type { SettingsStore, ThemeMode } from "../src/lib/config";
import { installOfficeEnvironment } from "./helpers";

(
  globalThis as typeof globalThis & {
    IS_REACT_ACT_ENVIRONMENT?: boolean;
  }
).IS_REACT_ACT_ENVIRONMENT = true;

class InMemorySettingsStore implements SettingsStore {
  public autoRender = false;
  public developerToolsEnabled = false;
  public introDismissed = false;
  public stylesheet = ".mo { color: rgb(1, 2, 3); }";
  public themeMode: ThemeMode = "system";
  public save = jest.fn().mockResolvedValue(undefined);

  public getAutoRender(): boolean {
    return this.autoRender;
  }

  public getDeveloperToolsEnabled(): boolean {
    return this.developerToolsEnabled;
  }

  public getIntroDismissed(): boolean {
    return this.introDismissed;
  }

  public getStylesheet(): string {
    return this.stylesheet;
  }

  public getThemeMode(): ThemeMode {
    return this.themeMode;
  }

  public setAutoRender(enabled: boolean): void {
    this.autoRender = enabled;
  }

  public setDeveloperToolsEnabled(enabled: boolean): void {
    this.developerToolsEnabled = enabled;
  }

  public setIntroDismissed(dismissed: boolean): void {
    this.introDismissed = dismissed;
  }

  public setStylesheet(stylesheet: string): void {
    this.stylesheet = stylesheet;
  }

  public setThemeMode(mode: ThemeMode): void {
    this.themeMode = mode;
  }
}

describe("taskpane app", () => {
  let container: HTMLDivElement;
  let root: Root;

  beforeEach(() => {
    installOfficeEnvironment();
    Object.defineProperty(window, "matchMedia", {
      configurable: true,
      value: jest.fn().mockImplementation((query: string) => ({
        addEventListener: jest.fn(),
        matches: query.includes("dark"),
        removeEventListener: jest.fn(),
      })),
    });
    container = document.createElement("div");
    document.body.appendChild(container);
    root = createRoot(container);
  });

  afterEach(() => {
    act(() => {
      root.unmount();
    });
    container.remove();
    jest.restoreAllMocks();
  });

  it("renders the intro by default and hides it after confirmation", async () => {
    const settingsStore = new InMemorySettingsStore();

    await renderApp(settingsStore);

    expect(container.textContent).toContain("What MarkOut does");
    clickButton("I have read this");
    await flushPromises();

    expect(container.textContent).toContain("Insert rendered Markdown");
    expect(container.textContent).not.toContain("What MarkOut does");
    expect(() => getButton("Intro")).toThrow('Button "Intro" was not found.');
  });

  it("keeps the developer panel hidden until the setting is enabled", async () => {
    const settingsStore = new InMemorySettingsStore();
    settingsStore.introDismissed = true;

    await renderApp(settingsStore);

    expect(container.textContent).not.toContain("Developer");

    clickButton("Settings");
    await flushPromises();
    toggleCheckbox("developer-tools-switch");
    await flushPromises();

    expect(container.textContent).toContain("Developer");
  });

  it("updates the effective theme when Outlook raises an Office theme change", async () => {
    const environment = installOfficeEnvironment();
    const settingsStore = new InMemorySettingsStore();
    settingsStore.introDismissed = true;

    await renderApp(settingsStore);

    expect(container.querySelector("[data-theme='light']")).not.toBeNull();
    await act(async () => {
      await environment.triggerOfficeThemeChange({
        bodyBackgroundColor: "#111111",
      });
    });
    await flushPromises();

    expect(container.querySelector("[data-theme='dark']")).not.toBeNull();
  });

  it("disables insert until Markdown text is present and then inserts rendered content", async () => {
    const settingsStore = new InMemorySettingsStore();
    settingsStore.introDismissed = true;
    const insertRenderedMarkdown = jest.fn().mockResolvedValue("inserted");

    await renderApp(settingsStore, {
      insertRenderedMarkdown,
      renderPreview: jest.fn().mockResolvedValue("<div>Preview</div>"),
    });

    const insertButton = getButton("Insert rendered markdown");
    expect(insertButton.disabled).toBe(true);

    const markdownInput =
      container.querySelector<HTMLTextAreaElement>("#markdown-input");
    if (markdownInput === null) {
      throw new Error('Textarea "#markdown-input" was not found.');
    }

    setTextareaValue(markdownInput, "## Fragment");
    await flushPromises();

    expect(getButton("Insert rendered markdown").disabled).toBe(false);

    clickButton("Insert rendered markdown");
    await flushPromises();

    expect(insertRenderedMarkdown).toHaveBeenCalledWith("## Fragment");
  });

  it("supports markdown file detection helpers", () => {
    expect(supportsMarkdownFile(new File(["x"], "sample.md"))).toBe(true);
    expect(supportsMarkdownFile(new File(["x"], "sample.markdown"))).toBe(true);
    expect(supportsMarkdownFile(new File(["x"], "sample.txt"))).toBe(true);
    expect(supportsMarkdownFile(new File(["x"], "sample.html"))).toBe(false);
  });

  it("reads dropped files and surfaces read failures", async () => {
    class SuccessfulFileReader {
      public onerror: (() => void) | null = null;
      public onload: (() => void) | null = null;
      public result: string | ArrayBuffer | null = "## Loaded";

      public readAsText(): void {
        this.onload?.();
      }
    }

    class FailingFileReader {
      public onerror: (() => void) | null = null;
      public onload: (() => void) | null = null;
      public result: string | ArrayBuffer | null = null;

      public readAsText(): void {
        this.onerror?.();
      }
    }

    Object.defineProperty(window, "FileReader", {
      configurable: true,
      value: SuccessfulFileReader,
    });

    await expect(
      readDroppedMarkdownFile(new File(["ignored"], "loaded.md"))
    ).resolves.toBe("## Loaded");

    Object.defineProperty(window, "FileReader", {
      configurable: true,
      value: FailingFileReader,
    });

    await expect(
      readDroppedMarkdownFile(new File(["ignored"], "broken.md"))
    ).rejects.toThrow("MarkOut could not read broken.md.");
  });

  async function renderApp(
    settingsStore: InMemorySettingsStore,
    overrides?: Partial<{
      getSelection: jest.Mock;
      insertRenderedMarkdown: jest.Mock;
      renderPreview: jest.Mock;
      renderSelection: jest.Mock;
      renderEntireDraft: jest.Mock;
    }>
  ): Promise<void> {
    act(() => {
      root.render(
        <TaskpaneApp
          services={{
            composeMarkdown: {
              getSelection:
                overrides?.getSelection ??
                jest.fn().mockResolvedValue({
                  hasSelection: false,
                  html: "",
                  source: "body",
                  text: "",
                }),
              insertRenderedMarkdown:
                overrides?.insertRenderedMarkdown ??
                jest.fn().mockResolvedValue("inserted"),
              renderPreview:
                overrides?.renderPreview ??
                jest.fn().mockResolvedValue("<div>Preview</div>"),
              renderSelection:
                overrides?.renderSelection ??
                jest.fn().mockResolvedValue(undefined),
            },
            renderEntireDraft:
              overrides?.renderEntireDraft ??
              jest.fn().mockResolvedValue("rendered"),
          }}
          settingsStore={settingsStore}
        />
      );
    });

    await flushPromises();
  }

  function clickButton(label: string): void {
    act(() => {
      getButton(label).click();
    });
  }

  function getButton(label: string): HTMLButtonElement {
    const button = Array.from(container.querySelectorAll("button")).find(
      (candidate) => candidate.textContent.trim() === label
    );

    if (!(button instanceof HTMLButtonElement)) {
      throw new Error(`Button "${label}" was not found.`);
    }

    return button;
  }

  function toggleCheckbox(id: string): void {
    const input = container.querySelector<HTMLInputElement>(`#${id}`);

    if (input === null) {
      throw new Error(`Checkbox "${id}" was not found.`);
    }

    act(() => {
      input.click();
    });
  }
});

async function flushPromises(): Promise<void> {
  await act(async () => {
    await Promise.resolve();
  });
}

function setTextareaValue(textarea: HTMLTextAreaElement, value: string): void {
  const setValue = Object.getOwnPropertyDescriptor(
    HTMLTextAreaElement.prototype,
    "value"
  )?.set;

  if (setValue === undefined) {
    throw new Error("Textarea value setter is unavailable.");
  }

  act(() => {
    setValue.call(textarea, value);
    textarea.dispatchEvent(new Event("input", { bubbles: true }));
    textarea.dispatchEvent(new Event("change", { bubbles: true }));
  });
}
