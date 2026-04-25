/** @jest-environment jsdom */

import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { act } from "react";
import type { ReactElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import type { ComposeNotificationService } from "../src/lib/compose-notifications";
import type { LanguagePreference, SettingsStore } from "../src/lib/config";
import {
  TaskpaneApp,
  type TaskpaneServices,
  buildToolbarPanels,
  getPanelAfterVisibilityChange,
  getRenderSelectionTooltip,
  isDarkColor,
  isInsertRenderedMarkdownDisabled,
  isRenderSelectionDisabled,
  readDroppedMarkdownFile,
  resolveSystemColorMode,
  resolveToolbarLayoutMode,
  supportsMarkdownFile,
} from "../src/taskpane/app";
import {
  usePreviewController,
  useSelectionStateController,
} from "../src/taskpane/controllers";
import { getStrings } from "../src/taskpane/i18n";
import { HelpPanel, IntroPanel, SettingsPanel } from "../src/taskpane/panels";
import { TaskpaneRuntimeErrorBoundary } from "../src/taskpane/runtime";

(
  globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true;

function createSettingsStore(
  overrides: Partial<{
    autoRender: boolean;
    creditsVisible: boolean;
    developerToolsEnabled: boolean;
    helpVisible: boolean;
    introDismissed: boolean;
    languagePreference: LanguagePreference;
    stylesheetMigrationPending: boolean;
    stylesheet: string;
    themeMode: "dark" | "light" | "system";
  }> = {}
): SettingsStore {
  return {
    getAutoRender: () => overrides.autoRender ?? false,
    getCreditsVisible: () => overrides.creditsVisible ?? true,
    getDeveloperToolsEnabled: () => overrides.developerToolsEnabled ?? false,
    getHelpVisible: () => overrides.helpVisible ?? true,
    getIntroDismissed: () => overrides.introDismissed ?? false,
    getLanguagePreference: () => overrides.languagePreference ?? "system",
    getStylesheet: () => overrides.stylesheet ?? "",
    getThemeMode: () => overrides.themeMode ?? "system",
    hasStylesheetMigrationPending: () =>
      overrides.stylesheetMigrationPending ?? false,
    save: () => Promise.resolve(),
    setAutoRender: () => undefined,
    setCreditsVisible: () => undefined,
    setDeveloperToolsEnabled: () => undefined,
    setHelpVisible: () => undefined,
    setIntroDismissed: () => undefined,
    setLanguagePreference: () => undefined,
    setStylesheet: () => undefined,
    setThemeMode: () => undefined,
  };
}

function createServices(): TaskpaneServices {
  return {
    composeMarkdown: {
      getSelection: () =>
        Promise.resolve({
          hasSelection: false,
          html: null,
          source: "body",
          text: "",
        }),
      insertRenderedMarkdown: () => Promise.resolve("inserted"),
      renderPreview: () => Promise.resolve("<p>preview</p>"),
      renderSelection: () => Promise.resolve(),
    },
    renderEntireDraft: () => Promise.resolve("rendered"),
  };
}

function createNotificationService(
  overrides: Partial<ComposeNotificationService> = {}
): ComposeNotificationService {
  return {
    clearAutoRenderDismissed: () => Promise.resolve(),
    clearAutoRenderNotification: () => Promise.resolve(),
    clearTransientNotification: () => Promise.resolve(),
    hasAutoRenderBeenDismissed: () => Promise.resolve(false),
    markAutoRenderDismissed: () => Promise.resolve(),
    onAutoRenderDismiss: () => undefined,
    showAutoRenderNotification: () => Promise.resolve("pane"),
    showTransientNotification: () => Promise.resolve("outlook"),
    ...overrides,
  };
}

function ensureMatchMedia(): () => void {
  const originalMatchMedia = window.matchMedia;

  Object.defineProperty(window, "matchMedia", {
    configurable: true,
    value: jest.fn().mockReturnValue({
      addEventListener: jest.fn(),
      addListener: jest.fn(),
      dispatchEvent: jest.fn().mockReturnValue(false),
      matches: false,
      media: "(prefers-color-scheme: dark)",
      onchange: null,
      removeEventListener: jest.fn(),
      removeListener: jest.fn(),
    }),
  });

  return () => {
    Object.defineProperty(window, "matchMedia", {
      configurable: true,
      value: originalMatchMedia,
    });
  };
}

function createPanelStyles(): Record<string, string> {
  const styles: Record<string, string> = {};

  return new Proxy(styles, {
    get: (_, key) => String(key),
  });
}

describe("taskpane app helpers", () => {
  afterEach(() => {
    jest.restoreAllMocks();
  });

  it("builds toolbar panels in the expected order and respects visibility toggles", () => {
    const strings = getStrings("en-US");

    expect(
      buildToolbarPanels(
        {
          autoRender: false,
          creditsVisible: true,
          developerToolsEnabled: false,
          helpVisible: true,
          introDismissed: false,
          languagePreference: "system",
          stylesheet: "",
          themeMode: "system",
        },
        strings
      ).map((panel) => panel.key)
    ).toEqual(["intro", "insert", "settings", "help", "credits"]);

    expect(
      buildToolbarPanels(
        {
          autoRender: false,
          creditsVisible: false,
          developerToolsEnabled: true,
          helpVisible: true,
          introDismissed: true,
          languagePreference: "system",
          stylesheet: "",
          themeMode: "system",
        },
        strings
      ).map((panel) => panel.key)
    ).toEqual(["insert", "settings", "help", "developer"]);
  });

  it("switches toolbar layout mode based on available width", () => {
    expect(resolveToolbarLayoutMode(480, 5)).toBe("regular");
    expect(resolveToolbarLayoutMode(300, 5)).toBe("compact");
    expect(getStrings("en-US").tooltips.toolbarCompactHint("Insert")).toBe(
      "Open Insert"
    );
  });

  it("returns the settings panel when a visible toolbar panel is hidden while active", () => {
    expect(getPanelAfterVisibilityChange("help", "help", false)).toBe(
      "settings"
    );
    expect(getPanelAfterVisibilityChange("credits", "credits", false)).toBe(
      "settings"
    );
    expect(getPanelAfterVisibilityChange("developer", "developer", false)).toBe(
      "settings"
    );
    expect(getPanelAfterVisibilityChange("insert", "help", false)).toBe(
      "insert"
    );
  });

  it("disables render selection unless the body selection is available", () => {
    const strings = getStrings("en-US");

    expect(isRenderSelectionDisabled(false, "body-selection")).toBe(false);
    expect(isRenderSelectionDisabled(true, "body-selection")).toBe(true);
    expect(isRenderSelectionDisabled(false, "body-none")).toBe(true);
    expect(isRenderSelectionDisabled(false, "subject")).toBe(true);
    expect(isRenderSelectionDisabled(false, "unknown")).toBe(true);

    expect(getRenderSelectionTooltip(strings, "body-selection")).toContain(
      "currently selected Markdown text"
    );
    expect(getRenderSelectionTooltip(strings, "body-none")).toContain(
      "Select Markdown text"
    );
    expect(getRenderSelectionTooltip(strings, "subject")).toContain(
      "message body"
    );
    expect(getRenderSelectionTooltip(strings, "unknown")).toContain(
      "Selection state could not be read"
    );
  });

  it("disables fragment insertion until Markdown input is present", () => {
    expect(isInsertRenderedMarkdownDisabled(false, "")).toBe(true);
    expect(isInsertRenderedMarkdownDisabled(false, "   ")).toBe(true);
    expect(isInsertRenderedMarkdownDisabled(true, "## Fragment")).toBe(true);
    expect(isInsertRenderedMarkdownDisabled(false, "## Fragment")).toBe(false);
  });

  it("resolves light and dark themes from Office theme colors and browser fallback", () => {
    const originalMatchMedia = window.matchMedia;

    Object.defineProperty(window, "matchMedia", {
      configurable: true,
      value: jest.fn().mockReturnValue({
        addEventListener: jest.fn(),
        matches: true,
        removeEventListener: jest.fn(),
      }),
    });

    expect(isDarkColor("#111111")).toBe(true);
    expect(isDarkColor("#f5f5f5")).toBe(false);
    expect(resolveSystemColorMode({ bodyBackgroundColor: "#111111" })).toBe(
      "dark"
    );
    expect(resolveSystemColorMode({ bodyBackgroundColor: "#f5f5f5" })).toBe(
      "light"
    );
    expect(resolveSystemColorMode(undefined)).toBe("dark");

    Object.defineProperty(window, "matchMedia", {
      configurable: true,
      value: originalMatchMedia,
    });
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

  it("renders the taskpane intro panel without tooltip-specific globals", async () => {
    const restoreMatchMedia = ensureMatchMedia();
    const originalNodeFilter = globalThis.NodeFilter;
    let root: Root | null = null;

    Object.defineProperty(globalThis, "NodeFilter", {
      configurable: true,
      value: undefined,
    });

    try {
      document.body.innerHTML = '<div id="root"></div>';
      const container = document.getElementById("root");

      if (container === null) {
        throw new Error("Expected a taskpane test container.");
      }

      root = createRoot(container);

      await act(async () => {
        root?.render(
          <TaskpaneApp
            locale="en-US"
            notificationService={createNotificationService()}
            services={createServices()}
            settingsStore={createSettingsStore()}
          />
        );
        await Promise.resolve();
      });

      expect(container.textContent).toContain("Intro");
      expect(container.textContent).toContain("I have read this");
      expect(container.querySelector("#panel-button-intro")).not.toBeNull();
    } finally {
      act(() => {
        root?.unmount();
      });
      Object.defineProperty(globalThis, "NodeFilter", {
        configurable: true,
        value: originalNodeFilter,
      });
      restoreMatchMedia();
    }
  });

  it("records preview controller diagnostics for successful renders", async () => {
    const restoreMatchMedia = ensureMatchMedia();
    const events: string[] = [];
    const service = createServices().composeMarkdown;
    let root: Root | null = null;

    function PreviewProbe(): ReactElement {
      const { previewHtml, previewState } = usePreviewController(
        service,
        "# Heading",
        "",
        "Preview failed.",
        () => undefined,
        (event) => {
          events.push(event.code);
        }
      );

      return (
        <div data-state={previewState} id="preview-probe">
          {previewHtml}
        </div>
      );
    }

    try {
      document.body.innerHTML = '<div id="root"></div>';
      const container = document.getElementById("root");

      if (container === null) {
        throw new Error("Expected a taskpane test container.");
      }

      root = createRoot(container);

      await act(async () => {
        root?.render(<PreviewProbe />);
        await Promise.resolve();
        await Promise.resolve();
      });

      expect(container.querySelector("#preview-probe")?.textContent).toContain(
        "preview"
      );
      expect(events).toEqual([
        "preview.render.started",
        "preview.render.succeeded",
      ]);
    } finally {
      act(() => {
        root?.unmount();
      });
      restoreMatchMedia();
    }
  });

  it("records preview controller diagnostics for render failures", async () => {
    const restoreMatchMedia = ensureMatchMedia();
    const consoleError = jest
      .spyOn(console, "error")
      .mockImplementation(() => undefined);
    const events: string[] = [];
    const panelErrors: string[] = [];
    const service = {
      ...createServices().composeMarkdown,
      renderPreview: () => Promise.reject(new TypeError("private failure")),
    };
    let root: Root | null = null;

    function PreviewProbe(): ReactElement {
      const { previewState } = usePreviewController(
        service,
        "# Heading",
        "",
        "Preview failed.",
        (message) => {
          panelErrors.push(message);
        },
        (event) => {
          events.push(event.code);
        }
      );

      return <div id="preview-probe">{previewState}</div>;
    }

    try {
      document.body.innerHTML = '<div id="root"></div>';
      const container = document.getElementById("root");

      if (container === null) {
        throw new Error("Expected a taskpane test container.");
      }

      root = createRoot(container);

      await act(async () => {
        root?.render(<PreviewProbe />);
        await Promise.resolve();
        await Promise.resolve();
      });

      expect(container.querySelector("#preview-probe")?.textContent).toBe(
        "empty"
      );
      expect(panelErrors).toEqual(["Preview failed."]);
      expect(events).toEqual([
        "preview.render.started",
        "preview.render.failed",
      ]);
      expect(consoleError).toHaveBeenCalled();
    } finally {
      act(() => {
        root?.unmount();
      });
      restoreMatchMedia();
    }
  });

  it("ships simplified panel copy for the reduced settings and insert layouts", () => {
    const strings = getStrings("en-US");

    expect(strings.settings.panelDescription).toBe("");
    expect(strings.settings.languageDescription).toBe("");
    expect(strings.settings.themeDescription).toBe("");
    expect(strings.insert.panelDescription).toBe("");
    expect(strings.insert.previewDescription).toBe("");
    expect(strings.help.panelDescription).toBe("");
    expect(strings.help.repoDescription).toBe("");
    expect(strings.help.docsDescription).toBe("");
    expect(strings.help.websiteDescription).toBe("");
    expect(strings.credits.panelDescription).toBe("");
    expect(strings.developer.panelDescription).toBe("");
    expect(strings.intro.panelDescription).toBe("");
  });

  it("does not leave the removed descriptions behind in intro, settings, or help", async () => {
    const restoreMatchMedia = ensureMatchMedia();
    let root: Root | null = null;
    const strings = getStrings("en-US");

    try {
      document.body.innerHTML = '<div id="root"></div>';
      const container = document.getElementById("root");

      if (container === null) {
        throw new Error("Expected a taskpane test container.");
      }

      root = createRoot(container);

      await act(async () => {
        root?.render(
          <FluentProvider theme={webLightTheme}>
            <>
              <IntroPanel
                onConfirm={() => undefined}
                strings={strings}
                styles={createPanelStyles()}
              />
              <SettingsPanel
                autoRenderEnabled={false}
                codeMirrorHostRef={{ current: null }}
                cssLintResult={null}
                developerToolsEnabled={false}
                helpVisible={true}
                introVisible={true}
                isCodeMirrorLoading={false}
                isWorking={false}
                languagePreference="system"
                onCreditsVisibilityChange={() => undefined}
                onDeveloperToolsChange={() => undefined}
                onHelpVisibilityChange={() => undefined}
                onIntroVisibilityChange={() => undefined}
                onLanguagePreferenceChange={() => undefined}
                onLintStylesheet={() => undefined}
                onResetStylesheet={() => undefined}
                onThemeModeChange={() => undefined}
                onToggleAutoRender={() => undefined}
                preferencesThemeMode="system"
                showCredits={true}
                strings={strings}
                styles={createPanelStyles()}
              />
              <HelpPanel strings={strings} styles={createPanelStyles()} />
            </>
          </FluentProvider>
        );
        await Promise.resolve();
      });

      expect(container.textContent).not.toContain(
        "MarkOut keeps compose work Markdown-first while staying inside Outlook's taskpane and Smart Alerts model."
      );

      expect(container.textContent).not.toContain(
        "System follows Outlook theme when the host provides it and falls back to the browser preference otherwise."
      );

      expect(container.textContent).not.toContain(
        "Track issues, releases, and the maintained Schoenfeld Solutions fork."
      );
      expect(container.textContent).not.toContain(
        "Open the GitHub Pages landing page with manifests, hosted docs, and deployment notes."
      );
      expect(container.textContent).not.toContain(
        "Open the Schoenfeld Solutions website. A support link can be added here later."
      );
      expect(
        Array.from(container.querySelectorAll("p")).every(
          (paragraph) => paragraph.textContent.trim().length > 0
        )
      ).toBe(true);
    } finally {
      act(() => {
        root?.unmount();
      });
      restoreMatchMedia();
    }
  });

  it("shows a runtime fallback instead of leaving the pane empty after a render crash", () => {
    const consoleError = jest
      .spyOn(console, "error")
      .mockImplementation(() => undefined);
    let root: Root | null = null;

    function BrokenPane(): never {
      throw new Error("Boom");
    }

    try {
      document.body.innerHTML = '<div id="root"></div>';
      const container = document.getElementById("root");

      if (container === null) {
        throw new Error("Expected a taskpane test container.");
      }

      root = createRoot(container);

      act(() => {
        root?.render(
          <TaskpaneRuntimeErrorBoundary strings={getStrings("en-US")}>
            <BrokenPane />
          </TaskpaneRuntimeErrorBoundary>
        );
      });

      expect(container.textContent).toContain(
        "MarkOut could not render the taskpane"
      );
      expect(container.textContent).toContain("Error: Boom");
      expect(container.querySelector("#taskpane-runtime-error")).not.toBeNull();
      expect(consoleError).toHaveBeenCalled();
    } finally {
      act(() => {
        root?.unmount();
      });
    }
  });
});
