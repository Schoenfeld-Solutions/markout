/** @jest-environment jsdom */

import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { act } from "react";
import type { ReactElement, ReactNode } from "react";
import { createRoot, type Root } from "react-dom/client";
import type { StylesheetLintResult } from "../src/lib/stylesheet-lint";
import { getStrings, type LocalizedStrings } from "../src/taskpane/i18n";
import {
  CreditsPanel,
  DeveloperPanel,
  InsertPanel,
  SettingsPanel,
  renderActivePanel,
} from "../src/taskpane/panels";

(
  globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true;

function createStyles(): Record<string, string> {
  const styles: Record<string, string> = {};

  return new Proxy(styles, {
    get: (_, key) => String(key),
  });
}

function mount(element: ReactElement): { container: HTMLElement; root: Root } {
  document.body.innerHTML = '<div id="root"></div>';
  const container = document.getElementById("root");

  if (container === null) {
    throw new Error("Expected a panel test container.");
  }

  const root = createRoot(container);
  act(() => {
    root.render(
      <FluentProvider theme={webLightTheme}>{element}</FluentProvider>
    );
  });

  return { container, root };
}

function readPanelText(element: ReactElement): ReactNode {
  return (element as ReactElement<{ children: ReactNode }>).props.children;
}

describe("taskpane panels", () => {
  let originalResizeObserver: typeof window.ResizeObserver | undefined;

  beforeEach(() => {
    originalResizeObserver = window.ResizeObserver;
    Object.defineProperty(window, "ResizeObserver", {
      configurable: true,
      value: class TestResizeObserver {
        public disconnect(): void {
          return undefined;
        }

        public observe(): void {
          return undefined;
        }

        public unobserve(): void {
          return undefined;
        }
      },
    });
  });

  afterEach(() => {
    Object.defineProperty(window, "ResizeObserver", {
      configurable: true,
      value: originalResizeObserver,
    });
    jest.restoreAllMocks();
  });

  it("renders the insert panel preview states and user actions", () => {
    const strings = getStrings("en-US");
    const callbacks = {
      onDrop: jest.fn(),
      onInsertRenderedMarkdown: jest.fn(),
      onMarkdownInputChange: jest.fn(),
      onRenderEntireDraft: jest.fn(),
      onRenderSelection: jest.fn(),
      setDropActive: jest.fn(),
    };
    const { container, root } = mount(
      <InsertPanel
        isDropActive={false}
        isInsertRenderedMarkdownDisabled={false}
        isWorking={false}
        markdownInput="## Input"
        onDrop={callbacks.onDrop}
        onInsertRenderedMarkdown={callbacks.onInsertRenderedMarkdown}
        onMarkdownInputChange={callbacks.onMarkdownInputChange}
        onRenderEntireDraft={callbacks.onRenderEntireDraft}
        onRenderSelection={callbacks.onRenderSelection}
        previewHtml="<p>Preview</p>"
        previewFrameStyle={{ colorScheme: "light" }}
        previewState="ready"
        renderSelectionDisabled={false}
        renderSelectionTooltip="Render selection"
        setDropActive={callbacks.setDropActive}
        strings={strings}
        styles={createStyles()}
      />
    );

    try {
      expect(container.textContent).toContain(strings.insert.previewTitle);
      expect(container.querySelector("#mo-preview")?.textContent).toContain(
        "Preview"
      );

      act(() => {
        container
          .querySelector<HTMLTextAreaElement>("#markdown-input")
          ?.dispatchEvent(
            new Event("change", { bubbles: true, cancelable: true })
          );
      });
      container.querySelector<HTMLTextAreaElement>("#markdown-input")!.value =
        "# Changed";
      container
        .querySelector<HTMLTextAreaElement>("#markdown-input")
        ?.dispatchEvent(
          new Event("input", { bubbles: true, cancelable: true })
        );

      container
        .querySelector<HTMLButtonElement>("#render-selection-button")
        ?.click();
      container
        .querySelector<HTMLButtonElement>("#render-entire-draft-button")
        ?.click();
      container
        .querySelector<HTMLButtonElement>("#insert-rendered-markdown-button")
        ?.click();

      expect(callbacks.onRenderSelection).toHaveBeenCalled();
      expect(callbacks.onRenderEntireDraft).toHaveBeenCalled();
      expect(callbacks.onInsertRenderedMarkdown).toHaveBeenCalled();

      const dropzone = container.querySelector<HTMLElement>(
        '[data-testid="taskpane-dropzone"]'
      );
      dropzone?.dispatchEvent(new Event("dragenter", { bubbles: true }));
      dropzone?.dispatchEvent(new Event("dragleave", { bubbles: true }));
      const dragOverEvent = new Event("dragover", {
        bubbles: true,
        cancelable: true,
      });
      const dragOverPreventDefault = jest.spyOn(
        dragOverEvent,
        "preventDefault"
      );
      dropzone?.dispatchEvent(dragOverEvent);
      const dropEvent = new Event("drop", {
        bubbles: true,
        cancelable: true,
      });
      const dropPreventDefault = jest.spyOn(dropEvent, "preventDefault");
      dropzone?.dispatchEvent(dropEvent);

      expect(callbacks.setDropActive).toHaveBeenCalledWith(true);
      expect(callbacks.setDropActive).toHaveBeenCalledWith(false);
      expect(dragOverPreventDefault).toHaveBeenCalled();
      expect(dropPreventDefault).toHaveBeenCalled();
      expect(callbacks.onDrop).toHaveBeenCalledTimes(1);
      expect(
        (
          callbacks.onDrop.mock.calls[0][0] as {
            nativeEvent: Event;
          }
        ).nativeEvent
      ).toBe(dropEvent);
    } finally {
      act(() => {
        root.unmount();
      });
    }
  });

  it("renders settings, developer, and credits panels with their visible states", () => {
    const strings = getStrings("en-US");
    const lintResult: StylesheetLintResult = {
      issues: [
        {
          code: "invalid-rule",
          message: 'The property "colour" is not allowed.',
          severity: "error",
        },
      ],
      validRuleCount: 0,
    };
    const styles = createStyles();
    const settingsPanel = (
      <SettingsPanel
        autoRenderEnabled={true}
        codeMirrorHostRef={{ current: null }}
        cssLintResult={lintResult}
        developerToolsEnabled={true}
        helpVisible={true}
        introVisible={false}
        isCodeMirrorLoading={true}
        isWorking={false}
        languagePreference="de-DE"
        onCreditsVisibilityChange={() => undefined}
        onDeveloperToolsChange={() => undefined}
        onHelpVisibilityChange={() => undefined}
        onIntroVisibilityChange={() => undefined}
        onLanguagePreferenceChange={() => undefined}
        onLintStylesheet={() => undefined}
        onResetStylesheet={() => undefined}
        onThemeModeChange={() => undefined}
        onToggleAutoRender={() => undefined}
        preferencesThemeMode="dark"
        showCredits={true}
        strings={strings}
        styles={styles}
      />
    );
    const developerPanel = (
      <DeveloperPanel
        diagnosticEvents={[
          {
            area: "render",
            code: "draft.render.failed",
            id: 1,
            level: "error",
            metadata: { errorName: "OfficeAsyncError" },
            timestamp: "2026-04-25T10:00:00.000Z",
          },
        ]}
        isInspectingSelection={false}
        onInspectSelection={() => undefined}
        resolvedColorMode="dark"
        selectionDebug={{
          hasSelection: true,
          source: "body",
          textPreview: "Preview text",
        }}
        strings={strings}
        styles={styles}
        themeMode="system"
      />
    );
    const creditsPanel = <CreditsPanel strings={strings} styles={styles} />;
    const { container, root } = mount(
      <>
        {renderActivePanel({
          activePanel: "settings",
          creditsPanel,
          developerPanel,
          helpPanel: <div>Help</div>,
          insertPanel: <div>Insert</div>,
          introPanel: <div>Intro</div>,
          settingsPanel,
        })}
        {developerPanel}
        {creditsPanel}
      </>
    );

    try {
      expect(container.textContent).toContain(strings.settings.panelTitle);
      expect(container.textContent).toContain(strings.editor.loading);
      expect(container.textContent).toContain(strings.editor.lintErrorLabel);
      expect(container.textContent).toContain("colour");
      expect(container.textContent).toContain(
        strings.developer.inspectSelection
      );
      expect(container.textContent).toContain("Preview text");
      expect(container.textContent).toContain(
        strings.developer.diagnosticsTitle
      );
      expect(container.textContent).toContain("draft.render.failed");
      expect(container.textContent).toContain(
        strings.credits.currentMaintenanceTitle
      );
    } finally {
      act(() => {
        root.unmount();
      });
    }
  });

  it("renders alternate insert, settings, and developer branches", () => {
    const baseStrings = getStrings("en-US");
    const strings: LocalizedStrings = {
      ...baseStrings,
      developer: {
        ...baseStrings.developer,
        panelDescription: "Developer description.",
      },
      insert: {
        ...baseStrings.insert,
        panelDescription: "Insert description.",
        previewDescription: "Preview description.",
      },
      settings: {
        ...baseStrings.settings,
        languageDescription: "Language description.",
        panelDescription: "Settings description.",
        themeDescription: "Theme description.",
      },
    };
    const styles = createStyles();
    const warningLintResult: StylesheetLintResult = {
      issues: [
        {
          code: "unsupported-selector",
          message: "Selector will be ignored.",
          severity: "warning",
        },
      ],
      validRuleCount: 1,
    };
    const noIssueLintResult: StylesheetLintResult = {
      issues: [],
      validRuleCount: 2,
    };
    const { container, root } = mount(
      <>
        <InsertPanel
          isDropActive={true}
          isInsertRenderedMarkdownDisabled={true}
          isWorking={true}
          markdownInput=""
          onDrop={() => undefined}
          onInsertRenderedMarkdown={() => undefined}
          onMarkdownInputChange={() => undefined}
          onRenderEntireDraft={() => undefined}
          onRenderSelection={() => undefined}
          previewHtml=""
          previewFrameStyle={{ colorScheme: "dark" }}
          previewState="loading"
          renderSelectionDisabled={true}
          renderSelectionTooltip="Selection unavailable"
          setDropActive={() => undefined}
          strings={strings}
          styles={styles}
        />
        <SettingsPanel
          autoRenderEnabled={false}
          codeMirrorHostRef={{ current: null }}
          cssLintResult={warningLintResult}
          developerToolsEnabled={false}
          helpVisible={false}
          introVisible={true}
          isCodeMirrorLoading={false}
          isWorking={true}
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
          showCredits={false}
          strings={strings}
          styles={styles}
        />
        <SettingsPanel
          autoRenderEnabled={false}
          codeMirrorHostRef={{ current: null }}
          cssLintResult={noIssueLintResult}
          developerToolsEnabled={false}
          helpVisible={false}
          introVisible={false}
          isCodeMirrorLoading={false}
          isWorking={false}
          languagePreference="en-US"
          onCreditsVisibilityChange={() => undefined}
          onDeveloperToolsChange={() => undefined}
          onHelpVisibilityChange={() => undefined}
          onIntroVisibilityChange={() => undefined}
          onLanguagePreferenceChange={() => undefined}
          onLintStylesheet={() => undefined}
          onResetStylesheet={() => undefined}
          onThemeModeChange={() => undefined}
          onToggleAutoRender={() => undefined}
          preferencesThemeMode="light"
          showCredits={false}
          strings={strings}
          styles={styles}
        />
        <DeveloperPanel
          diagnosticEvents={[]}
          isInspectingSelection={true}
          onInspectSelection={() => undefined}
          resolvedColorMode="light"
          selectionDebug={null}
          strings={strings}
          styles={styles}
          themeMode="light"
        />
      </>
    );

    try {
      expect(container.textContent).toContain("Insert description.");
      expect(container.textContent).toContain("Preview description.");
      expect(container.textContent).toContain(strings.insert.previewLoading);
      expect(container.textContent).toContain("Settings description.");
      expect(container.textContent).toContain("Theme description.");
      expect(container.textContent).toContain("Language description.");
      expect(container.textContent).toContain(strings.editor.lintWarningLabel);
      expect(container.textContent).toContain("Selector will be ignored.");
      expect(container.textContent).toContain(strings.editor.lintNoIssues);
      expect(container.textContent).toContain("Developer description.");
      expect(container.textContent).toContain(
        strings.developer.noSelectionSnapshot
      );
      expect(container.textContent).toContain(strings.developer.noDiagnostics);
    } finally {
      act(() => {
        root.unmount();
      });
    }
  });

  it("selects every active panel branch and falls back to insert", () => {
    const panels = {
      creditsPanel: <div>Credits</div>,
      developerPanel: <div>Developer</div>,
      helpPanel: <div>Help</div>,
      insertPanel: <div>Insert</div>,
      introPanel: <div>Intro</div>,
      settingsPanel: <div>Settings</div>,
    };

    expect(
      readPanelText(renderActivePanel({ ...panels, activePanel: "credits" }))
    ).toBe("Credits");
    expect(
      readPanelText(renderActivePanel({ ...panels, activePanel: "developer" }))
    ).toBe("Developer");
    expect(
      readPanelText(renderActivePanel({ ...panels, activePanel: "help" }))
    ).toBe("Help");
    expect(
      readPanelText(renderActivePanel({ ...panels, activePanel: "intro" }))
    ).toBe("Intro");
    expect(
      readPanelText(renderActivePanel({ ...panels, activePanel: "settings" }))
    ).toBe("Settings");
    expect(
      readPanelText(renderActivePanel({ ...panels, activePanel: "insert" }))
    ).toBe("Insert");
  });
});
