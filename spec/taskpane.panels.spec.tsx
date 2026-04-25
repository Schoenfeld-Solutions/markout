/** @jest-environment jsdom */

import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { act } from "react";
import type { ReactElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import type { StylesheetLintResult } from "../src/lib/stylesheet-lint";
import { getStrings } from "../src/taskpane/i18n";
import {
  CreditsPanel,
  DeveloperPanel,
  InsertPanel,
  SettingsPanel,
  renderActivePanel,
} from "../src/taskpane/panels";

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

describe("taskpane panels", () => {
  afterEach(() => {
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
      expect(callbacks.setDropActive).toHaveBeenCalledWith(true);
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
});
