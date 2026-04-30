/** @jest-environment jsdom */

import { SUBJECT_SELECTION_UNSUPPORTED_MESSAGE } from "../src/lib/compose-markdown";
import { getStrings } from "../src/taskpane/i18n";
import {
  createTaskpaneActionHandlers,
  persistTaskpanePreferences,
  runWithTaskpaneBusyState,
  type TaskpaneActionHandlers,
} from "../src/taskpane/taskpane-actions";
import type {
  PanelKey,
  PanelMessageState,
  PreferenceState,
} from "../src/taskpane/types";
import {
  createMutableSettingsStore,
  createTaskpaneServices,
  type MutableSettingsStore,
  type TaskpaneServiceMocks,
} from "./taskpane-app-harness";

const DEFAULT_PREFERENCES: PreferenceState = {
  autoRender: false,
  creditsVisible: true,
  developerToolsEnabled: false,
  helpVisible: true,
  introDismissed: true,
  languagePreference: "system",
  stylesheet: ".mo { color: inherit; }",
  themeMode: "system",
};

interface ActionHarness {
  activePanel: PanelKey;
  diagnostics: string[];
  handlers: TaskpaneActionHandlers;
  notifications: PanelMessageState[];
  panelMessages: (PanelMessageState | null)[];
  preferences: PreferenceState;
  services: TaskpaneServiceMocks;
  settingsStore: MutableSettingsStore;
  state: {
    cssLintResult: unknown;
    isInspectingSelection: boolean;
    isWorking: string | null;
    markdownInput: string;
  };
  updateSelectionState: jest.Mock<Promise<boolean>, []>;
}

describe("taskpane actions", () => {
  beforeEach(() => {
    jest.spyOn(console, "error").mockImplementation(() => undefined);
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  it("clears busy state after successful and failed operations", async () => {
    const setIsWorking = jest.fn();

    await runWithTaskpaneBusyState("insert-markdown", setIsWorking, () =>
      Promise.resolve()
    );

    await expect(
      runWithTaskpaneBusyState("render-selection", setIsWorking, () =>
        Promise.reject(new Error("boom"))
      )
    ).rejects.toThrow("boom");

    expect(setIsWorking.mock.calls).toEqual([
      ["insert-markdown"],
      [null],
      ["render-selection"],
      [null],
    ]);
  });

  it("persists preferences and rolls back on roaming settings failure", async () => {
    const strings = getStrings("en-US");
    const settingsStore = createMutableSettingsStore();
    let preferences = DEFAULT_PREFERENCES;
    const setPreferences = jest.fn(
      (
        next: PreferenceState | ((current: PreferenceState) => PreferenceState)
      ) => {
        preferences = typeof next === "function" ? next(preferences) : next;
      }
    );
    const panelMessages: (PanelMessageState | null)[] = [];

    await expect(
      persistTaskpanePreferences({
        localizedStrings: strings,
        nextPreferences: { ...preferences, autoRender: true },
        previousPreferences: preferences,
        settingsStore,
        setPanelMessage: (message) => {
          panelMessages.push(message);
        },
        setPreferences,
      })
    ).resolves.toBe(true);

    expect(settingsStore.state.autoRender).toBe(true);
    expect(preferences.autoRender).toBe(true);

    settingsStore.save.mockRejectedValueOnce(new Error("roaming down"));
    await expect(
      persistTaskpanePreferences({
        localizedStrings: strings,
        nextPreferences: { ...preferences, helpVisible: false },
        previousPreferences: preferences,
        settingsStore,
        setPanelMessage: (message) => {
          panelMessages.push(message);
        },
        setPreferences,
      })
    ).resolves.toBe(false);

    expect(settingsStore.state.helpVisible).toBe(true);
    expect(preferences.helpVisible).toBe(true);
    expect(panelMessages.at(-1)).toEqual({
      body: strings.status.settingsUpdateFailed,
      intent: "error",
    });
  });

  it("toggles visible settings and moves away from a hidden active panel", async () => {
    const harness = createActionHarness({ activePanel: "help" });

    await harness.handlers.toggleHelpVisibility(false);

    expect(harness.settingsStore.state.helpVisible).toBe(false);
    expect(harness.activePanel).toBe("settings");

    const visibleCredits = createActionHarness({ activePanel: "credits" });
    await visibleCredits.handlers.toggleCreditsVisibility(true);

    expect(visibleCredits.settingsStore.state.creditsVisible).toBe(true);
    expect(visibleCredits.activePanel).toBe("credits");

    const hiddenDeveloper = createActionHarness({ activePanel: "developer" });
    await hiddenDeveloper.handlers.toggleDeveloperTools(false);

    expect(hiddenDeveloper.settingsStore.state.developerToolsEnabled).toBe(
      false
    );
    expect(hiddenDeveloper.activePanel).toBe("settings");
  });

  it("persists simple preference actions and keeps failed intro visibility in place", async () => {
    const autoRender = createActionHarness();
    await autoRender.handlers.toggleAutoRender(true);

    expect(autoRender.settingsStore.state.autoRender).toBe(true);

    const language = createActionHarness();
    await language.handlers.setLanguagePreference("de-DE");

    expect(language.settingsStore.state.languagePreference).toBe("de-DE");

    const theme = createActionHarness();
    await theme.handlers.setThemeMode("dark");

    expect(theme.settingsStore.state.themeMode).toBe("dark");

    const failedIntro = createActionHarness();
    failedIntro.settingsStore.save.mockRejectedValueOnce(new Error("no save"));

    await failedIntro.handlers.toggleIntroVisibility(true);

    expect(failedIntro.activePanel).toBe("insert");
    expect(failedIntro.panelMessages.at(-1)).toEqual({
      body: "Settings could not be updated.",
      intent: "error",
    });
  });

  it("confirms intro once and navigates directly when it was already dismissed", async () => {
    const firstRun = createActionHarness({
      preferences: { ...DEFAULT_PREFERENCES, introDismissed: false },
    });
    await firstRun.handlers.confirmIntro();

    expect(firstRun.settingsStore.state.introDismissed).toBe(true);
    expect(firstRun.activePanel).toBe("insert");

    const dismissed = createActionHarness();
    await dismissed.handlers.confirmIntro();

    expect(dismissed.settingsStore.save).not.toHaveBeenCalled();
    expect(dismissed.activePanel).toBe("insert");
  });

  it("runs insert flow with diagnostics, notification copy, and selection refresh", async () => {
    const harness = createActionHarness({
      markdownInput: "# Heading",
      services: createTaskpaneServices({
        insertRenderedMarkdown: jest.fn().mockResolvedValue("replaced"),
      }),
    });

    await harness.handlers.insertRenderedMarkdown();

    expect(
      harness.services.composeMarkdown.insertRenderedMarkdown
    ).toHaveBeenCalledWith("# Heading");
    expect(harness.notifications.at(-1)).toEqual({
      body: "Rendered Markdown replaced the current selection.",
      intent: "success",
    });
    expect(harness.diagnostics).toEqual([
      "fragment.insert.started",
      "fragment.insert.replaced-selection",
    ]);
    expect(harness.updateSelectionState).toHaveBeenCalledTimes(1);
    expect(harness.state.isWorking).toBe(null);

    const inserted = createActionHarness({
      services: createTaskpaneServices({
        insertRenderedMarkdown: jest.fn().mockResolvedValue("inserted"),
      }),
    });

    await inserted.handlers.insertRenderedMarkdown();

    expect(inserted.notifications.at(-1)).toEqual({
      body: "Rendered Markdown was inserted at the current body cursor.",
      intent: "success",
    });
  });

  it("surfaces insert failures without leaving the action busy", async () => {
    const harness = createActionHarness({
      services: createTaskpaneServices({
        insertRenderedMarkdown: jest
          .fn()
          .mockRejectedValue(new Error(SUBJECT_SELECTION_UNSUPPORTED_MESSAGE)),
      }),
    });

    await harness.handlers.insertRenderedMarkdown();

    expect(harness.notifications.at(-1)).toEqual({
      body: "MarkOut can only update the message body. Move the cursor into the body or select text there first.",
      intent: "error",
    });
    expect(harness.diagnostics).toEqual([
      "fragment.insert.started",
      "fragment.insert.failed",
    ]);
    expect(harness.updateSelectionState).not.toHaveBeenCalled();
    expect(harness.state.isWorking).toBe(null);
  });

  it("runs selection rendering success and failure paths", async () => {
    const success = createActionHarness();
    await success.handlers.renderSelection();

    expect(success.services.composeMarkdown.renderSelection).toHaveBeenCalled();
    expect(success.notifications.at(-1)).toEqual({
      body: "The current body selection was rendered successfully.",
      intent: "success",
    });
    expect(success.diagnostics).toEqual([
      "selection.render.started",
      "selection.render.succeeded",
    ]);
    expect(success.updateSelectionState).toHaveBeenCalledTimes(1);

    const failure = createActionHarness({
      services: createTaskpaneServices({
        renderSelection: jest.fn().mockRejectedValue(new Error("nope")),
      }),
    });
    await failure.handlers.renderSelection();

    expect(failure.notifications.at(-1)).toEqual({
      body: "nope",
      intent: "error",
    });
    expect(failure.diagnostics).toEqual([
      "selection.render.started",
      "selection.render.failed",
    ]);
    expect(failure.updateSelectionState).toHaveBeenCalledTimes(1);
    expect(failure.state.isWorking).toBe(null);
  });

  it("runs draft render result and failure paths", async () => {
    const restored = createActionHarness({
      services: createTaskpaneServices({
        renderEntireDraft: jest.fn().mockResolvedValue("restored"),
      }),
    });
    await restored.handlers.renderEntireDraft();

    expect(restored.notifications.at(-1)).toEqual({
      body: "The original draft HTML was restored successfully.",
      intent: "success",
    });
    expect(restored.diagnostics).toEqual([
      "draft.render.started",
      "draft.restore.succeeded",
    ]);
    expect(restored.updateSelectionState).toHaveBeenCalledTimes(1);

    const failure = createActionHarness({
      services: createTaskpaneServices({
        renderEntireDraft: jest.fn().mockRejectedValue(new Error("draft down")),
      }),
    });
    await failure.handlers.renderEntireDraft();

    expect(failure.notifications.at(-1)).toEqual({
      body: "draft down",
      intent: "error",
    });
    expect(failure.diagnostics).toEqual([
      "draft.render.started",
      "draft.render.failed",
    ]);
    expect(failure.updateSelectionState).not.toHaveBeenCalled();
    expect(failure.state.isWorking).toBe(null);
  });

  it("reports manual selection inspection outcomes and failures", async () => {
    const unavailable = createActionHarness({
      updateSelectionState: jest.fn().mockResolvedValue(false),
    });
    await unavailable.handlers.inspectSelection();

    expect(unavailable.notifications.at(-1)).toEqual({
      body: "Selection state could not be read from Outlook.",
      intent: "error",
    });
    expect(unavailable.state.isInspectingSelection).toBe(false);

    const failure = createActionHarness({
      updateSelectionState: jest.fn().mockRejectedValue(new Error("boom")),
    });
    await failure.handlers.inspectSelection();

    expect(failure.notifications.at(-1)).toEqual({
      body: "Selection state could not be read from Outlook.",
      intent: "error",
    });
    expect(failure.state.isInspectingSelection).toBe(false);
  });

  it("loads dropped markdown, rejects unsupported files, and handles read/decode failures", async () => {
    const originalFileReader = window.FileReader;
    const harness = createActionHarness();

    try {
      await harness.handlers.loadDroppedMarkdownFile(null);
      expect(harness.notifications.at(-1)).toEqual({
        body: "Drop a Markdown or text file to load content into MarkOut.",
        intent: "warning",
      });

      await harness.handlers.loadDroppedMarkdownFile(
        new File(["ignored"], "index.html")
      );
      expect(harness.notifications.at(-1)).toEqual({
        body: "Only .md, .markdown, and .txt files are supported in the insert pane.",
        intent: "error",
      });

      Object.defineProperty(window, "FileReader", {
        configurable: true,
        value: createFileReaderClass("  - parent\n\u00a0\u00a0- child"),
      });
      await harness.handlers.loadDroppedMarkdownFile(
        new File(["ignored"], "loaded.md")
      );

      expect(harness.state.markdownInput).toBe("  - parent\n  - child");
      expect(harness.notifications.at(-1)).toEqual({
        body: "loaded.md loaded into the insert pane.",
        intent: "success",
      });

      Object.defineProperty(window, "FileReader", {
        configurable: true,
        value: createFileReaderClass(null),
      });
      await harness.handlers.loadDroppedMarkdownFile(
        new File(["ignored"], "broken.md")
      );
      expect(harness.notifications.at(-1)).toEqual({
        body: "broken.md could not be read.",
        intent: "error",
      });

      Object.defineProperty(window, "FileReader", {
        configurable: true,
        value: createFileReaderClass(new ArrayBuffer(0)),
      });
      await harness.handlers.loadDroppedMarkdownFile(
        new File(["ignored"], "binary.md")
      );
      expect(harness.notifications.at(-1)).toEqual({
        body: "binary.md could not be decoded.",
        intent: "error",
      });
    } finally {
      Object.defineProperty(window, "FileReader", {
        configurable: true,
        value: originalFileReader,
      });
    }
  });

  it("runs stylesheet lint and records the result without changing busy state", async () => {
    const harness = createActionHarness({
      preferences: {
        ...DEFAULT_PREFERENCES,
        stylesheet: "p { color: inherit; }\nscript { behavior: url(x); }",
      },
    });

    await harness.handlers.runStylesheetLint();

    expect(harness.state.cssLintResult).toEqual(
      expect.objectContaining({
        validRuleCount: 2,
      })
    );
    expect(harness.state.isWorking).toBe(null);
  });
});

function createActionHarness(
  overrides: Partial<{
    activePanel: PanelKey;
    markdownInput: string;
    preferences: PreferenceState;
    services: TaskpaneServiceMocks;
    updateSelectionState: jest.Mock<Promise<boolean>, []>;
  }> = {}
): ActionHarness {
  let activePanel = overrides.activePanel ?? "insert";
  let preferences = overrides.preferences ?? DEFAULT_PREFERENCES;
  const settingsStore = createMutableSettingsStore(preferences);
  const diagnostics: string[] = [];
  const notifications: PanelMessageState[] = [];
  const panelMessages: (PanelMessageState | null)[] = [];
  const services = overrides.services ?? createTaskpaneServices();
  const state = {
    cssLintResult: null as unknown,
    isInspectingSelection: false,
    isWorking: null as string | null,
    markdownInput: overrides.markdownInput ?? "# Heading",
  };
  const updateSelectionState =
    overrides.updateSelectionState ?? jest.fn().mockResolvedValue(true);

  const harness = {
    get activePanel() {
      return activePanel;
    },
    diagnostics,
    handlers: null as unknown as TaskpaneActionHandlers,
    notifications,
    panelMessages,
    get preferences() {
      return preferences;
    },
    services,
    settingsStore,
    state,
    updateSelectionState,
  };

  harness.handlers = createTaskpaneActionHandlers({
    activePanel,
    localizedStrings: getStrings("en-US"),
    markdownInput: state.markdownInput,
    preferences,
    recordDiagnostic: (event) => {
      diagnostics.push(event.code);
    },
    services,
    settingsStore,
    setActivePanel: (panel) => {
      activePanel = panel;
    },
    setCssLintResult: (result) => {
      state.cssLintResult = result;
    },
    setIsInspectingSelection: (isInspecting) => {
      state.isInspectingSelection = isInspecting;
    },
    setIsWorking: (busyKey) => {
      state.isWorking = busyKey;
    },
    setMarkdownInput: (markdown) => {
      state.markdownInput = markdown;
    },
    setPanelMessage: (message) => {
      panelMessages.push(message);
    },
    setPreferences: (next) => {
      preferences = typeof next === "function" ? next(preferences) : next;
    },
    showComposeNotification: (intent, message) => {
      notifications.push({ body: message, intent });
      return Promise.resolve();
    },
    updateSelectionState,
  });

  return harness;
}

function createFileReaderClass(
  result: string | ArrayBuffer | null
): typeof FileReader {
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
