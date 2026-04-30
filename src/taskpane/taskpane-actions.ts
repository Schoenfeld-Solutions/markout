import type { InsertMarkdownResult } from "../lib/compose-markdown";
import type { SettingsStore, ThemeMode } from "../lib/config";
import {
  getErrorDiagnosticMetadata,
  type DiagnosticEventInput,
} from "../lib/runtime";
import {
  lintStylesheet,
  type StylesheetLintResult,
} from "../lib/stylesheet-lint";
import { readDroppedMarkdownFile, supportsMarkdownFile } from "./file-drop";
import type { LocalizedStrings } from "./i18n";
import { normalizeMarkdownInput } from "./markdown-input";
import { writePreferences } from "./preferences";
import {
  getDraftRenderFeedback,
  localizeActionError,
} from "./taskpane-feedback";
import { getPanelAfterVisibilityChange } from "./toolbar";
import type {
  PanelKey,
  PanelMessageState,
  PreferenceState,
  TaskpaneServices,
} from "./types";

type StateUpdate<T> = T | ((current: T) => T);
type StateSetter<T> = (value: StateUpdate<T>) => void;

export type RecordDiagnostic = (event: DiagnosticEventInput) => void;

export type ShowComposeNotification = (
  intent: PanelMessageState["intent"],
  message: string
) => Promise<void>;

export interface TaskpaneActionDependencies {
  activePanel: PanelKey;
  localizedStrings: LocalizedStrings;
  markdownInput: string;
  preferences: PreferenceState;
  recordDiagnostic: RecordDiagnostic;
  services: TaskpaneServices;
  settingsStore: SettingsStore;
  setActivePanel: (panel: PanelKey) => void;
  setCssLintResult: (result: StylesheetLintResult | null) => void;
  setIsInspectingSelection: (isInspecting: boolean) => void;
  setIsWorking: (busyKey: string | null) => void;
  setMarkdownInput: (markdown: string) => void;
  setPanelMessage: (message: PanelMessageState | null) => void;
  setPreferences: StateSetter<PreferenceState>;
  showComposeNotification: ShowComposeNotification;
  updateSelectionState: () => Promise<boolean>;
}

export interface TaskpaneActionHandlers {
  confirmIntro(): Promise<void>;
  insertRenderedMarkdown(): Promise<void>;
  inspectSelection(): Promise<void>;
  loadDroppedMarkdownFile(file: File | null): Promise<void>;
  renderEntireDraft(): Promise<void>;
  renderSelection(): Promise<void>;
  runStylesheetLint(): Promise<void>;
  setLanguagePreference(
    preference: PreferenceState["languagePreference"]
  ): Promise<void>;
  setThemeMode(mode: ThemeMode): Promise<void>;
  toggleAutoRender(enabled: boolean): Promise<void>;
  toggleCreditsVisibility(visible: boolean): Promise<void>;
  toggleDeveloperTools(enabled: boolean): Promise<void>;
  toggleHelpVisibility(visible: boolean): Promise<void>;
  toggleIntroVisibility(showIntro: boolean): Promise<void>;
}

interface PersistTaskpanePreferencesOptions {
  localizedStrings: LocalizedStrings;
  nextPreferences: PreferenceState;
  previousPreferences: PreferenceState;
  settingsStore: SettingsStore;
  setPanelMessage: (message: PanelMessageState | null) => void;
  setPreferences: StateSetter<PreferenceState>;
}

export async function runWithTaskpaneBusyState(
  busyKey: string,
  setIsWorking: (busyKey: string | null) => void,
  operation: () => void | Promise<void>
): Promise<void> {
  setIsWorking(busyKey);

  try {
    await operation();
  } finally {
    setIsWorking(null);
  }
}

export async function persistTaskpanePreferences({
  localizedStrings,
  nextPreferences,
  previousPreferences,
  settingsStore,
  setPanelMessage,
  setPreferences,
}: PersistTaskpanePreferencesOptions): Promise<boolean> {
  setPreferences(nextPreferences);
  writePreferences(settingsStore, nextPreferences);

  try {
    await settingsStore.save();
    return true;
  } catch (error) {
    console.error("MarkOut failed to persist settings.", error);
    setPreferences(previousPreferences);
    writePreferences(settingsStore, previousPreferences);
    setPanelMessage({
      body: localizedStrings.status.settingsUpdateFailed,
      intent: "error",
    });
    return false;
  }
}

export function createTaskpaneActionHandlers(
  dependencies: TaskpaneActionDependencies
): TaskpaneActionHandlers {
  const {
    activePanel,
    localizedStrings,
    markdownInput,
    preferences,
    recordDiagnostic,
    services,
    settingsStore,
    setActivePanel,
    setCssLintResult,
    setIsInspectingSelection,
    setIsWorking,
    setMarkdownInput,
    setPanelMessage,
    setPreferences,
    showComposeNotification,
    updateSelectionState,
  } = dependencies;

  const persistPreferences = (nextPreferences: PreferenceState) =>
    persistTaskpanePreferences({
      localizedStrings,
      nextPreferences,
      previousPreferences: preferences,
      settingsStore,
      setPanelMessage,
      setPreferences,
    });

  const persistVisibilityPreference = async (
    panel: Extract<PanelKey, "credits" | "developer" | "help">,
    visible: boolean,
    nextPreferences: PreferenceState
  ): Promise<void> => {
    const didPersist = await persistPreferences(nextPreferences);

    if (didPersist && !visible && activePanel === panel) {
      setActivePanel(
        getPanelAfterVisibilityChange(activePanel, panel, visible)
      );
    }
  };

  return {
    confirmIntro: async () => {
      if (preferences.introDismissed) {
        setActivePanel("insert");
        return;
      }

      const didPersist = await persistPreferences({
        ...preferences,
        introDismissed: true,
      });

      if (didPersist) {
        setActivePanel("insert");
      }
    },
    insertRenderedMarkdown: async () => {
      await insertRenderedMarkdown({
        localizedStrings,
        markdownInput,
        recordDiagnostic,
        services,
        setIsWorking,
        showComposeNotification,
        updateSelectionState,
      });
    },
    inspectSelection: async () => {
      await inspectSelection({
        localizedStrings,
        setIsInspectingSelection,
        showComposeNotification,
        updateSelectionState,
      });
    },
    loadDroppedMarkdownFile: async (file) => {
      await loadDroppedMarkdownFile({
        file,
        localizedStrings,
        setMarkdownInput,
        showComposeNotification,
      });
    },
    renderEntireDraft: async () => {
      await renderEntireDraft({
        localizedStrings,
        recordDiagnostic,
        services,
        setIsWorking,
        showComposeNotification,
        updateSelectionState,
      });
    },
    renderSelection: async () => {
      await renderSelection({
        localizedStrings,
        recordDiagnostic,
        services,
        setIsWorking,
        showComposeNotification,
        updateSelectionState,
      });
    },
    runStylesheetLint: async () => {
      await runStylesheetLint({
        localizedStrings,
        preferences,
        setCssLintResult,
        setIsWorking,
        setPanelMessage,
      });
    },
    setLanguagePreference: async (preference) => {
      await persistPreferences({
        ...preferences,
        languagePreference: preference,
      });
    },
    setThemeMode: async (mode) => {
      await persistPreferences({
        ...preferences,
        themeMode: mode,
      });
    },
    toggleAutoRender: async (enabled) => {
      await persistPreferences({ ...preferences, autoRender: enabled });
    },
    toggleCreditsVisibility: async (visible) => {
      await persistVisibilityPreference("credits", visible, {
        ...preferences,
        creditsVisible: visible,
      });
    },
    toggleDeveloperTools: async (enabled) => {
      await persistVisibilityPreference("developer", enabled, {
        ...preferences,
        developerToolsEnabled: enabled,
      });
    },
    toggleHelpVisibility: async (visible) => {
      await persistVisibilityPreference("help", visible, {
        ...preferences,
        helpVisible: visible,
      });
    },
    toggleIntroVisibility: async (showIntro) => {
      const didPersist = await persistPreferences({
        ...preferences,
        introDismissed: !showIntro,
      });

      if (didPersist) {
        setActivePanel(showIntro ? "intro" : "insert");
      }
    },
  };
}

async function inspectSelection({
  localizedStrings,
  setIsInspectingSelection,
  showComposeNotification,
  updateSelectionState,
}: Pick<
  TaskpaneActionDependencies,
  | "localizedStrings"
  | "setIsInspectingSelection"
  | "showComposeNotification"
  | "updateSelectionState"
>): Promise<void> {
  setIsInspectingSelection(true);

  try {
    const loaded = await updateSelectionState();
    await showComposeNotification(
      loaded ? "success" : "error",
      loaded
        ? localizedStrings.status.selectionInspectionSuccess
        : localizedStrings.status.selectionInspectionFailed
    );
  } catch (error) {
    console.error("MarkOut failed to inspect the current selection.", error);
    await showComposeNotification(
      "error",
      localizedStrings.status.selectionInspectionFailed
    );
  } finally {
    setIsInspectingSelection(false);
  }
}

interface RenderActionDependencies {
  localizedStrings: LocalizedStrings;
  recordDiagnostic: RecordDiagnostic;
  services: TaskpaneServices;
  setIsWorking: (busyKey: string | null) => void;
  showComposeNotification: ShowComposeNotification;
  updateSelectionState: () => Promise<boolean>;
}

async function insertRenderedMarkdown({
  localizedStrings,
  markdownInput,
  recordDiagnostic,
  services,
  setIsWorking,
  showComposeNotification,
  updateSelectionState,
}: RenderActionDependencies & { markdownInput: string }): Promise<void> {
  await runWithTaskpaneBusyState("insert-markdown", setIsWorking, async () => {
    recordDiagnostic({
      area: "render",
      code: "fragment.insert.started",
      level: "debug",
      metadata: { inputLength: markdownInput.length },
    });

    try {
      const result =
        await services.composeMarkdown.insertRenderedMarkdown(markdownInput);
      await showComposeNotification(
        "success",
        getInsertRenderedMarkdownMessage(localizedStrings, result)
      );
      recordDiagnostic({
        area: "body-io",
        code:
          result === "replaced"
            ? "fragment.insert.replaced-selection"
            : "fragment.insert.inserted-at-cursor",
        level: "info",
      });
      await updateSelectionState();
    } catch (error) {
      console.error("MarkOut failed to insert rendered Markdown.", error);
      recordDiagnostic({
        area: "body-io",
        code: "fragment.insert.failed",
        level: "error",
        metadata: getErrorDiagnosticMetadata(error),
      });
      await showComposeNotification(
        "error",
        localizeActionError(localizedStrings, error)
      );
    }
  });
}

async function renderSelection({
  localizedStrings,
  recordDiagnostic,
  services,
  setIsWorking,
  showComposeNotification,
  updateSelectionState,
}: RenderActionDependencies): Promise<void> {
  await runWithTaskpaneBusyState("render-selection", setIsWorking, async () => {
    recordDiagnostic({
      area: "render",
      code: "selection.render.started",
      level: "debug",
    });

    try {
      await services.composeMarkdown.renderSelection();
      await showComposeNotification(
        "success",
        localizedStrings.status.selectionRendered
      );
      recordDiagnostic({
        area: "body-io",
        code: "selection.render.succeeded",
        level: "info",
      });
      await updateSelectionState();
    } catch (error) {
      console.error("MarkOut failed to render the current selection.", error);
      recordDiagnostic({
        area: "body-io",
        code: "selection.render.failed",
        level: "error",
        metadata: getErrorDiagnosticMetadata(error),
      });
      await showComposeNotification(
        "error",
        localizeActionError(localizedStrings, error)
      );
      await updateSelectionState();
    }
  });
}

async function renderEntireDraft({
  localizedStrings,
  recordDiagnostic,
  services,
  setIsWorking,
  showComposeNotification,
  updateSelectionState,
}: RenderActionDependencies): Promise<void> {
  await runWithTaskpaneBusyState(
    "render-entire-draft",
    setIsWorking,
    async () => {
      recordDiagnostic({
        area: "render",
        code: "draft.render.started",
        level: "debug",
      });

      try {
        const result = await services.renderEntireDraft();
        const feedback = getDraftRenderFeedback(localizedStrings, result);
        await showComposeNotification(feedback.intent, feedback.message);
        recordDiagnostic({
          area: feedback.diagnosticArea,
          code: feedback.diagnosticCode,
          level: "info",
        });
        await updateSelectionState();
      } catch (error) {
        console.error("MarkOut failed to render the current draft.", error);
        recordDiagnostic({
          area: "render",
          code: "draft.render.failed",
          level: "error",
          metadata: getErrorDiagnosticMetadata(error),
        });
        await showComposeNotification(
          "error",
          localizeActionError(localizedStrings, error)
        );
      }
    }
  );
}

async function runStylesheetLint({
  localizedStrings,
  preferences,
  setCssLintResult,
  setIsWorking,
  setPanelMessage,
}: Pick<
  TaskpaneActionDependencies,
  | "localizedStrings"
  | "preferences"
  | "setCssLintResult"
  | "setIsWorking"
  | "setPanelMessage"
>): Promise<void> {
  await runWithTaskpaneBusyState("lint-stylesheet", setIsWorking, () => {
    try {
      setCssLintResult(lintStylesheet(preferences.stylesheet));
    } catch (error) {
      console.error("MarkOut failed to lint the stylesheet.", error);
      setPanelMessage({
        body: localizedStrings.status.cssLintFailed,
        intent: "error",
      });
    }
  });
}

async function loadDroppedMarkdownFile({
  file,
  localizedStrings,
  setMarkdownInput,
  showComposeNotification,
}: Pick<
  TaskpaneActionDependencies,
  "localizedStrings" | "setMarkdownInput" | "showComposeNotification"
> & { file: File | null }): Promise<void> {
  if (file === null) {
    await showComposeNotification(
      "warning",
      localizedStrings.status.dropFileInstruction
    );
    return;
  }

  if (!supportsMarkdownFile(file)) {
    await showComposeNotification(
      "error",
      localizedStrings.status.unsupportedFileType
    );
    return;
  }

  try {
    const content = await readDroppedMarkdownFile(file);
    setMarkdownInput(normalizeMarkdownInput(content));
    await showComposeNotification(
      "success",
      localizedStrings.status.stylesheetLoaded(file.name)
    );
  } catch (error) {
    console.error("MarkOut failed to load a dropped file.", error);
    await showComposeNotification(
      "error",
      localizeActionError(localizedStrings, error)
    );
  }
}

function getInsertRenderedMarkdownMessage(
  strings: LocalizedStrings,
  result: InsertMarkdownResult
): string {
  return result === "replaced"
    ? strings.status.fragmentReplaced
    : strings.status.fragmentInserted;
}
