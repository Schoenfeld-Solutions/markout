import {
  Button,
  FluentProvider,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  makeStyles,
  mergeClasses,
  shorthands,
  tokens,
  webDarkTheme,
  webLightTheme,
} from "@fluentui/react-components";
import {
  type DragEvent,
  type ReactElement,
  useEffect,
  useEffectEvent,
  useRef,
  useState,
} from "react";
import {
  EMPTY_SELECTION_MESSAGE,
  FULL_DRAFT_ALREADY_RENDERED_MESSAGE,
  RENDERED_SELECTION_BLOCKED_MESSAGE,
  SUBJECT_SELECTION_UNSUPPORTED_MESSAGE,
} from "../lib/compose-markdown";
import { defaultStylesheet, type ThemeMode } from "../lib/config";
import {
  createInMemoryDiagnosticSink,
  getErrorDiagnosticMetadata,
  type DiagnosticEventInput,
} from "../lib/runtime";
import {
  lintStylesheet,
  type StylesheetLintResult,
} from "../lib/stylesheet-lint";
import {
  useAutoRenderNotificationController,
  usePreviewController,
  useSelectionStateController,
} from "./controllers";
import { useStylesheetEditor } from "./editor";
import { readDroppedMarkdownFile, supportsMarkdownFile } from "./file-drop";
import {
  getStrings,
  type LocalizedStrings,
  resolveLocale,
  resolveOfficeDisplayLanguage,
} from "./i18n";
import {
  CreditsPanel,
  DeveloperPanel,
  HelpPanel,
  InsertPanel,
  IntroPanel,
  SettingsPanel,
  renderActivePanel,
} from "./panels";
import { readPreferences, writePreferences } from "./preferences";
import { useResolvedColorMode } from "./theme";
import {
  buildToolbarPanels,
  getPanelAfterVisibilityChange,
  getRenderSelectionTooltip,
  isInsertRenderedMarkdownDisabled,
  isRenderSelectionDisabled,
  useToolbarLayoutMode,
  visibleToolbarPanelCount,
} from "./toolbar";
import type {
  PanelKey,
  PanelMessageState,
  PreferenceState,
  TaskpaneAppProps,
} from "./types";

const useStyles = makeStyles({
  appShell: {
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
    display: "grid",
    gridTemplateRows: "minmax(0, 1fr) auto",
    height: "100%",
    minHeight: 0,
    minWidth: 0,
    overflow: "hidden",
    width: "100%",
  },
  contentViewport: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    minHeight: 0,
    minWidth: 0,
    overflowX: "hidden",
    overflowY: "auto",
    paddingBlockEnd: tokens.spacingVerticalL,
    paddingBlockStart: tokens.spacingVerticalL,
    paddingInlineEnd: tokens.spacingHorizontalL,
    paddingInlineStart: tokens.spacingHorizontalL,
  },
  messageStack: {
    display: "grid",
    gap: tokens.spacingVerticalS,
    minWidth: 0,
  },
  panelRoot: {
    display: "grid",
    gap: tokens.spacingVerticalL,
    minWidth: 0,
  },
  sectionHeading: {
    display: "grid",
    gap: tokens.spacingVerticalXXS,
    minWidth: 0,
  },
  sectionTitle: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase500,
    margin: 0,
  },
  sectionBody: {
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    margin: 0,
  },
  card: {
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    backgroundColor: tokens.colorNeutralBackground2,
    display: "grid",
    gap: tokens.spacingVerticalM,
    minWidth: 0,
    paddingBlockEnd: tokens.spacingVerticalL,
    paddingBlockStart: tokens.spacingVerticalL,
    paddingInlineEnd: tokens.spacingHorizontalL,
    paddingInlineStart: tokens.spacingHorizontalL,
  },
  compactCard: {
    gap: tokens.spacingVerticalS,
  },
  dropzone: {
    ...shorthands.border("1px", "dashed", tokens.colorNeutralStroke2),
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    alignItems: "center",
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground2,
    display: "grid",
    gap: tokens.spacingVerticalS,
    justifyItems: "center",
    minHeight: "7.5rem",
    minWidth: 0,
    paddingBlockEnd: tokens.spacingVerticalL,
    paddingBlockStart: tokens.spacingVerticalL,
    paddingInlineEnd: tokens.spacingHorizontalL,
    paddingInlineStart: tokens.spacingHorizontalL,
    textAlign: "center",
  },
  dropzoneActive: {
    ...shorthands.borderColor(tokens.colorBrandStroke1),
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
  },
  dropzoneTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    margin: 0,
  },
  dropzoneCopy: {
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    margin: 0,
    maxWidth: "26rem",
  },
  textLabel: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    margin: 0,
  },
  textareaSurface: {
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    backgroundColor: tokens.colorNeutralBackground1,
    display: "grid",
    minHeight: "12rem",
    minWidth: 0,
    overflow: "hidden",
  },
  plainTextarea: {
    appearance: "none",
    backgroundColor: "transparent",
    border: "none",
    color: tokens.colorNeutralForeground1,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    minHeight: "12rem",
    minWidth: 0,
    outlineStyle: "none",
    overflowX: "hidden",
    overflowY: "auto",
    paddingBlockEnd: tokens.spacingVerticalM,
    paddingBlockStart: tokens.spacingVerticalM,
    paddingInlineEnd: tokens.spacingHorizontalM,
    paddingInlineStart: tokens.spacingHorizontalM,
    resize: "none",
    scrollbarWidth: "none",
    width: "100%",
    "&::-webkit-scrollbar": {
      display: "none",
    },
  },
  editorSurface: {
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    backgroundColor: tokens.colorNeutralBackground1,
    minHeight: "14rem",
    minWidth: 0,
    overflow: "hidden",
  },
  codeMirrorHost: {
    minHeight: "14rem",
    minWidth: 0,
  },
  codeMirrorLoading: {
    alignItems: "center",
    color: tokens.colorNeutralForeground3,
    display: "flex",
    justifyContent: "center",
    minHeight: "14rem",
  },
  previewFrame: {
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    backgroundColor: tokens.colorNeutralBackground1,
    minHeight: "12rem",
    minWidth: 0,
    overflowX: "hidden",
    overflowY: "auto",
    paddingBlockEnd: tokens.spacingVerticalL,
    paddingBlockStart: tokens.spacingVerticalL,
    paddingInlineEnd: tokens.spacingHorizontalL,
    paddingInlineStart: tokens.spacingHorizontalL,
  },
  previewFrameEmpty: {
    alignItems: "center",
    color: tokens.colorNeutralForeground3,
    display: "flex",
    justifyContent: "center",
    textAlign: "center",
  },
  previewContent: {
    color: tokens.colorNeutralForeground1,
    fontFamily: tokens.fontFamilyBase,
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    minWidth: 0,
    overflowWrap: "anywhere",
    wordBreak: "break-word",
    "& .markout-fragment-host, & .markout-fragment-host .mo": {
      color: "inherit",
      fontFamily: "inherit",
      fontSize: "inherit",
      lineHeight: "inherit",
    },
    "& a": {
      color: "inherit",
      textDecorationLine: "underline",
    },
    "& img": {
      maxWidth: "100%",
    },
    "& pre, & .hljs": {
      ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
      ...shorthands.borderRadius(tokens.borderRadiusMedium),
      backgroundColor: tokens.colorNeutralBackground3,
      color: "inherit",
      maxWidth: "100%",
      overflowX: "auto",
      whiteSpace: "pre-wrap",
      wordBreak: "break-word",
    },
    "& .markout-fragment-host .hljs-comment, & .markout-fragment-host .hljs-quote, & .markout-fragment-host .hljs-meta, & .markout-fragment-host .hljs-deletion":
      {
        color: tokens.colorNeutralForeground3,
      },
    "& .markout-fragment-host .hljs-string": {
      color: tokens.colorPaletteBerryForeground2,
    },
    "& .markout-fragment-host .hljs-variable, & .markout-fragment-host .hljs-template-variable, & .markout-fragment-host .hljs-symbol, & .markout-fragment-host .hljs-bullet, & .markout-fragment-host .hljs-section, & .markout-fragment-host .hljs-addition, & .markout-fragment-host .hljs-attribute, & .markout-fragment-host .hljs-link":
      {
        color: tokens.colorNeutralForeground2,
      },
    "& .markout-fragment-host .hljs-literal, & .markout-fragment-host .hljs-number":
      {
        color: tokens.colorBrandForeground1,
      },
    "& .markout-fragment-host .hljs-keyword, & .markout-fragment-host .hljs-selector-tag, & .markout-fragment-host .hljs-name, & .markout-fragment-host .hljs-type, & .markout-fragment-host .hljs-attr":
      {
        color: tokens.colorBrandForeground1,
      },
    "& code": {
      ...shorthands.borderRadius(tokens.borderRadiusSmall),
      backgroundColor: tokens.colorNeutralBackground3,
      color: "inherit",
      overflowWrap: "anywhere",
      padding: "0.08em 0.3em",
      wordBreak: "break-word",
    },
    "& pre code, & .hljs code": {
      backgroundColor: "transparent",
      padding: 0,
    },
    "& table": {
      display: "block",
      maxWidth: "100%",
      overflowX: "auto",
    },
  },
  actionRow: {
    display: "grid",
    gap: tokens.spacingHorizontalS,
    gridTemplateColumns: "repeat(3, minmax(0, 1fr))",
    minWidth: 0,
    "@media (max-width: 560px)": {
      gridTemplateColumns: "1fr",
    },
  },
  lintList: {
    display: "grid",
    gap: tokens.spacingVerticalXS,
    listStyleType: "none",
    margin: 0,
    padding: 0,
  },
  lintItem: {
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
    margin: 0,
    paddingBlockEnd: tokens.spacingVerticalXS,
    paddingBlockStart: tokens.spacingVerticalXS,
    paddingInlineEnd: tokens.spacingHorizontalS,
    paddingInlineStart: tokens.spacingHorizontalS,
  },
  lintItemError: {
    backgroundColor: tokens.colorPaletteRedBackground1,
    color: tokens.colorPaletteRedForeground1,
  },
  settingsRow: {
    alignItems: "center",
    display: "flex",
    gap: tokens.spacingHorizontalM,
    justifyContent: "space-between",
    minWidth: 0,
  },
  themeModeToolbar: {
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    backgroundColor: tokens.colorNeutralBackground1,
    display: "grid",
    padding: tokens.spacingHorizontalXXS,
    width: "100%",
  },
  themeModeToolbarGroup: {
    display: "grid",
    gap: tokens.spacingHorizontalXXS,
    gridTemplateColumns: "repeat(3, minmax(0, 1fr))",
    width: "100%",
  },
  themeModeToolbarButton: {
    justifyContent: "center",
    minWidth: 0,
    width: "100%",
  },
  linkList: {
    display: "grid",
    gap: tokens.spacingVerticalS,
    margin: 0,
    padding: 0,
  },
  linkCard: {
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    backgroundColor: tokens.colorNeutralBackground2,
    color: "inherit",
    display: "grid",
    gap: tokens.spacingVerticalXXS,
    minWidth: 0,
    paddingBlockEnd: tokens.spacingVerticalM,
    paddingBlockStart: tokens.spacingVerticalM,
    paddingInlineEnd: tokens.spacingHorizontalM,
    paddingInlineStart: tokens.spacingHorizontalM,
    textDecorationLine: "none",
  },
  linkCardHeader: {
    alignItems: "center",
    display: "flex",
    gap: tokens.spacingHorizontalS,
    minWidth: 0,
  },
  linkCardIcon: {
    color: tokens.colorBrandForeground1,
    flexShrink: 0,
  },
  introGrid: {
    display: "grid",
    gap: tokens.spacingHorizontalM,
    gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
    minWidth: 0,
    "@media (max-width: 640px)": {
      gridTemplateColumns: "1fr",
    },
  },
  introCard: {
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    backgroundColor: tokens.colorNeutralBackground2,
    display: "grid",
    gap: tokens.spacingVerticalS,
    minWidth: 0,
    paddingBlockEnd: tokens.spacingVerticalM,
    paddingBlockStart: tokens.spacingVerticalM,
    paddingInlineEnd: tokens.spacingHorizontalM,
    paddingInlineStart: tokens.spacingHorizontalM,
  },
  introIllustration: {
    color: tokens.colorBrandForeground1,
    height: "7rem",
    width: "100%",
  },
  creditsBox: {
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    backgroundColor: tokens.colorNeutralBackground2,
    display: "grid",
    gap: tokens.spacingVerticalS,
    minWidth: 0,
    paddingBlockEnd: tokens.spacingVerticalM,
    paddingBlockStart: tokens.spacingVerticalM,
    paddingInlineEnd: tokens.spacingHorizontalM,
    paddingInlineStart: tokens.spacingHorizontalM,
  },
  developerCode: {
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground2,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    margin: 0,
    minWidth: 0,
    overflowX: "auto",
    paddingBlockEnd: tokens.spacingVerticalS,
    paddingBlockStart: tokens.spacingVerticalS,
    paddingInlineEnd: tokens.spacingHorizontalS,
    paddingInlineStart: tokens.spacingHorizontalS,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },
  developerNoteList: {
    display: "grid",
    gap: tokens.spacingVerticalS,
    margin: 0,
    minWidth: 0,
    paddingInlineStart: tokens.spacingHorizontalL,
  },
  developerNoteItem: {
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    minWidth: 0,
  },
  inlineButtonRow: {
    display: "flex",
    flexWrap: "wrap",
    gap: tokens.spacingHorizontalS,
  },
  selectControl: {
    minWidth: 0,
    width: "100%",
  },
  toolbar: {
    ...shorthands.borderTop("1px", "solid", tokens.colorNeutralStroke2),
    alignItems: "stretch",
    backgroundColor: tokens.colorNeutralBackground1,
    display: "grid",
    gap: tokens.spacingHorizontalXXS,
    gridAutoColumns: "minmax(0, 1fr)",
    gridAutoFlow: "column",
    minWidth: 0,
    paddingBlockEnd: tokens.spacingVerticalXS,
    paddingBlockStart: tokens.spacingVerticalXS,
    paddingInlineEnd: tokens.spacingHorizontalS,
    paddingInlineStart: tokens.spacingHorizontalS,
  },
  toolbarButton: {
    justifyContent: "center",
    minWidth: 0,
  },
  toolbarButtonCompact: {
    paddingInlineEnd: tokens.spacingHorizontalXS,
    paddingInlineStart: tokens.spacingHorizontalXS,
  },
  toolbarLabel: {
    display: "block",
    fontSize: tokens.fontSizeBase100,
    lineHeight: tokens.lineHeightBase100,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
});

function TaskpaneContent({
  children,
}: {
  children: (styles: ReturnType<typeof useStyles>) => ReactElement;
}): ReactElement {
  const styles = useStyles();
  return children(styles);
}

export function TaskpaneApp({
  diagnosticSink,
  forcedToolbarLayoutMode,
  initialMarkdownInput = "",
  locale,
  notificationService,
  services,
  settingsStore,
}: TaskpaneAppProps): ReactElement {
  const diagnosticSinkRef = useRef(
    diagnosticSink ?? createInMemoryDiagnosticSink()
  );
  const [diagnosticEvents, setDiagnosticEvents] = useState(() =>
    diagnosticSinkRef.current.snapshot()
  );
  const [preferences, setPreferences] = useState<PreferenceState>(() =>
    readPreferences(settingsStore)
  );
  const resolvedLocale = resolveLocale(
    locale ?? resolveOfficeDisplayLanguage(),
    preferences.languagePreference
  );
  const localizedStrings = getStrings(resolvedLocale);
  const [activePanel, setActivePanel] = useState<PanelKey>(() =>
    preferences.introDismissed ? "insert" : "intro"
  );
  const [isDropActive, setIsDropActive] = useState(false);
  const [isWorking, setIsWorking] = useState<string | null>(null);
  const [markdownInput, setMarkdownInput] = useState(initialMarkdownInput);
  const [panelMessage, setPanelMessage] = useState<PanelMessageState | null>(
    null
  );
  const [cssLintResult, setCssLintResult] =
    useState<StylesheetLintResult | null>(null);
  const lastPersistedStylesheetRef = useRef(preferences.stylesheet);
  const migratedStylesheetSavedRef = useRef(false);
  const { mode: toolbarLayoutMode, ref: toolbarRef } = useToolbarLayoutMode(
    visibleToolbarPanelCount(preferences),
    forcedToolbarLayoutMode
  );
  const resolvedColorMode = useResolvedColorMode(preferences.themeMode);
  const currentTheme =
    resolvedColorMode === "dark" ? webDarkTheme : webLightTheme;
  const previewFrameStyle = {
    colorScheme: resolvedColorMode,
  } as const;

  const recordDiagnosticImplementationRef = useRef<
    (event: DiagnosticEventInput) => void
  >(() => undefined);
  recordDiagnosticImplementationRef.current = (event: DiagnosticEventInput) => {
    diagnosticSinkRef.current.record(event);
    setDiagnosticEvents(diagnosticSinkRef.current.snapshot());
  };
  const recordDiagnosticRef = useRef((event: DiagnosticEventInput) => {
    recordDiagnosticImplementationRef.current(event);
  });
  const recordDiagnostic = recordDiagnosticRef.current;

  const showComposeNotification = useEffectEvent(
    async (intent: PanelMessageState["intent"], message: string) => {
      if (notificationService === undefined) {
        console.warn(
          "MarkOut could not show a compose infobar because no notification service is available."
        );
        recordDiagnostic({
          area: "notification",
          code: "notification.transient.missing-service",
          level: "warning",
          metadata: { intent },
        });
        return;
      }

      const surface = await notificationService.showTransientNotification({
        intent,
        message,
      });

      if (surface === "pane") {
        console.warn(
          "MarkOut could not show the compose infobar and skipped the sidebar fallback."
        );
        recordDiagnostic({
          area: "notification",
          code: "notification.transient.fallback-pane",
          level: "warning",
          metadata: { intent },
        });
        return;
      }

      recordDiagnostic({
        area: "notification",
        code: "notification.transient.shown",
        level: "debug",
        metadata: { intent },
      });
    }
  );
  const handlePanelError = useEffectEvent((message: string) => {
    setPanelMessage({
      body: message,
      intent: "error",
    });
  });
  const handleStylesheetChange = useEffectEvent((stylesheet: string) => {
    setPreferences((currentPreferences) =>
      currentPreferences.stylesheet === stylesheet
        ? currentPreferences
        : {
            ...currentPreferences,
            stylesheet,
          }
    );
  });
  const { previewHtml, previewState } = usePreviewController(
    services.composeMarkdown,
    markdownInput,
    preferences.stylesheet,
    localizedStrings.status.previewFailed,
    handlePanelError,
    recordDiagnostic
  );
  const {
    isInspectingSelection,
    selectionState,
    setIsInspectingSelection,
    updateSelectionState,
  } = useSelectionStateController(
    services.composeMarkdown,
    activePanel,
    recordDiagnostic
  );
  const { dismissAutoRenderFallbackNotice, showAutoRenderFallbackNotice } =
    useAutoRenderNotificationController(
      notificationService,
      preferences.autoRender,
      localizedStrings.notifications.autoRenderStickyBody,
      recordDiagnostic
    );
  const { codeMirrorHostRef, isCodeMirrorLoading } = useStylesheetEditor(
    activePanel,
    preferences.stylesheet,
    cssLintResult,
    resolvedColorMode,
    handleStylesheetChange,
    handlePanelError,
    localizedStrings.editor.loadFailed
  );

  useEffect(() => {
    if (
      migratedStylesheetSavedRef.current ||
      !settingsStore.hasStylesheetMigrationPending()
    ) {
      return;
    }

    migratedStylesheetSavedRef.current = true;
    settingsStore.setStylesheet(preferences.stylesheet);

    void settingsStore.save().then(
      () => {
        lastPersistedStylesheetRef.current = settingsStore.getStylesheet();
      },
      (error: unknown) => {
        migratedStylesheetSavedRef.current = false;
        console.error(
          "MarkOut failed to persist the migrated default stylesheet.",
          error
        );
      }
    );
  }, [preferences.stylesheet, settingsStore]);

  useEffect(() => {
    if (preferences.stylesheet === lastPersistedStylesheetRef.current) {
      return;
    }

    const timeoutId = window.setTimeout(() => {
      settingsStore.setStylesheet(preferences.stylesheet);

      void settingsStore
        .save()
        .then(() => {
          lastPersistedStylesheetRef.current = settingsStore.getStylesheet();
        })
        .catch((error: unknown) => {
          console.error("MarkOut failed to persist stylesheet changes.", error);
          setPanelMessage({
            body: localizedStrings.status.stylesheetSaveFailed,
            intent: "error",
          });
        });
    }, 700);

    return () => {
      window.clearTimeout(timeoutId);
    };
  }, [
    localizedStrings.status.stylesheetSaveFailed,
    preferences.stylesheet,
    settingsStore,
  ]);

  useEffect(() => {
    setCssLintResult(null);
  }, [preferences.stylesheet]);

  async function withBusyState(
    busyKey: string,
    operation: () => void | Promise<void>
  ): Promise<void> {
    setIsWorking(busyKey);

    try {
      await operation();
    } finally {
      setIsWorking(null);
    }
  }

  async function persistPreferences(
    nextPreferences: PreferenceState
  ): Promise<boolean> {
    const previousPreferences = preferences;
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

  async function handleToggleAutoRender(enabled: boolean): Promise<void> {
    await persistPreferences({ ...preferences, autoRender: enabled });
  }

  async function handleToggleCreditsVisibility(
    visible: boolean
  ): Promise<void> {
    const didPersist = await persistPreferences({
      ...preferences,
      creditsVisible: visible,
    });

    if (didPersist && !visible && activePanel === "credits") {
      setActivePanel(
        getPanelAfterVisibilityChange(activePanel, "credits", visible)
      );
    }
  }

  async function handleToggleDeveloperTools(enabled: boolean): Promise<void> {
    const didPersist = await persistPreferences({
      ...preferences,
      developerToolsEnabled: enabled,
    });

    if (didPersist && !enabled && activePanel === "developer") {
      setActivePanel(
        getPanelAfterVisibilityChange(activePanel, "developer", enabled)
      );
    }
  }

  async function handleToggleHelpVisibility(visible: boolean): Promise<void> {
    const didPersist = await persistPreferences({
      ...preferences,
      helpVisible: visible,
    });

    if (didPersist && !visible && activePanel === "help") {
      setActivePanel(
        getPanelAfterVisibilityChange(activePanel, "help", visible)
      );
    }
  }

  async function handleToggleIntroVisibility(
    showIntro: boolean
  ): Promise<void> {
    const didPersist = await persistPreferences({
      ...preferences,
      introDismissed: !showIntro,
    });

    if (!didPersist) {
      return;
    }

    setActivePanel(showIntro ? "intro" : "insert");
  }

  async function handleThemeModeChange(mode: ThemeMode): Promise<void> {
    await persistPreferences({
      ...preferences,
      themeMode: mode,
    });
  }

  async function handleLanguagePreferenceChange(
    preference: PreferenceState["languagePreference"]
  ): Promise<void> {
    await persistPreferences({
      ...preferences,
      languagePreference: preference,
    });
  }

  async function handleConfirmIntro(): Promise<void> {
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
  }

  async function handleInspectSelection(): Promise<void> {
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

  async function handleInsertRenderedMarkdown(): Promise<void> {
    await withBusyState("insert-markdown", async () => {
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
          result === "replaced"
            ? localizedStrings.status.fragmentReplaced
            : localizedStrings.status.fragmentInserted
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

  async function handleRenderSelection(): Promise<void> {
    await withBusyState("render-selection", async () => {
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

  async function handleRenderEntireDraft(): Promise<void> {
    await withBusyState("render-entire-draft", async () => {
      recordDiagnostic({
        area: "render",
        code: "draft.render.started",
        level: "debug",
      });

      try {
        const result = await services.renderEntireDraft();
        await showComposeNotification(
          "success",
          result === "rendered"
            ? localizedStrings.status.draftRendered
            : localizedStrings.status.draftRestored
        );
        recordDiagnostic({
          area: result === "rendered" ? "render" : "restore",
          code:
            result === "rendered"
              ? "draft.render.succeeded"
              : "draft.restore.succeeded",
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
    });
  }

  async function handleLintStylesheet(): Promise<void> {
    await withBusyState("lint-stylesheet", () => {
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

  async function handleDrop(event: DragEvent<HTMLDivElement>): Promise<void> {
    setIsDropActive(false);
    const file = event.dataTransfer.files.item(0);

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
      setMarkdownInput(content);
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

  const toolbarPanels = buildToolbarPanels(preferences, localizedStrings);
  const renderSelectionDisabled = isRenderSelectionDisabled(
    isWorking !== null,
    selectionState.availability
  );
  const renderSelectionTooltip = getRenderSelectionTooltip(
    localizedStrings,
    selectionState.availability
  );

  return (
    <FluentProvider
      data-locale={resolvedLocale}
      data-theme={resolvedColorMode}
      id="taskpane-shell"
      theme={currentTheme}
    >
      <TaskpaneContent>
        {(styles) => (
          <div className={styles.appShell}>
            <main
              className={styles.contentViewport}
              data-testid="taskpane-content-viewport"
            >
              <div className={styles.messageStack}>
                {showAutoRenderFallbackNotice && preferences.autoRender ? (
                  <MessageBar intent="info">
                    <MessageBarBody>
                      <MessageBarTitle>
                        {localizedStrings.notifications.autoRenderFallbackTitle}
                      </MessageBarTitle>
                      {localizedStrings.notifications.autoRenderFallbackBody}
                    </MessageBarBody>
                    <div className={styles.inlineButtonRow}>
                      <Button
                        appearance="subtle"
                        onClick={() => {
                          void dismissAutoRenderFallbackNotice();
                        }}
                      >
                        {
                          localizedStrings.notifications
                            .autoRenderFallbackDismiss
                        }
                      </Button>
                    </div>
                  </MessageBar>
                ) : null}
                {panelMessage !== null ? (
                  <MessageBar intent={panelMessage.intent}>
                    <MessageBarBody>{panelMessage.body}</MessageBarBody>
                  </MessageBar>
                ) : null}
              </div>
              {renderActivePanel({
                activePanel,
                creditsPanel: (
                  <CreditsPanel strings={localizedStrings} styles={styles} />
                ),
                developerPanel: (
                  <DeveloperPanel
                    diagnosticEvents={diagnosticEvents}
                    isInspectingSelection={isInspectingSelection}
                    onInspectSelection={() => {
                      void handleInspectSelection();
                    }}
                    resolvedColorMode={resolvedColorMode}
                    selectionDebug={selectionState.debug}
                    strings={localizedStrings}
                    styles={styles}
                    themeMode={preferences.themeMode}
                  />
                ),
                helpPanel: (
                  <HelpPanel strings={localizedStrings} styles={styles} />
                ),
                insertPanel: (
                  <InsertPanel
                    isDropActive={isDropActive}
                    isInsertRenderedMarkdownDisabled={isInsertRenderedMarkdownDisabled(
                      isWorking !== null,
                      markdownInput
                    )}
                    isWorking={isWorking !== null}
                    markdownInput={markdownInput}
                    onDrop={(event) => {
                      void handleDrop(event);
                    }}
                    onInsertRenderedMarkdown={() => {
                      void handleInsertRenderedMarkdown();
                    }}
                    onMarkdownInputChange={setMarkdownInput}
                    onRenderEntireDraft={() => {
                      void handleRenderEntireDraft();
                    }}
                    onRenderSelection={() => {
                      void handleRenderSelection();
                    }}
                    previewHtml={previewHtml}
                    previewFrameStyle={previewFrameStyle}
                    previewState={previewState}
                    renderSelectionDisabled={renderSelectionDisabled}
                    renderSelectionTooltip={renderSelectionTooltip}
                    setDropActive={setIsDropActive}
                    strings={localizedStrings}
                    styles={styles}
                  />
                ),
                introPanel: (
                  <IntroPanel
                    onConfirm={() => {
                      void handleConfirmIntro();
                    }}
                    strings={localizedStrings}
                    styles={styles}
                  />
                ),
                settingsPanel: (
                  <SettingsPanel
                    autoRenderEnabled={preferences.autoRender}
                    codeMirrorHostRef={codeMirrorHostRef}
                    cssLintResult={cssLintResult}
                    developerToolsEnabled={preferences.developerToolsEnabled}
                    helpVisible={preferences.helpVisible}
                    introVisible={!preferences.introDismissed}
                    isCodeMirrorLoading={isCodeMirrorLoading}
                    isWorking={isWorking !== null}
                    languagePreference={preferences.languagePreference}
                    onCreditsVisibilityChange={(visible) => {
                      void handleToggleCreditsVisibility(visible);
                    }}
                    onDeveloperToolsChange={(enabled) => {
                      void handleToggleDeveloperTools(enabled);
                    }}
                    onHelpVisibilityChange={(visible) => {
                      void handleToggleHelpVisibility(visible);
                    }}
                    onIntroVisibilityChange={(visible) => {
                      void handleToggleIntroVisibility(visible);
                    }}
                    onLanguagePreferenceChange={(preference) => {
                      void handleLanguagePreferenceChange(preference);
                    }}
                    onLintStylesheet={() => {
                      void handleLintStylesheet();
                    }}
                    onResetStylesheet={() => {
                      setPreferences((currentPreferences) => ({
                        ...currentPreferences,
                        stylesheet: defaultStylesheet,
                      }));
                    }}
                    onThemeModeChange={(mode) => {
                      void handleThemeModeChange(mode);
                    }}
                    onToggleAutoRender={(enabled) => {
                      void handleToggleAutoRender(enabled);
                    }}
                    preferencesThemeMode={preferences.themeMode}
                    showCredits={preferences.creditsVisible}
                    strings={localizedStrings}
                    styles={styles}
                  />
                ),
              })}
            </main>
            <nav
              aria-label={localizedStrings.appTitle}
              className={styles.toolbar}
              data-testid="taskpane-toolbar"
              ref={toolbarRef}
            >
              {toolbarPanels.map((panel) => {
                const toolbarTitle =
                  toolbarLayoutMode === "compact"
                    ? localizedStrings.tooltips.toolbarCompactHint(panel.label)
                    : panel.label;

                return (
                  <Button
                    appearance={
                      activePanel === panel.key ? "primary" : "subtle"
                    }
                    aria-label={panel.label}
                    className={mergeClasses(
                      styles.toolbarButton,
                      toolbarLayoutMode === "compact"
                        ? styles.toolbarButtonCompact
                        : undefined
                    )}
                    icon={panel.icon}
                    id={`panel-button-${panel.key}`}
                    key={panel.key}
                    onClick={() => {
                      setActivePanel(panel.key);
                    }}
                    title={toolbarTitle}
                  >
                    {toolbarLayoutMode === "regular" ? (
                      <span className={styles.toolbarLabel}>{panel.label}</span>
                    ) : null}
                  </Button>
                );
              })}
            </nav>
          </div>
        )}
      </TaskpaneContent>
    </FluentProvider>
  );
}

function localizeActionError(
  strings: LocalizedStrings,
  error: unknown
): string {
  if (
    error instanceof Error &&
    error.message === SUBJECT_SELECTION_UNSUPPORTED_MESSAGE
  ) {
    return strings.tooltips.renderSelectionSubject;
  }

  if (error instanceof Error && error.message === EMPTY_SELECTION_MESSAGE) {
    return strings.tooltips.renderSelectionNoSelection;
  }

  if (
    error instanceof Error &&
    error.message === FULL_DRAFT_ALREADY_RENDERED_MESSAGE
  ) {
    return strings.tooltips.renderEntireDraft;
  }

  if (
    error instanceof Error &&
    error.message === RENDERED_SELECTION_BLOCKED_MESSAGE
  ) {
    return strings.tooltips.renderedFragmentBlocked;
  }

  if (error instanceof Error) {
    const readMatch = /^MarkOut could not read (.+)\.$/.exec(error.message);
    if (readMatch !== null) {
      const [, fileName = "the file"] = readMatch;
      return strings.status.fileReadFailed(fileName);
    }

    const decodeMatch = /^MarkOut could not decode (.+)\.$/.exec(error.message);
    if (decodeMatch !== null) {
      const [, fileName = "the file"] = decodeMatch;
      return strings.status.fileDecodeFailed(fileName);
    }

    return error.message;
  }

  return strings.status.unexpectedActionFailure;
}
