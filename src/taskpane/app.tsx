import {
  Button,
  FluentProvider,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Radio,
  RadioGroup,
  Switch,
  Tooltip,
  makeStyles,
  mergeClasses,
  shorthands,
  tokens,
  webDarkTheme,
  webLightTheme,
} from "@fluentui/react-components";
import hljs from "highlight.js/lib/core";
import cssLanguage from "highlight.js/lib/languages/css";
import {
  type ChangeEvent,
  type DragEvent,
  type RefObject,
  type ReactElement,
  type ReactNode,
  useDeferredValue,
  useEffect,
  useEffectEvent,
  useRef,
  useState,
} from "react";
import { type SelectionSource } from "../lib/body-accessor";
import {
  EMPTY_SELECTION_MESSAGE,
  FULL_DRAFT_ALREADY_RENDERED_MESSAGE,
  RENDERED_SELECTION_BLOCKED_MESSAGE,
  SUBJECT_SELECTION_UNSUPPORTED_MESSAGE,
  type ComposeMarkdownService,
} from "../lib/compose-markdown";
import type { ComposeNotificationService } from "../lib/compose-notifications";
import {
  defaultStylesheet,
  type SettingsStore,
  type ThemeMode,
} from "../lib/config";
import type { RenderItemResult } from "../lib/item";
import {
  lintStylesheet,
  type StylesheetLintResult,
} from "../lib/stylesheet-lint";
import {
  getStrings,
  resolveLocale,
  resolveOfficeDisplayLanguage,
  type LocalizedStrings,
  type SupportedLocale,
} from "./i18n";

hljs.registerLanguage("css", cssLanguage);

const DOCS_URL = "https://schoenfeld-solutions.github.io/markout/";
const REPOSITORY_URL = "https://github.com/Schoenfeld-Solutions/markout";
const STAR_URL = "https://github.com/Schoenfeld-Solutions/markout/stargazers";
const WEBSITE_URL = "https://schoenfeld.solutions";

const TOOLBAR_LABEL_MIN_WIDTH = 72;
const SELECTION_REFRESH_INTERVAL_MS = 1600;

const useStyles = makeStyles({
  appShell: {
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
    display: "grid",
    gridTemplateRows: "minmax(0, 1fr) auto",
    height: "100%",
    minHeight: 0,
    minWidth: 0,
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
  editorSurface: {
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    backgroundColor: tokens.colorNeutralBackground1,
    minHeight: "14rem",
    minWidth: 0,
    overflow: "hidden",
    position: "relative",
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
  codeMirror: {
    color: tokens.colorTransparentStroke,
    caretColor: tokens.colorNeutralForeground1,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase300,
    left: 0,
    lineHeight: tokens.lineHeightBase300,
    minHeight: "14rem",
    minWidth: 0,
    outlineStyle: "none",
    overflowX: "hidden",
    overflowY: "auto",
    paddingBlockEnd: tokens.spacingVerticalM,
    paddingBlockStart: tokens.spacingVerticalM,
    paddingInlineEnd: tokens.spacingHorizontalM,
    paddingInlineStart: tokens.spacingHorizontalM,
    position: "absolute",
    resize: "none",
    scrollbarWidth: "none",
    top: 0,
    width: "100%",
    zIndex: 1,
    "&::-webkit-scrollbar": {
      display: "none",
    },
  },
  codeHighlight: {
    color: tokens.colorNeutralForeground1,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase300,
    left: 0,
    lineHeight: tokens.lineHeightBase300,
    margin: 0,
    minHeight: "14rem",
    minWidth: 0,
    overflow: "hidden",
    paddingBlockEnd: tokens.spacingVerticalM,
    paddingBlockStart: tokens.spacingVerticalM,
    paddingInlineEnd: tokens.spacingHorizontalM,
    paddingInlineStart: tokens.spacingHorizontalM,
    pointerEvents: "none",
    position: "absolute",
    right: 0,
    top: 0,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    zIndex: 0,
    "& .hljs-comment": {
      color: tokens.colorNeutralForeground3,
    },
    "& .hljs-attribute, & .hljs-selector-class, & .hljs-selector-tag": {
      color: tokens.colorPaletteBlueForeground2,
    },
    "& .hljs-number, & .hljs-string": {
      color: tokens.colorPaletteGreenForeground1,
    },
    "& .hljs-keyword, & .hljs-literal, & .hljs-selector-pseudo": {
      color: tokens.colorPaletteBerryForeground2,
    },
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
    minWidth: 0,
    overflowWrap: "anywhere",
    wordBreak: "break-word",
    "& img": {
      maxWidth: "100%",
    },
    "& pre": {
      maxWidth: "100%",
      overflowX: "auto",
      whiteSpace: "pre-wrap",
      wordBreak: "break-word",
    },
    "& table": {
      display: "block",
      maxWidth: "100%",
      overflowX: "auto",
    },
    "& code": {
      overflowWrap: "anywhere",
      wordBreak: "break-word",
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
  radioGroup: {
    display: "grid",
    gap: tokens.spacingVerticalXS,
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
  inlineButtonRow: {
    display: "flex",
    flexWrap: "wrap",
    gap: tokens.spacingHorizontalS,
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

export type PanelKey =
  | "credits"
  | "developer"
  | "help"
  | "insert"
  | "intro"
  | "settings";
export type ToolbarLayoutMode = "compact" | "regular";

interface PreferenceState {
  autoRender: boolean;
  creditsVisible: boolean;
  developerToolsEnabled: boolean;
  helpVisible: boolean;
  introDismissed: boolean;
  stylesheet: string;
  themeMode: ThemeMode;
}

interface PanelMessageState {
  body: string;
  dismissible?: boolean;
  intent: "error" | "info" | "success" | "warning";
  title?: string;
}

interface SelectionDebugState {
  hasSelection: boolean;
  source: SelectionSource;
  textPreview: string;
}

type SelectionAvailability =
  | "body-none"
  | "body-selection"
  | "subject"
  | "unknown";

interface SelectionState {
  availability: SelectionAvailability;
  debug: SelectionDebugState | null;
}

export interface ToolbarPanelDescriptor {
  icon: ReactElement;
  key: PanelKey;
  label: string;
}

export interface TaskpaneServices {
  composeMarkdown: ComposeMarkdownService;
  renderEntireDraft(): Promise<RenderItemResult>;
}

export interface TaskpaneAppProps {
  forcedToolbarLayoutMode?: ToolbarLayoutMode;
  locale?: SupportedLocale;
  notificationService?: ComposeNotificationService;
  services: TaskpaneServices;
  settingsStore: SettingsStore;
  strings?: LocalizedStrings;
}

export function readDroppedMarkdownFile(file: File): Promise<string> {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();

    reader.onerror = () => {
      reject(new Error(`MarkOut could not read ${file.name}.`));
    };
    reader.onload = () => {
      if (typeof reader.result !== "string") {
        reject(new Error(`MarkOut could not decode ${file.name}.`));
        return;
      }

      resolve(reader.result);
    };

    reader.readAsText(file);
  });
}

export function supportsMarkdownFile(file: File): boolean {
  return /\.(md|markdown|txt)$/i.test(file.name);
}

export function TaskpaneApp({
  forcedToolbarLayoutMode,
  locale,
  notificationService,
  services,
  settingsStore,
  strings,
}: TaskpaneAppProps): ReactElement {
  const styles = useStyles();
  const resolvedLocale =
    locale ?? resolveLocale(resolveOfficeDisplayLanguage());
  const localizedStrings = strings ?? getStrings(resolvedLocale);
  const [preferences, setPreferences] = useState<PreferenceState>(() =>
    readPreferences(settingsStore)
  );
  const [activePanel, setActivePanel] = useState<PanelKey>(() =>
    preferences.introDismissed ? "insert" : "intro"
  );
  const [isDropActive, setIsDropActive] = useState(false);
  const [isInspectingSelection, setIsInspectingSelection] = useState(false);
  const [isWorking, setIsWorking] = useState<string | null>(null);
  const [markdownInput, setMarkdownInput] = useState("");
  const [panelMessage, setPanelMessage] = useState<PanelMessageState | null>(
    null
  );
  const [previewHtml, setPreviewHtml] = useState("");
  const [previewState, setPreviewState] = useState<
    "empty" | "loading" | "ready"
  >("empty");
  const [selectionState, setSelectionState] = useState<SelectionState>({
    availability: "unknown",
    debug: null,
  });
  const [showAutoRenderFallbackNotice, setShowAutoRenderFallbackNotice] =
    useState(false);
  const [cssLintResult, setCssLintResult] =
    useState<StylesheetLintResult | null>(null);
  const deferredMarkdownInput = useDeferredValue(markdownInput);
  const deferredStylesheet = useDeferredValue(preferences.stylesheet);
  const lastPersistedStylesheetRef = useRef(preferences.stylesheet);
  const previousAutoRenderRef = useRef(preferences.autoRender);
  const editorHighlightRef = useRef<HTMLPreElement | null>(null);
  const editorTextareaRef = useRef<HTMLTextAreaElement | null>(null);
  const { mode: toolbarLayoutMode, ref: toolbarRef } = useToolbarLayoutMode(
    visibleToolbarPanelCount(preferences),
    forcedToolbarLayoutMode
  );
  const resolvedColorMode = useResolvedColorMode(preferences.themeMode);
  const currentTheme =
    resolvedColorMode === "dark" ? webDarkTheme : webLightTheme;

  const updateSelectionState = useEffectEvent(async (): Promise<boolean> => {
    try {
      const selection = await services.composeMarkdown.getSelection();

      setSelectionState({
        availability:
          selection.source === "subject"
            ? "subject"
            : selection.hasSelection
              ? "body-selection"
              : "body-none",
        debug: {
          hasSelection: selection.hasSelection,
          source: selection.source,
          textPreview: selection.text.slice(0, 200),
        },
      });
      return true;
    } catch {
      setSelectionState((currentState) => ({
        availability: "unknown",
        debug: currentState.debug,
      }));
      return false;
    }
  });

  useEffect(() => {
    let ignore = false;

    if (deferredMarkdownInput.trim().length === 0) {
      setPreviewHtml("");
      setPreviewState("empty");
      return;
    }

    setPreviewState("loading");
    void services.composeMarkdown
      .renderPreview(deferredMarkdownInput, deferredStylesheet)
      .then((html) => {
        if (ignore) {
          return;
        }

        setPreviewHtml(html);
        setPreviewState("ready");
      })
      .catch((error: unknown) => {
        if (ignore) {
          return;
        }

        console.error("MarkOut failed to refresh the taskpane preview.", error);
        setPreviewHtml("");
        setPreviewState("empty");
        setPanelMessage({
          body: localizedStrings.status.previewFailed,
          intent: "error",
        });
      });

    return () => {
      ignore = true;
    };
  }, [
    deferredMarkdownInput,
    deferredStylesheet,
    localizedStrings.status.previewFailed,
    services.composeMarkdown,
  ]);

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
          setPanelMessage({
            body: localizedStrings.status.stylesheetSaved,
            intent: "success",
          });
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
    localizedStrings.status.stylesheetSaved,
    preferences.stylesheet,
    settingsStore,
  ]);

  useEffect(() => {
    setCssLintResult(null);
  }, [preferences.stylesheet]);

  useEffect(() => {
    if (activePanel !== "insert") {
      return;
    }

    const refreshSelection = () => {
      if (document.visibilityState === "hidden") {
        return;
      }

      void updateSelectionState();
    };

    refreshSelection();
    window.addEventListener("focus", refreshSelection);
    document.addEventListener("visibilitychange", refreshSelection);
    const intervalId = window.setInterval(
      refreshSelection,
      SELECTION_REFRESH_INTERVAL_MS
    );

    return () => {
      window.clearInterval(intervalId);
      window.removeEventListener("focus", refreshSelection);
      document.removeEventListener("visibilitychange", refreshSelection);
    };
  }, [activePanel, updateSelectionState]);

  useEffect(() => {
    if (notificationService === undefined) {
      return;
    }

    notificationService.onAutoRenderDismiss(() => {
      setShowAutoRenderFallbackNotice(false);
    });
  }, [notificationService]);

  useEffect(() => {
    if (notificationService === undefined) {
      return;
    }

    let cancelled = false;
    const isCancelled = () => cancelled;
    const wasEnabled = previousAutoRenderRef.current;
    previousAutoRenderRef.current = preferences.autoRender;

    void (async () => {
      if (!preferences.autoRender) {
        await notificationService.clearAutoRenderNotification();
        await notificationService.clearAutoRenderDismissed();
        if (!isCancelled()) {
          setShowAutoRenderFallbackNotice(false);
        }
        return;
      }

      if (!wasEnabled) {
        await notificationService.clearAutoRenderDismissed();
      }

      const dismissed = await notificationService.hasAutoRenderBeenDismissed();
      if (isCancelled() || dismissed) {
        setShowAutoRenderFallbackNotice(false);
        return;
      }

      const surface = await notificationService.showAutoRenderNotification({
        message: localizedStrings.notifications.autoRenderStickyBody,
      });

      if (!isCancelled()) {
        setShowAutoRenderFallbackNotice(surface === "pane");
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [
    localizedStrings.notifications.autoRenderStickyBody,
    notificationService,
    preferences.autoRender,
  ]);

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
    nextPreferences: PreferenceState,
    successMessage: string
  ): Promise<boolean> {
    const previousPreferences = preferences;
    setPreferences(nextPreferences);
    writePreferences(settingsStore, nextPreferences);

    try {
      await settingsStore.save();
      setPanelMessage({ body: successMessage, intent: "success" });
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
    await persistPreferences(
      { ...preferences, autoRender: enabled },
      enabled
        ? localizedStrings.status.autoRenderEnabled
        : localizedStrings.status.autoRenderDisabled
    );
  }

  async function handleToggleCreditsVisibility(
    visible: boolean
  ): Promise<void> {
    const didPersist = await persistPreferences(
      { ...preferences, creditsVisible: visible },
      visible
        ? localizedStrings.status.creditsShown
        : localizedStrings.status.creditsHidden
    );

    if (didPersist && !visible && activePanel === "credits") {
      setActivePanel(
        getPanelAfterVisibilityChange(activePanel, "credits", visible)
      );
    }
  }

  async function handleToggleDeveloperTools(enabled: boolean): Promise<void> {
    const didPersist = await persistPreferences(
      { ...preferences, developerToolsEnabled: enabled },
      enabled
        ? localizedStrings.status.developerEnabled
        : localizedStrings.status.developerDisabled
    );

    if (didPersist && !enabled && activePanel === "developer") {
      setActivePanel(
        getPanelAfterVisibilityChange(activePanel, "developer", enabled)
      );
    }
  }

  async function handleToggleHelpVisibility(visible: boolean): Promise<void> {
    const didPersist = await persistPreferences(
      { ...preferences, helpVisible: visible },
      visible
        ? localizedStrings.status.helpShown
        : localizedStrings.status.helpHidden
    );

    if (didPersist && !visible && activePanel === "help") {
      setActivePanel(
        getPanelAfterVisibilityChange(activePanel, "help", visible)
      );
    }
  }

  async function handleToggleIntroVisibility(
    showIntro: boolean
  ): Promise<void> {
    const didPersist = await persistPreferences(
      {
        ...preferences,
        introDismissed: !showIntro,
      },
      showIntro
        ? localizedStrings.status.introRestored
        : localizedStrings.status.introHidden
    );

    if (!didPersist) {
      return;
    }

    setActivePanel(showIntro ? "intro" : "insert");
  }

  async function handleThemeModeChange(mode: ThemeMode): Promise<void> {
    await persistPreferences(
      {
        ...preferences,
        themeMode: mode,
      },
      localizedStrings.status.themeUpdated(
        mode === "dark"
          ? localizedStrings.settings.themeModeDark
          : mode === "light"
            ? localizedStrings.settings.themeModeLight
            : localizedStrings.settings.themeModeSystem
      )
    );
  }

  async function handleConfirmIntro(): Promise<void> {
    if (preferences.introDismissed) {
      setActivePanel("insert");
      return;
    }

    const didPersist = await persistPreferences(
      {
        ...preferences,
        introDismissed: true,
      },
      localizedStrings.status.introHidden
    );

    if (didPersist) {
      setActivePanel("insert");
    }
  }

  async function handleInspectSelection(): Promise<void> {
    setIsInspectingSelection(true);

    try {
      const loaded = await updateSelectionState();
      setPanelMessage({
        body: loaded
          ? localizedStrings.status.selectionInspectionSuccess
          : localizedStrings.status.selectionInspectionFailed,
        intent: loaded ? "success" : "error",
      });
    } catch (error) {
      console.error("MarkOut failed to inspect the current selection.", error);
      setPanelMessage({
        body: localizedStrings.status.selectionInspectionFailed,
        intent: "error",
      });
    } finally {
      setIsInspectingSelection(false);
    }
  }

  async function handleInsertRenderedMarkdown(): Promise<void> {
    await withBusyState("insert-markdown", async () => {
      try {
        const result =
          await services.composeMarkdown.insertRenderedMarkdown(markdownInput);
        setPanelMessage({
          body:
            result === "replaced"
              ? localizedStrings.status.fragmentReplaced
              : localizedStrings.status.fragmentInserted,
          intent: "success",
        });
        await updateSelectionState();
      } catch (error) {
        console.error("MarkOut failed to insert rendered Markdown.", error);
        setPanelMessage({
          body: localizeActionError(localizedStrings, error),
          intent: "error",
        });
      }
    });
  }

  async function handleRenderSelection(): Promise<void> {
    await withBusyState("render-selection", async () => {
      try {
        await services.composeMarkdown.renderSelection();
        setPanelMessage({
          body: localizedStrings.status.selectionRendered,
          intent: "success",
        });
        await updateSelectionState();
      } catch (error) {
        console.error("MarkOut failed to render the current selection.", error);
        setPanelMessage({
          body: localizeActionError(localizedStrings, error),
          intent: "error",
        });
        await updateSelectionState();
      }
    });
  }

  async function handleRenderEntireDraft(): Promise<void> {
    await withBusyState("render-entire-draft", async () => {
      try {
        const result = await services.renderEntireDraft();
        setPanelMessage({
          body:
            result === "rendered"
              ? localizedStrings.status.draftRendered
              : localizedStrings.status.draftRestored,
          intent: "success",
        });
        await updateSelectionState();
      } catch (error) {
        console.error("MarkOut failed to render the current draft.", error);
        setPanelMessage({
          body: localizeActionError(localizedStrings, error),
          intent: "error",
        });
      }
    });
  }

  async function handleLintStylesheet(): Promise<void> {
    await withBusyState("lint-stylesheet", () => {
      try {
        setCssLintResult(lintStylesheet(preferences.stylesheet));
        setPanelMessage({
          body: localizedStrings.status.cssLintComplete,
          intent: "success",
        });
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
    event.preventDefault();
    setIsDropActive(false);
    const file = event.dataTransfer.files.item(0);

    if (file === null) {
      setPanelMessage({
        body: localizedStrings.status.dropFileInstruction,
        intent: "warning",
      });
      return;
    }

    if (!supportsMarkdownFile(file)) {
      setPanelMessage({
        body: localizedStrings.status.unsupportedFileType,
        intent: "error",
      });
      return;
    }

    try {
      const content = await readDroppedMarkdownFile(file);
      setMarkdownInput(content);
      setPanelMessage({
        body: localizedStrings.status.stylesheetLoaded(file.name),
        intent: "success",
      });
    } catch (error) {
      console.error("MarkOut failed to load a dropped file.", error);
      setPanelMessage({
        body: localizeActionError(localizedStrings, error),
        intent: "error",
      });
    }
  }

  function handleMarkdownInputChange(
    event: ChangeEvent<HTMLTextAreaElement>
  ): void {
    setMarkdownInput(event.target.value);
  }

  function handleStylesheetInputChange(
    event: ChangeEvent<HTMLTextAreaElement>
  ): void {
    setPreferences((currentPreferences) => ({
      ...currentPreferences,
      stylesheet: event.target.value,
    }));
  }

  function handleStylesheetScroll(): void {
    if (
      editorHighlightRef.current === null ||
      editorTextareaRef.current === null
    ) {
      return;
    }

    editorHighlightRef.current.scrollTop = editorTextareaRef.current.scrollTop;
    editorHighlightRef.current.scrollLeft =
      editorTextareaRef.current.scrollLeft;
  }

  async function handleDismissAutoRenderFallbackNotice(): Promise<void> {
    setShowAutoRenderFallbackNotice(false);
    await notificationService?.markAutoRenderDismissed();
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

  function renderActionMessage(): ReactElement | null {
    if (panelMessage === null) {
      return null;
    }

    return (
      <MessageBar intent={panelMessage.intent}>
        <MessageBarBody>
          {panelMessage.title !== undefined ? (
            <MessageBarTitle>{panelMessage.title}</MessageBarTitle>
          ) : null}
          {panelMessage.body}
        </MessageBarBody>
      </MessageBar>
    );
  }

  function renderAutoRenderFallbackMessage(): ReactElement | null {
    if (!showAutoRenderFallbackNotice || !preferences.autoRender) {
      return null;
    }

    return (
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
              void handleDismissAutoRenderFallbackNotice();
            }}
          >
            {localizedStrings.notifications.autoRenderFallbackDismiss}
          </Button>
        </div>
      </MessageBar>
    );
  }

  function renderPreview(): ReactElement {
    if (previewState === "loading") {
      return (
        <div
          className={mergeClasses(
            styles.previewFrame,
            styles.previewFrameEmpty
          )}
        >
          {localizedStrings.insert.previewLoading}
        </div>
      );
    }

    if (previewHtml.trim().length === 0) {
      return (
        <div
          className={mergeClasses(
            styles.previewFrame,
            styles.previewFrameEmpty
          )}
        >
          {localizedStrings.insert.emptyPreview}
        </div>
      );
    }

    return (
      <div id="mo-preview" className={styles.previewFrame} aria-live="polite">
        <div
          className={styles.previewContent}
          dangerouslySetInnerHTML={{ __html: previewHtml }}
        />
      </div>
    );
  }

  function renderInsertPanel(): ReactElement {
    return (
      <div className={styles.panelRoot}>
        <div className={styles.sectionHeading}>
          <h2 className={styles.sectionTitle}>
            {localizedStrings.insert.panelTitle}
          </h2>
          <p className={styles.sectionBody}>
            {localizedStrings.insert.panelDescription}
          </p>
        </div>
        <div
          className={mergeClasses(
            styles.dropzone,
            isDropActive ? styles.dropzoneActive : undefined
          )}
          data-testid="taskpane-dropzone"
          onDragEnter={() => setIsDropActive(true)}
          onDragLeave={() => setIsDropActive(false)}
          onDragOver={(event) => {
            event.preventDefault();
            setIsDropActive(true);
          }}
          onDrop={(event) => {
            void handleDrop(event);
          }}
        >
          <InsertIcon />
          <p className={styles.dropzoneTitle}>
            {localizedStrings.insert.dropzoneTitle}
          </p>
          <p className={styles.dropzoneCopy}>
            {localizedStrings.insert.dropzoneCopy}
          </p>
        </div>
        <div className={styles.card}>
          <div className={styles.sectionHeading}>
            <label className={styles.textLabel} htmlFor="markdown-input">
              {localizedStrings.insert.inputLabel}
            </label>
          </div>
          <div className={styles.textareaSurface}>
            <textarea
              className={styles.plainTextarea}
              id="markdown-input"
              onChange={handleMarkdownInputChange}
              placeholder={localizedStrings.insert.inputPlaceholder}
              spellCheck={false}
              value={markdownInput}
            />
          </div>
        </div>
        <div className={styles.card}>
          <div className={styles.sectionHeading}>
            <h3 className={styles.sectionTitle}>
              {localizedStrings.insert.previewTitle}
            </h3>
            <p className={styles.sectionBody}>
              {localizedStrings.insert.previewDescription}
            </p>
          </div>
          {renderPreview()}
          <div className={styles.actionRow}>
            <Tooltip
              content={renderSelectionTooltip}
              relationship="description"
            >
              <Button
                appearance="primary"
                aria-label={localizedStrings.insert.renderSelectionButton}
                disabled={renderSelectionDisabled}
                id="render-selection-button"
                onClick={() => {
                  void handleRenderSelection();
                }}
                title={renderSelectionTooltip}
              >
                {localizedStrings.insert.renderSelectionButton}
              </Button>
            </Tooltip>
            <Tooltip
              content={localizedStrings.tooltips.renderEntireDraft}
              relationship="description"
            >
              <Button
                appearance="secondary"
                disabled={isWorking !== null}
                id="render-entire-draft-button"
                onClick={() => {
                  void handleRenderEntireDraft();
                }}
                title={localizedStrings.tooltips.renderEntireDraft}
              >
                {localizedStrings.insert.renderEntireDraftButton}
              </Button>
            </Tooltip>
            <Tooltip
              content={localizedStrings.tooltips.insertRenderedMarkdown}
              relationship="description"
            >
              <Button
                appearance="secondary"
                disabled={isInsertRenderedMarkdownDisabled(
                  isWorking !== null,
                  markdownInput
                )}
                id="insert-rendered-markdown-button"
                onClick={() => {
                  void handleInsertRenderedMarkdown();
                }}
                title={localizedStrings.tooltips.insertRenderedMarkdown}
              >
                {localizedStrings.insert.insertButton}
              </Button>
            </Tooltip>
          </div>
        </div>
      </div>
    );
  }

  function renderSettingsPanel(): ReactElement {
    return (
      <div className={styles.panelRoot}>
        <div className={styles.sectionHeading}>
          <h2 className={styles.sectionTitle}>
            {localizedStrings.settings.panelTitle}
          </h2>
          <p className={styles.sectionBody}>
            {localizedStrings.settings.panelDescription}
          </p>
        </div>
        <div className={styles.card}>
          <h3 className={styles.sectionTitle}>
            {localizedStrings.settings.themeTitle}
          </h3>
          <p className={styles.sectionBody}>
            {localizedStrings.settings.themeDescription}
          </p>
          <RadioGroup
            className={styles.radioGroup}
            onChange={(_, data) => {
              void handleThemeModeChange(data.value as ThemeMode);
            }}
            value={preferences.themeMode}
          >
            <Radio
              label={localizedStrings.settings.themeModeLight}
              value="light"
            />
            <Radio
              label={localizedStrings.settings.themeModeDark}
              value="dark"
            />
            <Radio
              label={localizedStrings.settings.themeModeSystem}
              value="system"
            />
          </RadioGroup>
        </div>
        <div className={styles.card}>
          <div className={styles.settingsRow}>
            <div className={styles.sectionHeading}>
              <h3 className={styles.sectionTitle}>
                {localizedStrings.settings.autoRenderTitle}
              </h3>
              <p className={styles.sectionBody}>
                {localizedStrings.settings.autoRenderDescription}
              </p>
            </div>
            <Switch
              checked={preferences.autoRender}
              id="autorender-switch"
              label={
                preferences.autoRender
                  ? localizedStrings.general.on
                  : localizedStrings.general.off
              }
              onChange={(_, data) => {
                void handleToggleAutoRender(data.checked);
              }}
            />
          </div>
          <div className={styles.settingsRow}>
            <div className={styles.sectionHeading}>
              <h3 className={styles.sectionTitle}>
                {localizedStrings.settings.introTitle}
              </h3>
              <p className={styles.sectionBody}>
                {localizedStrings.settings.introDescription}
              </p>
            </div>
            <Switch
              checked={!preferences.introDismissed}
              id="show-intro-switch"
              label={
                !preferences.introDismissed
                  ? localizedStrings.general.shown
                  : localizedStrings.general.hidden
              }
              onChange={(_, data) => {
                void handleToggleIntroVisibility(data.checked);
              }}
            />
          </div>
          <div className={styles.settingsRow}>
            <div className={styles.sectionHeading}>
              <h3 className={styles.sectionTitle}>
                {localizedStrings.settings.helpTitle}
              </h3>
              <p className={styles.sectionBody}>
                {localizedStrings.settings.helpDescription}
              </p>
            </div>
            <Switch
              checked={preferences.helpVisible}
              id="show-help-switch"
              label={
                preferences.helpVisible
                  ? localizedStrings.general.shown
                  : localizedStrings.general.hidden
              }
              onChange={(_, data) => {
                void handleToggleHelpVisibility(data.checked);
              }}
            />
          </div>
          <div className={styles.settingsRow}>
            <div className={styles.sectionHeading}>
              <h3 className={styles.sectionTitle}>
                {localizedStrings.settings.creditsTitle}
              </h3>
              <p className={styles.sectionBody}>
                {localizedStrings.settings.creditsDescription}
              </p>
            </div>
            <Switch
              checked={preferences.creditsVisible}
              id="show-credits-switch"
              label={
                preferences.creditsVisible
                  ? localizedStrings.general.shown
                  : localizedStrings.general.hidden
              }
              onChange={(_, data) => {
                void handleToggleCreditsVisibility(data.checked);
              }}
            />
          </div>
          <div className={styles.settingsRow}>
            <div className={styles.sectionHeading}>
              <h3 className={styles.sectionTitle}>
                {localizedStrings.settings.developerTitle}
              </h3>
              <p className={styles.sectionBody}>
                {localizedStrings.settings.developerDescription}
              </p>
            </div>
            <Switch
              checked={preferences.developerToolsEnabled}
              id="developer-tools-switch"
              label={
                preferences.developerToolsEnabled
                  ? localizedStrings.general.shown
                  : localizedStrings.general.hidden
              }
              onChange={(_, data) => {
                void handleToggleDeveloperTools(data.checked);
              }}
            />
          </div>
        </div>
        <div className={styles.card}>
          <div className={styles.sectionHeading}>
            <h3 className={styles.sectionTitle}>
              {localizedStrings.editor.title}
            </h3>
            <p className={styles.sectionBody}>
              {localizedStrings.localization.supportedLanguagesNote}
            </p>
          </div>
          <div className={styles.editorSurface}>
            <pre
              aria-hidden="true"
              className={styles.codeHighlight}
              ref={editorHighlightRef}
            >
              <code
                dangerouslySetInnerHTML={{
                  __html: highlightCss(preferences.stylesheet),
                }}
              />
            </pre>
            <textarea
              className={styles.codeMirror}
              id="theme-editor"
              onChange={handleStylesheetInputChange}
              onScroll={handleStylesheetScroll}
              ref={editorTextareaRef}
              spellCheck={false}
              value={preferences.stylesheet}
            />
          </div>
          <div className={styles.inlineButtonRow}>
            <Button
              appearance="secondary"
              disabled={isWorking !== null}
              id="lint-stylesheet-button"
              onClick={() => {
                void handleLintStylesheet();
              }}
            >
              {localizedStrings.editor.lintButton}
            </Button>
            <Button
              appearance="secondary"
              onClick={() => {
                setPreferences((currentPreferences) => ({
                  ...currentPreferences,
                  stylesheet: defaultStylesheet,
                }));
              }}
            >
              {localizedStrings.editor.resetButton}
            </Button>
          </div>
          {renderLintResult(styles, localizedStrings, cssLintResult)}
        </div>
      </div>
    );
  }

  function renderHelpPanel(): ReactElement {
    return (
      <div className={styles.panelRoot}>
        <div className={styles.sectionHeading}>
          <h2 className={styles.sectionTitle}>
            {localizedStrings.help.panelTitle}
          </h2>
          <p className={styles.sectionBody}>
            {localizedStrings.help.panelDescription}
          </p>
        </div>
        <div className={styles.linkList}>
          <a
            className={styles.linkCard}
            href={REPOSITORY_URL}
            rel="noreferrer"
            target="_blank"
          >
            <strong>{localizedStrings.help.repoTitle}</strong>
            <p className={styles.sectionBody}>
              {localizedStrings.help.repoDescription}
            </p>
          </a>
          <a
            className={styles.linkCard}
            href={DOCS_URL}
            rel="noreferrer"
            target="_blank"
          >
            <strong>{localizedStrings.help.docsTitle}</strong>
            <p className={styles.sectionBody}>
              {localizedStrings.help.docsDescription}
            </p>
          </a>
          <a
            className={styles.linkCard}
            href={WEBSITE_URL}
            rel="noreferrer"
            target="_blank"
          >
            <strong>{localizedStrings.help.websiteTitle}</strong>
            <p className={styles.sectionBody}>
              {localizedStrings.help.websiteDescription}
            </p>
          </a>
          <div className={styles.linkCard}>
            <strong>{localizedStrings.buyMeACoffeePlaceholder}</strong>
            <p className={styles.sectionBody}>
              {localizedStrings.localization.buyMeACoffeePlaceholderDescription}
            </p>
          </div>
        </div>
      </div>
    );
  }

  function renderIntroPanel(): ReactElement {
    return (
      <div className={styles.panelRoot}>
        <div className={styles.sectionHeading}>
          <h2 className={styles.sectionTitle}>
            {localizedStrings.intro.panelTitle}
          </h2>
          <p className={styles.sectionBody}>
            {localizedStrings.intro.panelDescription}
          </p>
        </div>
        <div className={styles.introGrid}>
          <div className={styles.introCard}>
            <div className={styles.introIllustration}>
              <IntroComposeIllustration />
            </div>
            <h3 className={styles.sectionTitle}>
              {localizedStrings.intro.stepOneTitle}
            </h3>
            <p className={styles.sectionBody}>
              {localizedStrings.intro.stepOneBody}
            </p>
          </div>
          <div className={styles.introCard}>
            <div className={styles.introIllustration}>
              <IntroInsertIllustration />
            </div>
            <h3 className={styles.sectionTitle}>
              {localizedStrings.intro.stepTwoTitle}
            </h3>
            <p className={styles.sectionBody}>
              {localizedStrings.intro.stepTwoBody}
            </p>
          </div>
        </div>
        <div className={styles.inlineButtonRow}>
          <Button
            appearance="primary"
            id="intro-confirm-button"
            onClick={() => {
              void handleConfirmIntro();
            }}
          >
            {localizedStrings.intro.confirm}
          </Button>
        </div>
      </div>
    );
  }

  function renderCreditsPanel(): ReactElement {
    return (
      <div className={styles.panelRoot}>
        <div className={styles.sectionHeading}>
          <h2 className={styles.sectionTitle}>
            {localizedStrings.credits.panelTitle}
          </h2>
          <p className={styles.sectionBody}>
            {localizedStrings.credits.panelDescription}
          </p>
        </div>
        <div className={styles.creditsBox}>
          <h3 className={styles.sectionTitle}>
            {localizedStrings.credits.upstreamTitle}
          </h3>
          <p className={styles.sectionBody}>
            {localizedStrings.credits.upstreamBody}
          </p>
        </div>
        <div className={styles.creditsBox}>
          <h3 className={styles.sectionTitle}>
            {localizedStrings.credits.currentMaintenanceTitle}
          </h3>
          <p className={styles.sectionBody}>
            {localizedStrings.credits.currentMaintenanceBody}
          </p>
        </div>
        <div className={styles.inlineButtonRow}>
          <Button as="a" href={REPOSITORY_URL} rel="noreferrer" target="_blank">
            {localizedStrings.credits.openFork}
          </Button>
          <Button as="a" href={STAR_URL} rel="noreferrer" target="_blank">
            {localizedStrings.credits.starFork}
          </Button>
        </div>
      </div>
    );
  }

  function renderDeveloperPanel(): ReactElement {
    return (
      <div className={styles.panelRoot}>
        <div className={styles.sectionHeading}>
          <h2 className={styles.sectionTitle}>
            {localizedStrings.developer.panelTitle}
          </h2>
          <p className={styles.sectionBody}>
            {localizedStrings.developer.panelDescription}
          </p>
        </div>
        <div className={mergeClasses(styles.card, styles.compactCard)}>
          <div className={styles.settingsRow}>
            <div className={styles.sectionHeading}>
              <h3 className={styles.sectionTitle}>
                {localizedStrings.developer.hostNotesTitle}
              </h3>
              <p className={styles.sectionBody}>
                {localizedStrings.developer.resolvedTheme
                  .replace("{mode}", preferences.themeMode)
                  .replace("{resolvedMode}", resolvedColorMode)}
              </p>
            </div>
            <Button
              appearance="secondary"
              disabled={isInspectingSelection}
              onClick={() => {
                void handleInspectSelection();
              }}
            >
              {localizedStrings.developer.inspectSelection}
            </Button>
          </div>
          <pre className={styles.developerCode}>
            {selectionState.debug === null
              ? localizedStrings.developer.noSelectionSnapshot
              : JSON.stringify(selectionState.debug, null, 2)}
          </pre>
          <ul className={styles.linkList}>
            <li className={styles.sectionBody}>
              {localizedStrings.developer.subjectHint}
            </li>
            <li className={styles.sectionBody}>
              {localizedStrings.developer.ribbonHint}
            </li>
            <li className={styles.sectionBody}>
              {localizedStrings.developer.taskpaneHint}
            </li>
          </ul>
        </div>
      </div>
    );
  }

  function renderActivePanel(): ReactElement {
    switch (activePanel) {
      case "credits":
        return renderCreditsPanel();
      case "developer":
        return renderDeveloperPanel();
      case "help":
        return renderHelpPanel();
      case "intro":
        return renderIntroPanel();
      case "settings":
        return renderSettingsPanel();
      default:
        return renderInsertPanel();
    }
  }

  return (
    <FluentProvider
      className={styles.appShell}
      data-locale={resolvedLocale}
      data-theme={resolvedColorMode}
      id="taskpane-shell"
      theme={currentTheme}
    >
      <main className={styles.contentViewport}>
        <div className={styles.messageStack}>
          {renderAutoRenderFallbackMessage()}
          {renderActionMessage()}
        </div>
        {renderActivePanel()}
      </main>
      <nav
        aria-label={localizedStrings.appTitle}
        className={styles.toolbar}
        ref={toolbarRef}
      >
        {toolbarPanels.map((panel) => {
          const toolbarTitle =
            toolbarLayoutMode === "compact"
              ? localizedStrings.tooltips.toolbarCompactHint(panel.label)
              : panel.label;

          return (
            <Tooltip
              content={toolbarTitle}
              key={panel.key}
              relationship="description"
            >
              <Button
                appearance={activePanel === panel.key ? "primary" : "subtle"}
                aria-label={panel.label}
                className={mergeClasses(
                  styles.toolbarButton,
                  toolbarLayoutMode === "compact"
                    ? styles.toolbarButtonCompact
                    : undefined
                )}
                icon={panel.icon}
                id={`panel-button-${panel.key}`}
                onClick={() => {
                  setActivePanel(panel.key);
                }}
                title={toolbarTitle}
              >
                {toolbarLayoutMode === "regular" ? (
                  <span className={styles.toolbarLabel}>{panel.label}</span>
                ) : null}
              </Button>
            </Tooltip>
          );
        })}
      </nav>
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
    return resolvedFragmentBlockMessage(strings);
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

function renderLintResult(
  styles: ReturnType<typeof useStyles>,
  strings: LocalizedStrings,
  lintResult: StylesheetLintResult | null
): ReactElement | null {
  if (lintResult === null) {
    return null;
  }

  if (lintResult.issues.length === 0) {
    return (
      <MessageBar intent="success">
        <MessageBarBody>{strings.editor.lintNoIssues}</MessageBarBody>
      </MessageBar>
    );
  }

  return (
    <ul className={styles.lintList}>
      {lintResult.issues.map((issue, index) => (
        <li
          className={mergeClasses(
            styles.lintItem,
            issue.severity === "error" ? styles.lintItemError : undefined
          )}
          key={`${issue.code}-${index}`}
        >
          <strong>
            {issue.severity === "error"
              ? strings.editor.lintErrorLabel
              : strings.editor.lintWarningLabel}
            :
          </strong>{" "}
          {issue.message}
        </li>
      ))}
    </ul>
  );
}

function readPreferences(settingsStore: SettingsStore): PreferenceState {
  return {
    autoRender: settingsStore.getAutoRender(),
    creditsVisible: settingsStore.getCreditsVisible(),
    developerToolsEnabled: settingsStore.getDeveloperToolsEnabled(),
    helpVisible: settingsStore.getHelpVisible(),
    introDismissed: settingsStore.getIntroDismissed(),
    stylesheet: settingsStore.getStylesheet(),
    themeMode: settingsStore.getThemeMode(),
  };
}

function writePreferences(
  settingsStore: SettingsStore,
  preferences: PreferenceState
): void {
  settingsStore.setAutoRender(preferences.autoRender);
  settingsStore.setCreditsVisible(preferences.creditsVisible);
  settingsStore.setDeveloperToolsEnabled(preferences.developerToolsEnabled);
  settingsStore.setHelpVisible(preferences.helpVisible);
  settingsStore.setIntroDismissed(preferences.introDismissed);
  settingsStore.setStylesheet(preferences.stylesheet);
  settingsStore.setThemeMode(preferences.themeMode);
}

export function buildToolbarPanels(
  preferences: PreferenceState,
  strings: LocalizedStrings
): ToolbarPanelDescriptor[] {
  const panels: ToolbarPanelDescriptor[] = [];

  if (!preferences.introDismissed) {
    panels.push({
      icon: <InfoIcon />,
      key: "intro",
      label: strings.toolbar.intro,
    });
  }

  panels.push({
    icon: <InsertIcon />,
    key: "insert",
    label: strings.toolbar.insert,
  });
  panels.push({
    icon: <SettingsIcon />,
    key: "settings",
    label: strings.toolbar.settings,
  });

  if (preferences.helpVisible) {
    panels.push({
      icon: <HelpIcon />,
      key: "help",
      label: strings.toolbar.help,
    });
  }

  if (preferences.developerToolsEnabled) {
    panels.push({
      icon: <DeveloperIcon />,
      key: "developer",
      label: strings.toolbar.developer,
    });
  }

  if (preferences.creditsVisible) {
    panels.push({
      icon: <CreditsIcon />,
      key: "credits",
      label: strings.toolbar.credits,
    });
  }

  return panels;
}

export function visibleToolbarPanelCount(preferences: PreferenceState): number {
  return buildToolbarPanels(preferences, getStrings("en-US")).length;
}

export function getRenderSelectionTooltip(
  strings: LocalizedStrings,
  availability: SelectionAvailability
): string {
  switch (availability) {
    case "body-selection":
      return strings.tooltips.renderSelection;
    case "body-none":
      return localizeActionError(strings, new Error(EMPTY_SELECTION_MESSAGE));
    case "subject":
      return localizeActionError(
        strings,
        new Error(SUBJECT_SELECTION_UNSUPPORTED_MESSAGE)
      );
    default:
      return strings.tooltips.renderSelectionUnknown;
  }
}

export function isRenderSelectionDisabled(
  isBusy: boolean,
  availability: SelectionAvailability
): boolean {
  return isBusy || availability !== "body-selection";
}

export function isInsertRenderedMarkdownDisabled(
  isBusy: boolean,
  markdownInput: string
): boolean {
  return isBusy || markdownInput.trim().length === 0;
}

export function getPanelAfterVisibilityChange(
  activePanel: PanelKey,
  changedPanel: "credits" | "developer" | "help",
  visible: boolean
): PanelKey {
  return !visible && activePanel === changedPanel ? "settings" : activePanel;
}

export function resolveToolbarLayoutMode(
  availableWidth: number,
  itemCount: number
): ToolbarLayoutMode {
  return availableWidth / Math.max(itemCount, 1) >= TOOLBAR_LABEL_MIN_WIDTH
    ? "regular"
    : "compact";
}

function resolvedFragmentBlockMessage(strings: LocalizedStrings): string {
  return strings.tooltips.renderedFragmentBlocked;
}

function highlightCss(stylesheet: string): string {
  const source = stylesheet.length > 0 ? stylesheet : " ";

  try {
    return hljs.highlight(source, { language: "css" }).value;
  } catch {
    return escapeHtml(source);
  }
}

function escapeHtml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function useToolbarLayoutMode(
  itemCount: number,
  forcedMode: ToolbarLayoutMode | undefined
): { mode: ToolbarLayoutMode; ref: RefObject<HTMLElement | null> } {
  const ref = useRef<HTMLElement | null>(null);
  const [mode, setMode] = useState<ToolbarLayoutMode>(forcedMode ?? "regular");

  useEffect(() => {
    if (forcedMode !== undefined) {
      setMode(forcedMode);
      return;
    }

    const updateMode = () => {
      const availableWidth = ref.current?.clientWidth ?? window.innerWidth;
      const nextMode = resolveToolbarLayoutMode(availableWidth, itemCount);
      setMode(nextMode);
    };

    updateMode();

    if (typeof ResizeObserver === "function" && ref.current !== null) {
      const resizeObserver = new ResizeObserver(updateMode);
      resizeObserver.observe(ref.current);

      return () => {
        resizeObserver.disconnect();
      };
    }

    window.addEventListener("resize", updateMode);

    return () => {
      window.removeEventListener("resize", updateMode);
    };
  }, [forcedMode, itemCount]);

  return { mode, ref };
}

function useResolvedColorMode(themeMode: ThemeMode): "dark" | "light" {
  const [systemColorMode, setSystemColorMode] = useState<"dark" | "light">(() =>
    resolveSystemColorMode()
  );

  useEffect(() => {
    const mediaQuery =
      typeof window.matchMedia === "function"
        ? window.matchMedia("(prefers-color-scheme: dark)")
        : null;
    const mailbox =
      typeof Office === "undefined" ? null : Office.context.mailbox;
    const updateSystemColorMode = (
      officeTheme?: Partial<Office.OfficeTheme>
    ) => {
      setSystemColorMode(resolveSystemColorMode(officeTheme));
    };

    updateSystemColorMode();

    const handleMediaChange = () => {
      updateSystemColorMode();
    };

    mediaQuery?.addEventListener("change", handleMediaChange);

    if (mailbox !== null && typeof mailbox.addHandlerAsync === "function") {
      mailbox.addHandlerAsync(
        Office.EventType.OfficeThemeChanged,
        (event: Office.OfficeThemeChangedEventArgs) => {
          updateSystemColorMode(event.officeTheme);
        }
      );
    }

    return () => {
      mediaQuery?.removeEventListener("change", handleMediaChange);
    };
  }, []);

  return themeMode === "system" ? systemColorMode : themeMode;
}

export function resolveSystemColorMode(
  officeTheme: Partial<Office.OfficeTheme> | undefined = readOfficeTheme()
): "dark" | "light" {
  const officeThemeColor = officeTheme?.bodyBackgroundColor;

  if (typeof officeThemeColor === "string" && officeThemeColor.length > 0) {
    return isDarkColor(officeThemeColor) ? "dark" : "light";
  }

  if (
    typeof window.matchMedia === "function" &&
    window.matchMedia("(prefers-color-scheme: dark)").matches
  ) {
    return "dark";
  }

  return "light";
}

function readOfficeTheme(): Partial<Office.OfficeTheme> | undefined {
  if (typeof Office === "undefined") {
    return undefined;
  }

  const mailboxWithTheme = Office.context.mailbox as Office.Mailbox & {
    officeTheme?: Partial<Office.OfficeTheme>;
  };

  return mailboxWithTheme.officeTheme;
}

export function isDarkColor(color: string): boolean {
  const normalizedColor = normalizeHexColor(color);

  if (normalizedColor === null) {
    return false;
  }

  const red = Number.parseInt(normalizedColor.slice(0, 2), 16);
  const green = Number.parseInt(normalizedColor.slice(2, 4), 16);
  const blue = Number.parseInt(normalizedColor.slice(4, 6), 16);
  const luminance = (0.2126 * red + 0.7152 * green + 0.0722 * blue) / 255;

  return luminance < 0.55;
}

function normalizeHexColor(color: string): string | null {
  const trimmedColor = color.trim().replace(/^#/, "");

  if (/^[0-9a-f]{6}$/i.test(trimmedColor)) {
    return trimmedColor;
  }

  if (/^[0-9a-f]{3}$/i.test(trimmedColor)) {
    return trimmedColor
      .split("")
      .map((segment) => segment.repeat(2))
      .join("");
  }

  return null;
}

function ToolbarIcon({ children }: { children: ReactNode }): ReactElement {
  return (
    <svg
      aria-hidden="true"
      fill="none"
      height="18"
      stroke="currentColor"
      strokeLinecap="round"
      strokeLinejoin="round"
      strokeWidth="1.6"
      viewBox="0 0 20 20"
      width="18"
    >
      {children}
    </svg>
  );
}

function InsertIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <path d="M10 3v9" />
      <path d="M7 9.5 10 12.5l3-3" />
      <path d="M4 15.5h12" />
    </ToolbarIcon>
  );
}

function InfoIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <circle cx="10" cy="10" r="6.5" />
      <path d="M10 8v4" />
      <path d="M10 6.2h.01" />
    </ToolbarIcon>
  );
}

function HelpIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <path d="M7.8 7.8a2.2 2.2 0 1 1 3.5 1.7c-.8.6-1.3 1-1.3 2" />
      <path d="M10 14.3h.01" />
      <circle cx="10" cy="10" r="6.5" />
    </ToolbarIcon>
  );
}

function CreditsIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <path d="m10 4.3 1.5 3.1 3.4.5-2.5 2.5.6 3.5-3-1.6-3 1.6.6-3.5L5 7.9l3.5-.5Z" />
    </ToolbarIcon>
  );
}

function SettingsIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <circle cx="10" cy="10" r="2.2" />
      <path d="M10 4.5v1.4" />
      <path d="M10 14.1v1.4" />
      <path d="m5.8 5.8 1 1" />
      <path d="m13.2 13.2 1 1" />
      <path d="M4.5 10h1.4" />
      <path d="M14.1 10h1.4" />
      <path d="m5.8 14.2 1-1" />
      <path d="m13.2 6.8 1-1" />
    </ToolbarIcon>
  );
}

function DeveloperIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <path d="m7.5 6-3 4 3 4" />
      <path d="m12.5 6 3 4-3 4" />
      <path d="m11 5-2 10" />
    </ToolbarIcon>
  );
}

function IntroComposeIllustration(): ReactElement {
  return (
    <svg
      aria-hidden="true"
      fill="none"
      height="100%"
      viewBox="0 0 240 140"
      width="100%"
    >
      <rect
        fill="currentColor"
        fillOpacity="0.08"
        height="104"
        rx="18"
        width="208"
        x="16"
        y="18"
      />
      <rect
        fill="currentColor"
        fillOpacity="0.14"
        height="72"
        rx="14"
        width="84"
        x="30"
        y="34"
      />
      <rect
        fill="currentColor"
        fillOpacity="0.1"
        height="10"
        rx="5"
        width="84"
        x="128"
        y="38"
      />
      <rect
        fill="currentColor"
        fillOpacity="0.1"
        height="10"
        rx="5"
        width="66"
        x="128"
        y="58"
      />
      <rect
        fill="currentColor"
        fillOpacity="0.1"
        height="10"
        rx="5"
        width="74"
        x="128"
        y="78"
      />
      <rect
        fill="currentColor"
        fillOpacity="0.16"
        height="16"
        rx="8"
        width="48"
        x="128"
        y="100"
      />
    </svg>
  );
}

function IntroInsertIllustration(): ReactElement {
  return (
    <svg
      aria-hidden="true"
      fill="none"
      height="100%"
      viewBox="0 0 240 140"
      width="100%"
    >
      <rect
        fill="currentColor"
        fillOpacity="0.08"
        height="104"
        rx="18"
        width="208"
        x="16"
        y="18"
      />
      <rect
        fill="currentColor"
        fillOpacity="0.1"
        height="62"
        rx="12"
        width="120"
        x="32"
        y="38"
      />
      <path
        d="M92 50v18"
        stroke="currentColor"
        strokeLinecap="round"
        strokeWidth="6"
      />
      <path
        d="m82 60 10 10 10-10"
        stroke="currentColor"
        strokeLinecap="round"
        strokeLinejoin="round"
        strokeWidth="6"
      />
      <rect
        fill="currentColor"
        fillOpacity="0.16"
        height="62"
        rx="12"
        width="52"
        x="164"
        y="38"
      />
      <rect
        fill="currentColor"
        fillOpacity="0.08"
        height="10"
        rx="5"
        width="32"
        x="174"
        y="52"
      />
      <rect
        fill="currentColor"
        fillOpacity="0.08"
        height="10"
        rx="5"
        width="24"
        x="174"
        y="72"
      />
    </svg>
  );
}
