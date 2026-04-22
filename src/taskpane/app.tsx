import {
  Button,
  FluentProvider,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Select,
  Switch,
  Toolbar,
  ToolbarRadioButton,
  ToolbarRadioGroup,
  makeStyles,
  mergeClasses,
  shorthands,
  tokens,
  webDarkTheme,
  webLightTheme,
} from "@fluentui/react-components";
import type { Diagnostic } from "@codemirror/lint";
import type { HighlightStyle, LanguageSupport } from "@codemirror/language";
import type {
  EditorState as CodeMirrorEditorState,
  Extension,
  TransactionSpec,
} from "@codemirror/state";
import type {
  EditorView as CodeMirrorEditorView,
  ViewUpdate as CodeMirrorViewUpdate,
} from "@codemirror/view";
import type { StyleSpec } from "style-mod";
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
  type LanguagePreference,
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
  type LocalizedStrings,
  resolveLocale,
  resolveOfficeDisplayLanguage,
  type SupportedLocale,
} from "./i18n";

const DOCS_URL = "https://schoenfeld-solutions.github.io/markout/";
const REPOSITORY_URL = "https://github.com/Schoenfeld-Solutions/markout";
const STAR_URL = "https://github.com/Schoenfeld-Solutions/markout/stargazers";
const WEBSITE_URL = "https://schoenfeld.solutions";

const TOOLBAR_LABEL_MIN_WIDTH = 72;
const SELECTION_REFRESH_INTERVAL_MS = 1600;

interface CodeMirrorEditorStateConstructor {
  create(config: {
    doc: string;
    extensions: readonly Extension[];
  }): CodeMirrorEditorState;
}

interface CodeMirrorEditorViewConstructor {
  new (config: {
    parent: Element | DocumentFragment;
    state: CodeMirrorEditorState;
  }): CodeMirrorEditorView;
  lineWrapping: Extension;
  theme(
    spec: Record<string, StyleSpec>,
    options?: { dark?: boolean }
  ): Extension;
  updateListener: {
    of(listener: (update: CodeMirrorViewUpdate) => void): Extension;
  };
}

interface CodeMirrorModules {
  css: () => LanguageSupport;
  defaultHighlightStyle: HighlightStyle;
  EditorState: CodeMirrorEditorStateConstructor;
  EditorView: CodeMirrorEditorViewConstructor;
  lineNumbers: () => Extension;
  setDiagnostics: (
    state: CodeMirrorEditorState,
    diagnostics: readonly Diagnostic[]
  ) => TransactionSpec;
  syntaxHighlighting: (
    highlighter: HighlightStyle,
    options?: { fallback: boolean }
  ) => Extension;
}

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
  languagePreference: LanguagePreference;
  stylesheet: string;
  themeMode: ThemeMode;
}

interface PanelMessageState {
  body: string;
  dismissible?: boolean;
  intent: "error" | "info" | "success" | "warning";
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
}

async function loadCodeMirrorModules(): Promise<CodeMirrorModules> {
  const [cssModule, languageModule, lintModule, stateModule, viewModule] =
    await Promise.all([
      import("@codemirror/lang-css"),
      import("@codemirror/language"),
      import("@codemirror/lint"),
      import("@codemirror/state"),
      import("@codemirror/view"),
    ]);

  return {
    css: cssModule.css,
    defaultHighlightStyle: languageModule.defaultHighlightStyle,
    EditorState: stateModule.EditorState,
    EditorView: viewModule.EditorView,
    lineNumbers: viewModule.lineNumbers,
    setDiagnostics: lintModule.setDiagnostics,
    syntaxHighlighting: languageModule.syntaxHighlighting,
  };
}

function findLintIssueRange(
  stylesheet: string,
  issue: StylesheetLintResult["issues"][number]
): { from: number; to: number } {
  const normalizedStylesheet = stylesheet.length > 0 ? stylesheet : " ";
  const selectorMatch = /"([^"]+)"/.exec(issue.message);

  if (selectorMatch !== null) {
    const selectorText = selectorMatch[1];

    if (selectorText === undefined) {
      return {
        from: 0,
        to: normalizedStylesheet.length,
      };
    }

    const index = normalizedStylesheet.indexOf(selectorText);

    if (index !== -1) {
      return {
        from: index,
        to: index + selectorText.length,
      };
    }
  }

  const propertyMatch = /The property "([^"]+)"/.exec(issue.message);

  if (propertyMatch !== null) {
    const propertyName = propertyMatch[1];

    if (propertyName === undefined) {
      return {
        from: 0,
        to: normalizedStylesheet.length,
      };
    }

    const index = normalizedStylesheet.indexOf(propertyName);

    if (index !== -1) {
      return {
        from: index,
        to: index + propertyName.length,
      };
    }
  }

  if (issue.code === "empty-stylesheet") {
    return { from: 0, to: 0 };
  }

  return {
    from: 0,
    to: normalizedStylesheet.length,
  };
}

function toCodeMirrorDiagnostics(
  stylesheet: string,
  lintResult: StylesheetLintResult | null
): Diagnostic[] {
  if (lintResult === null) {
    return [];
  }

  return lintResult.issues.map((issue) => {
    const { from, to } = findLintIssueRange(stylesheet, issue);

    return {
      from,
      message: issue.message,
      severity: issue.severity === "error" ? "error" : "warning",
      to,
    };
  });
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
}: TaskpaneAppProps): ReactElement {
  const styles = useStyles();
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
  const [isCodeMirrorLoading, setIsCodeMirrorLoading] = useState(false);
  const deferredMarkdownInput = useDeferredValue(markdownInput);
  const deferredStylesheet = useDeferredValue(preferences.stylesheet);
  const lastPersistedStylesheetRef = useRef(preferences.stylesheet);
  const previousAutoRenderRef = useRef(preferences.autoRender);
  const migratedStylesheetSavedRef = useRef(false);
  const codeMirrorHostRef = useRef<HTMLDivElement | null>(null);
  const codeMirrorModulesRef = useRef<CodeMirrorModules | null>(null);
  const codeMirrorViewRef = useRef<CodeMirrorEditorView | null>(null);
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

  const showComposeNotification = useEffectEvent(
    async (intent: PanelMessageState["intent"], message: string) => {
      if (notificationService === undefined) {
        setPanelMessage({ body: message, intent });
        return;
      }

      const surface = await notificationService.showTransientNotification({
        intent,
        message,
      });

      if (surface === "pane") {
        setPanelMessage({ body: message, intent });
      }
    }
  );

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

  useEffect(() => {
    if (activePanel !== "settings" || codeMirrorHostRef.current === null) {
      return;
    }

    let cancelled = false;
    let editorView: CodeMirrorEditorView | null = null;

    setIsCodeMirrorLoading(true);

    void loadCodeMirrorModules()
      .then((modules) => {
        if (cancelled || codeMirrorHostRef.current === null) {
          return;
        }

        codeMirrorModulesRef.current = modules;
        const editorTheme = modules.EditorView.theme(
          {
            "&": {
              backgroundColor:
                resolvedColorMode === "dark" ? "transparent" : "transparent",
              color:
                resolvedColorMode === "dark"
                  ? tokens.colorNeutralForeground1
                  : tokens.colorNeutralForeground1,
              fontFamily: tokens.fontFamilyMonospace,
              fontSize: tokens.fontSizeBase300,
              minHeight: "14rem",
            },
            ".cm-scroller": {
              fontFamily: tokens.fontFamilyMonospace,
              lineHeight: tokens.lineHeightBase300,
              minHeight: "14rem",
            },
            ".cm-content": {
              caretColor: tokens.colorNeutralForeground1,
              minHeight: "14rem",
              padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalM}`,
            },
            ".cm-gutters": {
              backgroundColor:
                resolvedColorMode === "dark"
                  ? tokens.colorNeutralBackground1
                  : tokens.colorNeutralBackground1,
              borderRightColor: tokens.colorNeutralStroke2,
              color: tokens.colorNeutralForeground3,
            },
            ".cm-activeLine": {
              backgroundColor:
                resolvedColorMode === "dark"
                  ? "rgba(255, 255, 255, 0.04)"
                  : "rgba(15, 108, 189, 0.06)",
            },
            ".cm-activeLineGutter": {
              backgroundColor:
                resolvedColorMode === "dark"
                  ? "rgba(255, 255, 255, 0.04)"
                  : "rgba(15, 108, 189, 0.06)",
            },
            ".cm-selectionBackground": {
              backgroundColor:
                resolvedColorMode === "dark"
                  ? "rgba(96, 165, 250, 0.28) !important"
                  : "rgba(15, 108, 189, 0.24) !important",
            },
            ".cm-diagnostic": {
              fontFamily: tokens.fontFamilyBase,
            },
          },
          { dark: resolvedColorMode === "dark" }
        );
        const updateListener = modules.EditorView.updateListener.of(
          (update: CodeMirrorViewUpdate) => {
            if (!update.docChanged) {
              return;
            }

            const nextStylesheet = update.state.doc.toString();
            setPreferences((currentPreferences) =>
              currentPreferences.stylesheet === nextStylesheet
                ? currentPreferences
                : {
                    ...currentPreferences,
                    stylesheet: nextStylesheet,
                  }
            );
          }
        );

        editorView = new modules.EditorView({
          parent: codeMirrorHostRef.current,
          state: modules.EditorState.create({
            doc: preferences.stylesheet,
            extensions: [
              modules.lineNumbers(),
              modules.css(),
              modules.EditorView.lineWrapping,
              modules.syntaxHighlighting(modules.defaultHighlightStyle, {
                fallback: true,
              }),
              editorTheme,
              updateListener,
            ],
          }),
        });

        codeMirrorViewRef.current = editorView;
        setIsCodeMirrorLoading(false);
      })
      .catch((error: unknown) => {
        console.error(
          "MarkOut failed to initialize the stylesheet editor.",
          error
        );
        codeMirrorModulesRef.current = null;
        codeMirrorViewRef.current = null;
        setIsCodeMirrorLoading(false);
        if (!cancelled) {
          setPanelMessage({
            body: localizedStrings.editor.loadFailed,
            intent: "error",
          });
        }
      });

    return () => {
      cancelled = true;
      setIsCodeMirrorLoading(false);
      editorView?.destroy();
      codeMirrorViewRef.current = null;
    };
  }, [activePanel, localizedStrings.editor.loadFailed, resolvedColorMode]);

  useEffect(() => {
    const editorView = codeMirrorViewRef.current;

    if (editorView === null) {
      return;
    }

    const currentDocument = editorView.state.doc.toString();

    if (currentDocument === preferences.stylesheet) {
      return;
    }

    editorView.dispatch({
      changes: {
        from: 0,
        insert: preferences.stylesheet,
        to: currentDocument.length,
      },
    });
  }, [preferences.stylesheet]);

  useEffect(() => {
    const editorView = codeMirrorViewRef.current;
    const modules = codeMirrorModulesRef.current;

    if (editorView === null || modules === null) {
      return;
    }

    editorView.dispatch(
      modules.setDiagnostics(
        editorView.state,
        toCodeMirrorDiagnostics(preferences.stylesheet, cssLintResult)
      )
    );
  }, [cssLintResult, preferences.stylesheet]);

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
    preference: LanguagePreference
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
        await showComposeNotification(
          "success",
          result === "replaced"
            ? localizedStrings.status.fragmentReplaced
            : localizedStrings.status.fragmentInserted
        );
        await updateSelectionState();
      } catch (error) {
        console.error("MarkOut failed to insert rendered Markdown.", error);
        await showComposeNotification(
          "error",
          localizeActionError(localizedStrings, error)
        );
      }
    });
  }

  async function handleRenderSelection(): Promise<void> {
    await withBusyState("render-selection", async () => {
      try {
        await services.composeMarkdown.renderSelection();
        await showComposeNotification(
          "success",
          localizedStrings.status.selectionRendered
        );
        await updateSelectionState();
      } catch (error) {
        console.error("MarkOut failed to render the current selection.", error);
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
      try {
        const result = await services.renderEntireDraft();
        await showComposeNotification(
          "success",
          result === "rendered"
            ? localizedStrings.status.draftRendered
            : localizedStrings.status.draftRestored
        );
        await updateSelectionState();
      } catch (error) {
        console.error("MarkOut failed to render the current draft.", error);
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
    event.preventDefault();
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

  function handleMarkdownInputChange(
    event: ChangeEvent<HTMLTextAreaElement>
  ): void {
    setMarkdownInput(event.target.value);
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
        <MessageBarBody>{panelMessage.body}</MessageBarBody>
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

  function renderOptionalBody(copy: string): ReactElement | null {
    return copy.trim().length > 0 ? (
      <p className={styles.sectionBody}>{copy}</p>
    ) : null;
  }

  function renderPreview(): ReactElement {
    if (previewState === "loading") {
      return (
        <div
          className={mergeClasses(
            styles.previewFrame,
            styles.previewFrameEmpty
          )}
          style={previewFrameStyle}
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
          style={previewFrameStyle}
        >
          {localizedStrings.insert.emptyPreview}
        </div>
      );
    }

    return (
      <div
        id="mo-preview"
        className={styles.previewFrame}
        aria-live="polite"
        style={previewFrameStyle}
      >
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
          {renderOptionalBody(localizedStrings.insert.panelDescription)}
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
            {renderOptionalBody(localizedStrings.insert.previewDescription)}
          </div>
          {renderPreview()}
          <div className={styles.actionRow}>
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
          {renderOptionalBody(localizedStrings.settings.panelDescription)}
        </div>
        <div className={styles.card}>
          <h3 className={styles.sectionTitle}>
            {localizedStrings.settings.themeTitle}
          </h3>
          {renderOptionalBody(localizedStrings.settings.themeDescription)}
          <Toolbar
            aria-label={localizedStrings.settings.themeTitle}
            checkedValues={{ "theme-mode": [preferences.themeMode] }}
            className={styles.themeModeToolbar}
            onCheckedValueChange={(_, data) => {
              const nextMode = data.checkedItems[0];

              if (
                data.name === "theme-mode" &&
                (nextMode === "light" ||
                  nextMode === "dark" ||
                  nextMode === "system")
              ) {
                void handleThemeModeChange(nextMode);
              }
            }}
          >
            <ToolbarRadioGroup className={styles.themeModeToolbarGroup}>
              <ToolbarRadioButton
                appearance="subtle"
                className={styles.themeModeToolbarButton}
                id="theme-mode-light"
                name="theme-mode"
                value="light"
              >
                {localizedStrings.settings.themeModeLight}
              </ToolbarRadioButton>
              <ToolbarRadioButton
                appearance="subtle"
                className={styles.themeModeToolbarButton}
                id="theme-mode-dark"
                name="theme-mode"
                value="dark"
              >
                {localizedStrings.settings.themeModeDark}
              </ToolbarRadioButton>
              <ToolbarRadioButton
                appearance="subtle"
                className={styles.themeModeToolbarButton}
                id="theme-mode-system"
                name="theme-mode"
                value="system"
              >
                {localizedStrings.settings.themeModeSystem}
              </ToolbarRadioButton>
            </ToolbarRadioGroup>
          </Toolbar>
        </div>
        <div className={styles.card}>
          <h3 className={styles.sectionTitle}>
            {localizedStrings.settings.languageTitle}
          </h3>
          {renderOptionalBody(localizedStrings.settings.languageDescription)}
          <Select
            className={styles.selectControl}
            id="language-preference-select"
            onChange={(event) => {
              void handleLanguagePreferenceChange(
                event.currentTarget.value as LanguagePreference
              );
            }}
            value={preferences.languagePreference}
          >
            <option value="system">
              {localizedStrings.settings.languageSystem}
            </option>
            <option value="en-US">
              {localizedStrings.settings.languageEnglish}
            </option>
            <option value="de-DE">
              {localizedStrings.settings.languageGerman}
            </option>
          </Select>
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
          </div>
          <div className={styles.editorSurface}>
            {isCodeMirrorLoading ? (
              <div className={styles.codeMirrorLoading}>
                {localizedStrings.editor.loading}
              </div>
            ) : null}
            <div
              className={styles.codeMirrorHost}
              id="theme-editor"
              ref={codeMirrorHostRef}
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
          {renderOptionalBody(localizedStrings.help.panelDescription)}
        </div>
        <div className={styles.linkList}>
          <a
            className={styles.linkCard}
            href={REPOSITORY_URL}
            rel="noreferrer"
            target="_blank"
          >
            <div className={styles.linkCardHeader}>
              <span className={styles.linkCardIcon}>
                <RepositoryIcon />
              </span>
              <strong>{localizedStrings.help.repoTitle}</strong>
            </div>
            {renderOptionalBody(localizedStrings.help.repoDescription)}
          </a>
          <a
            className={styles.linkCard}
            href={DOCS_URL}
            rel="noreferrer"
            target="_blank"
          >
            <div className={styles.linkCardHeader}>
              <span className={styles.linkCardIcon}>
                <DocsIcon />
              </span>
              <strong>{localizedStrings.help.docsTitle}</strong>
            </div>
            {renderOptionalBody(localizedStrings.help.docsDescription)}
          </a>
          <a
            className={styles.linkCard}
            href={WEBSITE_URL}
            rel="noreferrer"
            target="_blank"
          >
            <div className={styles.linkCardHeader}>
              <span className={styles.linkCardIcon}>
                <CompanyIcon />
              </span>
              <strong>{localizedStrings.help.websiteTitle}</strong>
            </div>
            {renderOptionalBody(localizedStrings.help.websiteDescription)}
          </a>
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
          {renderOptionalBody(localizedStrings.intro.panelDescription)}
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
          {renderOptionalBody(localizedStrings.credits.panelDescription)}
        </div>
        <div className={styles.creditsBox}>
          <div className={styles.linkCardHeader}>
            <span className={styles.linkCardIcon}>
              <UpstreamIcon />
            </span>
            <h3 className={styles.sectionTitle}>
              {localizedStrings.credits.upstreamTitle}
            </h3>
          </div>
          <p className={styles.sectionBody}>
            {localizedStrings.credits.upstreamBody}
          </p>
        </div>
        <div className={styles.creditsBox}>
          <div className={styles.linkCardHeader}>
            <span className={styles.linkCardIcon}>
              <ForkIcon />
            </span>
            <h3 className={styles.sectionTitle}>
              {localizedStrings.credits.currentMaintenanceTitle}
            </h3>
          </div>
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
          {renderOptionalBody(localizedStrings.developer.panelDescription)}
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
          <ul className={styles.developerNoteList}>
            <li className={styles.developerNoteItem}>
              {localizedStrings.developer.subjectHint}
            </li>
            <li className={styles.developerNoteItem}>
              {localizedStrings.developer.ribbonHint}
            </li>
            <li className={styles.developerNoteItem}>
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
    languagePreference: settingsStore.getLanguagePreference(),
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
  settingsStore.setLanguagePreference(preferences.languagePreference);
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

function RepositoryIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <path d="M4.5 6.5A1.5 1.5 0 0 1 6 5h8a1.5 1.5 0 0 1 1.5 1.5v7A1.5 1.5 0 0 1 14 15H6a1.5 1.5 0 0 1-1.5-1.5Z" />
      <path d="M7 8.2h6" />
      <path d="M7 10.5h4" />
      <path d="M7 12.8h5" />
    </ToolbarIcon>
  );
}

function DocsIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <path d="M6.5 4.8h5.8l2.2 2.2v8.2H6.5Z" />
      <path d="M12.3 4.8v2.4h2.2" />
      <path d="M8.4 10h4.8" />
      <path d="M8.4 12.4h3.6" />
    </ToolbarIcon>
  );
}

function CompanyIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <path d="M6 15.2V6.2h8v9" />
      <path d="M4.5 15.2h11" />
      <path d="M8.2 8.4h1.2" />
      <path d="M10.6 8.4h1.2" />
      <path d="M8.2 10.8h1.2" />
      <path d="M10.6 10.8h1.2" />
      <path d="M9.3 15.2v-2.5h1.4v2.5" />
    </ToolbarIcon>
  );
}

function UpstreamIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <path d="m6.5 9.5 3-3 3 3" />
      <path d="M9.5 6.8v6.7" />
      <path d="M5.5 13.8h8" />
    </ToolbarIcon>
  );
}

function ForkIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <circle cx="6.5" cy="5.8" r="1.2" />
      <circle cx="13.5" cy="5.8" r="1.2" />
      <circle cx="10" cy="14.1" r="1.2" />
      <path d="M6.5 7v2.1c0 1.2 1 2.2 2.2 2.2H10" />
      <path d="M13.5 7v2.1c0 1.2-1 2.2-2.2 2.2H10" />
      <path d="M10 11.3v1.6" />
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
