import {
  Button,
  FluentProvider,
  Radio,
  RadioGroup,
  Switch,
  Textarea,
  makeStyles,
  shorthands,
  tokens,
  webDarkTheme,
  webLightTheme,
} from "@fluentui/react-components";
import {
  type ChangeEvent,
  type DragEvent,
  type ReactElement,
  type ReactNode,
  useDeferredValue,
  useEffect,
  useRef,
  useState,
} from "react";
import { type SelectionSource } from "../lib/body-accessor";
import {
  type ComposeMarkdownService,
  SUBJECT_SELECTION_UNSUPPORTED_MESSAGE,
} from "../lib/compose-markdown";
import {
  defaultStylesheet,
  type SettingsStore,
  type ThemeMode,
} from "../lib/config";
import type { RenderItemResult } from "../lib/item";

const DOCS_URL = "https://schoenfeld-solutions.github.io/markout/";
const REPOSITORY_URL = "https://github.com/Schoenfeld-Solutions/markout";
const STAR_URL = "https://github.com/Schoenfeld-Solutions/markout/stargazers";

const useStyles = makeStyles({
  appShell: {
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    minHeight: "100%",
    minWidth: 0,
    paddingBlockEnd: tokens.spacingVerticalM,
    paddingBlockStart: tokens.spacingVerticalM,
    paddingInlineEnd: tokens.spacingHorizontalM,
    paddingInlineStart: tokens.spacingHorizontalM,
  },
  header: {
    ...shorthands.borderRadius(tokens.borderRadiusXLarge),
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    backgroundColor: tokens.colorNeutralBackground1,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
    minWidth: 0,
    paddingBlockEnd: tokens.spacingVerticalL,
    paddingBlockStart: tokens.spacingVerticalL,
    paddingInlineEnd: tokens.spacingHorizontalL,
    paddingInlineStart: tokens.spacingHorizontalL,
  },
  eyebrow: {
    color: tokens.colorBrandForeground1,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: "0.08em",
    margin: 0,
    textTransform: "uppercase",
  },
  title: {
    fontSize: tokens.fontSizeHero700,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightHero700,
    margin: 0,
  },
  subtitle: {
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    margin: 0,
  },
  contentCard: {
    ...shorthands.borderRadius(tokens.borderRadiusXLarge),
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    backgroundColor: tokens.colorNeutralBackground1,
    display: "flex",
    flex: 1,
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    minHeight: 0,
    minWidth: 0,
    paddingBlockEnd: tokens.spacingVerticalL,
    paddingBlockStart: tokens.spacingVerticalL,
    paddingInlineEnd: tokens.spacingHorizontalL,
    paddingInlineStart: tokens.spacingHorizontalL,
  },
  sectionHeading: {
    display: "flex",
    flexDirection: "column",
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
  sectionStack: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    minWidth: 0,
  },
  inputArea: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
    minWidth: 0,
  },
  dropzone: {
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    ...shorthands.border("1px", "dashed", tokens.colorNeutralStroke2),
    alignItems: "center",
    backgroundColor: tokens.colorNeutralBackground2,
    color: tokens.colorNeutralForeground2,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
    justifyContent: "center",
    minHeight: "8rem",
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
  textarea: {
    minHeight: "10rem",
    width: "100%",
  },
  previewFrame: {
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
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
    "@media (max-width: 540px)": {
      gridTemplateColumns: "1fr",
    },
  },
  status: {
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase200,
    minWidth: 0,
    paddingBlockEnd: tokens.spacingVerticalXS,
    paddingBlockStart: tokens.spacingVerticalXS,
    paddingInlineEnd: tokens.spacingHorizontalM,
    paddingInlineStart: tokens.spacingHorizontalM,
  },
  statusInfo: {
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
  },
  statusSuccess: {
    backgroundColor: tokens.colorPaletteGreenBackground1,
    color: tokens.colorPaletteGreenForeground1,
  },
  statusError: {
    backgroundColor: tokens.colorPaletteRedBackground1,
    color: tokens.colorPaletteRedForeground1,
  },
  settingsGrid: {
    display: "grid",
    gap: tokens.spacingVerticalM,
    minWidth: 0,
  },
  settingsCard: {
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    backgroundColor: tokens.colorNeutralBackground2,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
    minWidth: 0,
    paddingBlockEnd: tokens.spacingVerticalM,
    paddingBlockStart: tokens.spacingVerticalM,
    paddingInlineEnd: tokens.spacingHorizontalM,
    paddingInlineStart: tokens.spacingHorizontalM,
  },
  radioGroup: {
    display: "grid",
    gap: tokens.spacingVerticalXS,
  },
  settingsRow: {
    alignItems: "center",
    display: "flex",
    flexWrap: "wrap",
    gap: tokens.spacingHorizontalM,
    justifyContent: "space-between",
    minWidth: 0,
  },
  toolbar: {
    ...shorthands.borderRadius(tokens.borderRadiusXLarge),
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    backgroundColor: tokens.colorNeutralBackground1,
    bottom: tokens.spacingVerticalM,
    display: "grid",
    gap: tokens.spacingHorizontalXS,
    gridAutoFlow: "column",
    gridAutoColumns: "minmax(0, 1fr)",
    minWidth: 0,
    paddingBlockEnd: tokens.spacingVerticalS,
    paddingBlockStart: tokens.spacingVerticalS,
    paddingInlineEnd: tokens.spacingHorizontalS,
    paddingInlineStart: tokens.spacingHorizontalS,
    position: "sticky",
    zIndex: 1,
  },
  toolbarButton: {
    justifyContent: "center",
    minWidth: 0,
  },
  toolbarLabel: {
    display: "block",
    fontSize: tokens.fontSizeBase100,
    lineHeight: tokens.lineHeightBase100,
  },
  list: {
    color: tokens.colorNeutralForeground2,
    display: "grid",
    gap: tokens.spacingVerticalXS,
    margin: 0,
    paddingInlineStart: "1.25rem",
  },
  linkList: {
    display: "grid",
    gap: tokens.spacingVerticalS,
    margin: 0,
    padding: 0,
  },
  linkCard: {
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    color: "inherit",
    display: "block",
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
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    backgroundColor: tokens.colorNeutralBackground2,
    display: "flex",
    flexDirection: "column",
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
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    backgroundColor: tokens.colorNeutralBackground2,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
    minWidth: 0,
    paddingBlockEnd: tokens.spacingVerticalM,
    paddingBlockStart: tokens.spacingVerticalM,
    paddingInlineEnd: tokens.spacingHorizontalM,
    paddingInlineStart: tokens.spacingHorizontalM,
  },
  developerBlock: {
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
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
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    margin: 0,
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
});

type PanelKey =
  | "credits"
  | "developer"
  | "help"
  | "insert"
  | "intro"
  | "settings";
type StatusTone = "error" | "info" | "success";

interface PreferenceState {
  autoRender: boolean;
  developerToolsEnabled: boolean;
  introDismissed: boolean;
  stylesheet: string;
  themeMode: ThemeMode;
}

interface StatusState {
  message: string;
  tone: StatusTone;
}

interface SelectionDebugState {
  hasSelection: boolean;
  source: SelectionSource;
  textPreview: string;
}

export interface TaskpaneServices {
  composeMarkdown: ComposeMarkdownService;
  renderEntireDraft(): Promise<RenderItemResult>;
}

export interface TaskpaneAppProps {
  services: TaskpaneServices;
  settingsStore: SettingsStore;
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
  services,
  settingsStore,
}: TaskpaneAppProps): ReactElement {
  const styles = useStyles();
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
  const [previewHtml, setPreviewHtml] = useState("");
  const [previewState, setPreviewState] = useState<
    "empty" | "loading" | "ready"
  >("empty");
  const [selectionDebug, setSelectionDebug] =
    useState<SelectionDebugState | null>(null);
  const [status, setStatus] = useState<StatusState>({
    message: preferences.introDismissed
      ? "Ready. Use the insert pane to render Markdown into the current draft."
      : "Welcome to MarkOut. Review the intro once, then switch to insert mode.",
    tone: "info",
  });
  const deferredMarkdownInput = useDeferredValue(markdownInput);
  const deferredStylesheet = useDeferredValue(preferences.stylesheet);
  const lastPersistedStylesheetRef = useRef(preferences.stylesheet);
  const resolvedColorMode = useResolvedColorMode(preferences.themeMode);

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
        setStatus({
          message:
            "Preview could not be rendered with the current Markdown or stylesheet.",
          tone: "error",
        });
      });

    return () => {
      ignore = true;
    };
  }, [deferredMarkdownInput, deferredStylesheet, services.composeMarkdown]);

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
          setStatus({
            message: "Stylesheet changes saved.",
            tone: "success",
          });
        })
        .catch((error: unknown) => {
          console.error("MarkOut failed to persist stylesheet changes.", error);
          setStatus({
            message: "Stylesheet changes could not be persisted.",
            tone: "error",
          });
        });
    }, 700);

    return () => {
      window.clearTimeout(timeoutId);
    };
  }, [preferences.stylesheet, settingsStore]);

  const toolbarPanels: {
    icon: ReactElement;
    key: PanelKey;
    label: string;
  }[] = [{ icon: <InsertIcon />, key: "insert", label: "Insert" }];

  if (!preferences.introDismissed) {
    toolbarPanels.push({ icon: <InfoIcon />, key: "intro", label: "Intro" });
  }

  toolbarPanels.push(
    { icon: <HelpIcon />, key: "help", label: "Help" },
    { icon: <CreditsIcon />, key: "credits", label: "Credits" },
    { icon: <SettingsIcon />, key: "settings", label: "Settings" }
  );

  if (preferences.developerToolsEnabled) {
    toolbarPanels.push({
      icon: <DeveloperIcon />,
      key: "developer",
      label: "Developer",
    });
  }

  const currentTheme =
    resolvedColorMode === "dark" ? webDarkTheme : webLightTheme;

  async function withBusyState(
    busyKey: string,
    operation: () => Promise<void>
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
  ): Promise<void> {
    const previousPreferences = preferences;
    setPreferences(nextPreferences);
    writePreferences(settingsStore, nextPreferences);

    try {
      await settingsStore.save();
      setStatus({ message: successMessage, tone: "success" });
    } catch (error) {
      console.error("MarkOut failed to persist settings.", error);
      setPreferences(previousPreferences);
      writePreferences(settingsStore, previousPreferences);
      setStatus({
        message: "Settings could not be updated.",
        tone: "error",
      });
    }
  }

  async function handleToggleAutoRender(enabled: boolean): Promise<void> {
    await persistPreferences(
      { ...preferences, autoRender: enabled },
      `Auto-render on send ${enabled ? "enabled" : "disabled"}.`
    );
  }

  async function handleToggleDeveloperTools(enabled: boolean): Promise<void> {
    const nextPreferences = {
      ...preferences,
      developerToolsEnabled: enabled,
    };

    if (!enabled && activePanel === "developer") {
      setActivePanel("insert");
    }

    await persistPreferences(
      nextPreferences,
      `Developer tools ${enabled ? "enabled" : "disabled"}.`
    );
  }

  async function handleToggleIntroVisibility(
    showIntro: boolean
  ): Promise<void> {
    const nextPreferences = {
      ...preferences,
      introDismissed: !showIntro,
    };

    if (showIntro) {
      setActivePanel("intro");
    } else if (activePanel === "intro") {
      setActivePanel("insert");
    }

    await persistPreferences(
      nextPreferences,
      showIntro
        ? "Intro restored to the toolbar."
        : "Intro hidden from the toolbar."
    );
  }

  async function handleThemeModeChange(mode: ThemeMode): Promise<void> {
    await persistPreferences(
      {
        ...preferences,
        themeMode: mode,
      },
      `Theme mode updated to ${mode}.`
    );
  }

  async function handleConfirmIntro(): Promise<void> {
    if (preferences.introDismissed) {
      setActivePanel("insert");
      return;
    }

    await persistPreferences(
      {
        ...preferences,
        introDismissed: true,
      },
      "Intro dismissed. You can restore it from settings later."
    );
    setActivePanel("insert");
  }

  async function handleInspectSelection(): Promise<void> {
    setIsInspectingSelection(true);

    try {
      const selection = await services.composeMarkdown.getSelection();
      setSelectionDebug({
        hasSelection: selection.hasSelection,
        source: selection.source,
        textPreview: selection.text.slice(0, 200),
      });
      setStatus({
        message: "Selection state refreshed from Outlook.",
        tone: "success",
      });
    } catch (error) {
      console.error("MarkOut failed to inspect the current selection.", error);
      setSelectionDebug(null);
      setStatus({
        message: "Selection state could not be read from Outlook.",
        tone: "error",
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
        setStatus({
          message:
            result === "replaced"
              ? "Rendered Markdown replaced the current selection."
              : "Rendered Markdown was inserted at the current body cursor.",
          tone: "success",
        });
      } catch (error) {
        console.error("MarkOut failed to insert rendered Markdown.", error);
        setStatus({
          message: formatActionError(error),
          tone: "error",
        });
      }
    });
  }

  async function handleRenderSelection(): Promise<void> {
    await withBusyState("render-selection", async () => {
      try {
        await services.composeMarkdown.renderSelection();
        setStatus({
          message: "The current body selection was rendered successfully.",
          tone: "success",
        });
      } catch (error) {
        console.error("MarkOut failed to render the current selection.", error);
        setStatus({
          message: formatActionError(error),
          tone: "error",
        });
      }
    });
  }

  async function handleRenderEntireDraft(): Promise<void> {
    await withBusyState("render-entire-draft", async () => {
      try {
        const result = await services.renderEntireDraft();
        setStatus({
          message:
            result === "rendered"
              ? "The current draft was rendered successfully."
              : "The original draft HTML was restored successfully.",
          tone: "success",
        });
      } catch (error) {
        console.error("MarkOut failed to render the current draft.", error);
        setStatus({
          message: formatActionError(error),
          tone: "error",
        });
      }
    });
  }

  async function handleDrop(event: DragEvent<HTMLDivElement>): Promise<void> {
    event.preventDefault();
    setIsDropActive(false);
    const file = event.dataTransfer.files.item(0);

    if (file === null) {
      setStatus({
        message: "Drop a Markdown or text file to load content into MarkOut.",
        tone: "error",
      });
      return;
    }

    if (!supportsMarkdownFile(file)) {
      setStatus({
        message:
          "Only .md, .markdown, and .txt files are supported in the insert pane.",
        tone: "error",
      });
      return;
    }

    try {
      const content = await readDroppedMarkdownFile(file);
      setMarkdownInput(content);
      setStatus({
        message: `${file.name} loaded into the insert pane.`,
        tone: "success",
      });
    } catch (error) {
      console.error("MarkOut failed to load a dropped file.", error);
      setStatus({
        message: formatActionError(error),
        tone: "error",
      });
    }
  }

  function handleMarkdownInputChange(
    event: ChangeEvent<HTMLTextAreaElement>,
    data?: { value?: string }
  ): void {
    const nextValue = data?.value ?? event.target.value;
    setMarkdownInput(nextValue);
  }

  function renderStatus(): ReactElement {
    return (
      <div
        id="status-message"
        className={[
          styles.status,
          status.tone === "error"
            ? styles.statusError
            : status.tone === "success"
              ? styles.statusSuccess
              : styles.statusInfo,
        ].join(" ")}
        role="status"
      >
        {status.message}
      </div>
    );
  }

  function renderPreview(): ReactElement {
    if (previewState === "loading") {
      return (
        <div className={`${styles.previewFrame} ${styles.previewFrameEmpty}`}>
          Rendering preview...
        </div>
      );
    }

    if (previewHtml.trim().length === 0) {
      return (
        <div className={`${styles.previewFrame} ${styles.previewFrameEmpty}`}>
          Paste or drop Markdown to preview the fragment that will be inserted
          into the draft.
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
      <div className={styles.sectionStack}>
        <div className={styles.sectionHeading}>
          <h2 className={styles.sectionTitle}>Insert rendered Markdown</h2>
          <p className={styles.sectionBody}>
            Build a fragment in the pane, replace a selected body range, or
            insert rendered content at the current body cursor.
          </p>
        </div>
        <div
          className={`${styles.dropzone} ${isDropActive ? styles.dropzoneActive : ""}`}
          data-testid="taskpane-dropzone"
          onDragEnter={() => setIsDropActive(true)}
          onDragLeave={() => setIsDropActive(false)}
          onDragOver={(event) => {
            event.preventDefault();
            setIsDropActive(true);
          }}
          onDrop={(event: DragEvent<HTMLDivElement>) => {
            void handleDrop(event);
          }}
        >
          <InsertIcon />
          <p className={styles.dropzoneTitle}>Drop a Markdown file here</p>
          <p className={styles.dropzoneCopy}>
            MarkOut accepts <code>.md</code>, <code>.markdown</code>, and{" "}
            <code>.txt</code> files. You can also paste Markdown directly into
            the editor below.
          </p>
        </div>
        <div className={styles.inputArea}>
          <label className={styles.textLabel} htmlFor="markdown-input">
            Markdown input
          </label>
          <Textarea
            id="markdown-input"
            className={styles.textarea}
            onChange={handleMarkdownInputChange}
            placeholder="Paste Markdown here, or drop a Markdown file into the pane."
            resize="vertical"
            value={markdownInput}
          />
        </div>
        <div className={styles.sectionHeading}>
          <h3 className={styles.sectionTitle}>Preview</h3>
          <p className={styles.sectionBody}>
            Preview uses the same sanitized fragment pipeline that MarkOut
            inserts into the draft body.
          </p>
        </div>
        {renderPreview()}
        <div className={styles.actionRow}>
          <Button
            appearance="primary"
            disabled={isWorking !== null}
            id="render-selection-button"
            onClick={() => {
              void handleRenderSelection();
            }}
          >
            Render selection
          </Button>
          <Button
            appearance="secondary"
            disabled={isWorking !== null}
            id="render-entire-draft-button"
            onClick={() => {
              void handleRenderEntireDraft();
            }}
          >
            Render entire draft
          </Button>
          <Button
            appearance="secondary"
            disabled={isWorking !== null || markdownInput.trim().length === 0}
            id="insert-rendered-markdown-button"
            onClick={() => {
              void handleInsertRenderedMarkdown();
            }}
          >
            Insert rendered markdown
          </Button>
        </div>
        <ul className={styles.list}>
          <li>
            Render selection only works on text selected inside the message
            body.
          </li>
          <li>
            Insert rendered markdown uses the current body selection if one
            exists, otherwise Outlook inserts at the last body cursor position
            it preserves.
          </li>
          <li>
            MarkOut blocks fragment insertion into content it already rendered.
          </li>
        </ul>
      </div>
    );
  }

  function renderSettingsPanel(): ReactElement {
    return (
      <div className={styles.settingsGrid}>
        <div className={styles.sectionHeading}>
          <h2 className={styles.sectionTitle}>Settings</h2>
          <p className={styles.sectionBody}>
            Control theme behavior, Smart Alerts rendering, first-run content,
            and developer visibility.
          </p>
        </div>
        <div className={styles.settingsCard}>
          <h3 className={styles.sectionTitle}>Theme mode</h3>
          <p className={styles.sectionBody}>
            System follows Outlook theme when the host provides it and falls
            back to the browser preference otherwise.
          </p>
          <RadioGroup
            className={styles.radioGroup}
            onChange={(_, data) => {
              void handleThemeModeChange(data.value as ThemeMode);
            }}
            value={preferences.themeMode}
          >
            <Radio label="Light" value="light" />
            <Radio label="Dark" value="dark" />
            <Radio label="System" value="system" />
          </RadioGroup>
        </div>
        <div className={styles.settingsCard}>
          <div className={styles.settingsRow}>
            <div>
              <h3 className={styles.sectionTitle}>Smart Alerts auto-render</h3>
              <p className={styles.sectionBody}>
                Render the entire draft before send when Smart Alerts run.
              </p>
            </div>
            <Switch
              checked={preferences.autoRender}
              id="autorender-switch"
              label={preferences.autoRender ? "On" : "Off"}
              onChange={(_, data) => {
                void handleToggleAutoRender(data.checked);
              }}
            />
          </div>
        </div>
        <div className={styles.settingsCard}>
          <div className={styles.settingsRow}>
            <div>
              <h3 className={styles.sectionTitle}>Intro visibility</h3>
              <p className={styles.sectionBody}>
                Restore or hide the intro icon in the bottom toolbar.
              </p>
            </div>
            <Switch
              checked={!preferences.introDismissed}
              id="show-intro-switch"
              label={!preferences.introDismissed ? "Shown" : "Hidden"}
              onChange={(_, data) => {
                void handleToggleIntroVisibility(data.checked);
              }}
            />
          </div>
          <div className={styles.settingsRow}>
            <div>
              <h3 className={styles.sectionTitle}>Developer tools</h3>
              <p className={styles.sectionBody}>
                Reveal the developer panel and additional host diagnostics.
              </p>
            </div>
            <Switch
              checked={preferences.developerToolsEnabled}
              id="developer-tools-switch"
              label={preferences.developerToolsEnabled ? "Shown" : "Hidden"}
              onChange={(_, data) => {
                void handleToggleDeveloperTools(data.checked);
              }}
            />
          </div>
        </div>
        <div className={styles.settingsCard}>
          <h3 className={styles.sectionTitle}>Inline stylesheet</h3>
          <p className={styles.sectionBody}>
            This stylesheet powers full-draft rendering and the fragment
            preview.
          </p>
          <Textarea
            className={styles.textarea}
            id="theme-editor"
            onChange={(_event, data) => {
              const value = data.value;
              setPreferences((currentPreferences) => ({
                ...currentPreferences,
                stylesheet: value,
              }));
            }}
            resize="vertical"
            value={preferences.stylesheet}
          />
          <div className={styles.inlineButtonRow}>
            <Button
              appearance="secondary"
              onClick={() => {
                setPreferences((currentPreferences) => ({
                  ...currentPreferences,
                  stylesheet: defaultStylesheet,
                }));
              }}
            >
              Reset default stylesheet
            </Button>
          </div>
        </div>
      </div>
    );
  }

  function renderHelpPanel(): ReactElement {
    return (
      <div className={styles.sectionStack}>
        <div className={styles.sectionHeading}>
          <h2 className={styles.sectionTitle}>Help</h2>
          <p className={styles.sectionBody}>
            MarkOut is taskpane-first in Outlook on the web. Selection work
            happens through the pane instead of a native Outlook context menu.
          </p>
        </div>
        <div className={styles.linkList}>
          <a
            className={styles.linkCard}
            href={REPOSITORY_URL}
            rel="noreferrer"
            target="_blank"
          >
            <strong>GitHub repository</strong>
            <p className={styles.sectionBody}>
              Track releases, issues, and the maintained Schoenfeld Solutions
              fork.
            </p>
          </a>
          <a
            className={styles.linkCard}
            href={DOCS_URL}
            rel="noreferrer"
            target="_blank"
          >
            <strong>Hosted project docs</strong>
            <p className={styles.sectionBody}>
              Open the GitHub Pages landing page with manifest links and
              deployment notes.
            </p>
          </a>
        </div>
        <ul className={styles.list}>
          <li>
            Use <strong>Render selection</strong> only after selecting Markdown
            in the message body.
          </li>
          <li>
            Use <strong>Render entire draft</strong> when the full draft is
            still unrendered and fragment-free.
          </li>
          <li>
            Use <strong>Insert rendered markdown</strong> for pasted or dropped
            Markdown content.
          </li>
          <li>
            Smart Alerts auto-render remains configurable in settings and
            continues to apply at send time.
          </li>
        </ul>
      </div>
    );
  }

  function renderIntroPanel(): ReactElement {
    return (
      <div className={styles.sectionStack}>
        <div className={styles.sectionHeading}>
          <h2 className={styles.sectionTitle}>What MarkOut does</h2>
          <p className={styles.sectionBody}>
            MarkOut keeps compose work Markdown-first while staying inside
            Outlook's taskpane and Smart Alerts model.
          </p>
        </div>
        <div className={styles.introGrid}>
          <div className={styles.introCard}>
            <div className={styles.introIllustration}>
              <IntroComposeIllustration />
            </div>
            <h3 className={styles.sectionTitle}>
              Render the draft when you are ready
            </h3>
            <p className={styles.sectionBody}>
              Render the whole draft or only a selected body range without
              leaving compose mode.
            </p>
          </div>
          <div className={styles.introCard}>
            <div className={styles.introIllustration}>
              <IntroInsertIllustration />
            </div>
            <h3 className={styles.sectionTitle}>
              Insert Markdown fragments safely
            </h3>
            <p className={styles.sectionBody}>
              Drop a file or paste Markdown, preview the fragment, then insert
              it where Outlook still preserves your body selection or cursor.
            </p>
          </div>
        </div>
        <ul className={styles.list}>
          <li>
            MarkOut blocks double conversion of its own rendered fragments.
          </li>
          <li>
            Developer tools stay hidden unless you turn them on in settings.
          </li>
          <li>
            The intro icon disappears after confirmation and can be restored
            later.
          </li>
        </ul>
        <div className={styles.inlineButtonRow}>
          <Button
            appearance="primary"
            id="intro-confirm-button"
            onClick={() => {
              void handleConfirmIntro();
            }}
          >
            I have read this
          </Button>
        </div>
      </div>
    );
  }

  function renderCreditsPanel(): ReactElement {
    return (
      <div className={styles.sectionStack}>
        <div className={styles.sectionHeading}>
          <h2 className={styles.sectionTitle}>Credits</h2>
          <p className={styles.sectionBody}>
            MarkOut continues as an independently maintained fork while
            preserving visible credit to the upstream project.
          </p>
        </div>
        <div className={styles.creditsBox}>
          <h3 className={styles.sectionTitle}>Upstream foundation</h3>
          <p className={styles.sectionBody}>
            Original product direction and implementation work came from{" "}
            <strong>SierraSoftworks/markout</strong>.
          </p>
        </div>
        <div className={styles.creditsBox}>
          <h3 className={styles.sectionTitle}>Current maintenance</h3>
          <p className={styles.sectionBody}>
            This fork is maintained by <strong>Schoenfeld Solutions</strong>,
            with ongoing work by <strong>Gabriel-Johannes Schönfeld</strong>.
          </p>
        </div>
        <div className={styles.inlineButtonRow}>
          <Button as="a" href={REPOSITORY_URL} rel="noreferrer" target="_blank">
            Open the fork
          </Button>
          <Button as="a" href={STAR_URL} rel="noreferrer" target="_blank">
            Leave a star
          </Button>
        </div>
      </div>
    );
  }

  function renderDeveloperPanel(): ReactElement {
    return (
      <div className={styles.sectionStack}>
        <div className={styles.sectionHeading}>
          <h2 className={styles.sectionTitle}>Developer tools</h2>
          <p className={styles.sectionBody}>
            Inspect host theme resolution and selection state without exposing
            debug noise to regular users.
          </p>
        </div>
        <div className={styles.developerBlock}>
          <div className={styles.settingsRow}>
            <div>
              <h3 className={styles.sectionTitle}>Resolved taskpane theme</h3>
              <p className={styles.sectionBody}>
                Preference: {preferences.themeMode}. Effective theme:{" "}
                {resolvedColorMode}.
              </p>
            </div>
            <Button
              appearance="secondary"
              disabled={isInspectingSelection}
              onClick={() => {
                void handleInspectSelection();
              }}
            >
              Inspect selection
            </Button>
          </div>
          <pre className={styles.developerCode}>
            {selectionDebug === null
              ? "No selection snapshot loaded yet."
              : JSON.stringify(selectionDebug, null, 2)}
          </pre>
        </div>
        <div className={styles.developerBlock}>
          <h3 className={styles.sectionTitle}>Host notes</h3>
          <ul className={styles.list}>
            <li>{SUBJECT_SELECTION_UNSUPPORTED_MESSAGE}</li>
            <li>
              Outlook decides whether a toolbar command appears directly on the
              ribbon or in overflow/app surfaces.
            </li>
            <li>
              Native Outlook context-menu commands are not the delivery path for
              this add-in.
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
      data-theme={resolvedColorMode}
      id="taskpane-shell"
      theme={currentTheme}
    >
      <header className={styles.header}>
        <p className={styles.eyebrow}>MarkOut</p>
        <h1 className={styles.title}>Markdown-first Outlook compose</h1>
        <p className={styles.subtitle}>
          Taskpane-first rendering, fragment insertion, and Smart Alerts support
          in a single Outlook-friendly workspace.
        </p>
        {renderStatus()}
      </header>
      <main className={styles.contentCard}>{renderActivePanel()}</main>
      <nav aria-label="MarkOut sections" className={styles.toolbar}>
        {toolbarPanels.map((panel) => (
          <Button
            appearance={activePanel === panel.key ? "primary" : "subtle"}
            className={styles.toolbarButton}
            icon={panel.icon}
            id={`panel-button-${panel.key}`}
            key={panel.key}
            onClick={() => setActivePanel(panel.key)}
          >
            <span className={styles.toolbarLabel}>{panel.label}</span>
          </Button>
        ))}
      </nav>
    </FluentProvider>
  );
}

function formatActionError(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  return "MarkOut could not complete that action.";
}

function readPreferences(settingsStore: SettingsStore): PreferenceState {
  return {
    autoRender: settingsStore.getAutoRender(),
    developerToolsEnabled: settingsStore.getDeveloperToolsEnabled(),
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
  settingsStore.setDeveloperToolsEnabled(preferences.developerToolsEnabled);
  settingsStore.setIntroDismissed(preferences.introDismissed);
  settingsStore.setStylesheet(preferences.stylesheet);
  settingsStore.setThemeMode(preferences.themeMode);
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

function resolveSystemColorMode(
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

function isDarkColor(color: string): boolean {
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
