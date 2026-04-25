import {
  Button,
  MessageBar,
  MessageBarBody,
  Select,
  Switch,
  Toolbar,
  ToolbarRadioButton,
  ToolbarRadioGroup,
  mergeClasses,
} from "@fluentui/react-components";
import type { DragEvent, ReactElement } from "react";
import type { StylesheetLintResult } from "../lib/stylesheet-lint";
import type { LanguagePreference, ThemeMode } from "../lib/config";
import type { LocalizedStrings } from "./i18n";
import {
  CompanyIcon,
  DocsIcon,
  ForkIcon,
  InsertIcon,
  IntroComposeIllustration,
  IntroInsertIllustration,
  RepositoryIcon,
  UpstreamIcon,
} from "./icons";
import type {
  DiagnosticEventRecord,
  PanelKey,
  SelectionDebugState,
} from "./types";

const DOCS_URL = "https://schoenfeld-solutions.github.io/markout/";
const REPOSITORY_URL = "https://github.com/Schoenfeld-Solutions/markout";
const STAR_URL = "https://github.com/Schoenfeld-Solutions/markout/stargazers";
const WEBSITE_URL = "https://schoenfeld.solutions";

function renderOptionalBody(
  styles: Record<string, string>,
  copy: string
): ReactElement | null {
  return copy.trim().length > 0 ? (
    <p className={styles.sectionBody}>{copy}</p>
  ) : null;
}

function renderLintResult(
  styles: Record<string, string>,
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

export function InsertPanel(props: {
  isDropActive: boolean;
  isInsertRenderedMarkdownDisabled: boolean;
  isWorking: boolean;
  markdownInput: string;
  onDrop: (event: DragEvent<HTMLDivElement>) => void;
  onInsertRenderedMarkdown: () => void;
  onMarkdownInputChange: (value: string) => void;
  onRenderEntireDraft: () => void;
  onRenderSelection: () => void;
  previewHtml: string;
  previewFrameStyle: { colorScheme: "dark" | "light" };
  previewState: "empty" | "loading" | "ready";
  renderSelectionDisabled: boolean;
  renderSelectionTooltip: string;
  setDropActive: (value: boolean) => void;
  strings: LocalizedStrings;
  styles: Record<string, string>;
}): ReactElement {
  const {
    isDropActive,
    isInsertRenderedMarkdownDisabled,
    isWorking,
    markdownInput,
    onDrop,
    onInsertRenderedMarkdown,
    onMarkdownInputChange,
    onRenderEntireDraft,
    onRenderSelection,
    previewHtml,
    previewFrameStyle,
    previewState,
    renderSelectionDisabled,
    renderSelectionTooltip,
    setDropActive,
    strings,
    styles,
  } = props;

  const renderPreview = (): ReactElement => {
    if (previewState === "loading") {
      return (
        <div
          className={mergeClasses(
            styles.previewFrame,
            styles.previewFrameEmpty
          )}
          style={previewFrameStyle}
        >
          {strings.insert.previewLoading}
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
          {strings.insert.emptyPreview}
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
  };

  return (
    <div className={styles.panelRoot}>
      <div className={styles.sectionHeading}>
        <h2 className={styles.sectionTitle}>{strings.insert.panelTitle}</h2>
        {renderOptionalBody(styles, strings.insert.panelDescription)}
      </div>
      <div
        className={mergeClasses(
          styles.dropzone,
          isDropActive ? styles.dropzoneActive : undefined
        )}
        data-testid="taskpane-dropzone"
        onDragEnter={() => setDropActive(true)}
        onDragLeave={() => setDropActive(false)}
        onDragOver={(event) => {
          event.preventDefault();
          setDropActive(true);
        }}
        onDrop={(event) => {
          event.preventDefault();
          onDrop(event);
        }}
      >
        <InsertIcon />
        <p className={styles.dropzoneTitle}>{strings.insert.dropzoneTitle}</p>
        <p className={styles.dropzoneCopy}>{strings.insert.dropzoneCopy}</p>
      </div>
      <div className={styles.card}>
        <div className={styles.sectionHeading}>
          <label className={styles.textLabel} htmlFor="markdown-input">
            {strings.insert.inputLabel}
          </label>
        </div>
        <div className={styles.textareaSurface}>
          <textarea
            className={styles.plainTextarea}
            id="markdown-input"
            onChange={(event) => {
              onMarkdownInputChange(event.target.value);
            }}
            placeholder={strings.insert.inputPlaceholder}
            spellCheck={false}
            value={markdownInput}
          />
        </div>
      </div>
      <div className={styles.card}>
        <div className={styles.sectionHeading}>
          <h3 className={styles.sectionTitle}>{strings.insert.previewTitle}</h3>
          {renderOptionalBody(styles, strings.insert.previewDescription)}
        </div>
        {renderPreview()}
        <div className={styles.actionRow}>
          <Button
            appearance="primary"
            aria-label={strings.insert.renderSelectionButton}
            disabled={renderSelectionDisabled}
            id="render-selection-button"
            onClick={onRenderSelection}
            title={renderSelectionTooltip}
          >
            {strings.insert.renderSelectionButton}
          </Button>
          <Button
            appearance="secondary"
            disabled={isWorking}
            id="render-entire-draft-button"
            onClick={onRenderEntireDraft}
            title={strings.tooltips.renderEntireDraft}
          >
            {strings.insert.renderEntireDraftButton}
          </Button>
          <Button
            appearance="secondary"
            disabled={isInsertRenderedMarkdownDisabled}
            id="insert-rendered-markdown-button"
            onClick={onInsertRenderedMarkdown}
            title={strings.tooltips.insertRenderedMarkdown}
          >
            {strings.insert.insertButton}
          </Button>
        </div>
      </div>
    </div>
  );
}

export function SettingsPanel(props: {
  autoRenderEnabled: boolean;
  codeMirrorHostRef: React.RefObject<HTMLDivElement | null>;
  cssLintResult: StylesheetLintResult | null;
  developerToolsEnabled: boolean;
  helpVisible: boolean;
  introVisible: boolean;
  isCodeMirrorLoading: boolean;
  isWorking: boolean;
  languagePreference: LanguagePreference;
  onCreditsVisibilityChange: (visible: boolean) => void;
  onDeveloperToolsChange: (enabled: boolean) => void;
  onHelpVisibilityChange: (visible: boolean) => void;
  onIntroVisibilityChange: (visible: boolean) => void;
  onLanguagePreferenceChange: (preference: LanguagePreference) => void;
  onLintStylesheet: () => void;
  onResetStylesheet: () => void;
  onThemeModeChange: (mode: ThemeMode) => void;
  onToggleAutoRender: (enabled: boolean) => void;
  preferencesThemeMode: ThemeMode;
  showCredits: boolean;
  strings: LocalizedStrings;
  styles: Record<string, string>;
}): ReactElement {
  const {
    autoRenderEnabled,
    codeMirrorHostRef,
    cssLintResult,
    developerToolsEnabled,
    helpVisible,
    introVisible,
    isCodeMirrorLoading,
    isWorking,
    languagePreference,
    onCreditsVisibilityChange,
    onDeveloperToolsChange,
    onHelpVisibilityChange,
    onIntroVisibilityChange,
    onLanguagePreferenceChange,
    onLintStylesheet,
    onResetStylesheet,
    onThemeModeChange,
    onToggleAutoRender,
    preferencesThemeMode,
    showCredits,
    strings,
    styles,
  } = props;

  return (
    <div className={styles.panelRoot}>
      <div className={styles.sectionHeading}>
        <h2 className={styles.sectionTitle}>{strings.settings.panelTitle}</h2>
        {renderOptionalBody(styles, strings.settings.panelDescription)}
      </div>
      <div className={styles.card}>
        <h3 className={styles.sectionTitle}>{strings.settings.themeTitle}</h3>
        {renderOptionalBody(styles, strings.settings.themeDescription)}
        <Toolbar
          aria-label={strings.settings.themeTitle}
          checkedValues={{ "theme-mode": [preferencesThemeMode] }}
          className={styles.themeModeToolbar}
          onCheckedValueChange={(_, data) => {
            const nextMode = data.checkedItems[0];

            if (
              data.name === "theme-mode" &&
              (nextMode === "light" ||
                nextMode === "dark" ||
                nextMode === "system")
            ) {
              onThemeModeChange(nextMode);
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
              {strings.settings.themeModeLight}
            </ToolbarRadioButton>
            <ToolbarRadioButton
              appearance="subtle"
              className={styles.themeModeToolbarButton}
              id="theme-mode-dark"
              name="theme-mode"
              value="dark"
            >
              {strings.settings.themeModeDark}
            </ToolbarRadioButton>
            <ToolbarRadioButton
              appearance="subtle"
              className={styles.themeModeToolbarButton}
              id="theme-mode-system"
              name="theme-mode"
              value="system"
            >
              {strings.settings.themeModeSystem}
            </ToolbarRadioButton>
          </ToolbarRadioGroup>
        </Toolbar>
      </div>
      <div className={styles.card}>
        <h3 className={styles.sectionTitle}>
          {strings.settings.languageTitle}
        </h3>
        {renderOptionalBody(styles, strings.settings.languageDescription)}
        <Select
          className={styles.selectControl}
          id="language-preference-select"
          onChange={(event) => {
            onLanguagePreferenceChange(
              event.currentTarget.value as LanguagePreference
            );
          }}
          value={languagePreference}
        >
          <option value="system">{strings.settings.languageSystem}</option>
          <option value="en-US">{strings.settings.languageEnglish}</option>
          <option value="de-DE">{strings.settings.languageGerman}</option>
        </Select>
      </div>
      <div className={styles.card}>
        <SettingsToggleRow
          body={strings.settings.autoRenderDescription}
          checked={autoRenderEnabled}
          id="autorender-switch"
          offLabel={strings.general.off}
          onChange={onToggleAutoRender}
          onLabel={strings.general.on}
          styles={styles}
          title={strings.settings.autoRenderTitle}
        />
        <SettingsToggleRow
          body={strings.settings.introDescription}
          checked={introVisible}
          id="show-intro-switch"
          offLabel={strings.general.hidden}
          onChange={onIntroVisibilityChange}
          onLabel={strings.general.shown}
          styles={styles}
          title={strings.settings.introTitle}
        />
        <SettingsToggleRow
          body={strings.settings.helpDescription}
          checked={helpVisible}
          id="show-help-switch"
          offLabel={strings.general.hidden}
          onChange={onHelpVisibilityChange}
          onLabel={strings.general.shown}
          styles={styles}
          title={strings.settings.helpTitle}
        />
        <SettingsToggleRow
          body={strings.settings.creditsDescription}
          checked={showCredits}
          id="show-credits-switch"
          offLabel={strings.general.hidden}
          onChange={onCreditsVisibilityChange}
          onLabel={strings.general.shown}
          styles={styles}
          title={strings.settings.creditsTitle}
        />
        <SettingsToggleRow
          body={strings.settings.developerDescription}
          checked={developerToolsEnabled}
          id="developer-tools-switch"
          offLabel={strings.general.hidden}
          onChange={onDeveloperToolsChange}
          onLabel={strings.general.shown}
          styles={styles}
          title={strings.settings.developerTitle}
        />
      </div>
      <div className={styles.card}>
        <div className={styles.sectionHeading}>
          <h3 className={styles.sectionTitle}>{strings.editor.title}</h3>
        </div>
        <div className={styles.editorSurface}>
          {isCodeMirrorLoading ? (
            <div className={styles.codeMirrorLoading}>
              {strings.editor.loading}
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
            disabled={isWorking}
            id="lint-stylesheet-button"
            onClick={onLintStylesheet}
          >
            {strings.editor.lintButton}
          </Button>
          <Button appearance="secondary" onClick={onResetStylesheet}>
            {strings.editor.resetButton}
          </Button>
        </div>
        {renderLintResult(styles, strings, cssLintResult)}
      </div>
    </div>
  );
}

function SettingsToggleRow(props: {
  body: string;
  checked: boolean;
  id: string;
  offLabel: string;
  onChange: (checked: boolean) => void;
  onLabel: string;
  styles: Record<string, string>;
  title: string;
}): ReactElement {
  return (
    <div className={props.styles.settingsRow}>
      <div className={props.styles.sectionHeading}>
        <h3 className={props.styles.sectionTitle}>{props.title}</h3>
        <p className={props.styles.sectionBody}>{props.body}</p>
      </div>
      <Switch
        checked={props.checked}
        id={props.id}
        label={props.checked ? props.onLabel : props.offLabel}
        onChange={(_, data) => {
          props.onChange(data.checked);
        }}
      />
    </div>
  );
}

export function HelpPanel(props: {
  strings: LocalizedStrings;
  styles: Record<string, string>;
}): ReactElement {
  const { strings, styles } = props;

  return (
    <div className={styles.panelRoot}>
      <div className={styles.sectionHeading}>
        <h2 className={styles.sectionTitle}>{strings.help.panelTitle}</h2>
        {renderOptionalBody(styles, strings.help.panelDescription)}
      </div>
      <div className={styles.linkList}>
        <ExternalLinkCard
          copy={strings.help.repoDescription}
          href={REPOSITORY_URL}
          icon={<RepositoryIcon />}
          styles={styles}
          title={strings.help.repoTitle}
        />
        <ExternalLinkCard
          copy={strings.help.docsDescription}
          href={DOCS_URL}
          icon={<DocsIcon />}
          styles={styles}
          title={strings.help.docsTitle}
        />
        <ExternalLinkCard
          copy={strings.help.websiteDescription}
          href={WEBSITE_URL}
          icon={<CompanyIcon />}
          styles={styles}
          title={strings.help.websiteTitle}
        />
      </div>
    </div>
  );
}

function ExternalLinkCard(props: {
  copy: string;
  href: string;
  icon: ReactElement;
  styles: Record<string, string>;
  title: string;
}): ReactElement {
  return (
    <a
      className={props.styles.linkCard}
      href={props.href}
      rel="noreferrer"
      target="_blank"
    >
      <div className={props.styles.linkCardHeader}>
        <span className={props.styles.linkCardIcon}>{props.icon}</span>
        <strong>{props.title}</strong>
      </div>
      {renderOptionalBody(props.styles, props.copy)}
    </a>
  );
}

export function IntroPanel(props: {
  onConfirm: () => void;
  strings: LocalizedStrings;
  styles: Record<string, string>;
}): ReactElement {
  const { onConfirm, strings, styles } = props;

  return (
    <div className={styles.panelRoot}>
      <div className={styles.sectionHeading}>
        <h2 className={styles.sectionTitle}>{strings.intro.panelTitle}</h2>
        {renderOptionalBody(styles, strings.intro.panelDescription)}
      </div>
      <div className={styles.introGrid}>
        <div className={styles.introCard}>
          <div className={styles.introIllustration}>
            <IntroComposeIllustration />
          </div>
          <h3 className={styles.sectionTitle}>{strings.intro.stepOneTitle}</h3>
          <p className={styles.sectionBody}>{strings.intro.stepOneBody}</p>
        </div>
        <div className={styles.introCard}>
          <div className={styles.introIllustration}>
            <IntroInsertIllustration />
          </div>
          <h3 className={styles.sectionTitle}>{strings.intro.stepTwoTitle}</h3>
          <p className={styles.sectionBody}>{strings.intro.stepTwoBody}</p>
        </div>
      </div>
      <div className={styles.inlineButtonRow}>
        <Button
          appearance="primary"
          id="intro-confirm-button"
          onClick={onConfirm}
        >
          {strings.intro.confirm}
        </Button>
      </div>
    </div>
  );
}

export function CreditsPanel(props: {
  strings: LocalizedStrings;
  styles: Record<string, string>;
}): ReactElement {
  const { strings, styles } = props;

  return (
    <div className={styles.panelRoot}>
      <div className={styles.sectionHeading}>
        <h2 className={styles.sectionTitle}>{strings.credits.panelTitle}</h2>
        {renderOptionalBody(styles, strings.credits.panelDescription)}
      </div>
      <CreditBox
        body={strings.credits.upstreamBody}
        icon={<UpstreamIcon />}
        styles={styles}
        title={strings.credits.upstreamTitle}
      />
      <CreditBox
        body={strings.credits.currentMaintenanceBody}
        icon={<ForkIcon />}
        styles={styles}
        title={strings.credits.currentMaintenanceTitle}
      />
      <div className={styles.inlineButtonRow}>
        <Button as="a" href={REPOSITORY_URL} rel="noreferrer" target="_blank">
          {strings.credits.openFork}
        </Button>
        <Button as="a" href={STAR_URL} rel="noreferrer" target="_blank">
          {strings.credits.starFork}
        </Button>
      </div>
    </div>
  );
}

function CreditBox(props: {
  body: string;
  icon: ReactElement;
  styles: Record<string, string>;
  title: string;
}): ReactElement {
  return (
    <div className={props.styles.creditsBox}>
      <div className={props.styles.linkCardHeader}>
        <span className={props.styles.linkCardIcon}>{props.icon}</span>
        <h3 className={props.styles.sectionTitle}>{props.title}</h3>
      </div>
      <p className={props.styles.sectionBody}>{props.body}</p>
    </div>
  );
}

export function DeveloperPanel(props: {
  diagnosticEvents: DiagnosticEventRecord[];
  isInspectingSelection: boolean;
  onInspectSelection: () => void;
  resolvedColorMode: "dark" | "light";
  selectionDebug: SelectionDebugState | null;
  strings: LocalizedStrings;
  styles: Record<string, string>;
  themeMode: ThemeMode;
}): ReactElement {
  const {
    diagnosticEvents,
    isInspectingSelection,
    onInspectSelection,
    resolvedColorMode,
    selectionDebug,
    strings,
    styles,
    themeMode,
  } = props;

  return (
    <div className={styles.panelRoot}>
      <div className={styles.sectionHeading}>
        <h2 className={styles.sectionTitle}>{strings.developer.panelTitle}</h2>
        {renderOptionalBody(styles, strings.developer.panelDescription)}
      </div>
      <div className={mergeClasses(styles.card, styles.compactCard)}>
        <div className={styles.settingsRow}>
          <div className={styles.sectionHeading}>
            <h3 className={styles.sectionTitle}>
              {strings.developer.hostNotesTitle}
            </h3>
            <p className={styles.sectionBody}>
              {strings.developer.resolvedTheme
                .replace("{mode}", themeMode)
                .replace("{resolvedMode}", resolvedColorMode)}
            </p>
          </div>
          <Button
            appearance="secondary"
            disabled={isInspectingSelection}
            onClick={onInspectSelection}
          >
            {strings.developer.inspectSelection}
          </Button>
        </div>
        <pre className={styles.developerCode}>
          {selectionDebug === null
            ? strings.developer.noSelectionSnapshot
            : JSON.stringify(selectionDebug, null, 2)}
        </pre>
        <div className={styles.sectionHeading}>
          <h3 className={styles.sectionTitle}>
            {strings.developer.diagnosticsTitle}
          </h3>
          <p className={styles.sectionBody}>
            {strings.developer.diagnosticsPrivacyNote}
          </p>
        </div>
        <pre className={styles.developerCode}>
          {diagnosticEvents.length === 0
            ? strings.developer.noDiagnostics
            : JSON.stringify(diagnosticEvents, null, 2)}
        </pre>
        <ul className={styles.developerNoteList}>
          <li className={styles.developerNoteItem}>
            {strings.developer.subjectHint}
          </li>
          <li className={styles.developerNoteItem}>
            {strings.developer.ribbonHint}
          </li>
          <li className={styles.developerNoteItem}>
            {strings.developer.taskpaneHint}
          </li>
        </ul>
      </div>
    </div>
  );
}

export function renderActivePanel(props: {
  activePanel: PanelKey;
  creditsPanel: ReactElement;
  developerPanel: ReactElement;
  helpPanel: ReactElement;
  insertPanel: ReactElement;
  introPanel: ReactElement;
  settingsPanel: ReactElement;
}): ReactElement {
  switch (props.activePanel) {
    case "credits":
      return props.creditsPanel;
    case "developer":
      return props.developerPanel;
    case "help":
      return props.helpPanel;
    case "intro":
      return props.introPanel;
    case "settings":
      return props.settingsPanel;
    default:
      return props.insertPanel;
  }
}
