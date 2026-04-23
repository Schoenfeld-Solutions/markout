// @ts-expect-error Webpack resolves taskpane CSS side-effect imports at build time.
import "../taskpane.css";
import { createRoot } from "react-dom/client";
import { createComposeMarkdownService } from "../../lib/compose-markdown";
import type {
  BodyAccessor,
  ComposeSelection,
  SelectionSource,
} from "../../lib/body-accessor";
import type {
  AutoRenderNotificationCopy,
  ComposeNotificationService,
  ComposeTransientNotificationCopy,
  NotificationSurface,
} from "../../lib/compose-notifications";
import {
  defaultStylesheet,
  type LanguagePreference,
  type SettingsStore,
  type ThemeMode,
} from "../../lib/config";
import { DefaultHtmlSanitizer } from "../../lib/html-sanitizer";
import { createItemRenderer, type RenderItemResult } from "../../lib/item";
import type {
  RenderState,
  RenderStateStore,
} from "../../lib/render-state-store";
import { createMarkdownRenderer } from "../../lib/renderer";
import { TaskpaneApp, type TaskpaneServices } from "../app";

interface MockSettingsSnapshot {
  autoRender: boolean;
  creditsVisible: boolean;
  developerToolsEnabled: boolean;
  helpVisible: boolean;
  introDismissed: boolean;
  languagePreference: LanguagePreference;
  stylesheet: string;
  themeMode: ThemeMode;
}

interface MockSnapshot {
  autoRenderNotifications: string[];
  bodyHtml: string;
  lastInsertedHtml: string | null;
  selection: ComposeSelection;
  settings: MockSettingsSnapshot;
  transientNotifications: {
    intent: ComposeTransientNotificationCopy["intent"];
    message: string;
  }[];
}

interface TaskpaneMockController {
  getState(): MockSnapshot;
  reset(): void;
  setBodyHtml(html: string): void;
  setSelection(
    nextSelection: Partial<ComposeSelection> & Pick<ComposeSelection, "text">
  ): void;
}

declare global {
  interface Window {
    __MARKOUT_TASKPANE_MOCK__?: TaskpaneMockController;
  }
}

const INITIAL_BODY_HTML = "<p>Mock draft body</p>";
const INITIAL_MARKDOWN_INPUT = `# Preview heading

Paragraph with [a link](https://example.com) and \`inline code\`.

> Blockquote content

| Column | Value |
| --- | --- |
| Alpha | Beta |

\`\`\`ts
const preview = "theme-aware";
\`\`\`
`;

function createInitialSelection(): ComposeSelection {
  return {
    hasSelection: false,
    html: null,
    source: "body",
    text: "",
  };
}

class MockSettingsStore implements SettingsStore {
  private readonly snapshot: MockSettingsSnapshot = {
    autoRender: false,
    creditsVisible: true,
    developerToolsEnabled: false,
    helpVisible: true,
    introDismissed: false,
    languagePreference: "system",
    stylesheet: defaultStylesheet,
    themeMode: "system",
  };

  public getAutoRender(): boolean {
    return this.snapshot.autoRender;
  }

  public getCreditsVisible(): boolean {
    return this.snapshot.creditsVisible;
  }

  public getDeveloperToolsEnabled(): boolean {
    return this.snapshot.developerToolsEnabled;
  }

  public getHelpVisible(): boolean {
    return this.snapshot.helpVisible;
  }

  public getIntroDismissed(): boolean {
    return this.snapshot.introDismissed;
  }

  public getLanguagePreference(): LanguagePreference {
    return this.snapshot.languagePreference;
  }

  public getStylesheet(): string {
    return this.snapshot.stylesheet;
  }

  public getThemeMode(): ThemeMode {
    return this.snapshot.themeMode;
  }

  public hasStylesheetMigrationPending(): boolean {
    return false;
  }

  public save(): Promise<void> {
    return Promise.resolve();
  }

  public setAutoRender(enabled: boolean): void {
    this.snapshot.autoRender = enabled;
  }

  public setCreditsVisible(visible: boolean): void {
    this.snapshot.creditsVisible = visible;
  }

  public setDeveloperToolsEnabled(enabled: boolean): void {
    this.snapshot.developerToolsEnabled = enabled;
  }

  public setHelpVisible(visible: boolean): void {
    this.snapshot.helpVisible = visible;
  }

  public setIntroDismissed(dismissed: boolean): void {
    this.snapshot.introDismissed = dismissed;
  }

  public setLanguagePreference(preference: LanguagePreference): void {
    this.snapshot.languagePreference = preference;
  }

  public setStylesheet(stylesheet: string): void {
    this.snapshot.stylesheet =
      stylesheet.trim().length > 0 ? stylesheet : defaultStylesheet;
  }

  public setThemeMode(mode: ThemeMode): void {
    this.snapshot.themeMode = mode;
  }

  public readSnapshot(): MockSettingsSnapshot {
    return { ...this.snapshot };
  }

  public reset(): void {
    this.snapshot.autoRender = false;
    this.snapshot.creditsVisible = true;
    this.snapshot.developerToolsEnabled = false;
    this.snapshot.helpVisible = true;
    this.snapshot.introDismissed = false;
    this.snapshot.languagePreference = "system";
    this.snapshot.stylesheet = defaultStylesheet;
    this.snapshot.themeMode = "system";
  }
}

class MockBodyAccessor implements BodyAccessor {
  public constructor(private readonly state: MockSnapshot) {}

  public getHtml(): Promise<string> {
    return Promise.resolve(this.state.bodyHtml);
  }

  public getSelection(): Promise<ComposeSelection> {
    return Promise.resolve({ ...this.state.selection });
  }

  public replaceSelectionWithHtml(html: string): Promise<void> {
    this.state.lastInsertedHtml = html;
    this.state.bodyHtml = html;
    this.state.selection = createInitialSelection();
    return Promise.resolve();
  }

  public setHtml(html: string): Promise<void> {
    this.state.bodyHtml = html;
    return Promise.resolve();
  }
}

class InMemoryRenderStateStore implements RenderStateStore {
  private renderState: RenderState | null = null;

  public clearRenderState(): Promise<void> {
    this.renderState = null;
    return Promise.resolve();
  }

  public getRenderState(): Promise<RenderState | null> {
    return Promise.resolve(this.renderState);
  }

  public setPendingRenderState(originalHtml: string): Promise<void> {
    this.renderState = {
      channelId: "local",
      originalHtml,
      phase: "pending",
      storedAt: new Date().toISOString(),
    };
    return Promise.resolve();
  }

  public setRenderedRenderState(originalHtml: string): Promise<void> {
    this.renderState = {
      channelId: "local",
      originalHtml,
      phase: "rendered",
      storedAt: new Date().toISOString(),
    };
    return Promise.resolve();
  }
}

class MockNotificationService implements ComposeNotificationService {
  public constructor(private readonly state: MockSnapshot) {}

  public clearAutoRenderDismissed(): Promise<void> {
    return Promise.resolve();
  }

  public clearAutoRenderNotification(): Promise<void> {
    this.state.autoRenderNotifications = [];
    return Promise.resolve();
  }

  public clearTransientNotification(): Promise<void> {
    this.state.transientNotifications = [];
    return Promise.resolve();
  }

  public hasAutoRenderBeenDismissed(): Promise<boolean> {
    return Promise.resolve(false);
  }

  public markAutoRenderDismissed(): Promise<void> {
    return Promise.resolve();
  }

  public onAutoRenderDismiss(handler: () => void): void {
    void handler;
    // The local harness doesn't emulate Outlook infobar dismissal events.
  }

  public showAutoRenderNotification(
    copy: AutoRenderNotificationCopy
  ): Promise<NotificationSurface> {
    this.state.autoRenderNotifications.push(copy.message);
    return Promise.resolve("outlook");
  }

  public showTransientNotification(
    copy: ComposeTransientNotificationCopy
  ): Promise<NotificationSurface> {
    this.state.transientNotifications.push({
      intent: copy.intent,
      message: copy.message,
    });
    return Promise.resolve("outlook");
  }
}

function createMockSnapshot(settingsStore: MockSettingsStore): MockSnapshot {
  return {
    autoRenderNotifications: [],
    bodyHtml: INITIAL_BODY_HTML,
    lastInsertedHtml: null,
    selection: createInitialSelection(),
    settings: settingsStore.readSnapshot(),
    transientNotifications: [],
  };
}

function createController(
  state: MockSnapshot,
  settingsStore: MockSettingsStore
): TaskpaneMockController {
  return {
    getState: () => ({
      autoRenderNotifications: [...state.autoRenderNotifications],
      bodyHtml: state.bodyHtml,
      lastInsertedHtml: state.lastInsertedHtml,
      selection: { ...state.selection },
      settings: settingsStore.readSnapshot(),
      transientNotifications: [...state.transientNotifications],
    }),
    reset: () => {
      settingsStore.reset();
      state.autoRenderNotifications = [];
      state.bodyHtml = INITIAL_BODY_HTML;
      state.lastInsertedHtml = null;
      state.selection = createInitialSelection();
      state.settings = settingsStore.readSnapshot();
      state.transientNotifications = [];
    },
    setBodyHtml: (html) => {
      state.bodyHtml = html;
    },
    setSelection: (nextSelection) => {
      const source: SelectionSource = nextSelection.source ?? "body";
      const trimmedText = nextSelection.text.trim();

      state.selection = {
        hasSelection: nextSelection.hasSelection ?? trimmedText.length > 0,
        html: nextSelection.html ?? null,
        source,
        text: nextSelection.text,
      };
    },
  };
}

function mountMockTaskpane(): void {
  const rootElement = document.getElementById("taskpane-root");

  if (rootElement === null) {
    throw new Error('MarkOut could not find the "taskpane-root" element.');
  }

  const settingsStore = new MockSettingsStore();
  const state = createMockSnapshot(settingsStore);
  const bodyAccessor = new MockBodyAccessor(state);
  const htmlSanitizer = new DefaultHtmlSanitizer();
  const markdownRenderer = createMarkdownRenderer();
  const notificationService = new MockNotificationService(state);
  const composeMarkdown = createComposeMarkdownService({
    bodyAccessor,
    htmlSanitizer,
    markdownRenderer,
    settingsStore,
  });
  const itemRenderer = createItemRenderer({
    bodyAccessor,
    htmlSanitizer,
    markdownRenderer,
    renderStateStore: new InMemoryRenderStateStore(),
    settingsStore,
  });
  const services: TaskpaneServices = {
    composeMarkdown,
    renderEntireDraft: async (): Promise<RenderItemResult> =>
      itemRenderer.renderItem(),
  };

  window.__MARKOUT_TASKPANE_MOCK__ = createController(state, settingsStore);
  document.documentElement.dataset.taskpaneMock = "ready";

  createRoot(rootElement).render(
    <TaskpaneApp
      initialMarkdownInput={INITIAL_MARKDOWN_INPUT}
      locale="en-US"
      notificationService={notificationService}
      services={services}
      settingsStore={settingsStore}
    />
  );
}

mountMockTaskpane();
