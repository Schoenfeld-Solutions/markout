import { act } from "react";
import { createRoot, type Root } from "react-dom/client";
import type { ComposeNotificationService } from "../src/lib/compose-notifications";
import type {
  LanguagePreference,
  SettingsStore,
  ThemeMode,
} from "../src/lib/config";
import type { RenderItemResult } from "../src/lib/item";
import type { DiagnosticSink } from "../src/lib/runtime";
import { TaskpaneApp, type TaskpaneServices } from "../src/taskpane/app";
import type { SupportedLocale } from "../src/taskpane/i18n";

interface SettingsState {
  autoRender: boolean;
  creditsVisible: boolean;
  developerToolsEnabled: boolean;
  helpVisible: boolean;
  introDismissed: boolean;
  languagePreference: LanguagePreference;
  stylesheet: string;
  stylesheetMigrationPending: boolean;
  themeMode: ThemeMode;
}

const DEFAULT_SETTINGS_STATE: SettingsState = {
  autoRender: false,
  creditsVisible: true,
  developerToolsEnabled: false,
  helpVisible: true,
  introDismissed: false,
  languagePreference: "system",
  stylesheet: "",
  stylesheetMigrationPending: false,
  themeMode: "system",
};

export type MutableSettingsStore = SettingsStore & {
  readonly state: SettingsState;
  save: jest.Mock;
  setAutoRender: jest.Mock;
  setCreditsVisible: jest.Mock;
  setDeveloperToolsEnabled: jest.Mock;
  setHelpVisible: jest.Mock;
  setIntroDismissed: jest.Mock;
  setLanguagePreference: jest.Mock;
  setStylesheet: jest.Mock;
  setThemeMode: jest.Mock;
};

export type TaskpaneServiceMocks = TaskpaneServices & {
  composeMarkdown: TaskpaneServices["composeMarkdown"] & {
    getSelection: jest.Mock;
    insertRenderedMarkdown: jest.Mock;
    renderPreview: jest.Mock;
    renderSelection: jest.Mock;
  };
  renderEntireDraft: jest.Mock;
};

export type NotificationServiceMocks = ComposeNotificationService & {
  clearAutoRenderDismissed: jest.Mock;
  clearAutoRenderNotification: jest.Mock;
  clearTransientNotification: jest.Mock;
  hasAutoRenderBeenDismissed: jest.Mock;
  markAutoRenderDismissed: jest.Mock;
  onAutoRenderDismiss: jest.Mock;
  showAutoRenderNotification: jest.Mock;
  showTransientNotification: jest.Mock;
};

export function createMutableSettingsStore(
  overrides: Partial<SettingsState> = {}
): MutableSettingsStore {
  const state: SettingsState = {
    ...DEFAULT_SETTINGS_STATE,
    ...overrides,
  };

  return {
    state,
    getAutoRender: () => state.autoRender,
    getCreditsVisible: () => state.creditsVisible,
    getDeveloperToolsEnabled: () => state.developerToolsEnabled,
    getHelpVisible: () => state.helpVisible,
    getIntroDismissed: () => state.introDismissed,
    getLanguagePreference: () => state.languagePreference,
    getStylesheet: () => state.stylesheet,
    getThemeMode: () => state.themeMode,
    hasStylesheetMigrationPending: () => state.stylesheetMigrationPending,
    save: jest.fn().mockResolvedValue(undefined),
    setAutoRender: jest.fn((value: boolean) => {
      state.autoRender = value;
    }),
    setCreditsVisible: jest.fn((value: boolean) => {
      state.creditsVisible = value;
    }),
    setDeveloperToolsEnabled: jest.fn((value: boolean) => {
      state.developerToolsEnabled = value;
    }),
    setHelpVisible: jest.fn((value: boolean) => {
      state.helpVisible = value;
    }),
    setIntroDismissed: jest.fn((value: boolean) => {
      state.introDismissed = value;
    }),
    setLanguagePreference: jest.fn((value: LanguagePreference) => {
      state.languagePreference = value;
    }),
    setStylesheet: jest.fn((value: string) => {
      state.stylesheet = value;
    }),
    setThemeMode: jest.fn((value: ThemeMode) => {
      state.themeMode = value;
    }),
  };
}

export function createTaskpaneServices(
  overrides: Partial<{
    getSelection: jest.Mock;
    insertRenderedMarkdown: jest.Mock;
    renderEntireDraft: jest.Mock;
    renderPreview: jest.Mock;
    renderSelection: jest.Mock;
  }> = {}
): TaskpaneServiceMocks {
  return {
    composeMarkdown: {
      getSelection:
        overrides.getSelection ??
        jest.fn().mockResolvedValue({
          hasSelection: false,
          html: null,
          source: "body",
          text: "",
        }),
      insertRenderedMarkdown:
        overrides.insertRenderedMarkdown ??
        jest.fn().mockResolvedValue("inserted"),
      renderPreview:
        overrides.renderPreview ??
        jest.fn().mockResolvedValue("<p>preview</p>"),
      renderSelection:
        overrides.renderSelection ?? jest.fn().mockResolvedValue(undefined),
    },
    renderEntireDraft:
      overrides.renderEntireDraft ??
      jest.fn<Promise<RenderItemResult>, []>().mockResolvedValue("rendered"),
  };
}

export function createNotificationService(
  overrides: Partial<NotificationServiceMocks> = {}
): NotificationServiceMocks {
  return {
    clearAutoRenderDismissed: jest.fn().mockResolvedValue(undefined),
    clearAutoRenderNotification: jest.fn().mockResolvedValue(undefined),
    clearTransientNotification: jest.fn().mockResolvedValue(undefined),
    hasAutoRenderBeenDismissed: jest.fn().mockResolvedValue(false),
    markAutoRenderDismissed: jest.fn().mockResolvedValue(undefined),
    onAutoRenderDismiss: jest.fn(),
    showAutoRenderNotification: jest.fn().mockResolvedValue("outlook"),
    showTransientNotification: jest.fn().mockResolvedValue("outlook"),
    ...overrides,
  };
}

export interface MountedTaskpaneApp {
  cleanup: () => void;
  click: (selector: string) => Promise<void>;
  container: HTMLElement;
  notificationService: NotificationServiceMocks;
  root: Root;
  services: TaskpaneServiceMocks;
  settingsStore: MutableSettingsStore;
  typeMarkdown: (value: string) => Promise<void>;
}

export async function mountTaskpaneApp(
  options: Partial<{
    diagnosticSink: DiagnosticSink;
    initialMarkdownInput: string;
    locale: SupportedLocale;
    notificationService: NotificationServiceMocks;
    services: TaskpaneServiceMocks;
    settingsStore: MutableSettingsStore;
  }> = {}
): Promise<MountedTaskpaneApp> {
  const settingsStore = options.settingsStore ?? createMutableSettingsStore();
  const services = options.services ?? createTaskpaneServices();
  const notificationService =
    options.notificationService ?? createNotificationService();
  const restoreBrowserLayoutApis = ensureBrowserLayoutApis();

  document.body.innerHTML = '<div id="root"></div>';
  const container = document.getElementById("root");

  if (container === null) {
    throw new Error("Expected a taskpane test container.");
  }

  const root = createRoot(container);

  const appProps = {
    locale: options.locale ?? "en-US",
    notificationService,
    services,
    settingsStore,
    ...(options.diagnosticSink === undefined
      ? {}
      : { diagnosticSink: options.diagnosticSink }),
    ...(options.initialMarkdownInput === undefined
      ? {}
      : { initialMarkdownInput: options.initialMarkdownInput }),
  };

  act(() => {
    root.render(<TaskpaneApp {...appProps} />);
  });
  await flushTaskpane();

  return {
    cleanup: () => {
      act(() => {
        root.unmount();
      });
      restoreBrowserLayoutApis();
    },
    click: async (selector: string) => {
      const element = container.querySelector<HTMLElement>(selector);
      if (element === null) {
        throw new Error(`Expected ${selector} to exist.`);
      }

      act(() => {
        element.click();
      });
      await flushTaskpane();
    },
    container,
    notificationService,
    root,
    services,
    settingsStore,
    typeMarkdown: async (value: string) => {
      const textarea =
        container.querySelector<HTMLTextAreaElement>("#markdown-input");
      if (textarea === null) {
        throw new Error("Expected markdown input to exist.");
      }

      await setNativeTextareaValue(textarea, value);
    },
  };
}

export async function flushTaskpane(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
  await Promise.resolve();
  await new Promise<void>((resolve) => {
    window.setTimeout(resolve, 0);
  });
  await Promise.resolve();
}

export async function setNativeTextareaValue(
  textarea: HTMLTextAreaElement,
  value: string
): Promise<void> {
  const valueSetter = Object.getOwnPropertyDescriptor(
    HTMLTextAreaElement.prototype,
    "value"
  )?.set;

  act(() => {
    valueSetter?.call(textarea, value);
    textarea.dispatchEvent(new Event("input", { bubbles: true }));
    textarea.dispatchEvent(new Event("change", { bubbles: true }));
  });
  await flushTaskpane();
}

function ensureBrowserLayoutApis(): () => void {
  const originalMatchMedia = window.matchMedia;
  const originalResizeObserver = window.ResizeObserver;
  const originalClearInterval = window.clearInterval;
  const originalSetInterval = window.setInterval;
  const originalVisibilityStateDescriptor = Object.getOwnPropertyDescriptor(
    Document.prototype,
    "visibilityState"
  );

  Object.defineProperty(window, "matchMedia", {
    configurable: true,
    value: jest.fn().mockReturnValue({
      addEventListener: jest.fn(),
      addListener: jest.fn(),
      dispatchEvent: jest.fn().mockReturnValue(false),
      matches: false,
      media: "(prefers-color-scheme: dark)",
      onchange: null,
      removeEventListener: jest.fn(),
      removeListener: jest.fn(),
    }),
  });

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
  Object.defineProperty(window, "setInterval", {
    configurable: true,
    value: jest.fn(() => 0),
  });
  Object.defineProperty(window, "clearInterval", {
    configurable: true,
    value: jest.fn(),
  });
  Object.defineProperty(document, "visibilityState", {
    configurable: true,
    value: "visible",
  });

  return () => {
    Object.defineProperty(window, "matchMedia", {
      configurable: true,
      value: originalMatchMedia,
    });
    Object.defineProperty(window, "ResizeObserver", {
      configurable: true,
      value: originalResizeObserver,
    });
    Object.defineProperty(window, "setInterval", {
      configurable: true,
      value: originalSetInterval,
    });
    Object.defineProperty(window, "clearInterval", {
      configurable: true,
      value: originalClearInterval,
    });
    if (originalVisibilityStateDescriptor === undefined) {
      Reflect.deleteProperty(document, "visibilityState");
    } else {
      Object.defineProperty(
        Document.prototype,
        "visibilityState",
        originalVisibilityStateDescriptor
      );
      Reflect.deleteProperty(document, "visibilityState");
    }
  };
}
