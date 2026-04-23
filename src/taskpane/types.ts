import type { ReactElement } from "react";
import type { SelectionSource } from "../lib/body-accessor";
import type { ComposeMarkdownService } from "../lib/compose-markdown";
import type { ComposeNotificationService } from "../lib/compose-notifications";
import type {
  LanguagePreference,
  SettingsStore,
  ThemeMode,
} from "../lib/config";
import type { RenderItemResult } from "../lib/item";
import type { SupportedLocale } from "./i18n";

export type PanelKey =
  | "credits"
  | "developer"
  | "help"
  | "insert"
  | "intro"
  | "settings";

export type ToolbarLayoutMode = "compact" | "regular";

export interface PreferenceState {
  autoRender: boolean;
  creditsVisible: boolean;
  developerToolsEnabled: boolean;
  helpVisible: boolean;
  introDismissed: boolean;
  languagePreference: LanguagePreference;
  stylesheet: string;
  themeMode: ThemeMode;
}

export interface PanelMessageState {
  body: string;
  dismissible?: boolean;
  intent: "error" | "info" | "success" | "warning";
}

export interface SelectionDebugState {
  hasSelection: boolean;
  source: SelectionSource;
  textPreview: string;
}

export type SelectionAvailability =
  | "body-none"
  | "body-selection"
  | "subject"
  | "unknown";

export interface SelectionState {
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
  initialMarkdownInput?: string;
  locale?: SupportedLocale;
  notificationService?: ComposeNotificationService;
  services: TaskpaneServices;
  settingsStore: SettingsStore;
}
