import type { SettingsStore } from "../lib/config";
import type { PreferenceState } from "./types";

export function readPreferences(settingsStore: SettingsStore): PreferenceState {
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

export function writePreferences(
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
