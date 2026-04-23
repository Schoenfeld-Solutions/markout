import type { SettingsStore } from "../src/lib/config";
import { readPreferences, writePreferences } from "../src/taskpane/preferences";

function createSettingsStore(
  overrides: Partial<{
    autoRender: boolean;
    creditsVisible: boolean;
    developerToolsEnabled: boolean;
    helpVisible: boolean;
    introDismissed: boolean;
    languagePreference: "de-DE" | "en-US" | "system";
    stylesheet: string;
    themeMode: "dark" | "light" | "system";
  }> = {}
): SettingsStore {
  return {
    getAutoRender: () => overrides.autoRender ?? false,
    getCreditsVisible: () => overrides.creditsVisible ?? true,
    getDeveloperToolsEnabled: () => overrides.developerToolsEnabled ?? false,
    getHelpVisible: () => overrides.helpVisible ?? true,
    getIntroDismissed: () => overrides.introDismissed ?? false,
    getLanguagePreference: () => overrides.languagePreference ?? "system",
    getStylesheet: () => overrides.stylesheet ?? "",
    getThemeMode: () => overrides.themeMode ?? "system",
    hasStylesheetMigrationPending: () => false,
    save: jest.fn().mockResolvedValue(undefined),
    setAutoRender: jest.fn(),
    setCreditsVisible: jest.fn(),
    setDeveloperToolsEnabled: jest.fn(),
    setHelpVisible: jest.fn(),
    setIntroDismissed: jest.fn(),
    setLanguagePreference: jest.fn(),
    setStylesheet: jest.fn(),
    setThemeMode: jest.fn(),
  };
}

describe("taskpane preferences", () => {
  it("reads taskpane preference state from the settings store", () => {
    const settingsStore = createSettingsStore({
      autoRender: true,
      creditsVisible: false,
      developerToolsEnabled: true,
      helpVisible: false,
      introDismissed: true,
      languagePreference: "de-DE",
      stylesheet: ".mo { color: red; }",
      themeMode: "dark",
    });

    expect(readPreferences(settingsStore)).toEqual({
      autoRender: true,
      creditsVisible: false,
      developerToolsEnabled: true,
      helpVisible: false,
      introDismissed: true,
      languagePreference: "de-DE",
      stylesheet: ".mo { color: red; }",
      themeMode: "dark",
    });
  });

  it("writes every preference field back to the settings store", () => {
    const settingsStore = createSettingsStore();

    writePreferences(settingsStore, {
      autoRender: true,
      creditsVisible: false,
      developerToolsEnabled: true,
      helpVisible: false,
      introDismissed: true,
      languagePreference: "en-US",
      stylesheet: ".mo { color: blue; }",
      themeMode: "light",
    });

    expect(settingsStore.setAutoRender).toHaveBeenCalledWith(true);
    expect(settingsStore.setCreditsVisible).toHaveBeenCalledWith(false);
    expect(settingsStore.setDeveloperToolsEnabled).toHaveBeenCalledWith(true);
    expect(settingsStore.setHelpVisible).toHaveBeenCalledWith(false);
    expect(settingsStore.setIntroDismissed).toHaveBeenCalledWith(true);
    expect(settingsStore.setLanguagePreference).toHaveBeenCalledWith("en-US");
    expect(settingsStore.setStylesheet).toHaveBeenCalledWith(
      ".mo { color: blue; }"
    );
    expect(settingsStore.setThemeMode).toHaveBeenCalledWith("light");
  });
});
