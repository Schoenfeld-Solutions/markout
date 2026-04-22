import {
  createOfficeSettingsStore,
  defaultStylesheet,
} from "../src/lib/config";
import { FakeRoamingSettings, installOfficeEnvironment } from "./helpers";

describe("settings store", () => {
  beforeEach(() => {
    installOfficeEnvironment();
  });

  it("falls back to the default stylesheet when no custom theme is stored", () => {
    const settingsStore = createOfficeSettingsStore(new FakeRoamingSettings());

    expect(settingsStore.getStylesheet()).toBe(defaultStylesheet);
    expect(settingsStore.getAutoRender()).toBe(false);
    expect(settingsStore.getCreditsVisible()).toBe(true);
    expect(settingsStore.getDeveloperToolsEnabled()).toBe(false);
    expect(settingsStore.getHelpVisible()).toBe(true);
    expect(settingsStore.getIntroDismissed()).toBe(false);
    expect(settingsStore.getLanguagePreference()).toBe("system");
    expect(settingsStore.getThemeMode()).toBe("system");
  });

  it("falls back to defaults when invalid values are stored", () => {
    const roamingSettings = new FakeRoamingSettings();
    const settingsStore = createOfficeSettingsStore(roamingSettings);

    roamingSettings.set("markout.stylesheet", 42);
    roamingSettings.set("markout.autorender", "yes");
    roamingSettings.set("markout.creditsVisible", "yes");
    roamingSettings.set("markout.developerToolsEnabled", "yes");
    roamingSettings.set("markout.helpVisible", "yes");
    roamingSettings.set("markout.introDismissed", "yes");
    roamingSettings.set("markout.languagePreference", "fr-FR");
    roamingSettings.set("markout.themeMode", "sepia");

    expect(settingsStore.getStylesheet()).toBe(defaultStylesheet);
    expect(settingsStore.getAutoRender()).toBe(false);
    expect(settingsStore.getCreditsVisible()).toBe(true);
    expect(settingsStore.getDeveloperToolsEnabled()).toBe(false);
    expect(settingsStore.getHelpVisible()).toBe(true);
    expect(settingsStore.getIntroDismissed()).toBe(false);
    expect(settingsStore.getLanguagePreference()).toBe("system");
    expect(settingsStore.getThemeMode()).toBe("system");
  });

  it("persists stylesheet, auto-render, theme, intro, and developer settings", async () => {
    const roamingSettings = new FakeRoamingSettings();
    const settingsStore = createOfficeSettingsStore(roamingSettings);

    settingsStore.setStylesheet(".mo { color: rgb(1, 2, 3); }");
    settingsStore.setAutoRender(true);
    settingsStore.setCreditsVisible(false);
    settingsStore.setDeveloperToolsEnabled(true);
    settingsStore.setHelpVisible(false);
    settingsStore.setIntroDismissed(true);
    settingsStore.setLanguagePreference("de-DE");
    settingsStore.setThemeMode("dark");
    await settingsStore.save();

    expect(settingsStore.getStylesheet()).toBe(".mo { color: rgb(1, 2, 3); }");
    expect(settingsStore.getAutoRender()).toBe(true);
    expect(settingsStore.getCreditsVisible()).toBe(false);
    expect(settingsStore.getDeveloperToolsEnabled()).toBe(true);
    expect(settingsStore.getHelpVisible()).toBe(false);
    expect(settingsStore.getIntroDismissed()).toBe(true);
    expect(settingsStore.getLanguagePreference()).toBe("de-DE");
    expect(settingsStore.getThemeMode()).toBe("dark");
  });

  it("falls back to in-memory settings when roaming settings are unavailable", async () => {
    const settingsStore = createOfficeSettingsStore(undefined);

    expect(settingsStore.getStylesheet()).toBe(defaultStylesheet);
    expect(settingsStore.getAutoRender()).toBe(false);
    expect(settingsStore.getCreditsVisible()).toBe(true);
    expect(settingsStore.getDeveloperToolsEnabled()).toBe(false);
    expect(settingsStore.getHelpVisible()).toBe(true);
    expect(settingsStore.getIntroDismissed()).toBe(false);
    expect(settingsStore.getLanguagePreference()).toBe("system");
    expect(settingsStore.getThemeMode()).toBe("system");

    settingsStore.setStylesheet(".mo { color: rgb(4, 5, 6); }");
    settingsStore.setAutoRender(true);
    settingsStore.setCreditsVisible(false);
    settingsStore.setDeveloperToolsEnabled(true);
    settingsStore.setHelpVisible(false);
    settingsStore.setIntroDismissed(true);
    settingsStore.setLanguagePreference("en-US");
    settingsStore.setThemeMode("light");
    await expect(settingsStore.save()).resolves.toBeUndefined();

    expect(settingsStore.getStylesheet()).toBe(".mo { color: rgb(4, 5, 6); }");
    expect(settingsStore.getAutoRender()).toBe(true);
    expect(settingsStore.getCreditsVisible()).toBe(false);
    expect(settingsStore.getDeveloperToolsEnabled()).toBe(true);
    expect(settingsStore.getHelpVisible()).toBe(false);
    expect(settingsStore.getIntroDismissed()).toBe(true);
    expect(settingsStore.getLanguagePreference()).toBe("en-US");
    expect(settingsStore.getThemeMode()).toBe("light");
  });

  it("normalizes empty stylesheet updates back to the default stylesheet", () => {
    const settingsStore = createOfficeSettingsStore(new FakeRoamingSettings());

    settingsStore.setStylesheet("   ");

    expect(settingsStore.getStylesheet()).toBe(defaultStylesheet);
    expect(defaultStylesheet).toContain("line-height: 1.5;");
    expect(defaultStylesheet).not.toContain("font: inherit;");
    expect(defaultStylesheet).not.toContain("color: inherit;");
  });

  it("surfaces save failures from roaming settings", async () => {
    const roamingSettings = new FakeRoamingSettings();
    const settingsStore = createOfficeSettingsStore(roamingSettings);
    roamingSettings.failNextSave = true;

    await expect(settingsStore.save()).rejects.toMatchObject({
      message: "Roaming settings save failed.",
      name: "RoamingSettingsSaveError",
    });
  });
});
