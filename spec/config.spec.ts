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
    expect(settingsStore.hasStylesheetMigrationPending()).toBe(false);
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
    expect(settingsStore.hasStylesheetMigrationPending()).toBe(false);
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
    expect(settingsStore.hasStylesheetMigrationPending()).toBe(false);
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
    expect(settingsStore.hasStylesheetMigrationPending()).toBe(false);

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
    expect(settingsStore.hasStylesheetMigrationPending()).toBe(false);
  });

  it("normalizes empty stylesheet updates back to the default stylesheet", () => {
    const settingsStore = createOfficeSettingsStore(new FakeRoamingSettings());

    settingsStore.setStylesheet("   ");

    expect(settingsStore.getStylesheet()).toBe(defaultStylesheet);
    expect(settingsStore.hasStylesheetMigrationPending()).toBe(false);
    expect(defaultStylesheet).toContain("color: inherit;");
    expect(defaultStylesheet).toContain("font-family: inherit;");
    expect(defaultStylesheet).toContain("font-size: 1em;");
    expect(defaultStylesheet).not.toContain("font-size: 14px;");
    expect(defaultStylesheet).not.toContain("rgb(36,41,46)");
  });

  it("aggressively migrates legacy MarkOut defaults to the current host-inherit preset", async () => {
    const roamingSettings = new FakeRoamingSettings();
    const legacyDefault = `
      .mo {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
        font-size: 14px;
        color: rgb(36, 41, 46);
      }

      blockquote::before, blockquote::after, q::before, q::after {
        content: none;
      }

      table tr:nth-child(2n) {
        background-color: #F8F8F8;
      }
    `;

    roamingSettings.set("markout.stylesheet", legacyDefault);

    const settingsStore = createOfficeSettingsStore(roamingSettings);

    expect(settingsStore.getStylesheet()).toBe(defaultStylesheet);
    expect(settingsStore.hasStylesheetMigrationPending()).toBe(true);

    settingsStore.setStylesheet(settingsStore.getStylesheet());
    await settingsStore.save();

    const persistedStore = createOfficeSettingsStore(roamingSettings);

    expect(roamingSettings.get("markout.stylesheetPreset")).toBe(
      "default-host-inherit-v1"
    );
    expect(persistedStore.getStylesheet()).toBe(defaultStylesheet);
    expect(persistedStore.hasStylesheetMigrationPending()).toBe(false);
  });

  it("keeps obvious user custom css instead of migrating it", () => {
    const roamingSettings = new FakeRoamingSettings();
    const customStylesheet = `
      .mo {
        line-height: 1.8;
      }

      .signature-note {
        color: rgb(120, 40, 160);
        font-style: italic;
      }
    `;

    roamingSettings.set("markout.stylesheet", customStylesheet);

    const settingsStore = createOfficeSettingsStore(roamingSettings);

    expect(settingsStore.getStylesheet()).toBe(customStylesheet);
    expect(settingsStore.hasStylesheetMigrationPending()).toBe(false);
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
