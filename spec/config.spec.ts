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
  });

  it("falls back to defaults when invalid values are stored", () => {
    const roamingSettings = new FakeRoamingSettings();
    const settingsStore = createOfficeSettingsStore(roamingSettings);

    roamingSettings.set("markout.stylesheet", 42);
    roamingSettings.set("markout.autorender", "yes");

    expect(settingsStore.getStylesheet()).toBe(defaultStylesheet);
    expect(settingsStore.getAutoRender()).toBe(false);
  });

  it("persists stylesheet and auto-render settings", async () => {
    const roamingSettings = new FakeRoamingSettings();
    const settingsStore = createOfficeSettingsStore(roamingSettings);

    settingsStore.setStylesheet(".mo { color: rgb(1, 2, 3); }");
    settingsStore.setAutoRender(true);
    await settingsStore.save();

    expect(settingsStore.getStylesheet()).toBe(".mo { color: rgb(1, 2, 3); }");
    expect(settingsStore.getAutoRender()).toBe(true);
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
