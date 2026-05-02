import { readFileSync } from "fs";
import path from "path";
import {
  getStrings,
  resolveLocale,
  resolveOfficeDisplayLanguage,
} from "../src/taskpane/i18n";
import { installOfficeEnvironment } from "./helpers";

describe("localization", () => {
  afterEach(() => {
    jest.restoreAllMocks();
    delete (globalThis as { Office?: typeof Office }).Office;
  });

  it("resolves German and English locales with an English fallback", () => {
    expect(resolveLocale("de-DE", "system", "en-US")).toBe("de-DE");
    expect(resolveLocale("de", "system", "en-US")).toBe("de-DE");
    expect(resolveLocale("en-GB", "system", "de-DE")).toBe("en-US");
    expect(resolveLocale("fr-FR", "system", "fr-FR")).toBe("en-US");
    expect(resolveLocale(undefined, "system", "de-DE")).toBe("de-DE");
  });

  it("honors explicit runtime language overrides", () => {
    expect(resolveLocale("fr-FR", "de-DE", "en-US")).toBe("de-DE");
    expect(resolveLocale("de-DE", "en-US", "de-DE")).toBe("en-US");
  });

  it("returns localized runtime strings", () => {
    expect(getStrings("de-DE").toolbar.settings).toBe("Einstellungen");
    expect(getStrings("en-US").toolbar.settings).toBe("Settings");
    expect(getStrings("de-DE").settings.languageTitle).toBe("Sprache");
    expect(getStrings("en-US").settings.languageSystem).toBe("Browser default");
    expect(getStrings("en-US").status.fileDecodeFailed("broken.md")).toBe(
      "broken.md could not be decoded."
    );
    expect(getStrings("de-DE").status.fileDecodeFailed("broken.md")).toBe(
      "broken.md konnte nicht dekodiert werden."
    );
    expect(getStrings("de-DE").status.fileReadFailed("broken.md")).toBe(
      "broken.md konnte nicht gelesen werden."
    );
    expect(getStrings("de-DE").status.dropFileInstruction).toContain(
      "Textdatei"
    );
    expect(getStrings("de-DE").status.fragmentInserted).toContain(
      "Body-Cursor"
    );
    expect(getStrings("de-DE").status.fragmentReplaced).toContain("Selektion");
    expect(getStrings("de-DE").status.helpHidden).toContain("ausgeblendet");
    expect(getStrings("de-DE").status.helpShown).toContain("wiederhergestellt");
    expect(getStrings("de-DE").status.introHidden).toContain("ausgeblendet");
    expect(getStrings("de-DE").status.introRestored).toContain(
      "wiederhergestellt"
    );
    expect(getStrings("de-DE").status.previewFailed).toContain("Vorschau");
    expect(getStrings("de-DE").status.selectionInspectionFailed).toContain(
      "Outlook"
    );
    expect(getStrings("de-DE").status.selectionInspectionSuccess).toContain(
      "Outlook"
    );
    expect(getStrings("de-DE").status.selectionRendered).toContain(
      "Body-Selektion"
    );
    expect(getStrings("de-DE").status.settingsUpdateFailed).toContain(
      "Einstellungen"
    );
    expect(getStrings("de-DE").status.stylesheetLoaded("style.md")).toBe(
      "style.md wurde in die Insert-Pane geladen."
    );
    expect(getStrings("en-US").status.themeUpdated("dark")).toBe(
      "Theme mode updated to dark."
    );
    expect(getStrings("de-DE").status.themeUpdated("dark")).toBe(
      "Theme-Modus auf dark gesetzt."
    );
    expect(getStrings("de-DE").status.draftUnchanged).toContain(
      "Nachrichten-Body"
    );
  });

  it("reads the Office display language when the host is available", () => {
    expect(resolveOfficeDisplayLanguage()).toBeUndefined();

    installOfficeEnvironment({ displayLanguage: "de-DE" });

    expect(resolveOfficeDisplayLanguage()).toBe("de-DE");
  });

  it("adds German locale overrides to every manifest variant", () => {
    const manifestFiles = [
      "manifest.xml",
      "manifest.beta.xml",
      "manifest-localhost.xml",
    ];

    for (const manifestFile of manifestFiles) {
      const manifest = readFileSync(
        path.join(__dirname, "..", manifestFile),
        "utf8"
      );

      expect(manifest).toContain('Override Locale="de-de"');
      expect(manifest).toContain("TaskpaneButton.Label");
      expect(manifest).toContain("TaskpaneButton.Tooltip");
    }
  });
});
