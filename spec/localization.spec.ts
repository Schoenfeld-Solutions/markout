import { readFileSync } from "fs";
import path from "path";
import { getStrings, resolveLocale } from "../src/taskpane/i18n";

describe("localization", () => {
  it("resolves German and English locales with an English fallback", () => {
    expect(resolveLocale("de-DE", "system", "en-US")).toBe("de-DE");
    expect(resolveLocale("en-GB", "system", "de-DE")).toBe("en-US");
    expect(resolveLocale("fr-FR", "system", "fr-FR")).toBe("en-US");
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
