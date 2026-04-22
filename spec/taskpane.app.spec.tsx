/** @jest-environment jsdom */

import {
  buildToolbarPanels,
  getPanelAfterVisibilityChange,
  getRenderSelectionTooltip,
  isDarkColor,
  isInsertRenderedMarkdownDisabled,
  isRenderSelectionDisabled,
  readDroppedMarkdownFile,
  resolveSystemColorMode,
  resolveToolbarLayoutMode,
  supportsMarkdownFile,
} from "../src/taskpane/app";
import { getStrings } from "../src/taskpane/i18n";

describe("taskpane app helpers", () => {
  afterEach(() => {
    jest.restoreAllMocks();
  });

  it("builds toolbar panels in the expected order and respects visibility toggles", () => {
    const strings = getStrings("en-US");

    expect(
      buildToolbarPanels(
        {
          autoRender: false,
          creditsVisible: true,
          developerToolsEnabled: false,
          helpVisible: true,
          introDismissed: false,
          stylesheet: "",
          themeMode: "system",
        },
        strings
      ).map((panel) => panel.key)
    ).toEqual(["intro", "insert", "settings", "help", "credits"]);

    expect(
      buildToolbarPanels(
        {
          autoRender: false,
          creditsVisible: false,
          developerToolsEnabled: true,
          helpVisible: true,
          introDismissed: true,
          stylesheet: "",
          themeMode: "system",
        },
        strings
      ).map((panel) => panel.key)
    ).toEqual(["insert", "settings", "help", "developer"]);
  });

  it("switches toolbar layout mode based on available width", () => {
    expect(resolveToolbarLayoutMode(480, 5)).toBe("regular");
    expect(resolveToolbarLayoutMode(300, 5)).toBe("compact");
    expect(getStrings("en-US").tooltips.toolbarCompactHint("Insert")).toBe(
      "Open Insert"
    );
  });

  it("returns the settings panel when a visible toolbar panel is hidden while active", () => {
    expect(getPanelAfterVisibilityChange("help", "help", false)).toBe(
      "settings"
    );
    expect(getPanelAfterVisibilityChange("credits", "credits", false)).toBe(
      "settings"
    );
    expect(getPanelAfterVisibilityChange("developer", "developer", false)).toBe(
      "settings"
    );
    expect(getPanelAfterVisibilityChange("insert", "help", false)).toBe(
      "insert"
    );
  });

  it("disables render selection unless the body selection is available", () => {
    const strings = getStrings("en-US");

    expect(isRenderSelectionDisabled(false, "body-selection")).toBe(false);
    expect(isRenderSelectionDisabled(true, "body-selection")).toBe(true);
    expect(isRenderSelectionDisabled(false, "body-none")).toBe(true);
    expect(isRenderSelectionDisabled(false, "subject")).toBe(true);
    expect(isRenderSelectionDisabled(false, "unknown")).toBe(true);

    expect(getRenderSelectionTooltip(strings, "body-selection")).toContain(
      "currently selected Markdown text"
    );
    expect(getRenderSelectionTooltip(strings, "body-none")).toContain(
      "Select Markdown text"
    );
    expect(getRenderSelectionTooltip(strings, "subject")).toContain(
      "message body"
    );
    expect(getRenderSelectionTooltip(strings, "unknown")).toContain(
      "Selection state could not be read"
    );
  });

  it("disables fragment insertion until Markdown input is present", () => {
    expect(isInsertRenderedMarkdownDisabled(false, "")).toBe(true);
    expect(isInsertRenderedMarkdownDisabled(false, "   ")).toBe(true);
    expect(isInsertRenderedMarkdownDisabled(true, "## Fragment")).toBe(true);
    expect(isInsertRenderedMarkdownDisabled(false, "## Fragment")).toBe(false);
  });

  it("resolves light and dark themes from Office theme colors and browser fallback", () => {
    const originalMatchMedia = window.matchMedia;

    Object.defineProperty(window, "matchMedia", {
      configurable: true,
      value: jest.fn().mockReturnValue({
        addEventListener: jest.fn(),
        matches: true,
        removeEventListener: jest.fn(),
      }),
    });

    expect(isDarkColor("#111111")).toBe(true);
    expect(isDarkColor("#f5f5f5")).toBe(false);
    expect(resolveSystemColorMode({ bodyBackgroundColor: "#111111" })).toBe(
      "dark"
    );
    expect(resolveSystemColorMode({ bodyBackgroundColor: "#f5f5f5" })).toBe(
      "light"
    );
    expect(resolveSystemColorMode(undefined)).toBe("dark");

    Object.defineProperty(window, "matchMedia", {
      configurable: true,
      value: originalMatchMedia,
    });
  });

  it("supports markdown file detection helpers", () => {
    expect(supportsMarkdownFile(new File(["x"], "sample.md"))).toBe(true);
    expect(supportsMarkdownFile(new File(["x"], "sample.markdown"))).toBe(true);
    expect(supportsMarkdownFile(new File(["x"], "sample.txt"))).toBe(true);
    expect(supportsMarkdownFile(new File(["x"], "sample.html"))).toBe(false);
  });

  it("reads dropped files and surfaces read failures", async () => {
    class SuccessfulFileReader {
      public onerror: (() => void) | null = null;
      public onload: (() => void) | null = null;
      public result: string | ArrayBuffer | null = "## Loaded";

      public readAsText(): void {
        this.onload?.();
      }
    }

    class FailingFileReader {
      public onerror: (() => void) | null = null;
      public onload: (() => void) | null = null;
      public result: string | ArrayBuffer | null = null;

      public readAsText(): void {
        this.onerror?.();
      }
    }

    Object.defineProperty(window, "FileReader", {
      configurable: true,
      value: SuccessfulFileReader,
    });

    await expect(
      readDroppedMarkdownFile(new File(["ignored"], "loaded.md"))
    ).resolves.toBe("## Loaded");

    Object.defineProperty(window, "FileReader", {
      configurable: true,
      value: FailingFileReader,
    });

    await expect(
      readDroppedMarkdownFile(new File(["ignored"], "broken.md"))
    ).rejects.toThrow("MarkOut could not read broken.md.");
  });
});
