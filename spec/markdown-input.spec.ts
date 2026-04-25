import { normalizeMarkdownInput } from "../src/taskpane/markdown-input";

describe("markdown input normalization", () => {
  it("normalizes paste-only non-breaking indentation to markdown spaces", () => {
    expect(
      normalizeMarkdownInput(
        [
          "# hi",
          "",
          "ich bin",
          "- cool",
          "\u00a0\u00a0- super cool",
          "\u202f\u202f- narrow cool",
          "\u2007\u2007- figure cool",
          "- cool",
        ].join("\n")
      )
    ).toBe(
      [
        "# hi",
        "",
        "ich bin",
        "- cool",
        "  - super cool",
        "  - narrow cool",
        "  - figure cool",
        "- cool",
      ].join("\n")
    );
  });

  it("preserves non-breaking spaces outside line-leading indentation", () => {
    expect(normalizeMarkdownInput("Text\u00a0inside\n- item\u00a0suffix")).toBe(
      "Text\u00a0inside\n- item\u00a0suffix"
    );
  });

  it("preserves existing line endings while normalizing leading paste spaces", () => {
    expect(normalizeMarkdownInput("- item\r\n\u00a0\u00a0- child")).toBe(
      "- item\r\n  - child"
    );
  });
});
