import { defaultStylesheet } from "../src/lib/config";
import { lintStylesheet } from "../src/lib/stylesheet-lint";

describe("stylesheet lint", () => {
  it("reports an empty stylesheet as a warning", () => {
    const result = lintStylesheet(" \n\t ");

    expect(result).toEqual({
      issues: [
        expect.objectContaining({
          code: "empty-stylesheet",
          severity: "warning",
        }),
      ],
      validRuleCount: 0,
    });
  });

  it("reports pseudo selectors and sanitizer-unsafe properties", () => {
    const result = lintStylesheet(`
      .mo::before { content: "x"; }
      a[href] { position: absolute; }
      p > a { color: inherit; }
      p + a { color: inherit; }
      p ~ a { color: inherit; }
    `);

    expect(result.issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: "pseudo-selector",
          severity: "warning",
        }),
        expect.objectContaining({
          code: "unsupported-selector",
          severity: "warning",
        }),
        expect.objectContaining({
          code: "sanitizer-unsafe-property",
          severity: "warning",
        }),
      ])
    );
  });

  it("reports invalid rules with unmatched braces", () => {
    const result = lintStylesheet(".mo { color: red; stray-fragment");

    expect(result.issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ code: "invalid-rule", severity: "error" }),
      ])
    );
  });

  it("reports rules with no valid declarations", () => {
    const result = lintStylesheet(".mo { color }");

    expect(result.issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ code: "invalid-rule", severity: "error" }),
      ])
    );
  });

  it("returns no findings for a supported stylesheet", () => {
    const result = lintStylesheet(`
      .mo { color: inherit; }
      p { margin-bottom: 1em; }
    `);

    expect(result.issues).toEqual([]);
    expect(result.validRuleCount).toBe(2);
  });

  it("keeps the shipped default stylesheet free of lint findings", () => {
    const result = lintStylesheet(defaultStylesheet);

    expect(result.issues).toEqual([]);
    expect(result.validRuleCount).toBeGreaterThan(0);
  });
});
