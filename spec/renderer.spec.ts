import { renderMarkdown } from "../src/lib/renderer";
import { installDomParser } from "./helpers";

describe("renderer", () => {
  beforeEach(() => {
    installDomParser();
  });

  it("generates HTML for basic markdown", async () => {
    const input = "# Example\nThis is a test";
    const expected =
      `<div class="mo markout-rendered">\n<h1>Example</h1>\n` +
      `<p>This is a test</p>\n</div>\n`;
    const output = await renderMarkdown({ css: "html {}", markdown: input });

    expect(output).toBe(expected);
  });

  it("renders supported HTML that is present inside markdown", async () => {
    const input = `# Example\nThis is a test with HTML elements\n<img src="http://example.com/img.png">`;
    const expected =
      `<div class="mo markout-rendered">\n<h1>Example</h1>\n<p>This is a test with HTML elements\n` +
      `<img src="http://example.com/img.png"></p>\n</div>\n`;
    const output = await renderMarkdown({ css: "html {}", markdown: input });

    expect(output).toBe(expected);
  });

  it("inlines stylesheet rules that target supported selectors", async () => {
    const output = await renderMarkdown({
      css: ".mo { color: rgb(1, 2, 3); } p { margin-top: 12px; }",
      markdown: "Paragraph text",
    });

    expect(output).toContain(
      `<div class="mo markout-rendered" style="color: rgb(1, 2, 3);">`
    );
    expect(output).toContain(`<p style="margin-top: 12px;">Paragraph text</p>`);
  });

  it("ignores pseudo selectors and invalid stylesheet fragments", async () => {
    const output = await renderMarkdown({
      css: ".mo::before { content: 'x'; } p { color: rgb(1, 2, 3); } .mo {",
      markdown: "Paragraph text",
    });

    expect(output).toContain(
      `<p style="color: rgb(1, 2, 3);">Paragraph text</p>`
    );
    expect(output).not.toContain("content:");
  });
});
