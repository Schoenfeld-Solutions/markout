import {
  containsMarkOutFragmentMarker,
  containsMarkOutFullRenderMarker,
  renderMarkdown,
} from "../src/lib/renderer";
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

  it("renders fragment markup with a scoped stylesheet host instead of inline styles", async () => {
    const output = await renderMarkdown({
      css: ".mo { color: rgb(1, 2, 3); } p { margin-top: 12px; }",
      markdown: "Paragraph text",
      mode: "fragment",
    });

    expect(output).toContain('<div class="markout-fragment-host">');
    expect(output).toContain('data-markout-styles="fragment"');
    expect(output).toContain(".markout-fragment-host .mo");
    expect(output).toContain('<div class="mo markout-fragment-rendered">');
    expect(output).not.toContain('style="color: rgb(1, 2, 3);"');
  });

  it("keeps default output host-inherit friendly for base text styles", async () => {
    const output = await renderMarkdown({
      css: `
        .mo {
          color: inherit;
          font-family: inherit;
          font-size: 1em;
          line-height: 1.5;
        }
        a {
          color: inherit;
        }
        p { margin-bottom: 1em; }
      `,
      markdown: "Body text",
    });

    expect(output).toContain('<div class="mo markout-rendered"');
    expect(output).toContain('<p style="margin-bottom: 1em;">Body text</p>');
    expect(output).toContain("font-family: inherit;");
    expect(output).toContain("font-size: 1em;");
    expect(output).not.toContain("rgb(36, 41, 46)");
    expect(output).not.toContain("font-size: 14px");
    expect(output).not.toContain("-apple-system");
    expect(output).not.toContain("nth-child");
  });

  it("detects full-render and fragment markers independently", () => {
    expect(
      containsMarkOutFullRenderMarker(
        '<div class="mo markout-rendered"><p>Rendered</p></div>'
      )
    ).toBe(true);
    expect(
      containsMarkOutFragmentMarker(
        '<div class="markout-fragment-host"><div class="mo markout-fragment-rendered">Fragment</div></div>'
      )
    ).toBe(true);
  });
});
