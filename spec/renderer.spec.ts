/** @jest-environment jsdom */
import {
  containsMarkOutFragmentMarker,
  containsMarkOutFullRenderMarker,
  renderMarkdown,
} from "../src/lib/renderer";
import { defaultStylesheet } from "../src/lib/config";
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

  it("renders emoji shortcodes and shortcuts with the markdown-it emoji plugin", async () => {
    const output = await renderMarkdown({
      css: "html {}",
      markdown: "Emoji :smile: :-)",
    });

    expect(output).toContain("<p>Emoji 😄 😃</p>");
    expect(output).not.toContain(":smile:");
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
          line-height: 1.5;
        }
        a {
          text-decoration: underline;
        }
        h1 {
          font-size: 1.75em;
          font-weight: bold;
        }
        p { margin-bottom: 1em; }
      `,
      markdown: "# Title\n\nBody text with [a link](https://example.com).",
    });

    expect(output).toContain(
      '<div class="mo markout-rendered" style="line-height: 1.5;">'
    );
    expect(output).toContain(
      '<h1 style="font-size: 1.75em; font-weight: bold;">Title</h1>'
    );
    expect(output).toContain(
      '<p style="margin-bottom: 1em;">Body text with <a href="https://example.com" style="text-decoration: underline;">a link</a>.</p>'
    );
    expect(output).not.toContain("font-family: inherit;");
    expect(output).not.toContain("font-size: 1em;");
    expect(output).not.toContain("color: inherit;");
    expect(output).not.toContain("rgb(36, 41, 46)");
    expect(output).not.toContain("font-size: 14px");
    expect(output).not.toContain("-apple-system");
    expect(output).not.toContain("border-bottom");
    expect(output).not.toContain("nth-child");
  });

  it("keeps lists compact without adding artificial spacing between items", async () => {
    const output = await renderMarkdown({
      css: `
        ul,
        ol {
          margin: 0.9em 0;
          padding-left: 1.5em;
        }

        li {
          margin: 0;
        }

        li p {
          margin: 0;
        }
      `,
      markdown: "- one\n- two\n- three",
    });

    expect(output).toContain(
      '<ul style="margin: 0.9em 0px; padding-left: 1.5em;">'
    );
    expect(output).toContain('<li style="margin: 0px;">one</li>');
    expect(output).toContain('<li style="margin: 0px;">two</li>');
    expect(output).toContain('<li style="margin: 0px;">three</li>');
  });

  it("keeps nested lists visually attached to their parent item", async () => {
    const output = await renderMarkdown({
      css: defaultStylesheet,
      markdown: "- parent\n  - child\n    - grandchild",
    });
    const documentFragment = new DOMParser().parseFromString(
      output,
      "text/html"
    );
    const nestedList = documentFragment.body.querySelector("li > ul");

    expect(nestedList?.getAttribute("style")).toContain(
      "margin: 0px !important"
    );
    expect(nestedList?.getAttribute("style")).toContain("padding-left: 1em");
  });

  it("keeps blank-line list grouping under markdown-it semantics", async () => {
    const output = await renderMarkdown({
      css: "html {}",
      markdown: "- one\n\n\n- two",
    });

    expect(output.match(/<ul>/g) ?? []).toHaveLength(1);
    expect(output).toContain("<p>one</p>");
    expect(output).toContain("<p>two</p>");
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
