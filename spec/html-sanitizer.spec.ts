import { JSDOM } from "jsdom";
import { DefaultHtmlSanitizer } from "../src/lib/html-sanitizer";

const runtime = globalThis as typeof globalThis & {
  DOMParser: typeof DOMParser;
};
runtime.DOMParser = new JSDOM().window.DOMParser;

describe("html sanitizer", () => {
  it("removes active content and unsafe URLs", () => {
    const input =
      `<div onclick="alert(1)"><img src="https://example.com/example.png" onerror="alert(1)">` +
      `<a href="javascript:alert(1)" style="color: red; position: absolute">Test</a>` +
      `<script>alert(1)</script></div>`;

    const output = new DefaultHtmlSanitizer().sanitize(input);

    expect(output).toContain(
      `<div><img src="https://example.com/example.png"><a style="color: red">Test</a></div>`
    );
    expect(output).not.toContain("onclick");
    expect(output).not.toContain("onerror");
    expect(output).not.toContain("javascript:");
    expect(output).not.toContain("position");
    expect(output).not.toContain("<script");
  });

  it("drops SVG payloads and keeps supported cid image sources", () => {
    const input = `<div><svg onload="alert(1)"><circle></circle></svg><img src="cid:inline-image" alt="inline"></div>`;
    const output = new DefaultHtmlSanitizer().sanitize(input);

    expect(output).toContain(`<img src="cid:inline-image" alt="inline">`);
    expect(output).not.toContain("<svg");
  });

  it("keeps MarkOut fragment stylesheets and strips unsafe declarations", () => {
    const input =
      `<div class="markout-fragment-host">` +
      `<style data-markout-styles="fragment">` +
      `.markout-fragment-host .mo { color: rgb(1, 2, 3); background-image: url(javascript:alert(1)); }` +
      `</style><div class="mo markout-fragment-rendered">Text</div></div>`;
    const output = new DefaultHtmlSanitizer().sanitize(input);

    expect(output).toContain(`<style data-markout-styles="fragment">`);
    expect(output).toContain(`color: rgb(1, 2, 3)`);
    expect(output).not.toContain("background-image");
    expect(output).not.toContain("javascript:");
  });
});
