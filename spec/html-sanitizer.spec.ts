/** @jest-environment jsdom */
import { DefaultHtmlSanitizer } from "../src/lib/html-sanitizer";

describe("html sanitizer", () => {
  const sanitizer = new DefaultHtmlSanitizer();

  it("removes active content and unsafe URLs", () => {
    const input =
      `<div onclick="alert(1)"><img src="https://example.com/example.png" onerror="alert(1)">` +
      `<a href="javascript:alert(1)" style="color: red; position: absolute">Test</a>` +
      `<script>alert(1)</script></div>`;

    const output = sanitizer.sanitize(input);

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
    const output = sanitizer.sanitize(input);

    expect(output).toContain(`<img src="cid:inline-image" alt="inline">`);
    expect(output).not.toContain("<svg");
  });

  it("keeps MarkOut fragment stylesheets and strips unsafe declarations", () => {
    const input =
      `<div class="markout-fragment-host">` +
      `<style data-markout-styles="fragment">` +
      `.markout-fragment-host .mo { color: rgb(1, 2, 3); background-image: url(javascript:alert(1)); }` +
      `</style><div class="mo markout-fragment-rendered">Text</div></div>`;
    const output = sanitizer.sanitize(input);

    expect(output).toContain(`<style data-markout-styles="fragment">`);
    expect(output).toContain(`color: rgb(1, 2, 3)`);
    expect(output).not.toContain("background-image");
    expect(output).not.toContain("javascript:");
  });

  it.each([
    [`<a href="jav&#x61;script:alert(1)">encoded</a>`, `<a>encoded</a>`],
    [`<a href="java&#10;script:alert(1)">newline</a>`, `<a>newline</a>`],
    [`<a href="vbscript:msgbox(1)">legacy</a>`, `<a>legacy</a>`],
    [
      `<a href="mailto:user@example.com">mail</a>`,
      `<a href="mailto:user@example.com">mail</a>`,
    ],
    [
      `<a href="/relative/path">relative</a>`,
      `<a href="/relative/path">relative</a>`,
    ],
  ])("sanitizes URL payload %s", (input, expected) => {
    expect(sanitizer.sanitize(input)).toBe(expected);
  });

  it("allows only safe raster data image sources", () => {
    const input =
      `<p>` +
      `<img src="data:image/png;base64,iVBORw0KGgo=" alt="safe">` +
      `<img src="data:image/svg+xml;base64,PHN2ZyBvbmxvYWQ9YWxlcnQoMSk+" alt="svg">` +
      `<img src="data:text/html;base64,PHNjcmlwdD5hbGVydCgxKTwvc2NyaXB0Pg==" alt="html">` +
      `</p>`;

    const output = sanitizer.sanitize(input);

    expect(output).toContain(
      `<img src="data:image/png;base64,iVBORw0KGgo=" alt="safe">`
    );
    expect(output).toContain(`<img alt="svg">`);
    expect(output).toContain(`<img alt="html">`);
    expect(output).not.toContain("svg+xml");
    expect(output).not.toContain("data:text/html");
  });

  it("strips CSS exfiltration and execution primitives while keeping safe declarations", () => {
    const input =
      `<p style="` +
      `color: rgb(10, 20, 30); ` +
      `background: url(https://attacker.test/pixel); ` +
      `width: expression(alert(1)); ` +
      `behavior: url(#default#time2); ` +
      `-moz-binding: url(https://attacker.test/xbl); ` +
      `@import: url(https://attacker.test/style.css);` +
      `">Styled</p>`;

    const output = sanitizer.sanitize(input);

    expect(output).toBe(`<p style="color: rgb(10, 20, 30)">Styled</p>`);
  });

  it("drops unsafe SVG, MathML, and embedded document content", () => {
    const input =
      `<div>before` +
      `<svg><a href="javascript:alert(1)">svg text</a></svg>` +
      `<math><mtext onclick="alert(1)">math text</mtext></math>` +
      `<iframe srcdoc="<script>alert(1)</script>"></iframe>` +
      `<object data="javascript:alert(1)">object text</object>` +
      `after</div>`;

    const output = sanitizer.sanitize(input);

    expect(output).toBe(`<div>beforeafter</div>`);
    expect(output).not.toMatch(/<svg|<math|<iframe|<object|javascript:/i);
  });

  it("normalizes safe global and table attributes while dropping invalid values", () => {
    const input =
      `<table><tbody><tr>` +
      `<td align="RIGHT" valign="center" colspan="abc" rowspan="2" title=" Cell ">Cell</td>` +
      `</tr></tbody></table>` +
      `<p class="safe-name unsafe:bad" dir="RTL" lang="de-DE">Text</p>`;

    const output = sanitizer.sanitize(input);

    expect(output).toContain(
      `<td align="right" rowspan="2" title="Cell">Cell</td>`
    );
    expect(output).toContain(`<p dir="rtl" lang="de-DE">Text</p>`);
    expect(output).not.toContain("valign");
    expect(output).not.toContain("colspan");
    expect(output).not.toContain("class=");
  });

  it("keeps Outlook namespaced text but drops document-level active content", () => {
    const input =
      `<div><o:p>Outlook marker</o:p><meta http-equiv="refresh" content="0;url=javascript:alert(1)">` +
      `<link rel="stylesheet" href="https://attacker.test/style.css">` +
      `<b onclick="alert(1)">Bold text</b></div>`;

    const output = sanitizer.sanitize(input);

    expect(output).toBe(`<div>Outlook markerBold text</div>`);
  });

  it("keeps only safe MarkOut fragment stylesheet rules", () => {
    const input =
      `<div><style data-markout-styles="fragment">` +
      `.mo:hover { color: red; }` +
      `.mo { color: rgb(1, 2, 3); }` +
      `.mo a { background-image: url(javascript:alert(1)); }` +
      `</style>` +
      `<style>.mo { color: red; }</style></div>`;

    const output = sanitizer.sanitize(input);

    expect(output).toBe(
      `<div><style data-markout-styles="fragment">.mo { color: rgb(1, 2, 3) }</style></div>`
    );
  });

  it("drops malformed and non-element active nodes without losing safe text", () => {
    const input =
      `safe<!-- comment --><script>alert(1)</script>` +
      `<span title="ok"><broken onclick="alert(1)">nested</span>`;

    const output = sanitizer.sanitize(input);

    expect(output).toBe(`safe<span title="ok">nested</span>`);
  });

  it("satisfies deterministic safety invariants across mixed payloads", () => {
    const payloads = [
      `<img src="x" onerror="alert(1)">`,
      `<a href="javascript:alert(1)" onclick="alert(1)">click</a>`,
      `<p style="background-image:url(javascript:alert(1));color:blue">x</p>`,
      `<svg><script>alert(1)</script></svg>`,
      `<math><mi href="javascript:alert(1)">x</mi></math>`,
      `<style data-markout-styles="fragment">.mo { behavior: url(x); }</style>`,
      `<iframe src="https://attacker.test"></iframe>`,
      `<object data="data:text/html,<script>alert(1)</script>"></object>`,
    ];

    for (let index = 0; index < 64; index += 1) {
      const input = payloads
        .filter((_, payloadIndex) => ((index >> payloadIndex) & 1) === 1)
        .join("");
      const output = sanitizer.sanitize(input).toLowerCase();

      expect(output).not.toMatch(
        /<script|<svg|<math|<iframe|<object|onerror=|onclick=|javascript:|vbscript:|expression\(|behavior:|-moz-binding|url\(/
      );
    }
  });
});
