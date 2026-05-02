/** @jest-environment jsdom */
import { cleanse, extractMarkdownSourceFromHtml } from "../src/lib/cleanser";
import { readFile } from "./helpers";

const tests: {
  inputFile: string;
  name: string;
  outputFile: string;
  stripNewlines?: boolean;
}[] = [
  {
    inputFile: "cleanser/test1.input.html",
    name: "handles a simple line correctly",
    outputFile: "cleanser/test1.output.md",
    stripNewlines: true,
  },
  {
    inputFile: "cleanser/test2.input.html",
    name: "handles line breaks correctly",
    outputFile: "cleanser/test2.output.md",
    stripNewlines: true,
  },
  {
    inputFile: "cleanser/test3.input.html",
    name: "handles HTML escaped characters correctly",
    outputFile: "cleanser/test3.output.md",
    stripNewlines: true,
  },
  {
    inputFile: "cleanser/test4.input.html",
    name: "handles a more complex example correctly",
    outputFile: "cleanser/test4.output.md",
    stripNewlines: true,
  },
  {
    inputFile: "cleanser/test5.input.html",
    name: "handles image tags correctly",
    outputFile: "cleanser/test5.output.md",
    stripNewlines: true,
  },
  {
    inputFile: "cleanser/test6.input.html",
    name: "does not modify containers with IDs",
    outputFile: "cleanser/test6.output.md",
    stripNewlines: true,
  },
  {
    inputFile: "cleanser/test7.input.html",
    name: "handles a weirdly formatted complex input file",
    outputFile: "cleanser/test7.output.md",
  },
  {
    inputFile: "cleanser/test8.input.html",
    name: "extracts auto-linked text correctly",
    outputFile: "cleanser/test8.output.md",
    stripNewlines: true,
  },
  {
    inputFile: "cleanser/test9.input.html",
    name: "handles a complex example with multiple images",
    outputFile: "cleanser/test9.output.md",
    stripNewlines: true,
  },
  {
    inputFile: "cleanser/test10.input.html",
    name: "handles complex code blocks with indentation",
    outputFile: "cleanser/test10.output.md",
    stripNewlines: true,
  },
  {
    inputFile: "cleanser/test11.input.html",
    name: "handles text from Outlook for macOS",
    outputFile: "cleanser/test11.output.md",
    stripNewlines: false,
  },
  {
    inputFile: "cleanser/test12.input.html",
    name: "handles a complex Outlook for macOS example including tables",
    outputFile: "cleanser/test12.output.md",
    stripNewlines: false,
  },
];

describe("cleanser", () => {
  for (const testCase of tests) {
    it(testCase.name, () => {
      const input = readFile(testCase.inputFile, testCase.stripNewlines);
      const expected = readFile(testCase.outputFile)
        .replace(/\n$/, "")
        .split("\n");

      expect(cleanse(input).split("\n")).toEqual(expected);
    });
  }

  it("returns the inner text for text nodes", () => {
    expect(cleanse("this is a test")).toBe("this is a test");
  });

  it("replaces inline newlines with spaces", () => {
    expect(cleanse("this is a test\nwith inline spaces")).toBe(
      "this is a test with inline spaces"
    );
  });

  it("drops comment nodes", () => {
    expect(cleanse("<!-- this is a comment -->")).toBe("");
  });

  it("drops script and style nodes", () => {
    expect(cleanse("<script>this is a script</script>")).toBe("");
    expect(cleanse("<style>this is a style</style>")).toBe("");
    expect(cleanse("<br>")).toBe("");
    expect(cleanse("<strong>authored html</strong>")).toBe(
      "<strong>authored html</strong>"
    );
  });

  it("keeps image nodes unchanged", () => {
    expect(
      cleanse(`<img src="https://google.com/favicon.ico" id="test">`)
    ).toBe(`<img src="https://google.com/favicon.ico" id="test">`);
  });

  it("handles escaped properties in attributes", () => {
    expect(
      cleanse(
        `<img src="https://google.com/favicon.ico" alt="Test&#10;Newline">`
      )
    ).toBe(`<img src="https://google.com/favicon.ico" alt="Test\nNewline">`);
  });

  it("separates divs with newlines", () => {
    expect(
      cleanse(`<div>this is a div</div><div>this is another div</div>`)
    ).toBe("this is a div\nthis is another div");
  });

  it("treats divs that contain only a br as a single newline", () => {
    expect(
      cleanse(
        `<div>this is a div<br></div><div><br></div><div>this is another div</div>`
      )
    ).toBe("this is a div\n\nthis is another div");
  });

  it("merges child trees with newlines", () => {
    expect(cleanse(`<div><div>a</div><div>b</div></div>`)).toBe("a\nb");
    expect(cleanse(`<div><div>a</div><div><br></div><div>b</div></div>`)).toBe(
      "a\n\nb"
    );
  });

  it("preserves auto-linked anchors as markdown source but leaves authored anchors intact", () => {
    expect(
      cleanse(`<a href="https://example.test">https://example.test</a>`)
    ).toBe("https://example.test");
    expect(cleanse(`<a href="https://example.test">Example</a>`)).toBe(
      `<a href="https://example.test">Example</a>`
    );
  });

  it("only preserves ascii-id containers as authored html", () => {
    expect(cleanse(`<div id="safe-id">Keep <strong>HTML</strong></div>`)).toBe(
      `<div id="safe-id">Keep <strong>HTML</strong></div>`
    );
    expect(cleanse(`<div id="ümlaut">Flatten <span>text</span></div>`)).toBe(
      "Flatten text"
    );
  });

  it("extracts structured markdown from Outlook-like html blocks", () => {
    expect(
      extractMarkdownSourceFromHtml(`
        <div class="WordSection1">
          <!-- Outlook comment -->
          <h1>Release <span>Notes</span></h1>
          <p>Intro&nbsp;text<br>continued</p>
          <script>alert("x")</script>
          <style>.x { color: red; }</style>
          <ul>
            <li>Parent
              <ol>
                <li>Child one</li>
                <li><p>Child two</p></li>
              </ol>
            </li>
            <li>Next</li>
          </ul>
          <div><br></div>
          <h2> </h2>
          <li>Loose item</li>
        </div>
      `)
    ).toBe(
      [
        "# Release Notes",
        "",
        "Intro text",
        "",
        "continued",
        "",
        "- Parent",
        "  1. Child one",
        "  2. Child two",
        "- Next",
        "",
        "- Loose item",
      ].join("\n")
    );
  });

  it("extracts markdown from non-paragraph containers, empty list items, and inline heading children", () => {
    expect(
      extractMarkdownSourceFromHtml(`
        <blockquote>Quoted<br>line</blockquote>
        <ul>
          <div>ignored list wrapper</div>
          <li><ul><li>Nested only</li></ul></li>
          <li>Second</li>
        </ul>
        <h2>Inline<!-- hidden --><script>ignored()</script><br>heading</h2>
      `)
    ).toBe(
      [
        "Quoted",
        "",
        "line",
        "-",
        "  - Nested only",
        "- Second",
        "",
        "## Inline",
        "heading",
      ].join("\n")
    );
  });

  it("normalizes text and empty html while extracting markdown source", () => {
    expect(
      extractMarkdownSourceFromHtml(
        `<p>\u007fAlpha&nbsp;Beta\u0085</p><div></div><p></p><p>Gamma</p>`
      )
    ).toBe(["Alpha Beta", "", "Gamma"].join("\n"));
  });
});
