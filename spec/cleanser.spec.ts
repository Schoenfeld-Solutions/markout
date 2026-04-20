import { JSDOM } from "jsdom";
import { cleanse } from "../src/lib/cleanser";
import { readFile } from "./helpers";

const runtime = globalThis as typeof globalThis & {
  DOMParser: typeof DOMParser;
};
runtime.DOMParser = new JSDOM().window.DOMParser;

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
});
