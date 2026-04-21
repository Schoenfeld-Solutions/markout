import hljs from "highlight.js/lib/core";
import bash from "highlight.js/lib/languages/bash";
import css from "highlight.js/lib/languages/css";
import javascript from "highlight.js/lib/languages/javascript";
import json from "highlight.js/lib/languages/json";
import MarkdownIt from "markdown-it";
import markdown from "highlight.js/lib/languages/markdown";
import typescript from "highlight.js/lib/languages/typescript";
import xml from "highlight.js/lib/languages/xml";
import { getStylesheet } from "./config";

// eslint-disable-next-line @typescript-eslint/no-require-imports
const markdownItEmoji = require("markdown-it-emoji") as (
  markdownIt: MarkdownIt
) => void;
// eslint-disable-next-line @typescript-eslint/no-require-imports
const markdownItFootnote = require("markdown-it-footnote") as (
  markdownIt: MarkdownIt
) => void;

hljs.registerLanguage("bash", bash);
hljs.registerLanguage("css", css);
hljs.registerLanguage("javascript", javascript);
hljs.registerLanguage("json", json);
hljs.registerLanguage("markdown", markdown);
hljs.registerLanguage("typescript", typescript);
hljs.registerLanguage("xml", xml);

function escapeHtml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

const markdownIt = new MarkdownIt({
  breaks: false,
  highlight(str, lang): string {
    if (lang !== "" && hljs.getLanguage(lang)) {
      try {
        return `<pre class="hljs"><code>${hljs.highlight(str, { language: lang }).value}</code></pre>`;
      } catch {
        // Fall back to escaped code below.
      }
    }

    return `<pre class="hljs"><code>${escapeHtml(str)}</code></pre>`;
  },
  html: true,
})
  .use(markdownItFootnote)
  .use(markdownItEmoji);

export interface RenderOptions {
  css?: string;
  markdown: string;
}

export interface MarkdownRenderer {
  render(options: RenderOptions): Promise<string>;
}

export const MARKOUT_RENDERED_CLASS = "markout-rendered";
export const MO_CONTENT_PREFIX = (): string =>
  `<div class="mo ${MARKOUT_RENDERED_CLASS}">\n`;
export const MO_CONTENT_SUFFIX = (): string => `</div>\n`;

class InlineStyleMarkdownRenderer implements MarkdownRenderer {
  public render({
    markdown,
    css = getStylesheet(),
  }: RenderOptions): Promise<string> {
    const rawHtml = `${MO_CONTENT_PREFIX()}${markdownIt.render(
      markdown
    )}${MO_CONTENT_SUFFIX()}`;
    const documentFragment = new DOMParser().parseFromString(
      rawHtml,
      "text/html"
    );
    applyInlineStyles(documentFragment, css);
    return Promise.resolve(documentFragment.body.innerHTML);
  }
}

function applyInlineStyles(
  documentFragment: Document,
  stylesheet: string
): void {
  if (stylesheet.trim().length === 0) {
    return;
  }

  const styleRules = parseInlineStyleRules(stylesheet);

  for (const styleRule of styleRules) {
    if (!isInlineableSelector(styleRule.selectorText)) {
      continue;
    }

    let matchingElements: NodeListOf<Element>;

    try {
      matchingElements = documentFragment.body.querySelectorAll(
        styleRule.selectorText
      );
    } catch {
      continue;
    }

    for (const matchingElement of Array.from(matchingElements)) {
      mergeInlineStyles(matchingElement, styleRule.declarationText);
    }
  }
}

function isInlineableSelector(selectorText: string): boolean {
  return selectorText
    .split(",")
    .map((selector) => selector.trim())
    .every((selector) => selector.length > 0 && !selector.includes(":"));
}

function parseInlineStyleRules(
  stylesheet: string
): { declarationText: string; selectorText: string }[] {
  return stylesheet
    .replaceAll(/\/\*[\s\S]*?\*\//g, "")
    .split("}")
    .flatMap((ruleFragment) => {
      const separatorIndex = ruleFragment.indexOf("{");

      if (separatorIndex === -1) {
        return [];
      }

      const selectorText = ruleFragment.slice(0, separatorIndex).trim();
      const declarationText = ruleFragment.slice(separatorIndex + 1).trim();

      if (selectorText.length === 0 || declarationText.length === 0) {
        return [];
      }

      return [{ declarationText, selectorText }];
    });
}

function mergeInlineStyles(element: Element, declarationText: string): void {
  const workingElement = element.ownerDocument.createElement("div");
  const declarationElement = element.ownerDocument.createElement("div");
  const existingStyle = element.getAttribute("style");

  if (existingStyle !== null) {
    workingElement.setAttribute("style", existingStyle);
  }

  declarationElement.setAttribute("style", declarationText);

  for (const propertyName of Array.from(declarationElement.style)) {
    const propertyValue = declarationElement.style
      .getPropertyValue(propertyName)
      .trim();

    if (propertyValue.length === 0) {
      continue;
    }

    workingElement.style.setProperty(
      propertyName,
      propertyValue,
      declarationElement.style.getPropertyPriority(propertyName)
    );
  }

  const cssText = workingElement.style.cssText.trim();

  if (cssText.length === 0) {
    element.removeAttribute("style");
    return;
  }

  element.setAttribute("style", cssText);
}

export function createMarkdownRenderer(): MarkdownRenderer {
  return new InlineStyleMarkdownRenderer();
}

export async function renderMarkdown({
  markdown,
  css,
}: RenderOptions): Promise<string> {
  return css === undefined
    ? createMarkdownRenderer().render({ markdown })
    : createMarkdownRenderer().render({ css, markdown });
}
