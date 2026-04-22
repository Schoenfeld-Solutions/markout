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
import {
  MARKOUT_FRAGMENT_HOST_CLASS,
  MARKOUT_FRAGMENT_RENDERED_CLASS,
  MARKOUT_RENDERED_CLASS,
} from "./render-markers";
import { isInlineableSelector, parseStyleRules } from "./stylesheet-rules";

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

export type RenderMode = "fragment" | "full";

export interface RenderOptions {
  css?: string;
  markdown: string;
  mode?: RenderMode;
}

export interface MarkdownRenderer {
  render(options: RenderOptions): Promise<string>;
}

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

function createWrapperClasses(mode: RenderMode): string {
  return mode === "full"
    ? `mo ${MARKOUT_RENDERED_CLASS}`
    : `mo ${MARKOUT_FRAGMENT_RENDERED_CLASS}`;
}

export const MO_CONTENT_PREFIX = (mode: RenderMode = "full"): string =>
  `<div class="${createWrapperClasses(mode)}">\n`;

export const MO_CONTENT_SUFFIX = (): string => `</div>\n`;

class InlineStyleMarkdownRenderer implements MarkdownRenderer {
  public render({
    markdown,
    css = getStylesheet(),
    mode = "full",
  }: RenderOptions): Promise<string> {
    const rawHtml = `${MO_CONTENT_PREFIX(mode)}${markdownIt.render(
      markdown
    )}${MO_CONTENT_SUFFIX()}`;
    const documentFragment = new DOMParser().parseFromString(
      rawHtml,
      "text/html"
    );

    if (mode === "fragment") {
      return Promise.resolve(
        renderFragmentHtml(documentFragment.body.innerHTML, css)
      );
    }

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

  const styleRules = parseStyleRules(stylesheet);

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

function buildScopedStylesheet(
  stylesheet: string,
  rootSelector: string
): string {
  return parseStyleRules(stylesheet)
    .filter((styleRule) => isInlineableSelector(styleRule.selectorText))
    .map((styleRule) => {
      const scopedSelector = styleRule.selectorText
        .split(",")
        .map((selector) => selector.trim())
        .filter((selector) => selector.length > 0)
        .map((selector) => `${rootSelector} ${selector}`)
        .join(", ");

      return `${scopedSelector} { ${styleRule.declarationText} }`;
    })
    .join("\n");
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

function renderFragmentHtml(contentHtml: string, stylesheet: string): string {
  const scopedStylesheet = buildScopedStylesheet(
    stylesheet,
    `.${MARKOUT_FRAGMENT_HOST_CLASS}`
  );
  const styleTag =
    scopedStylesheet.trim().length === 0
      ? ""
      : `<style data-markout-styles="fragment">${escapeHtml(
          scopedStylesheet
        )}</style>\n`;

  return `<div class="${MARKOUT_FRAGMENT_HOST_CLASS}">\n${styleTag}${contentHtml}</div>\n`;
}

export function createMarkdownRenderer(): MarkdownRenderer {
  return new InlineStyleMarkdownRenderer();
}

export {
  containsMarkOutFragmentMarker,
  containsMarkOutFullRenderMarker,
} from "./render-markers";

export async function renderMarkdown({
  markdown,
  css,
  mode,
}: RenderOptions): Promise<string> {
  if (css === undefined) {
    if (mode === undefined) {
      return createMarkdownRenderer().render({ markdown });
    }

    return createMarkdownRenderer().render({ markdown, mode });
  }

  if (mode === undefined) {
    return createMarkdownRenderer().render({ css, markdown });
  }

  return createMarkdownRenderer().render({ css, markdown, mode });
}
