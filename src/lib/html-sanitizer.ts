const ELEMENT_NODE = 1;
const TEXT_NODE = 3;

export interface HtmlSanitizer {
  sanitize(html: string): string;
}

const ALLOWED_TAGS = new Set([
  "a",
  "blockquote",
  "br",
  "caption",
  "code",
  "dd",
  "del",
  "div",
  "dl",
  "dt",
  "em",
  "figcaption",
  "figure",
  "h1",
  "h2",
  "h3",
  "h4",
  "h5",
  "h6",
  "hr",
  "img",
  "kbd",
  "li",
  "ol",
  "p",
  "pre",
  "q",
  "s",
  "span",
  "strong",
  "sub",
  "sup",
  "style",
  "table",
  "tbody",
  "td",
  "tfoot",
  "th",
  "thead",
  "tr",
  "ul",
]);

const DROP_CONTENT_TAGS = new Set([
  "embed",
  "iframe",
  "link",
  "math",
  "meta",
  "object",
  "script",
  "style",
  "svg",
]);
const SAFE_URL_PROTOCOLS = new Set(["http:", "https:", "mailto:", "tel:"]);
const SAFE_STYLE_PROPERTIES = new Set([
  "background",
  "background-color",
  "border",
  "border-bottom",
  "border-collapse",
  "border-left",
  "border-radius",
  "border-right",
  "border-spacing",
  "border-top",
  "color",
  "display",
  "font",
  "font-family",
  "font-size",
  "font-style",
  "font-weight",
  "line-height",
  "list-style-type",
  "margin",
  "margin-bottom",
  "margin-left",
  "margin-right",
  "margin-top",
  "overflow",
  "overflow-x",
  "padding",
  "padding-bottom",
  "padding-left",
  "padding-right",
  "padding-top",
  "quotes",
  "text-decoration",
  "white-space",
]);
const SAFE_STYLE_VALUE = /^[#(),.%!'"/+:\w\s-]+$/;
const SAFE_CLASS_NAME = /^[A-Za-z0-9_\- ]+$/;
const SAFE_LANG_VALUE = /^[A-Za-z0-9-]+$/;
const SAFE_NUMERIC_VALUE = /^\d{1,4}$/;
const SAFE_RELATIVE_URL = /^(#|\/|\.\/|\.\.\/)/;

const TAG_ATTRIBUTES: Record<string, ReadonlySet<string>> = {
  a: new Set(["href"]),
  img: new Set(["alt", "height", "src", "width"]),
  ol: new Set(["start"]),
  style: new Set(["data-markout-styles"]),
  td: new Set(["align", "colspan", "rowspan", "valign"]),
  th: new Set(["align", "colspan", "rowspan", "valign"]),
};

export class DefaultHtmlSanitizer implements HtmlSanitizer {
  public sanitize(html: string): string {
    const inputDocument = new DOMParser().parseFromString(html, "text/html");
    const outputDocument = inputDocument.implementation.createHTMLDocument("");

    for (const child of Array.from(inputDocument.body.childNodes)) {
      for (const sanitizedChild of this.sanitizeNode(outputDocument, child)) {
        outputDocument.body.appendChild(sanitizedChild);
      }
    }

    return outputDocument.body.innerHTML;
  }

  private sanitizeAttributeValue(
    tagName: string,
    attributeName: string,
    value: string
  ): string | null {
    const trimmedValue = value.trim();

    if (trimmedValue.length === 0) {
      return null;
    }

    switch (attributeName) {
      case "align":
        return ["center", "justify", "left", "right"].includes(
          trimmedValue.toLowerCase()
        )
          ? trimmedValue.toLowerCase()
          : null;
      case "alt":
      case "title":
        return trimmedValue;
      case "class":
        return SAFE_CLASS_NAME.test(trimmedValue) ? trimmedValue : null;
      case "colspan":
      case "height":
      case "rowspan":
      case "start":
      case "width":
        return SAFE_NUMERIC_VALUE.test(trimmedValue) ? trimmedValue : null;
      case "dir":
        return ["auto", "ltr", "rtl"].includes(trimmedValue.toLowerCase())
          ? trimmedValue.toLowerCase()
          : null;
      case "href":
        return sanitizeUrl(trimmedValue, false);
      case "lang":
        return SAFE_LANG_VALUE.test(trimmedValue) ? trimmedValue : null;
      case "src":
        return sanitizeUrl(trimmedValue, tagName === "img");
      case "style":
        return sanitizeStyle(trimmedValue);
      case "valign":
        return ["bottom", "middle", "top"].includes(trimmedValue.toLowerCase())
          ? trimmedValue.toLowerCase()
          : null;
      default:
        return null;
    }
  }

  private sanitizeChildren(outputDocument: Document, element: Element): Node[] {
    return Array.from(element.childNodes).flatMap((child) =>
      this.sanitizeNode(outputDocument, child)
    );
  }

  private sanitizeElement(outputDocument: Document, element: Element): Node[] {
    const tagName = element.tagName.toLowerCase();

    if (tagName === "style") {
      return this.sanitizeStyleElement(outputDocument, element);
    }

    if (DROP_CONTENT_TAGS.has(tagName)) {
      return [];
    }

    if (!ALLOWED_TAGS.has(tagName)) {
      return this.sanitizeChildren(outputDocument, element);
    }

    const sanitizedElement = outputDocument.createElement(tagName);
    const allowedAttributes = TAG_ATTRIBUTES[tagName] ?? new Set<string>();

    for (const attribute of Array.from(element.attributes)) {
      const attributeName = attribute.name.toLowerCase();

      if (attributeName.startsWith("on")) {
        continue;
      }

      if (
        !["class", "dir", "lang", "style", "title"].includes(attributeName) &&
        !allowedAttributes.has(attributeName)
      ) {
        continue;
      }

      const safeValue = this.sanitizeAttributeValue(
        tagName,
        attributeName,
        attribute.value
      );

      if (safeValue !== null) {
        sanitizedElement.setAttribute(attributeName, safeValue);
      }
    }

    for (const child of Array.from(element.childNodes)) {
      for (const sanitizedChild of this.sanitizeNode(outputDocument, child)) {
        sanitizedElement.appendChild(sanitizedChild);
      }
    }

    return [sanitizedElement];
  }

  private sanitizeStyleElement(
    outputDocument: Document,
    element: Element
  ): Node[] {
    if (element.getAttribute("data-markout-styles") !== "fragment") {
      return [];
    }

    const sanitizedStylesheet = sanitizeStylesheetText(element.textContent);

    if (sanitizedStylesheet === null) {
      return [];
    }

    const sanitizedElement = outputDocument.createElement("style");
    sanitizedElement.setAttribute("data-markout-styles", "fragment");
    sanitizedElement.textContent = sanitizedStylesheet;
    return [sanitizedElement];
  }

  private sanitizeNode(outputDocument: Document, node: Node): Node[] {
    switch (node.nodeType) {
      case ELEMENT_NODE:
        return this.sanitizeElement(outputDocument, node as Element);
      case TEXT_NODE:
        return [outputDocument.createTextNode(node.textContent ?? "")];
      default:
        return [];
    }
  }
}

export function sanitizeHtml(html: string): string {
  return new DefaultHtmlSanitizer().sanitize(html);
}

function sanitizeStyle(style: string): string | null {
  const safeDeclarations = style
    .split(";")
    .map((declaration) => declaration.trim())
    .filter((declaration) => declaration.length > 0)
    .flatMap((declaration) => {
      const separatorIndex = declaration.indexOf(":");

      if (separatorIndex === -1) {
        return [];
      }

      const propertyName = declaration
        .slice(0, separatorIndex)
        .trim()
        .toLowerCase();
      const propertyValue = declaration.slice(separatorIndex + 1).trim();
      const normalizedValue = propertyValue.replace(/\s+/g, " ");

      if (!SAFE_STYLE_PROPERTIES.has(propertyName)) {
        return [];
      }

      if (
        !SAFE_STYLE_VALUE.test(normalizedValue) ||
        /(expression|javascript:|vbscript:|behavior|@import|-moz-binding|url\()/i.test(
          normalizedValue
        )
      ) {
        return [];
      }

      return [`${propertyName}: ${normalizedValue}`];
    });

  return safeDeclarations.length > 0 ? safeDeclarations.join("; ") : null;
}

function sanitizeStylesheetText(stylesheet: string): string | null {
  const safeRules = stylesheet
    .split("}")
    .map((rule) => rule.trim())
    .filter((rule) => rule.length > 0)
    .flatMap((rule) => {
      const separatorIndex = rule.indexOf("{");

      if (separatorIndex === -1) {
        return [];
      }

      const selectorText = rule.slice(0, separatorIndex).trim();
      const declarationText = rule.slice(separatorIndex + 1).trim();

      if (
        selectorText.length === 0 ||
        declarationText.length === 0 ||
        selectorText.includes(":") ||
        /[<>]/.test(selectorText)
      ) {
        return [];
      }

      const safeDeclarationText = sanitizeStyle(declarationText);

      if (safeDeclarationText === null) {
        return [];
      }

      return [`${selectorText} { ${safeDeclarationText} }`];
    });

  return safeRules.length > 0 ? safeRules.join("\n") : null;
}

function sanitizeUrl(value: string, allowDataImage: boolean): string | null {
  const lowercaseValue = value.toLowerCase();

  if (SAFE_RELATIVE_URL.test(value)) {
    return value;
  }

  if (lowercaseValue.startsWith("cid:")) {
    return value;
  }

  if (
    allowDataImage &&
    /^data:image\/[a-z0-9.+-]+;base64,[a-z0-9+/=]+$/i.test(value)
  ) {
    return value;
  }

  try {
    const parsedUrl = new URL(value);
    return SAFE_URL_PROTOCOLS.has(parsedUrl.protocol) ? value : null;
  } catch {
    return null;
  }
}
