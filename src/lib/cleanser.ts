const ELEMENT_NODE = 1;
const TEXT_NODE = 3;
const COMMENT_NODE = 8;

export function cleanse(text: string): string {
  const documentFragment = new DOMParser().parseFromString(text, "text/html");
  const lines = Array.from(documentFragment.body.childNodes).flatMap((node) =>
    cleanseNode(documentFragment, node)
  );

  return lines.join("").trim();
}

export function extractMarkdownSourceFromHtml(html: string): string {
  const documentFragment = new DOMParser().parseFromString(html, "text/html");
  const lines = Array.from(documentFragment.body.childNodes).flatMap((node) =>
    extractMarkdownLines(node, 0)
  );

  return normalizeMarkdownLines(lines);
}

function cleanseNode(documentFragment: Document, node: Node): string[] {
  switch (node.nodeType) {
    case ELEMENT_NODE:
      return cleanseElement(documentFragment, node as Element);
    case TEXT_NODE:
      return [cleanseText(documentFragment, node.textContent ?? "")];
    case COMMENT_NODE:
      return [];
    default:
      return [];
  }
}

function cleanseElement(
  documentFragment: Document,
  element: Element
): string[] {
  switch (element.tagName.toLowerCase()) {
    case "script":
    case "style":
      return [];
    case "br":
      return [];
    case "a": {
      const href = element.getAttribute("href");
      if (href !== null && element.innerHTML === href) {
        return [href];
      }

      return [element.outerHTML];
    }
    case "img":
      return [element.outerHTML.replace(/\n+/g, "\n")];
    case "div":
    case "p":
      return [...cleanseElementContainer(documentFragment, element), "\n"];
    case "span":
      return cleanseElementContainer(documentFragment, element);
    default:
      return [element.outerHTML];
  }
}

function cleanseElementContainer(
  documentFragment: Document,
  container: Element
): string[] {
  if (
    container.id !== "" &&
    Array.from(container.id).every((character) => character.charCodeAt(0) < 128)
  ) {
    return [container.outerHTML];
  }

  return Array.from(container.childNodes).flatMap((node) =>
    cleanseNode(documentFragment, node)
  );
}

function cleanseText(documentFragment: Document, text: string): string {
  const container = documentFragment.createElement("span");
  container.innerHTML = text
    .replace(/[\u007F-\u009F]/g, "")
    .replace(/[\u00a0]/g, " ");

  if (!container.innerHTML.trim()) {
    return "";
  }

  return container.textContent
    .replace(/^(\r?\n)+/, "\n")
    .replace(/(\r?\n)+$/, "\n")
    .replace(/(.+)[\r\n]+(.+)/g, "$1 $2");
}

function extractMarkdownLines(node: Node, listDepth: number): string[] {
  switch (node.nodeType) {
    case ELEMENT_NODE:
      return extractMarkdownElementLines(node as Element, listDepth);
    case TEXT_NODE:
      return splitTextLines(node.textContent ?? "");
    case COMMENT_NODE:
      return [];
    default:
      return [];
  }
}

function extractMarkdownElementLines(
  element: Element,
  listDepth: number
): string[] {
  const tagName = element.tagName.toLowerCase();

  if (tagName === "script" || tagName === "style") {
    return [];
  }

  if (tagName === "br") {
    return [""];
  }

  if (/^h[1-6]$/.test(tagName)) {
    const headingLevel = Number.parseInt(tagName.slice(1), 10);
    const headingText = collectInlineText(element).trim();

    return headingText.length === 0
      ? []
      : [`${"#".repeat(headingLevel)} ${headingText}`, ""];
  }

  if (tagName === "ul" || tagName === "ol") {
    return [...extractListLines(element, listDepth), ""];
  }

  if (tagName === "li") {
    return extractListItemLines(element, false, 1, listDepth);
  }

  if (tagName === "div") {
    const lines = extractContainerLines(element, listDepth);

    return lines.length === 0 ? [""] : lines;
  }

  if (tagName === "p") {
    const lines = extractContainerLines(element, listDepth);

    return lines.length === 0 ? [""] : [...lines, ""];
  }

  return extractContainerLines(element, listDepth);
}

function extractContainerLines(element: Element, listDepth: number): string[] {
  const lines = Array.from(element.childNodes).flatMap((childNode) =>
    extractMarkdownLines(childNode, listDepth)
  );
  const normalized = normalizeMarkdownLines(lines);

  return normalized.length === 0 ? [] : normalized.split("\n");
}

function extractListLines(listElement: Element, listDepth: number): string[] {
  const isOrdered = listElement.tagName.toLowerCase() === "ol";
  const lines: string[] = [];
  let itemIndex = 1;

  for (const childElement of Array.from(listElement.children)) {
    if (childElement.tagName.toLowerCase() !== "li") {
      continue;
    }

    lines.push(
      ...extractListItemLines(childElement, isOrdered, itemIndex, listDepth)
    );
    itemIndex += 1;
  }

  return lines;
}

function extractListItemLines(
  listItemElement: Element,
  isOrdered: boolean,
  itemIndex: number,
  listDepth: number
): string[] {
  const indent = "  ".repeat(listDepth);
  const marker = isOrdered ? `${itemIndex}.` : "-";
  const contentLines: string[] = [];
  const nestedListLines: string[] = [];

  for (const childNode of Array.from(listItemElement.childNodes)) {
    if (
      childNode.nodeType === ELEMENT_NODE &&
      isListElement(childNode as Element)
    ) {
      nestedListLines.push(
        ...extractListLines(childNode as Element, listDepth + 1)
      );
      continue;
    }

    contentLines.push(...extractMarkdownLines(childNode, listDepth));
  }

  const content = normalizeMarkdownLines(contentLines);
  const lines =
    content.length === 0
      ? [`${indent}${marker}`]
      : content
          .split("\n")
          .map((line, index) =>
            index === 0 ? `${indent}${marker} ${line}` : `${indent}  ${line}`
          );

  return [...lines, ...nestedListLines];
}

function isListElement(element: Element): boolean {
  const tagName = element.tagName.toLowerCase();

  return tagName === "ul" || tagName === "ol";
}

function collectInlineText(element: Element): string {
  return normalizeInlineText(
    Array.from(element.childNodes)
      .flatMap((node) => {
        if (node.nodeType === TEXT_NODE) {
          return [node.textContent ?? ""];
        }

        if (node.nodeType !== ELEMENT_NODE) {
          return [];
        }

        const childElement = node as Element;
        const tagName = childElement.tagName.toLowerCase();

        if (tagName === "script" || tagName === "style") {
          return [];
        }

        if (tagName === "br") {
          return ["\n"];
        }

        return [collectInlineText(childElement)];
      })
      .join("")
  );
}

function splitTextLines(text: string): string[] {
  const normalizedText = text
    .replace(/[\u007F-\u009F]/g, "")
    .replace(/[\u00a0]/g, " ");

  return normalizedText
    .split(/\r?\n/)
    .map((line) => normalizeMarkdownTextLine(line))
    .filter((line, index, lines) => {
      if (line.length > 0) {
        return true;
      }

      return index > 0 && index < lines.length - 1;
    });
}

function normalizeInlineText(text: string): string {
  return text.replace(/[ \t\f\v]+/g, " ").trim();
}

function normalizeMarkdownTextLine(text: string): string {
  return text.replace(/[\t\f\v]+/g, " ").replace(/[ ]+$/g, "");
}

function normalizeMarkdownLines(lines: string[]): string {
  const normalizedLines = lines.map((line) => line.replace(/[ \t\f\v]+$/g, ""));

  while (
    normalizedLines.length > 0 &&
    (normalizedLines[0] ?? "").trim().length === 0
  ) {
    normalizedLines.shift();
  }

  while (
    normalizedLines.length > 0 &&
    (normalizedLines[normalizedLines.length - 1] ?? "").trim().length === 0
  ) {
    normalizedLines.pop();
  }

  return normalizedLines.join("\n").replace(/\n{3,}/g, "\n\n");
}
