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
