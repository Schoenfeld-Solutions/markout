import { extractMarkdownSourceFromHtml } from "./cleanser";
import { createOfficeBodyAccessor, type BodyAccessor } from "./body-accessor";
import { createOfficeSettingsStore, type SettingsStore } from "./config";
import { DefaultHtmlSanitizer, type HtmlSanitizer } from "./html-sanitizer";
import {
  containsMarkOutFragmentMarker,
  containsMarkOutFullRenderMarker,
  MARKOUT_RENDERED_CLASS,
} from "./render-markers";
import { createLazyMarkdownRenderer } from "./lazy-markdown-renderer";
import type { MarkdownRenderer } from "./renderer";
import {
  createOfficeRenderStateStore,
  type RenderState,
  type RenderStateStore,
} from "./render-state-store";
import { resolveRuntimeChannelConfig } from "./runtime";

export type RenderItemResult = "rendered" | "restored" | "unchanged";

export interface RenderDependencies {
  bodyAccessor: BodyAccessor;
  htmlSanitizer: HtmlSanitizer;
  markdownRenderer: MarkdownRenderer;
  renderStateStore: RenderStateStore;
  settingsStore: Pick<SettingsStore, "getStylesheet">;
}

export interface ItemRenderer {
  ensureRendered(): Promise<boolean>;
  renderItem(): Promise<RenderItemResult>;
}

export const FULL_RENDER_BLOCKED_BY_FRAGMENT_MESSAGE =
  "MarkOut can't render the entire draft while it already contains inserted MarkOut fragments. Keep working with fragments or restore an unrendered draft first.";
export const LARGE_DRAFT_RESTORE_MESSAGE =
  "This draft was rendered in an earlier compose session, but Outlook didn't preserve the original HTML for restore. Reopen the unrendered draft or continue editing the rendered version.";
export { MARKOUT_RENDERED_CLASS };

export function createItemRenderer(
  dependencies: RenderDependencies
): ItemRenderer {
  return {
    ensureRendered: async () => ensureRenderedInternal(dependencies),
    renderItem: async () => renderItemInternal(dependencies),
  };
}

export async function renderItem(): Promise<RenderItemResult> {
  return createItemRenderer(createDefaultDependencies()).renderItem();
}

export async function ensureRendered(): Promise<boolean> {
  return createItemRenderer(createDefaultDependencies()).ensureRendered();
}

async function applyRenderedContent(
  dependencies: RenderDependencies,
  originalHtml: string
): Promise<boolean> {
  const renderedHtml = await renderDraftMarkdownSegments(
    dependencies,
    originalHtml
  );

  if (renderedHtml === null) {
    return false;
  }

  await dependencies.renderStateStore.setPendingRenderState(originalHtml);

  try {
    await dependencies.bodyAccessor.setHtml(renderedHtml);
  } catch (error) {
    await clearRenderStateQuietly(dependencies.renderStateStore);
    throw error;
  }

  await dependencies.renderStateStore.setRenderedRenderState(originalHtml);
  return true;
}

async function clearRenderStateQuietly(
  renderStateStore: RenderStateStore
): Promise<void> {
  try {
    await renderStateStore.clearRenderState();
  } catch {
    // Preserve the original error when recovery cleanup also fails.
  }
}

function createDefaultDependencies(): RenderDependencies {
  const runtimeChannelConfig = resolveRuntimeChannelConfig();

  return {
    bodyAccessor: createOfficeBodyAccessor(),
    htmlSanitizer: new DefaultHtmlSanitizer(),
    markdownRenderer: createLazyMarkdownRenderer(),
    renderStateStore: createOfficeRenderStateStore(
      undefined,
      runtimeChannelConfig
    ),
    settingsStore: createOfficeSettingsStore(undefined, runtimeChannelConfig),
  };
}

async function ensureRenderedInternal(
  dependencies: RenderDependencies
): Promise<boolean> {
  const renderState = await dependencies.renderStateStore.getRenderState();

  if (renderState?.phase === "rendered") {
    return false;
  }

  if (renderState?.phase === "pending") {
    const originalHtml = await recoverPendingRenderState(
      dependencies,
      renderState
    );
    return applyRenderedContent(dependencies, originalHtml);
  }

  const currentHtml = await dependencies.bodyAccessor.getHtml();
  if (containsMarkOutFragmentMarker(currentHtml)) {
    throw new Error(FULL_RENDER_BLOCKED_BY_FRAGMENT_MESSAGE);
  }

  if (containsMarkOutFullRenderMarker(currentHtml)) {
    return false;
  }

  return applyRenderedContent(dependencies, currentHtml);
}

async function recoverPendingRenderState(
  dependencies: RenderDependencies,
  renderState: RenderState
): Promise<string> {
  const currentHtml = await dependencies.bodyAccessor.getHtml();

  if (currentHtml !== renderState.originalHtml) {
    await dependencies.bodyAccessor.setHtml(renderState.originalHtml);
  }

  await clearRenderStateQuietly(dependencies.renderStateStore);
  return renderState.originalHtml;
}

async function renderItemInternal(
  dependencies: RenderDependencies
): Promise<RenderItemResult> {
  const renderState = await dependencies.renderStateStore.getRenderState();

  if (renderState?.phase === "rendered") {
    await dependencies.bodyAccessor.setHtml(renderState.originalHtml);
    await dependencies.renderStateStore.clearRenderState();
    return "restored";
  }

  if (renderState?.phase === "pending") {
    const originalHtml = await recoverPendingRenderState(
      dependencies,
      renderState
    );
    return (await applyRenderedContent(dependencies, originalHtml))
      ? "rendered"
      : "unchanged";
  }

  const currentHtml = await dependencies.bodyAccessor.getHtml();
  assertFullDraftRenderAllowed(currentHtml);
  return (await applyRenderedContent(dependencies, currentHtml))
    ? "rendered"
    : "unchanged";
}

function assertFullDraftRenderAllowed(html: string): void {
  if (containsMarkOutFragmentMarker(html)) {
    throw new Error(FULL_RENDER_BLOCKED_BY_FRAGMENT_MESSAGE);
  }

  if (containsMarkOutFullRenderMarker(html)) {
    throw new Error(LARGE_DRAFT_RESTORE_MESSAGE);
  }
}

async function renderDraftMarkdownSegments(
  dependencies: RenderDependencies,
  originalHtml: string
): Promise<string | null> {
  const documentFragment = new DOMParser().parseFromString(
    originalHtml,
    "text/html"
  );
  const outputSegments: string[] = [];
  const markdownSegments: string[] = [];
  let renderedAnySegment = false;
  let preserveRemainingNodes = false;

  async function flushMarkdownSegments(): Promise<string | null> {
    const markdownSource = markdownSegments.join("\n").trim();
    markdownSegments.length = 0;

    if (markdownSource.length === 0) {
      return null;
    }

    const renderedHtml = await dependencies.markdownRenderer.render({
      css: dependencies.settingsStore.getStylesheet(),
      markdown: markdownSource,
      mode: "full",
    });

    return dependencies.htmlSanitizer.sanitize(renderedHtml);
  }

  async function appendMarkdownSegments(): Promise<boolean> {
    const renderedSegment = await flushMarkdownSegments();

    if (renderedSegment === null) {
      return false;
    }

    outputSegments.push(renderedSegment);
    return true;
  }

  for (const node of Array.from(documentFragment.body.childNodes)) {
    if (preserveRemainingNodes) {
      outputSegments.push(serializeNode(node));
      continue;
    }

    if (isSignatureBoundaryNode(node)) {
      if (await appendMarkdownSegments()) {
        renderedAnySegment = true;
      }
      outputSegments.push(serializeNode(node));
      preserveRemainingNodes = true;
      continue;
    }

    const markdownSource = extractMarkdownSourceFromNode(node);

    if (isMarkdownSourceRenderable(markdownSource)) {
      markdownSegments.push(markdownSource);
      continue;
    }

    if (await appendMarkdownSegments()) {
      renderedAnySegment = true;
    }
    outputSegments.push(serializeNode(node));
  }

  if (await appendMarkdownSegments()) {
    renderedAnySegment = true;
  }

  return renderedAnySegment ? outputSegments.join("") : null;
}

function extractMarkdownSourceFromNode(node: Node): string {
  return extractMarkdownSourceFromHtml(serializeNode(node));
}

function isMarkdownSourceRenderable(markdownSource: string): boolean {
  const lines = markdownSource
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  if (lines.length === 0) {
    return false;
  }

  if (hasMarkdownTable(lines)) {
    return true;
  }

  return lines.some((line) => {
    return (
      /^#{1,6}\s+\S/.test(line) ||
      /^>\s?\S/.test(line) ||
      /^[-+*]\s+\S/.test(line) ||
      /^\d+[.)]\s+\S/.test(line) ||
      /^(```|~~~)/.test(line) ||
      /^ {0,3}([-*_])(?:\s*\1){2,}\s*$/.test(line) ||
      /(?:^|[^\w])(?:\*\*|__|`)[^\s].*(?:\*\*|__|`)/.test(line) ||
      /!?\[[^\]]+\]\([^)]+\)/.test(line)
    );
  });
}

function hasMarkdownTable(lines: string[]): boolean {
  return lines.some((line, index) => {
    const nextLine = lines[index + 1];

    return (
      line.includes("|") &&
      nextLine !== undefined &&
      /^\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?$/.test(nextLine)
    );
  });
}

function isSignatureBoundaryNode(node: Node): boolean {
  if (node.nodeType !== 1) {
    return false;
  }

  const element = node as Element;
  const signatureMetadata = [
    element.id,
    element.className,
    element.getAttribute("aria-label") ?? "",
    element.getAttribute("title") ?? "",
  ]
    .join(" ")
    .toLowerCase();

  if (signatureMetadata.includes("signature")) {
    return true;
  }

  return element.textContent.trim() === "--";
}

function serializeNode(node: Node): string {
  const container = node.ownerDocument?.createElement("div");

  if (container === undefined) {
    return node.textContent ?? "";
  }

  container.appendChild(node.cloneNode(true));
  return container.innerHTML;
}
