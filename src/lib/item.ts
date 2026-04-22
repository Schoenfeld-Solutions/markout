import { cleanse } from "./cleanser";
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

export type RenderItemResult = "rendered" | "restored";

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
): Promise<void> {
  const markdownSource = cleanse(originalHtml);
  const renderedHtml = await dependencies.markdownRenderer.render({
    css: dependencies.settingsStore.getStylesheet(),
    markdown: markdownSource,
    mode: "full",
  });
  const sanitizedHtml = dependencies.htmlSanitizer.sanitize(renderedHtml);

  await dependencies.renderStateStore.setPendingRenderState(originalHtml);

  try {
    await dependencies.bodyAccessor.setHtml(sanitizedHtml);
  } catch (error) {
    await clearRenderStateQuietly(dependencies.renderStateStore);
    throw error;
  }

  await dependencies.renderStateStore.setRenderedRenderState(originalHtml);
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
  return {
    bodyAccessor: createOfficeBodyAccessor(),
    htmlSanitizer: new DefaultHtmlSanitizer(),
    markdownRenderer: createLazyMarkdownRenderer(),
    renderStateStore: createOfficeRenderStateStore(),
    settingsStore: createOfficeSettingsStore(),
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
    await applyRenderedContent(dependencies, originalHtml);
    return true;
  }

  const currentHtml = await dependencies.bodyAccessor.getHtml();
  if (containsMarkOutFragmentMarker(currentHtml)) {
    throw new Error(FULL_RENDER_BLOCKED_BY_FRAGMENT_MESSAGE);
  }

  if (containsMarkOutFullRenderMarker(currentHtml)) {
    return false;
  }

  await applyRenderedContent(dependencies, currentHtml);
  return true;
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
    await applyRenderedContent(dependencies, originalHtml);
    return "rendered";
  }

  const currentHtml = await dependencies.bodyAccessor.getHtml();
  assertFullDraftRenderAllowed(currentHtml);
  await applyRenderedContent(dependencies, currentHtml);
  return "rendered";
}

function assertFullDraftRenderAllowed(html: string): void {
  if (containsMarkOutFragmentMarker(html)) {
    throw new Error(FULL_RENDER_BLOCKED_BY_FRAGMENT_MESSAGE);
  }

  if (containsMarkOutFullRenderMarker(html)) {
    throw new Error(LARGE_DRAFT_RESTORE_MESSAGE);
  }
}
