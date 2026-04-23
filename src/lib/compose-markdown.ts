import {
  createOfficeBodyAccessor,
  type BodyAccessor,
  type ComposeSelection,
} from "./body-accessor";
import { createOfficeSettingsStore, type SettingsStore } from "./config";
import { DefaultHtmlSanitizer, type HtmlSanitizer } from "./html-sanitizer";
import {
  containsMarkOutFragmentMarker,
  containsMarkOutFullRenderMarker,
} from "./render-markers";
import { createLazyMarkdownRenderer } from "./lazy-markdown-renderer";
import type { MarkdownRenderer } from "./renderer";
import { resolveRuntimeChannelConfig } from "./runtime";

export const SUBJECT_SELECTION_UNSUPPORTED_MESSAGE =
  "MarkOut can only update the message body. Move the cursor into the body or select text there first.";
export const EMPTY_SELECTION_MESSAGE =
  "Select Markdown text in the message body before using Render selection.";
export const FULL_DRAFT_ALREADY_RENDERED_MESSAGE =
  "This draft was already rendered by MarkOut. Restore the draft before inserting or rendering more Markdown fragments.";
export const RENDERED_SELECTION_BLOCKED_MESSAGE =
  "MarkOut won't replace content that already contains rendered MarkOut markup.";

export type InsertMarkdownResult = "inserted" | "replaced";

export interface ComposeMarkdownDependencies {
  bodyAccessor: BodyAccessor;
  htmlSanitizer: HtmlSanitizer;
  markdownRenderer: MarkdownRenderer;
  settingsStore: Pick<SettingsStore, "getStylesheet">;
}

export interface ComposeMarkdownService {
  getSelection(): Promise<ComposeSelection>;
  insertRenderedMarkdown(markdown: string): Promise<InsertMarkdownResult>;
  renderPreview(markdown: string, stylesheet?: string): Promise<string>;
  renderSelection(): Promise<void>;
}

export function createComposeMarkdownService(
  dependencies: ComposeMarkdownDependencies = createDefaultDependencies()
): ComposeMarkdownService {
  return {
    getSelection: async () => dependencies.bodyAccessor.getSelection(),
    insertRenderedMarkdown: async (markdown) =>
      insertRenderedMarkdownInternal(markdown, dependencies),
    renderPreview: async (markdown, stylesheet) =>
      renderPreviewInternal(markdown, stylesheet, dependencies),
    renderSelection: async () => renderSelectionInternal(dependencies),
  };
}

async function insertRenderedMarkdownInternal(
  markdown: string,
  dependencies: ComposeMarkdownDependencies
): Promise<InsertMarkdownResult> {
  if (markdown.trim().length === 0) {
    throw new Error("Paste or drop Markdown content before inserting it.");
  }

  const selection = await dependencies.bodyAccessor.getSelection();
  assertInsertTarget(selection);
  await assertDraftIsNotFullyRendered(dependencies.bodyAccessor);

  const renderedHtml = await renderFragment(markdown, dependencies);
  await dependencies.bodyAccessor.replaceSelectionWithHtml(renderedHtml);

  return selection.hasSelection && selection.source === "body"
    ? "replaced"
    : "inserted";
}

async function renderPreviewInternal(
  markdown: string,
  stylesheet: string | undefined,
  dependencies: ComposeMarkdownDependencies
): Promise<string> {
  if (markdown.trim().length === 0) {
    return "";
  }

  return renderFragment(markdown, dependencies, stylesheet);
}

async function renderSelectionInternal(
  dependencies: ComposeMarkdownDependencies
): Promise<void> {
  const selection = await dependencies.bodyAccessor.getSelection();
  assertBodySelection(selection.source);

  if (!selection.hasSelection) {
    throw new Error(EMPTY_SELECTION_MESSAGE);
  }

  assertSelectionIsNotAlreadyRendered(selection.html);
  await assertDraftIsNotFullyRendered(dependencies.bodyAccessor);

  const renderedHtml = await renderFragment(selection.text, dependencies);
  await dependencies.bodyAccessor.replaceSelectionWithHtml(renderedHtml);
}

async function renderFragment(
  markdown: string,
  dependencies: ComposeMarkdownDependencies,
  stylesheet: string | undefined = dependencies.settingsStore.getStylesheet()
): Promise<string> {
  const renderedHtml = await dependencies.markdownRenderer.render({
    css: stylesheet,
    markdown,
    mode: "fragment",
  });

  return dependencies.htmlSanitizer.sanitize(renderedHtml);
}

async function assertDraftIsNotFullyRendered(
  bodyAccessor: BodyAccessor
): Promise<void> {
  const currentHtml = await bodyAccessor.getHtml();

  if (containsMarkOutFullRenderMarker(currentHtml)) {
    throw new Error(FULL_DRAFT_ALREADY_RENDERED_MESSAGE);
  }
}

function assertBodySelection(source: "body" | "subject"): void {
  if (source !== "body") {
    throw new Error(SUBJECT_SELECTION_UNSUPPORTED_MESSAGE);
  }
}

function assertInsertTarget(selection: ComposeSelection): void {
  if (selection.hasSelection) {
    assertBodySelection(selection.source);
  }

  assertSelectionIsNotAlreadyRendered(selection.html);
}

function assertSelectionIsNotAlreadyRendered(
  selectionHtml: string | null
): void {
  if (
    selectionHtml !== null &&
    (containsMarkOutFragmentMarker(selectionHtml) ||
      containsMarkOutFullRenderMarker(selectionHtml))
  ) {
    throw new Error(RENDERED_SELECTION_BLOCKED_MESSAGE);
  }
}

function createDefaultDependencies(): ComposeMarkdownDependencies {
  const runtimeChannelConfig = resolveRuntimeChannelConfig();

  return {
    bodyAccessor: createOfficeBodyAccessor(),
    htmlSanitizer: new DefaultHtmlSanitizer(),
    markdownRenderer: createLazyMarkdownRenderer(),
    settingsStore: createOfficeSettingsStore(undefined, runtimeChannelConfig),
  };
}
