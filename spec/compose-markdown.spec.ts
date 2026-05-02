/** @jest-environment jsdom */
import type { BodyAccessor } from "../src/lib/body-accessor";
import {
  EMPTY_SELECTION_MESSAGE,
  FULL_DRAFT_ALREADY_RENDERED_MESSAGE,
  RENDERED_SELECTION_BLOCKED_MESSAGE,
  SUBJECT_SELECTION_UNSUPPORTED_MESSAGE,
  createComposeMarkdownService,
} from "../src/lib/compose-markdown";
import { DefaultHtmlSanitizer } from "../src/lib/html-sanitizer";
import {
  installDomParser,
  FakeMailboxItem,
  installOfficeEnvironment,
} from "./helpers";

describe("compose markdown service", () => {
  beforeEach(() => {
    installDomParser();
    installOfficeEnvironment();
  });

  it("renders the current body selection into a MarkOut fragment", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.selectionSource = "body";
    mailboxItem.selectionText = "# Selected";
    mailboxItem.selectionHtml = "<p># Selected</p>";
    installOfficeEnvironment({ mailboxItem });

    const composeMarkdownService = createComposeMarkdownService();

    await composeMarkdownService.renderSelection();

    expect(mailboxItem.body.lastSelectedHtml).toContain(
      "markout-fragment-host"
    );
    expect(mailboxItem.body.lastSelectedHtml).toContain(
      "markout-fragment-rendered"
    );
    expect(mailboxItem.body.lastSelectedHtml).toContain(
      'data-markout-styles="fragment"'
    );
  });

  it("uses selection html structure instead of flattened selection text", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.selectionSource = "body";
    mailboxItem.selectionText = "# Title Paragraph - parent - child";
    mailboxItem.selectionHtml = [
      "<div># Title</div>",
      "<div>Paragraph text</div>",
      "<div>- parent</div>",
      "<div>&nbsp;&nbsp;- child</div>",
    ].join("");
    installOfficeEnvironment({ mailboxItem });

    const composeMarkdownService = createComposeMarkdownService();

    await composeMarkdownService.renderSelection();

    expect(mailboxItem.body.lastSelectedHtml).toContain("<h1>Title</h1>");
    expect(mailboxItem.body.lastSelectedHtml).toContain(
      "<p>Paragraph text</p>"
    );
    expect(mailboxItem.body.lastSelectedHtml).toContain("<li>parent");
    expect(mailboxItem.body.lastSelectedHtml).toContain("<li>child</li>");
    expect(mailboxItem.body.lastSelectedHtml).not.toContain(
      "# Title Paragraph - parent - child"
    );
  });

  it("falls back to selection text when Outlook does not provide selection html", async () => {
    const bodyAccessor: BodyAccessor = {
      getHtml: jest.fn().mockResolvedValue("<div>Original</div>"),
      getSelection: jest.fn().mockResolvedValue({
        hasSelection: true,
        html: null,
        source: "body",
        text: "## Fallback heading",
      }),
      replaceSelectionWithHtml: jest.fn().mockResolvedValue(undefined),
      setHtml: jest.fn().mockResolvedValue(undefined),
    };
    const markdownRenderer = {
      render: jest
        .fn()
        .mockResolvedValue(
          '<div class="markout-fragment-host"><div class="mo markout-fragment-rendered"><h2>Fallback heading</h2></div></div>'
        ),
    };
    const composeMarkdownService = createComposeMarkdownService({
      bodyAccessor,
      htmlSanitizer: new DefaultHtmlSanitizer(),
      markdownRenderer,
      settingsStore: {
        getStylesheet(): string {
          return "";
        },
      },
    });

    await composeMarkdownService.renderSelection();

    expect(markdownRenderer.render).toHaveBeenCalledWith(
      expect.objectContaining({
        markdown: "## Fallback heading",
      })
    );
    expect(bodyAccessor.replaceSelectionWithHtml).toHaveBeenCalledWith(
      expect.stringContaining("Fallback heading")
    );
  });

  it("inserts rendered markdown at the current body cursor when no selection exists", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.selectionSource = "body";
    mailboxItem.selectionText = "";
    mailboxItem.selectionHtml = "";
    installOfficeEnvironment({ mailboxItem });

    const composeMarkdownService = createComposeMarkdownService();

    await expect(
      composeMarkdownService.insertRenderedMarkdown("## Insert me")
    ).resolves.toBe("inserted");
    expect(mailboxItem.body.lastSelectedHtml).toContain("Insert me");
  });

  it("replaces an active body selection when inserting rendered markdown", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.selectionSource = "body";
    mailboxItem.selectionText = "replace me";
    mailboxItem.selectionHtml = "<p>replace me</p>";
    installOfficeEnvironment({ mailboxItem });

    const composeMarkdownService = createComposeMarkdownService();

    await expect(
      composeMarkdownService.insertRenderedMarkdown("## Replacement")
    ).resolves.toBe("replaced");
    expect(mailboxItem.body.lastSelectedHtml).toContain("Replacement");
  });

  it("still inserts at the body when the host reports a stale subject source without an active selection", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.selectionSource = "subject";
    mailboxItem.selectionText = "";
    mailboxItem.selectionHtml = "";
    installOfficeEnvironment({ mailboxItem });

    const composeMarkdownService = createComposeMarkdownService();

    await expect(
      composeMarkdownService.insertRenderedMarkdown("## Insert me")
    ).resolves.toBe("inserted");
    expect(mailboxItem.body.lastSelectedHtml).toContain("Insert me");
  });

  it("renders a fragment preview with sanitized scoped styles", async () => {
    installOfficeEnvironment({
      mailboxItem: new FakeMailboxItem("<div>Original</div>"),
    });
    const composeMarkdownService = createComposeMarkdownService();

    const preview = await composeMarkdownService.renderPreview("# Preview");

    expect(preview).toContain("markout-fragment-host");
    expect(preview).toContain('data-markout-styles="fragment"');
  });

  it("returns the current selection and keeps empty preview/insert inputs safe", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.selectionSource = "body";
    mailboxItem.selectionText = "Selected text";
    mailboxItem.selectionHtml = "<p>Selected text</p>";
    installOfficeEnvironment({ mailboxItem });

    const composeMarkdownService = createComposeMarkdownService();

    await expect(composeMarkdownService.getSelection()).resolves.toMatchObject({
      hasSelection: true,
      source: "body",
      text: "Selected text",
    });
    await expect(composeMarkdownService.renderPreview("   ")).resolves.toBe("");
    await expect(
      composeMarkdownService.insertRenderedMarkdown("   ")
    ).rejects.toThrow("Paste or drop Markdown content before inserting it.");
  });

  it("fails when rendering a selection without selected body text", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.selectionSource = "body";
    mailboxItem.selectionText = "";
    installOfficeEnvironment({ mailboxItem });

    const composeMarkdownService = createComposeMarkdownService();

    await expect(composeMarkdownService.renderSelection()).rejects.toThrow(
      EMPTY_SELECTION_MESSAGE
    );
  });

  it("fails when the current selection is in the subject", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.selectionSource = "subject";
    mailboxItem.selectionText = "Subject selection";
    installOfficeEnvironment({ mailboxItem });

    const composeMarkdownService = createComposeMarkdownService();

    await expect(composeMarkdownService.renderSelection()).rejects.toThrow(
      SUBJECT_SELECTION_UNSUPPORTED_MESSAGE
    );
  });

  it("still blocks fragment insertion when the host reports a subject selection", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.selectionSource = "subject";
    mailboxItem.selectionText = "Subject selection";
    installOfficeEnvironment({ mailboxItem });

    const composeMarkdownService = createComposeMarkdownService();

    await expect(
      composeMarkdownService.insertRenderedMarkdown("## Insert me")
    ).rejects.toThrow(SUBJECT_SELECTION_UNSUPPORTED_MESSAGE);
  });

  it("blocks fragment work when the entire draft was already rendered", async () => {
    const mailboxItem = new FakeMailboxItem(
      '<div class="mo markout-rendered"><p>Rendered body</p></div>'
    );
    mailboxItem.selectionSource = "body";
    mailboxItem.selectionText = "# Selected";
    mailboxItem.selectionHtml = "<p># Selected</p>";
    installOfficeEnvironment({ mailboxItem });

    const composeMarkdownService = createComposeMarkdownService();

    await expect(composeMarkdownService.renderSelection()).rejects.toThrow(
      FULL_DRAFT_ALREADY_RENDERED_MESSAGE
    );
  });

  it("blocks fragment work when the selection already contains rendered MarkOut content", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.selectionSource = "body";
    mailboxItem.selectionText = "Rendered";
    mailboxItem.selectionHtml =
      '<div class="markout-fragment-host"><div class="mo markout-fragment-rendered">Rendered</div></div>';
    installOfficeEnvironment({ mailboxItem });

    const composeMarkdownService = createComposeMarkdownService();

    await expect(composeMarkdownService.renderSelection()).rejects.toThrow(
      RENDERED_SELECTION_BLOCKED_MESSAGE
    );
  });

  it("surfaces selection replacement failures", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.selectionSource = "body";
    mailboxItem.selectionText = "# Selected";
    mailboxItem.selectionHtml = "<p># Selected</p>";
    mailboxItem.body.failNextSetSelected = true;
    installOfficeEnvironment({ mailboxItem });

    const composeMarkdownService = createComposeMarkdownService();

    await expect(
      composeMarkdownService.renderSelection()
    ).rejects.toMatchObject({
      message: "Selected body write failed.",
      name: "BodySetSelectedError",
    });
  });
});
