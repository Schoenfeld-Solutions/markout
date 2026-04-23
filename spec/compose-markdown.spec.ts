import {
  EMPTY_SELECTION_MESSAGE,
  FULL_DRAFT_ALREADY_RENDERED_MESSAGE,
  RENDERED_SELECTION_BLOCKED_MESSAGE,
  SUBJECT_SELECTION_UNSUPPORTED_MESSAGE,
  createComposeMarkdownService,
} from "../src/lib/compose-markdown";
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
