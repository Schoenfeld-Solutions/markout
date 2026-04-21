import { createOfficeBodyAccessor } from "../src/lib/body-accessor";
import { FakeMailboxItem, installOfficeEnvironment } from "./helpers";

describe("body accessor", () => {
  beforeEach(() => {
    installOfficeEnvironment();
  });

  it("reads and writes the current draft html", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    const bodyAccessor = createOfficeBodyAccessor(mailboxItem);

    expect(await bodyAccessor.getHtml()).toBe("<div>Original</div>");

    await bodyAccessor.setHtml("<div>Updated</div>");

    expect(mailboxItem.body.currentHtml).toBe("<div>Updated</div>");
  });

  it("reads selection metadata for body selections", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.selectionSource = "body";
    mailboxItem.selectionText = "# Heading";
    mailboxItem.selectionHtml = "<p># Heading</p>";

    const bodyAccessor = createOfficeBodyAccessor(mailboxItem);

    await expect(bodyAccessor.getSelection()).resolves.toEqual({
      hasSelection: true,
      html: "<p># Heading</p>",
      source: "body",
      text: "# Heading",
    });
  });

  it("surfaces subject selections without html payloads", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.selectionSource = "subject";
    mailboxItem.selectionText = "Subject text";

    const bodyAccessor = createOfficeBodyAccessor(mailboxItem);

    await expect(bodyAccessor.getSelection()).resolves.toEqual({
      hasSelection: true,
      html: null,
      source: "subject",
      text: "Subject text",
    });
  });

  it("falls back cleanly when html selection data cannot be read", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.selectionSource = "body";
    mailboxItem.selectionText = "plain text";
    mailboxItem.nextHtmlSelectionError = {
      message: "The current selection cannot be read as html.",
      name: "InvalidFormatError",
    };

    const bodyAccessor = createOfficeBodyAccessor(mailboxItem);

    await expect(bodyAccessor.getSelection()).resolves.toEqual({
      hasSelection: true,
      html: null,
      source: "body",
      text: "plain text",
    });
  });

  it("replaces the current selection with html", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    const bodyAccessor = createOfficeBodyAccessor(mailboxItem);

    await bodyAccessor.replaceSelectionWithHtml("<div>Rendered fragment</div>");

    expect(mailboxItem.body.lastSelectedHtml).toBe(
      "<div>Rendered fragment</div>"
    );
  });

  it("surfaces getAsync failures", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.body.failNextGet = true;
    const bodyAccessor = createOfficeBodyAccessor(mailboxItem);

    await expect(bodyAccessor.getHtml()).rejects.toMatchObject({
      message: "Body read failed.",
      name: "BodyGetError",
    });
  });

  it("surfaces selection read failures", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.nextTextSelectionError = {
      message: "Selection read failed.",
      name: "SelectionReadError",
    };
    const bodyAccessor = createOfficeBodyAccessor(mailboxItem);

    await expect(bodyAccessor.getSelection()).rejects.toMatchObject({
      message: "Selection read failed.",
      name: "SelectionReadError",
    });
  });

  it("surfaces setAsync failures", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.body.failNextSet = true;
    const bodyAccessor = createOfficeBodyAccessor(mailboxItem);

    await expect(
      bodyAccessor.setHtml("<div>Updated</div>")
    ).rejects.toMatchObject({
      message: "Body write failed.",
      name: "BodySetError",
    });
  });

  it("surfaces body type read failures when replacing a selection", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.body.failNextGetType = true;
    const bodyAccessor = createOfficeBodyAccessor(mailboxItem);

    await expect(
      bodyAccessor.replaceSelectionWithHtml("<div>Updated</div>")
    ).rejects.toMatchObject({
      message: "Body type read failed.",
      name: "BodyTypeError",
    });
  });

  it("fails when the compose body is plain text", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.body.type = "text" as Office.CoercionType;
    const bodyAccessor = createOfficeBodyAccessor(mailboxItem);

    await expect(
      bodyAccessor.replaceSelectionWithHtml("<div>Updated</div>")
    ).rejects.toMatchObject({
      message:
        "MarkOut can only insert rendered content into an HTML compose body.",
      name: "UnsupportedBodyType",
    });
  });

  it("surfaces setSelectedDataAsync failures", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.body.failNextSetSelected = true;
    const bodyAccessor = createOfficeBodyAccessor(mailboxItem);

    await expect(
      bodyAccessor.replaceSelectionWithHtml("<div>Updated</div>")
    ).rejects.toMatchObject({
      message: "Selected body write failed.",
      name: "BodySetSelectedError",
    });
  });

  it("fails when no active compose item is available", () => {
    installOfficeEnvironment({ mailboxItem: undefined });

    expect(() => createOfficeBodyAccessor()).toThrow(
      "MarkOut requires an active Outlook compose item."
    );
  });
});
