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

  it("surfaces getAsync failures", async () => {
    const mailboxItem = new FakeMailboxItem("<div>Original</div>");
    mailboxItem.body.failNextGet = true;
    const bodyAccessor = createOfficeBodyAccessor(mailboxItem);

    await expect(bodyAccessor.getHtml()).rejects.toMatchObject({
      message: "Body read failed.",
      name: "BodyGetError",
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

  it("fails when no active compose item is available", () => {
    installOfficeEnvironment({ mailboxItem: undefined });

    expect(() => createOfficeBodyAccessor()).toThrow(
      "MarkOut requires an active Outlook compose item."
    );
  });
});
