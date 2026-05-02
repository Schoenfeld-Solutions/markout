/** @jest-environment jsdom */
import type { BodyAccessor } from "../src/lib/body-accessor";
import { DefaultHtmlSanitizer } from "../src/lib/html-sanitizer";
import {
  FULL_RENDER_BLOCKED_BY_FRAGMENT_MESSAGE,
  ensureRendered,
  createItemRenderer,
  renderItem,
} from "../src/lib/item";
import {
  createMarkdownRenderer,
  type MarkdownRenderer,
} from "../src/lib/renderer";
import type {
  RenderState,
  RenderStateStore,
} from "../src/lib/render-state-store";
import {
  FakeMailboxItem,
  installDomParser,
  installOfficeEnvironment,
} from "./helpers";

class InMemoryBodyAccessor implements BodyAccessor {
  public failNextSetHtml = false;

  public constructor(private html: string) {}

  public getHtml(): Promise<string> {
    return Promise.resolve(this.html);
  }

  public getSelection(): Promise<{
    hasSelection: boolean;
    html: string | null;
    source: "body";
    text: string;
  }> {
    return Promise.resolve({
      hasSelection: false,
      html: null,
      source: "body",
      text: "",
    });
  }

  public replaceSelectionWithHtml(html: string): Promise<void> {
    this.html = html;
    return Promise.resolve();
  }

  public setHtml(html: string): Promise<void> {
    if (this.failNextSetHtml) {
      this.failNextSetHtml = false;
      return Promise.reject(new Error("Body set failed."));
    }

    this.html = html;
    return Promise.resolve();
  }
}

class InMemoryRenderStateStore implements RenderStateStore {
  public failNextClear = false;
  private renderState: RenderState | null = null;

  public clearRenderState(): Promise<void> {
    if (this.failNextClear) {
      this.failNextClear = false;
      return Promise.reject(new Error("Render state clear failed."));
    }

    this.renderState = null;
    return Promise.resolve();
  }

  public getRenderState(): Promise<RenderState | null> {
    return Promise.resolve(this.renderState);
  }

  public setPendingRenderState(originalHtml: string): Promise<void> {
    this.renderState = {
      channelId: "production",
      originalHtml,
      phase: "pending",
      storedAt: new Date().toISOString(),
    };
    return Promise.resolve();
  }

  public setRenderedRenderState(originalHtml: string): Promise<void> {
    this.renderState = {
      channelId: "production",
      originalHtml,
      phase: "rendered",
      storedAt: new Date().toISOString(),
    };
    return Promise.resolve();
  }
}

describe("item renderer", () => {
  beforeEach(() => {
    installDomParser();
  });

  afterEach(() => {
    delete (globalThis as { Office?: typeof Office }).Office;
  });

  it("wires default Office dependencies for exported render helpers", async () => {
    const mailboxItem = new FakeMailboxItem(
      "<div># Default dependencies</div>"
    );
    installOfficeEnvironment({ mailboxItem });

    await expect(renderItem()).resolves.toBe("rendered");
    expect(mailboxItem.body.currentHtml).toContain("Default dependencies");

    await expect(ensureRendered()).resolves.toBe(false);
  });

  it("renders the draft and restores the original html on the next toggle", async () => {
    const originalHtml = "<div># Hello team</div>";
    const bodyAccessor = new InMemoryBodyAccessor(originalHtml);
    const renderStateStore = new InMemoryRenderStateStore();
    const markdownRenderer: MarkdownRenderer = {
      render(): Promise<string> {
        return Promise.resolve(
          `<div class="mo"><p>Rendered output</p><img src="https://example.com/safe.png" onerror="alert(1)"></div>`
        );
      },
    };

    const itemRenderer = createItemRenderer({
      bodyAccessor,
      htmlSanitizer: new DefaultHtmlSanitizer(),
      markdownRenderer,
      renderStateStore,
      settingsStore: {
        getStylesheet(): string {
          return ".mo { color: rgb(1, 2, 3); }";
        },
      },
    });

    expect(await itemRenderer.renderItem()).toBe("rendered");
    expect(await renderStateStore.getRenderState()).toEqual(
      expect.objectContaining({
        channelId: "production",
        originalHtml,
        phase: "rendered",
      })
    );
    expect(await bodyAccessor.getHtml()).toContain(
      `<img src="https://example.com/safe.png">`
    );
    expect(await bodyAccessor.getHtml()).not.toContain("onerror");

    expect(await itemRenderer.renderItem()).toBe("restored");
    expect(await renderStateStore.getRenderState()).toBeNull();
    expect(await bodyAccessor.getHtml()).toBe(originalHtml);
  });

  it("skips ensureRendered when the item is already rendered", async () => {
    const bodyAccessor = new InMemoryBodyAccessor(
      "<div>Already rendered</div>"
    );
    const renderStateStore = new InMemoryRenderStateStore();
    await renderStateStore.setRenderedRenderState("<div>Original</div>");
    let renderCalls = 0;

    const itemRenderer = createItemRenderer({
      bodyAccessor,
      htmlSanitizer: new DefaultHtmlSanitizer(),
      markdownRenderer: {
        render(): Promise<string> {
          renderCalls += 1;
          return Promise.resolve('<div class="mo">Should not run</div>');
        },
      },
      renderStateStore,
      settingsStore: {
        getStylesheet(): string {
          return "";
        },
      },
    });

    expect(await itemRenderer.ensureRendered()).toBe(false);
    expect(renderCalls).toBe(0);
  });

  it("recovers pending state by restoring the original html before re-rendering", async () => {
    const originalHtml = "<div>## Original draft</div>";
    const bodyAccessor = new InMemoryBodyAccessor("<div>Half rendered</div>");
    const renderStateStore = new InMemoryRenderStateStore();
    await renderStateStore.setPendingRenderState(originalHtml);
    let renderInput = "";

    const itemRenderer = createItemRenderer({
      bodyAccessor,
      htmlSanitizer: new DefaultHtmlSanitizer(),
      markdownRenderer: {
        render({ markdown }): Promise<string> {
          renderInput = markdown;
          return Promise.resolve(
            '<div class="mo"><p>Recovered output</p></div>'
          );
        },
      },
      renderStateStore,
      settingsStore: {
        getStylesheet(): string {
          return "";
        },
      },
    });

    expect(await itemRenderer.ensureRendered()).toBe(true);
    expect(renderInput).toContain("Original draft");
    expect(await renderStateStore.getRenderState()).toEqual(
      expect.objectContaining({
        channelId: "production",
        originalHtml,
        phase: "rendered",
      })
    );
    expect(await bodyAccessor.getHtml()).toContain("Recovered output");
  });

  it("returns unchanged when a pending recovery has no renderable markdown", async () => {
    const originalHtml = "<div>Hello team</div>";
    const bodyAccessor = new InMemoryBodyAccessor(originalHtml);
    const renderStateStore = new InMemoryRenderStateStore();
    await renderStateStore.setPendingRenderState(originalHtml);
    let renderCalls = 0;

    const itemRenderer = createItemRenderer({
      bodyAccessor,
      htmlSanitizer: new DefaultHtmlSanitizer(),
      markdownRenderer: {
        render(): Promise<string> {
          renderCalls += 1;
          return Promise.resolve("<div>Should not render</div>");
        },
      },
      renderStateStore,
      settingsStore: {
        getStylesheet(): string {
          return "";
        },
      },
    });

    expect(await itemRenderer.renderItem()).toBe("unchanged");
    expect(renderCalls).toBe(0);
    expect(await bodyAccessor.getHtml()).toBe(originalHtml);
  });

  it("skips ensureRendered when the current draft already contains a MarkOut render marker", async () => {
    const bodyAccessor = new InMemoryBodyAccessor(
      '<div class="mo markout-rendered"><p>Rendered output</p></div>'
    );
    const renderStateStore = new InMemoryRenderStateStore();
    let renderCalls = 0;

    const itemRenderer = createItemRenderer({
      bodyAccessor,
      htmlSanitizer: new DefaultHtmlSanitizer(),
      markdownRenderer: {
        render(): Promise<string> {
          renderCalls += 1;
          return Promise.resolve('<div class="mo markout-rendered">noop</div>');
        },
      },
      renderStateStore,
      settingsStore: {
        getStylesheet(): string {
          return "";
        },
      },
    });

    expect(await itemRenderer.ensureRendered()).toBe(false);
    expect(renderCalls).toBe(0);
  });

  it("blocks ensureRendered when the current draft already contains a MarkOut fragment", async () => {
    const bodyAccessor = new InMemoryBodyAccessor(
      '<div class="markout-fragment-host"><div class="mo markout-fragment-rendered">Rendered fragment</div></div>'
    );
    const renderStateStore = new InMemoryRenderStateStore();
    const itemRenderer = createItemRenderer({
      bodyAccessor,
      htmlSanitizer: new DefaultHtmlSanitizer(),
      markdownRenderer: {
        render(): Promise<string> {
          return Promise.resolve('<div class="mo markout-rendered">noop</div>');
        },
      },
      renderStateStore,
      settingsStore: {
        getStylesheet(): string {
          return "";
        },
      },
    });

    await expect(itemRenderer.ensureRendered()).rejects.toThrow(
      FULL_RENDER_BLOCKED_BY_FRAGMENT_MESSAGE
    );
  });

  it("clears render state quietly when body writes fail", async () => {
    const bodyAccessor = new InMemoryBodyAccessor("<div># Broken write</div>");
    bodyAccessor.failNextSetHtml = true;
    const renderStateStore = new InMemoryRenderStateStore();
    renderStateStore.failNextClear = true;
    const itemRenderer = createItemRenderer({
      bodyAccessor,
      htmlSanitizer: new DefaultHtmlSanitizer(),
      markdownRenderer: {
        render(): Promise<string> {
          return Promise.resolve(
            '<div class="mo markout-rendered"><h1>Broken write</h1></div>'
          );
        },
      },
      renderStateStore,
      settingsStore: {
        getStylesheet(): string {
          return "";
        },
      },
    });

    await expect(itemRenderer.renderItem()).rejects.toThrow("Body set failed.");
  });

  it("fails restore when the rendered marker is present but the original html is unavailable", async () => {
    const bodyAccessor = new InMemoryBodyAccessor(
      '<div class="mo markout-rendered"><p>Rendered output</p></div>'
    );
    const renderStateStore = new InMemoryRenderStateStore();

    const itemRenderer = createItemRenderer({
      bodyAccessor,
      htmlSanitizer: new DefaultHtmlSanitizer(),
      markdownRenderer: {
        render(): Promise<string> {
          return Promise.resolve('<div class="mo markout-rendered">noop</div>');
        },
      },
      renderStateStore,
      settingsStore: {
        getStylesheet(): string {
          return "";
        },
      },
    });

    await expect(itemRenderer.renderItem()).rejects.toThrow(
      "Outlook didn't preserve the original HTML for restore"
    );
  });

  it("blocks full-draft rendering when the draft already contains a rendered MarkOut fragment", async () => {
    const bodyAccessor = new InMemoryBodyAccessor(
      '<div class="markout-fragment-host"><div class="mo markout-fragment-rendered">Rendered fragment</div></div>'
    );
    const renderStateStore = new InMemoryRenderStateStore();

    const itemRenderer = createItemRenderer({
      bodyAccessor,
      htmlSanitizer: new DefaultHtmlSanitizer(),
      markdownRenderer: {
        render(): Promise<string> {
          return Promise.resolve('<div class="mo markout-rendered">noop</div>');
        },
      },
      renderStateStore,
      settingsStore: {
        getStylesheet(): string {
          return "";
        },
      },
    });

    await expect(itemRenderer.renderItem()).rejects.toThrow(
      FULL_RENDER_BLOCKED_BY_FRAGMENT_MESSAGE
    );
  });

  it("renders only markdown-looking draft blocks and preserves signatures", async () => {
    const signatureHtml =
      '<div id="owa-signature" class="signature"><p>Kind regards,<br>Gabriel</p><img src="https://example.com/logo.png"></div>';
    const originalHtml = [
      "<div># Release notes</div>",
      "<div>- fixed selection rendering</div>",
      "<div>&nbsp;&nbsp;- kept nested list spacing tight</div>",
      signatureHtml,
    ].join("");
    const bodyAccessor = new InMemoryBodyAccessor(originalHtml);
    const renderStateStore = new InMemoryRenderStateStore();

    const itemRenderer = createItemRenderer({
      bodyAccessor,
      htmlSanitizer: new DefaultHtmlSanitizer(),
      markdownRenderer: createMarkdownRenderer(),
      renderStateStore,
      settingsStore: {
        getStylesheet(): string {
          return "";
        },
      },
    });

    expect(await itemRenderer.renderItem()).toBe("rendered");

    const renderedHtml = await bodyAccessor.getHtml();
    expect(renderedHtml).toContain("<h1>Release notes</h1>");
    expect(renderedHtml).toContain("<li>fixed selection rendering");
    expect(renderedHtml).toContain("<li>kept nested list spacing tight</li>");
    expect(renderedHtml).toContain(signatureHtml);

    expect(await itemRenderer.renderItem()).toBe("restored");
    expect(await bodyAccessor.getHtml()).toBe(originalHtml);
  });

  it("leaves non-markdown draft html unchanged", async () => {
    const originalHtml =
      '<div>Hello team,<br>please review the attached file.</div><div class="signature">Kind regards,<br>Gabriel</div>';
    const bodyAccessor = new InMemoryBodyAccessor(originalHtml);
    const renderStateStore = new InMemoryRenderStateStore();
    let renderCalls = 0;

    const itemRenderer = createItemRenderer({
      bodyAccessor,
      htmlSanitizer: new DefaultHtmlSanitizer(),
      markdownRenderer: {
        render(): Promise<string> {
          renderCalls += 1;
          return Promise.resolve("<div>Should not render</div>");
        },
      },
      renderStateStore,
      settingsStore: {
        getStylesheet(): string {
          return "";
        },
      },
    });

    expect(await itemRenderer.renderItem()).toBe("unchanged");
    expect(await bodyAccessor.getHtml()).toBe(originalHtml);
    expect(await renderStateStore.getRenderState()).toBeNull();
    expect(renderCalls).toBe(0);
  });

  it.each([
    ["table", "<div>| A | B |\n| --- | --- |</div>"],
    ["blockquote", "<div>&gt; quoted</div>"],
    ["ordered list", "<div>1. first</div>"],
    ["fenced code", "<div>```</div>"],
    ["thematic break", "<div>---</div>"],
    ["inline emphasis", "<div>**strong**</div>"],
    ["inline code", "<div>`code`</div>"],
    ["link", "<div>[docs](https://example.test)</div>"],
    ["image", "<div>![alt](https://example.test/a.png)</div>"],
  ])(
    "renders markdown-looking %s draft segments",
    async (_name, originalHtml) => {
      const bodyAccessor = new InMemoryBodyAccessor(originalHtml);
      const renderStateStore = new InMemoryRenderStateStore();
      let renderInput = "";

      const itemRenderer = createItemRenderer({
        bodyAccessor,
        htmlSanitizer: new DefaultHtmlSanitizer(),
        markdownRenderer: {
          render({ markdown }): Promise<string> {
            renderInput = markdown;
            return Promise.resolve(
              '<div class="mo markout-rendered"><p>Rendered marker</p></div>'
            );
          },
        },
        renderStateStore,
        settingsStore: {
          getStylesheet(): string {
            return "";
          },
        },
      });

      expect(await itemRenderer.renderItem()).toBe("rendered");
      expect(renderInput.length).toBeGreaterThan(0);
      expect(await bodyAccessor.getHtml()).toContain("Rendered marker");
    }
  );
});
