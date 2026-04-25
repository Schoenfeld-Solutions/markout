import assert from "node:assert/strict";
import { existsSync } from "node:fs";
import { mkdir } from "node:fs/promises";
import path from "node:path";
import { chromium, type Browser, type Page } from "playwright-core";

interface TaskpaneUiConfig {
  baseUrl: string;
  browserExecutable: string;
  outputDirectory: string;
  timeoutMs: number;
}

interface PreviewScenario {
  colorScheme: "dark" | "light";
  height: number;
  name: string;
  width: number;
}

interface MockSnapshot {
  bodyHtml: string;
}

const DEFAULT_BASE_URL = "http://localhost:3000/taskpane-mock.html";
const DEFAULT_OUTPUT_DIRECTORY = path.join(
  process.cwd(),
  "output",
  "playwright"
);
const DEFAULT_TIMEOUT_MS = 30_000;
const MARKDOWN_SAMPLE = `# Preview heading

Paragraph with [a link](https://example.com) and \`inline code\`.

> Blockquote content

| Column | Value |
| --- | --- |
| Alpha | Beta |

\`\`\`ts
const preview = "theme-aware";
\`\`\`
`;
const NESTED_LIST_SAMPLE = `# List spacing

- parent
  - child
    - grandchild
`;
const RAPID_MARKDOWN_SAMPLE = `# Stable preview

Typing should keep the toolbar usable while the preview settles.

- first
- second
  - nested
`;
const PREVIEW_LOADING_TEXT = "Rendering preview...";

export async function runTaskpaneUiPlaywright(
  partialConfig: Partial<TaskpaneUiConfig> = {}
): Promise<void> {
  const config: TaskpaneUiConfig = {
    baseUrl: partialConfig.baseUrl ?? DEFAULT_BASE_URL,
    browserExecutable:
      partialConfig.browserExecutable ?? findBrowserExecutable(),
    outputDirectory: partialConfig.outputDirectory ?? DEFAULT_OUTPUT_DIRECTORY,
    timeoutMs: partialConfig.timeoutMs ?? DEFAULT_TIMEOUT_MS,
  };

  await mkdir(config.outputDirectory, { recursive: true });

  const browser = await chromium.launch({
    executablePath: config.browserExecutable,
    headless: true,
  });

  try {
    console.log("Verifying theme mode control in the local taskpane harness.");
    await verifyThemeModeControl(browser, config);

    const scenarios: PreviewScenario[] = [
      {
        colorScheme: "light",
        height: 844,
        name: "light-390x844",
        width: 390,
      },
      {
        colorScheme: "dark",
        height: 844,
        name: "dark-390x844",
        width: 390,
      },
      {
        colorScheme: "light",
        height: 570,
        name: "light-320x570",
        width: 320,
      },
      {
        colorScheme: "dark",
        height: 570,
        name: "dark-320x570",
        width: 320,
      },
    ];

    for (const scenario of scenarios) {
      console.log(`Verifying preview scenario ${scenario.name}.`);
      await verifyPreviewScenario(browser, config, scenario);
    }

    console.log("Verifying rapid input and toolbar interaction regressions.");
    await verifyRapidInputAndToolbarScenario(browser, config);
  } finally {
    await browser.close();
  }
}

function findBrowserExecutable(): string {
  const configuredExecutable = process.env.MARKOUT_TASKPANE_UI_BROWSER;

  if (
    configuredExecutable !== undefined &&
    configuredExecutable.trim().length > 0
  ) {
    return configuredExecutable;
  }

  const candidates = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
    "/usr/bin/google-chrome",
    "/usr/bin/chromium-browser",
    "/usr/bin/chromium",
  ];

  for (const candidate of candidates) {
    if (existsSync(candidate)) {
      return candidate;
    }
  }

  throw new Error(
    "Set MARKOUT_TASKPANE_UI_BROWSER to a Chrome or Chromium executable."
  );
}

async function openMockPage(
  browser: Browser,
  config: TaskpaneUiConfig,
  scenario: PreviewScenario,
  queryParams: Record<string, string> = {}
): Promise<Page> {
  const context = await browser.newContext({
    colorScheme: scenario.colorScheme,
    ignoreHTTPSErrors: true,
    viewport: {
      height: scenario.height,
      width: scenario.width,
    },
  });
  const page = await context.newPage();
  await page.goto(buildPageUrl(config.baseUrl, queryParams), {
    timeout: config.timeoutMs,
    waitUntil: "domcontentloaded",
  });
  await page.locator("#taskpane-shell").waitFor({ timeout: config.timeoutMs });
  return page;
}

function buildPageUrl(
  baseUrl: string,
  queryParams: Record<string, string>
): string {
  const pageUrl = new URL(baseUrl);

  for (const [key, value] of Object.entries(queryParams)) {
    pageUrl.searchParams.set(key, value);
  }

  return pageUrl.toString();
}

async function openInsertPanel(page: Page): Promise<void> {
  await clickElement(page, "#panel-button-insert");
}

async function openSettingsPanel(page: Page): Promise<void> {
  await clickElement(page, "#panel-button-settings");
}

async function openHelpPanel(page: Page): Promise<void> {
  await clickElement(page, "#panel-button-help");
}

async function verifyThemeModeControl(
  browser: Browser,
  config: TaskpaneUiConfig
): Promise<void> {
  const page = await openMockPage(browser, config, {
    colorScheme: "light",
    height: 844,
    name: "theme-control",
    width: 390,
  });

  try {
    await openSettingsPanel(page);

    await clickElement(page, "#theme-mode-dark");
    await assertTogglePressed(page, "#theme-mode-dark");

    await clickElement(page, "#theme-mode-light");
    await assertTogglePressed(page, "#theme-mode-light");

    await clickElement(page, "#theme-mode-system");
    await assertTogglePressed(page, "#theme-mode-system");

    await page.locator("#theme-editor .cm-content").waitFor({
      timeout: config.timeoutMs,
    });

    const editorText =
      (await page.locator("#theme-editor").textContent())?.replace(
        /\s+/g,
        " "
      ) ?? "";
    assert.match(editorText, /\.mo/);
    assert.match(editorText, /line-height:\s*1\.5/);
    assert.doesNotMatch(editorText, /font-family:/);
    assert.doesNotMatch(editorText, /font-size:\s*1em/);
    assert.doesNotMatch(editorText, /color:\s*inherit/);

    await openHelpPanel(page);
    const pageText =
      (await page.locator("#taskpane-shell").textContent()) ?? "";

    assert.ok(!pageText.includes("Track issues, releases"));
    assert.ok(
      !pageText.includes(
        "Open the GitHub Pages landing page with manifests, hosted docs, and deployment notes."
      )
    );
    assert.ok(
      !pageText.includes(
        "Open the Schoenfeld Solutions website. A support link can be added here later."
      )
    );
    assert.ok(
      !pageText.includes(
        "MarkOut keeps compose work Markdown-first while staying inside Outlook's taskpane and Smart Alerts model."
      )
    );
    assert.ok(
      !pageText.includes(
        "System follows Outlook theme when the host provides it and falls back to the browser preference otherwise."
      )
    );
  } finally {
    await page.context().close();
  }
}

async function verifyPreviewScenario(
  browser: Browser,
  config: TaskpaneUiConfig,
  scenario: PreviewScenario
): Promise<void> {
  const page = await openMockPage(browser, config, scenario);

  try {
    await openInsertPanel(page);
    await setTextareaValue(page, "#markdown-input", MARKDOWN_SAMPLE);
    await page.locator("#mo-preview").waitFor({ timeout: config.timeoutMs });
    await page
      .locator("#mo-preview")
      .getByText("Preview heading", { exact: false })
      .waitFor({ timeout: config.timeoutMs });

    const previewMetrics = await page.evaluate(() => {
      const frame = document.getElementById("mo-preview");
      const previewContent = frame?.firstElementChild as HTMLElement | null;

      if (frame === null || previewContent === null) {
        throw new Error("Preview frame is missing.");
      }

      const frameStyle = window.getComputedStyle(frame);
      const contentStyle = window.getComputedStyle(previewContent);

      return {
        backgroundColor: frameStyle.backgroundColor,
        color: contentStyle.color,
      };
    });

    assert.ok(
      colorsDoNotCollide(previewMetrics.color, previewMetrics.backgroundColor),
      `Preview text collides with the preview background in ${scenario.name}.`
    );

    await assertToolbarPinnedToViewport(page, scenario.name);

    await page.screenshot({
      fullPage: false,
      path: path.join(
        config.outputDirectory,
        `taskpane-preview-${scenario.name}.png`
      ),
    });

    await verifyNestedListSpacing(page, config, scenario.name);
    await clickElement(page, "#insert-rendered-markdown-button");

    const snapshot = await page.evaluate(() => {
      return window.__MARKOUT_TASKPANE_MOCK__?.getState() ?? null;
    });

    assert.ok(snapshot !== null, "Taskpane mock state is unavailable.");
    assertHostInheritOutput(snapshot);
  } finally {
    await page.context().close();
  }
}

async function verifyRapidInputAndToolbarScenario(
  browser: Browser,
  config: TaskpaneUiConfig
): Promise<void> {
  const page = await openMockPage(
    browser,
    config,
    {
      colorScheme: "dark",
      height: 570,
      name: "rapid-input-toolbar",
      width: 320,
    },
    { previewDelayMs: "120" }
  );

  try {
    await openInsertPanel(page);
    const textarea = page.locator("#markdown-input");
    await setTextareaValueInChunks(
      page,
      "#markdown-input",
      RAPID_MARKDOWN_SAMPLE
    );

    await openSettingsPanel(page);
    await page
      .locator("#theme-mode-system")
      .waitFor({ timeout: config.timeoutMs });

    await scrollElementIntoView(page, "#developer-tools-switch");
    await clickElement(page, "#developer-tools-switch");
    await page
      .locator("#panel-button-developer")
      .waitFor({ timeout: config.timeoutMs });

    await clickElement(page, "#panel-button-developer");
    await page
      .locator("#taskpane-shell")
      .getByText("Developer tools", { exact: true })
      .waitFor({ timeout: config.timeoutMs });

    await openHelpPanel(page);
    await page
      .locator("#taskpane-shell")
      .getByText("GitHub repository", { exact: true })
      .waitFor({ timeout: config.timeoutMs });

    await openInsertPanel(page);
    await page
      .locator("#mo-preview")
      .getByText("Stable preview", { exact: false })
      .waitFor({ timeout: config.timeoutMs });

    assert.equal(await textarea.inputValue(), RAPID_MARKDOWN_SAMPLE);

    for (let attempt = 0; attempt < 5; attempt += 1) {
      const shellText =
        (await page.locator("#taskpane-shell").textContent()) ?? "";
      assert.ok(
        !shellText.includes(PREVIEW_LOADING_TEXT),
        "Preview returned to the loading state after rapid input settled."
      );
      assert.equal(await textarea.inputValue(), RAPID_MARKDOWN_SAMPLE);
      await page.waitForTimeout(100);
    }

    await assertToolbarPinnedToViewport(page, "rapid-input-toolbar");
  } finally {
    await page.context().close();
  }
}

async function verifyNestedListSpacing(
  page: Page,
  config: TaskpaneUiConfig,
  scenarioName: string
): Promise<void> {
  await setTextareaValue(page, "#markdown-input", NESTED_LIST_SAMPLE);
  await page.locator("#mo-preview li > ul").first().waitFor({
    timeout: config.timeoutMs,
  });

  const nestedListMetrics = await page.evaluate(() => {
    const nestedList = document.querySelector<HTMLElement>(
      "#mo-preview li > ul"
    );

    if (nestedList === null) {
      throw new Error("Nested preview list is missing.");
    }

    const nestedListStyle = window.getComputedStyle(nestedList);

    return {
      marginBottom: Number.parseFloat(nestedListStyle.marginBottom),
      marginTop: Number.parseFloat(nestedListStyle.marginTop),
    };
  });

  assert.equal(
    nestedListMetrics.marginTop,
    0,
    `Nested list has an unexpected top margin in ${scenarioName}.`
  );
  assert.equal(
    nestedListMetrics.marginBottom,
    0,
    `Nested list has an unexpected bottom margin in ${scenarioName}.`
  );
}

async function assertToolbarPinnedToViewport(
  page: Page,
  scenarioName: string
): Promise<void> {
  const metrics = await page.evaluate(() => {
    const contentViewport = document.querySelector<HTMLElement>(
      '[data-testid="taskpane-content-viewport"]'
    );
    const toolbar = document.querySelector<HTMLElement>(
      '[data-testid="taskpane-toolbar"]'
    );

    if (contentViewport === null || toolbar === null) {
      throw new Error("Taskpane scroll regions are missing.");
    }

    contentViewport.scrollTop = 0;
    const initialToolbarRect = toolbar.getBoundingClientRect();
    const initialContentRect = contentViewport.getBoundingClientRect();
    const initialDocumentScrollTop =
      document.scrollingElement?.scrollTop ??
      document.documentElement.scrollTop;

    contentViewport.scrollTop = contentViewport.scrollHeight;
    const scrolledToolbarRect = toolbar.getBoundingClientRect();
    const scrolledDocumentScrollTop =
      document.scrollingElement?.scrollTop ??
      document.documentElement.scrollTop;

    return {
      contentBottom: initialContentRect.bottom,
      contentClientHeight: contentViewport.clientHeight,
      contentScrollHeight: contentViewport.scrollHeight,
      contentScrollTop: contentViewport.scrollTop,
      initialDocumentScrollTop,
      scrolledDocumentScrollTop,
      toolbarBottom: initialToolbarRect.bottom,
      toolbarTop: initialToolbarRect.top,
      toolbarTopAfterContentScroll: scrolledToolbarRect.top,
      viewportHeight: window.innerHeight,
    };
  });

  assert.ok(
    metrics.contentScrollHeight > metrics.contentClientHeight,
    `Content viewport does not expose its own scroll range in ${scenarioName}.`
  );
  assert.equal(
    metrics.initialDocumentScrollTop,
    0,
    `Document scrolled before content scroll in ${scenarioName}.`
  );
  assert.equal(
    metrics.scrolledDocumentScrollTop,
    0,
    `Document scrolled instead of the content viewport in ${scenarioName}.`
  );
  assert.ok(
    metrics.contentScrollTop > 0,
    `Content viewport did not scroll in ${scenarioName}.`
  );
  assert.ok(
    Math.abs(metrics.toolbarBottom - metrics.viewportHeight) <= 1,
    `Toolbar is not pinned to the viewport bottom in ${scenarioName}.`
  );
  assert.ok(
    Math.abs(metrics.contentBottom - metrics.toolbarTop) <= 1,
    `Content viewport does not end flush at the toolbar in ${scenarioName}.`
  );
  assert.ok(
    Math.abs(metrics.toolbarTopAfterContentScroll - metrics.toolbarTop) <= 1,
    `Toolbar moved while content scrolled in ${scenarioName}.`
  );
}

async function assertTogglePressed(
  page: Page,
  selector: string
): Promise<void> {
  const state = await page.locator(selector).evaluate((node) => {
    return (
      node.getAttribute("aria-checked") ??
      node.getAttribute("aria-pressed") ??
      node.getAttribute("data-checked")
    );
  });

  assert.equal(state, "true");
}

async function clickElement(page: Page, selector: string): Promise<void> {
  const clickTarget = await page.locator(selector).evaluate((node) => {
    const rect = node.getBoundingClientRect();
    const x = rect.left + rect.width / 2;
    const y = rect.top + rect.height / 2;
    const elementAtPoint = document.elementFromPoint(x, y);

    return {
      hitTestTarget:
        elementAtPoint === null
          ? null
          : {
              id: elementAtPoint.id,
              tagName: elementAtPoint.tagName,
            },
      height: rect.height,
      isHitTarget:
        elementAtPoint !== null &&
        (node === elementAtPoint || node.contains(elementAtPoint)),
      viewportHeight: window.innerHeight,
      viewportWidth: window.innerWidth,
      width: rect.width,
      x,
      y,
    };
  });

  assert.ok(clickTarget.width > 0, `${selector} has no clickable width.`);
  assert.ok(clickTarget.height > 0, `${selector} has no clickable height.`);
  assert.ok(clickTarget.x >= 0, `${selector} is left of the viewport.`);
  assert.ok(clickTarget.y >= 0, `${selector} is above the viewport.`);
  assert.ok(
    clickTarget.x <= clickTarget.viewportWidth,
    `${selector} is right of the viewport.`
  );
  assert.ok(
    clickTarget.y <= clickTarget.viewportHeight,
    `${selector} is below the viewport.`
  );
  assert.ok(
    clickTarget.isHitTarget,
    `${selector} is covered by ${JSON.stringify(clickTarget.hitTestTarget)}.`
  );

  await page.locator(selector).evaluate((node) => {
    (node as HTMLElement).click();
  });
}

async function setTextareaValue(
  page: Page,
  selector: string,
  value: string
): Promise<void> {
  await page.locator(selector).evaluate((node, nextValue) => {
    const textarea = node as HTMLTextAreaElement;
    const valueSetter = Object.getOwnPropertyDescriptor(
      window.HTMLTextAreaElement.prototype,
      "value"
    )?.set;

    textarea.focus();
    valueSetter?.call(textarea, nextValue);
    textarea.dispatchEvent(
      new InputEvent("input", {
        bubbles: true,
        data: nextValue,
        inputType: "insertText",
      })
    );
    textarea.dispatchEvent(new Event("change", { bubbles: true }));
  }, value);
}

async function setTextareaValueInChunks(
  page: Page,
  selector: string,
  value: string
): Promise<void> {
  await setTextareaValue(page, selector, "");

  for (let index = 1; index <= value.length; index += 1) {
    await setTextareaValue(page, selector, value.slice(0, index));
  }
}

async function scrollElementIntoView(
  page: Page,
  selector: string
): Promise<void> {
  await page.locator(selector).evaluate((node) => {
    node.scrollIntoView({ block: "center", inline: "nearest" });
  });
}

function assertHostInheritOutput(snapshot: MockSnapshot): void {
  assert.ok(snapshot.bodyHtml.includes("markout-fragment-host"));
  assert.ok(snapshot.bodyHtml.includes(".markout-fragment-host .mo"));
  assert.ok(!snapshot.bodyHtml.includes("-apple-system"));
  assert.ok(!snapshot.bodyHtml.includes("font-size: 14px"));
  assert.doesNotMatch(
    snapshot.bodyHtml,
    /\.markout-fragment-host \.mo\s*\{[^}]*\b(?:color|font-family|font-size|font)\b/i
  );
  assert.doesNotMatch(
    snapshot.bodyHtml,
    /\.markout-fragment-host h[1-6]\s*\{[^}]*\b(?:color|font-family)\b/i
  );
  assert.doesNotMatch(
    snapshot.bodyHtml,
    /\.markout-fragment-host a\s*\{[^}]*\bcolor\b/i
  );
}

function colorsDoNotCollide(
  foregroundColor: string,
  backgroundColor: string
): boolean {
  const foreground = parseRgbColor(foregroundColor);
  const background = parseRgbColor(backgroundColor);

  if (foreground === null || background === null) {
    return false;
  }

  const foregroundLuminance = getLuminance(foreground);
  const backgroundLuminance = getLuminance(background);
  const lighter = Math.max(foregroundLuminance, backgroundLuminance);
  const darker = Math.min(foregroundLuminance, backgroundLuminance);
  const contrastRatio = (lighter + 0.05) / (darker + 0.05);

  return contrastRatio >= 3;
}

function getLuminance(color: {
  blue: number;
  green: number;
  red: number;
}): number {
  const channels = [color.red, color.green, color.blue].map((value) => {
    const normalized = value / 255;

    return normalized <= 0.03928
      ? normalized / 12.92
      : ((normalized + 0.055) / 1.055) ** 2.4;
  });

  const [red = 0, green = 0, blue = 0] = channels;
  return 0.2126 * red + 0.7152 * green + 0.0722 * blue;
}

function parseRgbColor(
  color: string
): { blue: number; green: number; red: number } | null {
  const match = /rgba?\(\s*(\d+),\s*(\d+),\s*(\d+)/i.exec(color);

  if (match === null) {
    return null;
  }

  const [, red = "0", green = "0", blue = "0"] = match;
  return {
    blue: Number.parseInt(blue, 10),
    green: Number.parseInt(green, 10),
    red: Number.parseInt(red, 10),
  };
}
