import assert from "node:assert/strict";
import { existsSync } from "node:fs";
import { mkdir } from "node:fs/promises";
import path from "node:path";
import {
  chromium,
  type Browser,
  type Frame,
  type LaunchOptions,
  type Page,
} from "playwright-core";

interface TaskpaneUiConfig {
  baseUrl: string;
  browserExecutable: string;
  headless: boolean;
  outputDirectory: string;
  owaHostUrl: string;
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
  lastInsertedHtml: string | null;
  transientNotifications: {
    intent: "error" | "info" | "success" | "warning";
    message: string;
  }[];
}

interface OwaHostPage {
  page: Page;
  taskpane: Frame;
}

interface ToolbarLayoutOptions {
  expectScrollableContent: boolean;
}

type TaskpaneSurface = Frame | Page;

const DEFAULT_BASE_URL = "http://localhost:3000/taskpane-mock.html";
const DEFAULT_OWA_HOST_URL = "http://localhost:3000/owa-taskpane-host.html";
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
const PASTED_NBSP_LIST_SAMPLE = [
  "# hi",
  "",
  "ich bin",
  "- cool",
  "\u00a0\u00a0- super cool",
  "- cool",
].join("\n");
const RAPID_MARKDOWN_SAMPLE = `# Stable preview

Typing should keep the toolbar usable while the preview settles.

- first
- second
  - nested
`;
const LONG_DRAWER_MARKDOWN_SAMPLE = `# Long drawer content

${Array.from(
  { length: 24 },
  (_value, index) =>
    `- Drawer line ${String(index + 1).padStart(2, "0")} keeps the content viewport scrollable.`
).join("\n")}
`;
const COMPLEX_SELECTION_TEXT = "# Selection Title Paragraph - parent - child";
const COMPLEX_SELECTION_HTML = [
  "<div># Selection Title</div>",
  "<div>Paragraph from Outlook selection</div>",
  "<div>- parent</div>",
  "<div>&nbsp;&nbsp;- child</div>",
].join("");
const SIGNATURE_HTML =
  '<div id="owa-signature" class="signature"><p>Kind regards,<br>Gabriel</p><img src="https://example.com/logo.png"></div>';
const DRAFT_WITH_SIGNATURE_HTML = [
  "<div># Draft Title</div>",
  "<div>- parent</div>",
  "<div>&nbsp;&nbsp;- child</div>",
  SIGNATURE_HTML,
].join("");
const DRAFT_WITHOUT_MARKDOWN_HTML =
  '<div>Hello team,<br>please review the attached file.</div><div class="signature">Kind regards,<br>Gabriel</div>';
const PREVIEW_LOADING_TEXT = "Rendering preview...";

export async function runTaskpaneUiPlaywright(
  partialConfig: Partial<TaskpaneUiConfig> = {}
): Promise<void> {
  const config: TaskpaneUiConfig = {
    baseUrl: partialConfig.baseUrl ?? DEFAULT_BASE_URL,
    browserExecutable:
      partialConfig.browserExecutable ?? findBrowserExecutable(),
    headless: partialConfig.headless ?? true,
    owaHostUrl: partialConfig.owaHostUrl ?? DEFAULT_OWA_HOST_URL,
    outputDirectory: partialConfig.outputDirectory ?? DEFAULT_OUTPUT_DIRECTORY,
    timeoutMs: partialConfig.timeoutMs ?? DEFAULT_TIMEOUT_MS,
  };

  await mkdir(config.outputDirectory, { recursive: true });

  const launchOptions: LaunchOptions = {
    executablePath: config.browserExecutable,
    headless: config.headless,
  };

  if (!config.headless) {
    launchOptions.slowMo = 50;
  }

  const browser = await chromium.launch(launchOptions);

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

    console.log("Verifying selection and draft rendering regressions.");
    await verifySelectionAndDraftRenderingScenario(browser, config);

    console.log("Verifying OWA-like drawer host layout regressions.");
    await verifyOwaLikeDrawerHostScenario(browser, config);
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

async function openOwaHostPage(
  browser: Browser,
  config: TaskpaneUiConfig,
  scenario: PreviewScenario,
  queryParams: Record<string, string> = {}
): Promise<OwaHostPage> {
  const context = await browser.newContext({
    colorScheme: scenario.colorScheme,
    ignoreHTTPSErrors: true,
    viewport: {
      height: scenario.height,
      width: scenario.width,
    },
  });
  const page = await context.newPage();

  await page.goto(buildPageUrl(config.owaHostUrl, queryParams), {
    timeout: config.timeoutMs,
    waitUntil: "domcontentloaded",
  });

  const frameLocator = page.locator('[data-testid="owa-taskpane-frame"]');
  await frameLocator.waitFor({ state: "visible", timeout: config.timeoutMs });

  const frameHandle = await frameLocator.elementHandle({
    timeout: config.timeoutMs,
  });
  const taskpane = await frameHandle?.contentFrame();

  if (taskpane === null || taskpane === undefined) {
    throw new Error("OWA-like taskpane iframe did not expose a frame.");
  }

  await taskpane
    .locator("#taskpane-shell")
    .waitFor({ timeout: config.timeoutMs });

  return { page, taskpane };
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

async function openInsertPanel(page: TaskpaneSurface): Promise<void> {
  await clickElement(page, "#panel-button-insert");
}

async function openSettingsPanel(page: TaskpaneSurface): Promise<void> {
  await clickElement(page, "#panel-button-settings");
}

async function openHelpPanel(page: TaskpaneSurface): Promise<void> {
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

    await assertToolbarPinnedToViewport(page, scenario.name, {
      expectScrollableContent: true,
    });

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

    await assertToolbarPinnedToViewport(page, "rapid-input-toolbar", {
      expectScrollableContent: true,
    });
  } finally {
    await page.context().close();
  }
}

async function verifySelectionAndDraftRenderingScenario(
  browser: Browser,
  config: TaskpaneUiConfig
): Promise<void> {
  const page = await openMockPage(browser, config, {
    colorScheme: "light",
    height: 570,
    name: "selection-draft-rendering",
    width: 320,
  });

  try {
    await openInsertPanel(page);
    await page.evaluate(
      ({ html, text }) => {
        window.__MARKOUT_TASKPANE_MOCK__?.setSelection({
          hasSelection: true,
          html,
          source: "body",
          text,
        });
        window.dispatchEvent(new Event("focus"));
      },
      {
        html: COMPLEX_SELECTION_HTML,
        text: COMPLEX_SELECTION_TEXT,
      }
    );
    await page.locator("#render-selection-button:not([disabled])").waitFor({
      timeout: config.timeoutMs,
    });
    await scrollElementIntoView(page, "#render-selection-button");
    await clickElement(page, "#render-selection-button");
    await page.waitForFunction(() => {
      return (
        window.__MARKOUT_TASKPANE_MOCK__
          ?.getState()
          .lastInsertedHtml?.includes("Selection Title") ?? false
      );
    });

    const selectionSnapshot = await readMockSnapshot(page);
    assert.ok(
      selectionSnapshot.lastInsertedHtml !== null,
      "Selection render did not write any HTML."
    );
    assert.ok(
      selectionSnapshot.lastInsertedHtml.includes("<h1>Selection Title</h1>"),
      "Selection heading was not preserved as a heading."
    );
    assert.ok(
      selectionSnapshot.lastInsertedHtml.includes(
        "<p>Paragraph from Outlook selection</p>"
      ),
      "Selection paragraph was not preserved as its own block."
    );
    assert.ok(
      selectionSnapshot.lastInsertedHtml.includes("<li>parent"),
      "Selection parent list item was not rendered as a list item."
    );
    assert.ok(
      selectionSnapshot.lastInsertedHtml.includes("<li>child</li>"),
      "Selection child list item was not rendered as a nested list item."
    );
    assert.ok(
      !selectionSnapshot.lastInsertedHtml.includes(COMPLEX_SELECTION_TEXT),
      "Selection render used the flattened Outlook text instead of HTML structure."
    );

    await page.evaluate((bodyHtml) => {
      window.__MARKOUT_TASKPANE_MOCK__?.reset();
      window.__MARKOUT_TASKPANE_MOCK__?.setBodyHtml(bodyHtml);
    }, DRAFT_WITH_SIGNATURE_HTML);
    await openInsertPanel(page);
    await scrollElementIntoView(page, "#render-entire-draft-button");
    await clickElement(page, "#render-entire-draft-button");
    await page.waitForFunction(() => {
      return (
        window.__MARKOUT_TASKPANE_MOCK__
          ?.getState()
          .bodyHtml.includes("Draft Title") ?? false
      );
    });

    const renderedDraftSnapshot = await readMockSnapshot(page);
    assert.match(
      renderedDraftSnapshot.bodyHtml,
      /<h1\b[^>]*>Draft Title<\/h1>/
    );
    assert.match(renderedDraftSnapshot.bodyHtml, /<li\b[^>]*>parent/);
    assert.match(renderedDraftSnapshot.bodyHtml, /<li\b[^>]*>child<\/li>/);
    assert.ok(
      renderedDraftSnapshot.bodyHtml.includes(SIGNATURE_HTML),
      "Draft render changed the non-Markdown signature HTML."
    );

    await scrollElementIntoView(page, "#render-entire-draft-button");
    await clickElement(page, "#render-entire-draft-button");
    await page.waitForFunction((originalBodyHtml) => {
      return (
        window.__MARKOUT_TASKPANE_MOCK__?.getState().bodyHtml ===
        originalBodyHtml
      );
    }, DRAFT_WITH_SIGNATURE_HTML);

    await page.evaluate((bodyHtml) => {
      window.__MARKOUT_TASKPANE_MOCK__?.reset();
      window.__MARKOUT_TASKPANE_MOCK__?.setBodyHtml(bodyHtml);
    }, DRAFT_WITHOUT_MARKDOWN_HTML);
    await openInsertPanel(page);
    await scrollElementIntoView(page, "#render-entire-draft-button");
    await clickElement(page, "#render-entire-draft-button");

    const unchangedDraftSnapshot = await readMockSnapshot(page);
    assert.equal(unchangedDraftSnapshot.bodyHtml, DRAFT_WITHOUT_MARKDOWN_HTML);
    assert.ok(
      unchangedDraftSnapshot.transientNotifications.some(
        (notification) =>
          notification.intent === "info" &&
          notification.message.includes("No Markdown-looking draft blocks")
      ),
      "No-op draft render did not surface an informational notification."
    );
  } finally {
    await page.context().close();
  }
}

async function verifyOwaLikeDrawerHostScenario(
  browser: Browser,
  config: TaskpaneUiConfig
): Promise<void> {
  const { page, taskpane } = await openOwaHostPage(
    browser,
    config,
    {
      colorScheme: "dark",
      height: 1024,
      name: "owa-like-drawer",
      width: 644,
    },
    { previewDelayMs: "120" }
  );

  try {
    await openHelpPanel(taskpane);
    await taskpane
      .locator("#taskpane-shell")
      .getByText("GitHub repository", { exact: true })
      .waitFor({ timeout: config.timeoutMs });

    await assertOwaHostFrameLayout(page, "owa-like-help-short-content");
    await assertToolbarPinnedToViewport(
      taskpane,
      "owa-like-help-short-content",
      {
        expectScrollableContent: false,
      }
    );

    await page.screenshot({
      fullPage: false,
      path: path.join(config.outputDirectory, "taskpane-owa-like-help.png"),
    });

    await openInsertPanel(taskpane);
    await setTextareaValue(
      taskpane,
      "#markdown-input",
      LONG_DRAWER_MARKDOWN_SAMPLE
    );
    await waitForPreviewText(taskpane, "Long drawer content", config.timeoutMs);

    await assertOwaHostFrameLayout(page, "owa-like-insert-long-content");
    await assertToolbarPinnedToViewport(
      taskpane,
      "owa-like-insert-long-content",
      {
        expectScrollableContent: true,
      }
    );

    await setTextareaValueInChunks(
      taskpane,
      "#markdown-input",
      RAPID_MARKDOWN_SAMPLE
    );
    await openSettingsPanel(taskpane);
    await taskpane
      .locator("#theme-mode-system")
      .waitFor({ timeout: config.timeoutMs });
    await openHelpPanel(taskpane);
    await taskpane
      .locator("#taskpane-shell")
      .getByText("GitHub repository", { exact: true })
      .waitFor({ timeout: config.timeoutMs });
    await openInsertPanel(taskpane);
    await waitForPreviewText(taskpane, "Stable preview", config.timeoutMs);
    assert.equal(
      await taskpane.locator("#markdown-input").inputValue(),
      RAPID_MARKDOWN_SAMPLE
    );
  } finally {
    await page.context().close();
  }
}

async function verifyNestedListSpacing(
  page: TaskpaneSurface,
  config: TaskpaneUiConfig,
  scenarioName: string
): Promise<void> {
  await setTextareaValue(page, "#markdown-input", NESTED_LIST_SAMPLE);
  await page.locator("#mo-preview li > ul").first().waitFor({
    timeout: config.timeoutMs,
  });
  await assertNestedListSpacing(page, scenarioName);

  await setTextareaValue(page, "#markdown-input", PASTED_NBSP_LIST_SAMPLE);
  await page.waitForFunction(() => {
    return (
      document
        .querySelector<HTMLTextAreaElement>("#markdown-input")
        ?.value.includes("  - super cool") ?? false
    );
  });
  await page.locator("#mo-preview li > ul li").getByText("super cool").waitFor({
    timeout: config.timeoutMs,
  });

  const pastedListSnapshot = await page.evaluate(() => {
    const textarea =
      document.querySelector<HTMLTextAreaElement>("#markdown-input");
    const parentListItem =
      document.querySelector<HTMLElement>("#mo-preview li");
    const nestedListItem = document.querySelector<HTMLElement>(
      "#mo-preview li > ul li"
    );

    if (
      textarea === null ||
      parentListItem === null ||
      nestedListItem === null
    ) {
      throw new Error("Pasted nested list preview is missing.");
    }

    return {
      nestedText: nestedListItem.textContent.trim(),
      parentText: parentListItem.textContent,
      textareaValue: textarea.value,
    };
  });

  assert.ok(
    pastedListSnapshot.textareaValue.includes("  - super cool"),
    `Pasted non-breaking indentation was not normalized in ${scenarioName}.`
  );
  assert.ok(
    !pastedListSnapshot.textareaValue.includes("\u00a0\u00a0- super cool"),
    `Pasted non-breaking indentation remained in ${scenarioName}.`
  );
  assert.equal(pastedListSnapshot.nestedText, "super cool");
  assert.ok(
    !pastedListSnapshot.parentText.includes("- super cool"),
    `Pasted sublist rendered as parent item text in ${scenarioName}.`
  );

  await assertNestedListSpacing(page, `${scenarioName}-pasted-nbsp`);
}

async function assertNestedListSpacing(
  page: TaskpaneSurface,
  scenarioName: string
): Promise<void> {
  const nestedListMetrics = await page.evaluate(() => {
    const nestedList = document.querySelector<HTMLElement>(
      "#mo-preview li > ul"
    );
    const parentListItem =
      document.querySelector<HTMLElement>("#mo-preview li");

    if (nestedList === null || parentListItem === null) {
      throw new Error("Nested preview list is missing.");
    }

    const nestedListStyle = window.getComputedStyle(nestedList);
    const parentTextNode = Array.from(parentListItem.childNodes).find(
      (childNode) =>
        childNode.nodeType === Node.TEXT_NODE &&
        (childNode.textContent ?? "").trim().length > 0
    );

    if (parentTextNode === undefined) {
      throw new Error("Nested preview list parent text is missing.");
    }

    const textRange = document.createRange();
    textRange.selectNodeContents(parentTextNode);
    const textRect = textRange.getBoundingClientRect();
    const nestedListRect = nestedList.getBoundingClientRect();

    return {
      visualGap: nestedListRect.top - textRect.bottom,
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
  assert.ok(
    nestedListMetrics.visualGap <= 3,
    `Nested list is visually detached from its parent in ${scenarioName}.`
  );
}

async function readMockSnapshot(page: TaskpaneSurface): Promise<MockSnapshot> {
  const snapshot = await page.evaluate(() => {
    return window.__MARKOUT_TASKPANE_MOCK__?.getState() ?? null;
  });

  assert.ok(snapshot !== null, "Taskpane mock state is unavailable.");
  return snapshot;
}

async function assertOwaHostFrameLayout(
  page: Page,
  scenarioName: string
): Promise<void> {
  const metrics = await page.evaluate(() => {
    const drawer = document.querySelector<HTMLElement>(
      '[data-testid="owa-drawer"]'
    );
    const drawerBody = document.querySelector<HTMLElement>(
      '[data-testid="owa-drawer-body"]'
    );
    const frame = document.querySelector<HTMLIFrameElement>(
      '[data-testid="owa-taskpane-frame"]'
    );
    const frameHost = document.querySelector<HTMLElement>(
      '[data-testid="owa-frame-host"]'
    );

    if (
      drawer === null ||
      drawerBody === null ||
      frame === null ||
      frameHost === null
    ) {
      throw new Error("OWA-like drawer host elements are missing.");
    }

    const drawerRect = drawer.getBoundingClientRect();
    const drawerBodyRect = drawerBody.getBoundingClientRect();
    const frameRect = frame.getBoundingClientRect();
    const frameHostRect = frameHost.getBoundingClientRect();
    const documentScrollTop =
      document.scrollingElement?.scrollTop ??
      document.documentElement.scrollTop;

    return {
      documentScrollTop,
      drawerBodyBottom: drawerBodyRect.bottom,
      drawerBodyTop: drawerBodyRect.top,
      drawerBottom: drawerRect.bottom,
      drawerHeight: drawerRect.height,
      drawerWidth: drawerRect.width,
      frameBottom: frameRect.bottom,
      frameHeight: frameRect.height,
      frameHostBottom: frameHostRect.bottom,
      frameHostHeight: frameHostRect.height,
      frameHostTop: frameHostRect.top,
      frameTop: frameRect.top,
      frameWidth: frameRect.width,
      viewportHeight: window.innerHeight,
    };
  });

  assert.equal(
    metrics.documentScrollTop,
    0,
    `OWA-like host document scrolled in ${scenarioName}.`
  );
  assert.ok(
    Math.abs(metrics.drawerWidth - 320) <= 1,
    `OWA-like drawer width drifted in ${scenarioName}.`
  );
  assert.ok(
    Math.abs(metrics.drawerBottom - metrics.viewportHeight) <= 1,
    `OWA-like drawer does not fill the viewport in ${scenarioName}.`
  );
  assert.ok(
    metrics.drawerHeight > metrics.frameHeight,
    `OWA-like drawer header is not outside the iframe in ${scenarioName}.`
  );
  assert.ok(
    Math.abs(metrics.drawerBodyTop - metrics.frameTop) <= 1,
    `OWA-like iframe does not start flush with the drawer body in ${scenarioName}.`
  );
  assert.ok(
    Math.abs(metrics.drawerBodyBottom - metrics.frameBottom) <= 1,
    `OWA-like iframe does not end flush with the drawer body in ${scenarioName}.`
  );
  assert.ok(
    Math.abs(metrics.frameHostTop - metrics.frameTop) <= 1,
    `OWA-like frame host does not start flush with the iframe in ${scenarioName}.`
  );
  assert.ok(
    Math.abs(metrics.frameHostBottom - metrics.frameBottom) <= 1,
    `OWA-like frame host does not end flush with the iframe in ${scenarioName}.`
  );
  assert.ok(
    Math.abs(metrics.frameHostHeight - metrics.frameHeight) <= 1,
    `OWA-like frame host height differs from the iframe in ${scenarioName}.`
  );
  assert.ok(
    Math.abs(metrics.frameWidth - 320) <= 1,
    `OWA-like iframe width drifted in ${scenarioName}.`
  );
}

async function assertToolbarPinnedToViewport(
  page: TaskpaneSurface,
  scenarioName: string,
  options: ToolbarLayoutOptions
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

  if (options.expectScrollableContent) {
    assert.ok(
      metrics.contentScrollHeight > metrics.contentClientHeight,
      `Content viewport does not expose its own scroll range in ${scenarioName}.`
    );
    assert.ok(
      metrics.contentScrollTop > 0,
      `Content viewport did not scroll in ${scenarioName}.`
    );
  }

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

async function clickElement(
  page: TaskpaneSurface,
  selector: string
): Promise<void> {
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
  page: TaskpaneSurface,
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
  page: TaskpaneSurface,
  selector: string,
  value: string
): Promise<void> {
  await setTextareaValue(page, selector, "");

  for (let index = 1; index <= value.length; index += 1) {
    await setTextareaValue(page, selector, value.slice(0, index));
  }
}

async function waitForPreviewText(
  page: TaskpaneSurface,
  expectedText: string,
  timeoutMs: number
): Promise<void> {
  await page.waitForFunction(
    (text) => {
      const preview = document.querySelector("#mo-preview");
      return preview?.textContent.includes(text) ?? false;
    },
    expectedText,
    { timeout: timeoutMs }
  );
}

async function scrollElementIntoView(
  page: TaskpaneSurface,
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
