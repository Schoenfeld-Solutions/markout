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

const DEFAULT_BASE_URL = "https://localhost:3000/taskpane-mock.html";
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
  scenario: PreviewScenario
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
  await page.goto(config.baseUrl, {
    timeout: config.timeoutMs,
    waitUntil: "domcontentloaded",
  });
  await page.locator("#taskpane-shell").waitFor({ timeout: config.timeoutMs });
  return page;
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

    await page.screenshot({
      fullPage: true,
      path: path.join(
        config.outputDirectory,
        `taskpane-preview-${scenario.name}.png`
      ),
    });

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

    valueSetter?.call(textarea, nextValue);
    textarea.dispatchEvent(new Event("input", { bubbles: true }));
    textarea.dispatchEvent(new Event("change", { bubbles: true }));
  }, value);
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
