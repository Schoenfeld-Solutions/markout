import { mkdir } from "fs/promises";
import { existsSync } from "fs";
import path from "path";
import {
  chromium,
  type FrameLocator,
  type Locator,
  type Page,
} from "playwright-core";

interface HostSmokeConfig {
  autoRenderButtonPrefix: string;
  browserExecutable: string;
  composeUrl: string;
  headless: boolean;
  messageBodySelector: string;
  openButtonSelector: string | null;
  openButtonText: string;
  outputDirectory: string;
  previewSelector: string;
  recipient: string;
  renderButtonText: string;
  sendButtonSelector: string | null;
  sendButtonText: string;
  sentConfirmationText: string;
  storageStatePath: string;
  taskpaneFrameSelector: string;
  timeoutMs: number;
  toFieldSelector: string;
}

const DEFAULT_COMPOSE_URL = "https://outlook.office.com/mail/deeplink/compose";
const DEFAULT_OUTPUT_DIRECTORY = path.join(
  process.cwd(),
  "output",
  "playwright"
);
const DEFAULT_TASKPANE_FRAME_SELECTOR = 'iframe[src*="taskpane.html"]';
const DEFAULT_TIMEOUT_MS = 30_000;

void runHostSmoke().catch((error: unknown) => {
  console.error("MarkOut host smoke failed.", error);
  process.exitCode = 1;
});

async function runHostSmoke(): Promise<void> {
  const config = readHostSmokeConfig();
  await mkdir(config.outputDirectory, { recursive: true });

  const browser = await chromium.launch({
    executablePath: config.browserExecutable,
    headless: config.headless,
  });

  try {
    const context = await browser.newContext({
      storageState: config.storageStatePath,
    });
    const page = await context.newPage();

    await page.goto(config.composeUrl, {
      timeout: config.timeoutMs,
      waitUntil: "domcontentloaded",
    });

    await openTaskpane(page, config);
    let taskpane = await waitForTaskpane(page, config);
    await ensureAutoRenderEnabled(taskpane, config);

    await page.goto(config.composeUrl, {
      timeout: config.timeoutMs,
      waitUntil: "domcontentloaded",
    });

    await openTaskpane(page, config);
    taskpane = await waitForTaskpane(page, config);
    await assertAutoRenderEnabled(taskpane, config);

    await page.locator(config.toFieldSelector).first().fill(config.recipient);
    await page
      .locator(config.messageBodySelector)
      .first()
      .fill("# Smoke Heading\n\nParagraph text");

    await taskpane
      .getByRole("button", {
        name: new RegExp(`^${escapeForRegex(config.renderButtonText)}$`, "i"),
      })
      .click();

    await waitFor(async () => {
      const bodyText =
        (
          await page.locator(config.messageBodySelector).first().textContent()
        )?.replace(/\s+/g, " ") ?? "";

      return (
        bodyText.includes("Smoke Heading") &&
        !bodyText.includes("# Smoke Heading")
      );
    }, config.timeoutMs);

    await getSendButton(page, config).click();
    await waitFor(async () => {
      return await page
        .getByText(config.sentConfirmationText, { exact: false })
        .first()
        .isVisible()
        .catch(() => false);
    }, config.timeoutMs);

    console.log("MarkOut host smoke passed.");
  } catch (error) {
    const context = browser.contexts().at(0);
    const page = context?.pages().at(0);

    if (page !== undefined) {
      await page.screenshot({
        path: path.join(
          config.outputDirectory,
          "markout-host-smoke-failure.png"
        ),
        fullPage: true,
      });
    }

    throw error;
  } finally {
    await browser.close();
  }
}

async function assertAutoRenderEnabled(
  taskpane: FrameLocator,
  config: HostSmokeConfig
): Promise<void> {
  const autoRenderButton = taskpane
    .getByRole("button", {
      name: new RegExp(
        `^${escapeForRegex(config.autoRenderButtonPrefix)}`,
        "i"
      ),
    })
    .first();

  await waitForLocatorText(autoRenderButton, /On$/i, config.timeoutMs);
}

function escapeForRegex(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

async function ensureAutoRenderEnabled(
  taskpane: FrameLocator,
  config: HostSmokeConfig
): Promise<void> {
  const autoRenderButton = taskpane
    .getByRole("button", {
      name: new RegExp(
        `^${escapeForRegex(config.autoRenderButtonPrefix)}`,
        "i"
      ),
    })
    .first();
  const currentLabel = (await autoRenderButton.textContent()) ?? "";

  if (!/On$/i.test(currentLabel)) {
    await autoRenderButton.click();
  }

  await waitForLocatorText(autoRenderButton, /On$/i, config.timeoutMs);
}

function findBrowserExecutable(): string {
  const configuredExecutable =
    process.env.MARKOUT_HOST_SMOKE_BROWSER_EXECUTABLE;

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
    "Set MARKOUT_HOST_SMOKE_BROWSER_EXECUTABLE to a Chrome or Chromium binary."
  );
}

function getSendButton(page: Page, config: HostSmokeConfig): Locator {
  if (config.sendButtonSelector !== null) {
    return page.locator(config.sendButtonSelector).first();
  }

  return page.getByRole("button", {
    name: new RegExp(`^${escapeForRegex(config.sendButtonText)}$`, "i"),
  });
}

async function openTaskpane(
  page: Page,
  config: HostSmokeConfig
): Promise<void> {
  if (config.openButtonSelector !== null) {
    await page.locator(config.openButtonSelector).first().click();
    return;
  }

  await page
    .getByRole("button", {
      name: new RegExp(`^${escapeForRegex(config.openButtonText)}$`, "i"),
    })
    .click();
}

function readBooleanEnv(name: string, fallbackValue: boolean): boolean {
  const rawValue = process.env[name];

  if (rawValue === undefined) {
    return fallbackValue;
  }

  return !["0", "false", "no"].includes(rawValue.toLowerCase());
}

function readHostSmokeConfig(): HostSmokeConfig {
  return {
    autoRenderButtonPrefix:
      process.env.MARKOUT_HOST_SMOKE_AUTORENDER_BUTTON_PREFIX ??
      "Auto-render on send",
    browserExecutable: findBrowserExecutable(),
    composeUrl:
      process.env.MARKOUT_HOST_SMOKE_COMPOSE_URL ?? DEFAULT_COMPOSE_URL,
    headless: readBooleanEnv("MARKOUT_HOST_SMOKE_HEADLESS", true),
    messageBodySelector:
      process.env.MARKOUT_HOST_SMOKE_MESSAGE_BODY_SELECTOR ??
      '[aria-label="Message body"], div[contenteditable="true"][role="textbox"]',
    openButtonSelector:
      process.env.MARKOUT_HOST_SMOKE_OPEN_BUTTON_SELECTOR ?? null,
    openButtonText:
      process.env.MARKOUT_HOST_SMOKE_OPEN_BUTTON_TEXT ?? "Open MarkOut",
    outputDirectory:
      process.env.MARKOUT_HOST_SMOKE_OUTPUT_DIRECTORY ??
      DEFAULT_OUTPUT_DIRECTORY,
    previewSelector:
      process.env.MARKOUT_HOST_SMOKE_PREVIEW_SELECTOR ?? "#mo-preview",
    recipient: readRequiredEnv("MARKOUT_HOST_SMOKE_RECIPIENT"),
    renderButtonText:
      process.env.MARKOUT_HOST_SMOKE_RENDER_BUTTON_TEXT ??
      "Render current draft",
    sendButtonSelector:
      process.env.MARKOUT_HOST_SMOKE_SEND_BUTTON_SELECTOR ?? null,
    sendButtonText: process.env.MARKOUT_HOST_SMOKE_SEND_BUTTON_TEXT ?? "Send",
    sentConfirmationText:
      process.env.MARKOUT_HOST_SMOKE_SENT_CONFIRMATION_TEXT ?? "Sent",
    storageStatePath: readRequiredEnv("MARKOUT_HOST_SMOKE_STORAGE_STATE"),
    taskpaneFrameSelector:
      process.env.MARKOUT_HOST_SMOKE_TASKPANE_FRAME_SELECTOR ??
      DEFAULT_TASKPANE_FRAME_SELECTOR,
    timeoutMs: Number(
      process.env.MARKOUT_HOST_SMOKE_TIMEOUT_MS ?? DEFAULT_TIMEOUT_MS
    ),
    toFieldSelector:
      process.env.MARKOUT_HOST_SMOKE_TO_FIELD_SELECTOR ??
      'input[aria-label*="To"]',
  };
}

function readRequiredEnv(name: string): string {
  const value = process.env[name];

  if (value === undefined || value.trim().length === 0) {
    throw new Error(`Missing required environment variable: ${name}`);
  }

  return value;
}

async function waitFor(
  predicate: () => Promise<boolean>,
  timeoutMs: number
): Promise<void> {
  const deadline = Date.now() + timeoutMs;

  while (Date.now() < deadline) {
    if (await predicate()) {
      return;
    }

    await new Promise((resolve) => {
      setTimeout(resolve, 200);
    });
  }

  throw new Error(
    "Timed out while waiting for the Outlook host smoke assertion."
  );
}

async function waitForLocatorText(
  locator: Locator,
  expectedText: RegExp,
  timeoutMs: number
): Promise<void> {
  await waitFor(async () => {
    const text = (await locator.textContent()) ?? "";
    return expectedText.test(text);
  }, timeoutMs);
}

async function waitForTaskpane(
  page: Page,
  config: HostSmokeConfig
): Promise<FrameLocator> {
  await page
    .locator(config.taskpaneFrameSelector)
    .first()
    .waitFor({ state: "visible", timeout: config.timeoutMs });

  const taskpane = page.frameLocator(config.taskpaneFrameSelector).first();
  await taskpane
    .locator(config.previewSelector)
    .waitFor({ state: "visible", timeout: config.timeoutMs });
  return taskpane;
}
