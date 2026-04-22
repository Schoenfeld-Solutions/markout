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
  autoRenderSwitchSelector: string;
  browserExecutable: string;
  composeUrl: string;
  expectedTaskpaneUrlPrefix: string | null;
  headless: boolean;
  insertPanelButtonSelector: string;
  introConfirmButtonSelector: string;
  introPanelButtonSelector: string;
  messageBodySelector: string;
  openButtonSelector: string | null;
  openButtonText: string;
  outputDirectory: string;
  previewSelector: string;
  recipient: string;
  renderButtonSelector: string;
  sendButtonSelector: string | null;
  sendButtonText: string;
  settingsPanelButtonSelector: string;
  sentConfirmationText: string;
  storageStatePath: string;
  taskpaneFrameSelector: string;
  taskpaneReadySelector: string;
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
    await dismissIntroIfVisible(taskpane, config);
    await ensureAutoRenderEnabled(taskpane, config);

    await page.goto(config.composeUrl, {
      timeout: config.timeoutMs,
      waitUntil: "domcontentloaded",
    });

    await openTaskpane(page, config);
    taskpane = await waitForTaskpane(page, config);
    await assertIntroDismissed(taskpane, config);
    await assertAutoRenderEnabled(taskpane, config);
    await openInsertPanel(taskpane, config);

    await page.locator(config.toFieldSelector).first().fill(config.recipient);
    await page
      .locator(config.messageBodySelector)
      .first()
      .fill("# Smoke Heading\n\nParagraph text");

    await taskpane.locator(config.renderButtonSelector).first().click();

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
  await openSettingsPanel(taskpane, config);
  const autoRenderSwitch = taskpane
    .locator(config.autoRenderSwitchSelector)
    .first();

  await waitFor(async () => {
    return await isSwitchChecked(autoRenderSwitch);
  }, config.timeoutMs);
}

function escapeForRegex(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

async function ensureAutoRenderEnabled(
  taskpane: FrameLocator,
  config: HostSmokeConfig
): Promise<void> {
  await openSettingsPanel(taskpane, config);
  const autoRenderSwitch = taskpane
    .locator(config.autoRenderSwitchSelector)
    .first();

  if (!(await isSwitchChecked(autoRenderSwitch))) {
    await autoRenderSwitch.click();
  }

  await waitFor(async () => {
    return await isSwitchChecked(autoRenderSwitch);
  }, config.timeoutMs);
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

  if (rawValue === undefined || rawValue.trim().length === 0) {
    return fallbackValue;
  }

  return !["0", "false", "no"].includes(rawValue.toLowerCase());
}

function readNumberEnv(name: string, fallbackValue: number): number {
  const rawValue = process.env[name];

  if (rawValue === undefined || rawValue.trim().length === 0) {
    return fallbackValue;
  }

  const parsedValue = Number(rawValue);

  if (!Number.isFinite(parsedValue)) {
    throw new Error(`Expected ${name} to be a finite number.`);
  }

  return parsedValue;
}

function readOptionalEnv(name: string): string | null {
  const rawValue = process.env[name];

  if (rawValue === undefined || rawValue.trim().length === 0) {
    return null;
  }

  return rawValue;
}

function readStringEnv(name: string, fallbackValue: string): string {
  const rawValue = process.env[name];

  if (rawValue === undefined || rawValue.trim().length === 0) {
    return fallbackValue;
  }

  return rawValue;
}

function readHostSmokeConfig(): HostSmokeConfig {
  return {
    autoRenderSwitchSelector: readStringEnv(
      "MARKOUT_HOST_SMOKE_AUTORENDER_SWITCH_SELECTOR",
      "#autorender-switch"
    ),
    browserExecutable: findBrowserExecutable(),
    composeUrl: readStringEnv(
      "MARKOUT_HOST_SMOKE_COMPOSE_URL",
      DEFAULT_COMPOSE_URL
    ),
    expectedTaskpaneUrlPrefix: readOptionalEnv(
      "MARKOUT_HOST_SMOKE_EXPECTED_TASKPANE_URL_PREFIX"
    ),
    headless: readBooleanEnv("MARKOUT_HOST_SMOKE_HEADLESS", true),
    insertPanelButtonSelector: readStringEnv(
      "MARKOUT_HOST_SMOKE_INSERT_PANEL_BUTTON_SELECTOR",
      "#panel-button-insert"
    ),
    introConfirmButtonSelector: readStringEnv(
      "MARKOUT_HOST_SMOKE_INTRO_CONFIRM_BUTTON_SELECTOR",
      "#intro-confirm-button"
    ),
    introPanelButtonSelector: readStringEnv(
      "MARKOUT_HOST_SMOKE_INTRO_PANEL_BUTTON_SELECTOR",
      "#panel-button-intro"
    ),
    messageBodySelector: readStringEnv(
      "MARKOUT_HOST_SMOKE_MESSAGE_BODY_SELECTOR",
      '[aria-label="Message body"], div[contenteditable="true"][role="textbox"]'
    ),
    openButtonSelector: readOptionalEnv(
      "MARKOUT_HOST_SMOKE_OPEN_BUTTON_SELECTOR"
    ),
    openButtonText: readStringEnv(
      "MARKOUT_HOST_SMOKE_OPEN_BUTTON_TEXT",
      "Open MarkOut"
    ),
    outputDirectory: readStringEnv(
      "MARKOUT_HOST_SMOKE_OUTPUT_DIRECTORY",
      DEFAULT_OUTPUT_DIRECTORY
    ),
    previewSelector: readStringEnv(
      "MARKOUT_HOST_SMOKE_PREVIEW_SELECTOR",
      "#mo-preview"
    ),
    recipient: readRequiredEnv("MARKOUT_HOST_SMOKE_RECIPIENT"),
    renderButtonSelector: readStringEnv(
      "MARKOUT_HOST_SMOKE_RENDER_BUTTON_SELECTOR",
      "#render-entire-draft-button"
    ),
    sendButtonSelector: readOptionalEnv(
      "MARKOUT_HOST_SMOKE_SEND_BUTTON_SELECTOR"
    ),
    sendButtonText: readStringEnv(
      "MARKOUT_HOST_SMOKE_SEND_BUTTON_TEXT",
      "Send"
    ),
    settingsPanelButtonSelector: readStringEnv(
      "MARKOUT_HOST_SMOKE_SETTINGS_PANEL_BUTTON_SELECTOR",
      "#panel-button-settings"
    ),
    sentConfirmationText: readStringEnv(
      "MARKOUT_HOST_SMOKE_SENT_CONFIRMATION_TEXT",
      "Sent"
    ),
    storageStatePath: readRequiredEnv("MARKOUT_HOST_SMOKE_STORAGE_STATE"),
    taskpaneFrameSelector: readStringEnv(
      "MARKOUT_HOST_SMOKE_TASKPANE_FRAME_SELECTOR",
      DEFAULT_TASKPANE_FRAME_SELECTOR
    ),
    taskpaneReadySelector: readStringEnv(
      "MARKOUT_HOST_SMOKE_TASKPANE_READY_SELECTOR",
      "#taskpane-shell"
    ),
    timeoutMs: readNumberEnv(
      "MARKOUT_HOST_SMOKE_TIMEOUT_MS",
      DEFAULT_TIMEOUT_MS
    ),
    toFieldSelector: readStringEnv(
      "MARKOUT_HOST_SMOKE_TO_FIELD_SELECTOR",
      'input[aria-label*="To"]'
    ),
  };
}

function readRequiredEnv(name: string): string {
  const value = process.env[name];

  if (value === undefined || value.trim().length === 0) {
    throw new Error(`Missing required environment variable: ${name}`);
  }

  return value;
}

async function isSwitchChecked(locator: Locator): Promise<boolean> {
  const checkedProperty = await locator
    .evaluate((element) => {
      if (
        element instanceof HTMLInputElement &&
        typeof element.checked === "boolean"
      ) {
        return element.checked;
      }

      return null;
    })
    .catch(() => null);

  if (typeof checkedProperty === "boolean") {
    return checkedProperty;
  }

  return (await locator.getAttribute("aria-checked")) === "true";
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

async function openInsertPanel(
  taskpane: FrameLocator,
  config: HostSmokeConfig
): Promise<void> {
  await taskpane.locator(config.insertPanelButtonSelector).first().click();
  await taskpane
    .locator(config.previewSelector)
    .first()
    .waitFor({ state: "visible", timeout: config.timeoutMs });
}

async function openSettingsPanel(
  taskpane: FrameLocator,
  config: HostSmokeConfig
): Promise<void> {
  await taskpane.locator(config.settingsPanelButtonSelector).first().click();
  await taskpane
    .locator(config.autoRenderSwitchSelector)
    .first()
    .waitFor({ state: "visible", timeout: config.timeoutMs });
}

async function waitForTaskpane(
  page: Page,
  config: HostSmokeConfig
): Promise<FrameLocator> {
  const taskpaneFrame = page.locator(config.taskpaneFrameSelector).first();
  await taskpaneFrame.waitFor({ state: "visible", timeout: config.timeoutMs });
  await assertTaskpaneFrameSource(taskpaneFrame, config);

  const taskpane = page.frameLocator(config.taskpaneFrameSelector).first();
  await taskpane
    .locator(config.taskpaneReadySelector)
    .waitFor({ state: "visible", timeout: config.timeoutMs });
  return taskpane;
}

async function assertTaskpaneFrameSource(
  taskpaneFrame: Locator,
  config: HostSmokeConfig
): Promise<void> {
  const expectedTaskpaneUrlPrefix = config.expectedTaskpaneUrlPrefix;

  if (expectedTaskpaneUrlPrefix === null) {
    return;
  }

  await waitFor(async () => {
    const source = await taskpaneFrame.getAttribute("src");
    return source?.startsWith(expectedTaskpaneUrlPrefix) ?? false;
  }, config.timeoutMs);
}

async function dismissIntroIfVisible(
  taskpane: FrameLocator,
  config: HostSmokeConfig
): Promise<void> {
  const introConfirmButton = taskpane
    .locator(config.introConfirmButtonSelector)
    .first();

  if (!(await introConfirmButton.isVisible().catch(() => false))) {
    return;
  }

  await introConfirmButton.click();
  await introConfirmButton.waitFor({
    state: "hidden",
    timeout: config.timeoutMs,
  });
}

async function assertIntroDismissed(
  taskpane: FrameLocator,
  config: HostSmokeConfig
): Promise<void> {
  await waitFor(async () => {
    return (
      (await taskpane.locator(config.introPanelButtonSelector).count()) === 0
    );
  }, config.timeoutMs);
}
