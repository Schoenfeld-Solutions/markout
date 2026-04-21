// @ts-expect-error Webpack resolves taskpane CSS side-effect imports at build time.
import "./taskpane.css";

interface BootStateProps {
  copy: string;
  details?: string;
  stage: string;
  title: string;
}

const taskpaneRootElement = document.getElementById("taskpane-root");

if (taskpaneRootElement === null) {
  throw new Error('MarkOut could not find the "taskpane-root" element.');
}

const rootElement: HTMLElement = taskpaneRootElement;

function appendTextElement(
  parent: HTMLElement,
  tagName: keyof HTMLElementTagNameMap,
  className: string,
  text: string
): void {
  const element = document.createElement(tagName);
  element.className = className;
  element.textContent = text;
  parent.appendChild(element);
}

function renderBootState({
  copy,
  details,
  stage,
  title,
}: BootStateProps): void {
  rootElement.replaceChildren();

  const shell = document.createElement("section");
  shell.className = "boot-shell";
  const card = document.createElement("div");
  card.className = "boot-card";

  appendTextElement(card, "p", "boot-eyebrow", "MarkOut");
  appendTextElement(card, "h1", "boot-title", title);
  appendTextElement(card, "p", "boot-copy", copy);

  const stageElement = document.createElement("p");
  stageElement.className = "boot-stage";
  stageElement.setAttribute("role", "status");
  stageElement.textContent = stage;
  card.appendChild(stageElement);

  if (details !== undefined && details.length > 0) {
    appendTextElement(card, "pre", "boot-details", details);
  }

  shell.appendChild(card);
  rootElement.appendChild(shell);
}

function formatError(error: unknown): string {
  if (error instanceof Error) {
    return `${error.name}: ${error.message}`;
  }

  return String(error);
}

function getInitializationFailureMessage(): string {
  if (window.location.hostname === "localhost") {
    return "Use Outlook's Add from File flow with manifest-localhost.xml to load the local development build.";
  }

  return "Reload Outlook and reopen MarkOut from Apps. If the pane still fails, reinstall the current manifest and try again.";
}

function reportBootFailure(error: unknown): void {
  renderBootState({
    copy: getInitializationFailureMessage(),
    details: formatError(error),
    stage: "Initialization failed.",
    title: "MarkOut could not initialize",
  });
}

window.addEventListener("error", (event) => {
  console.error(
    "MarkOut captured a window error.",
    event.error ?? event.message
  );
  reportBootFailure(event.error ?? event.message);
});

window.addEventListener("unhandledrejection", (event) => {
  console.error("MarkOut captured an unhandled rejection.", event.reason);
  reportBootFailure(event.reason);
});

async function bootTaskpane(): Promise<void> {
  console.info("[MarkOut] taskpane script loaded");
  renderBootState({
    copy: "Waiting for Office.onReady before the Outlook compose workspace is mounted.",
    stage: "Task pane script loaded.",
    title: "MarkOut is loading",
  });

  if (typeof Office === "undefined") {
    throw new Error("Office.js did not load before taskpane startup.");
  }

  const info = await Office.onReady();
  console.info("[MarkOut] Office.onReady", { host: info.host });

  if (info.host !== Office.HostType.Outlook) {
    renderBootState({
      copy: "This task pane only runs when MarkOut is opened inside Outlook compose.",
      stage: "Wrong host detected.",
      title: "Open MarkOut from Outlook",
    });
    return;
  }

  const { mountTaskpane } = await import(
    /* webpackChunkName: "taskpane-runtime" */ "./runtime"
  );

  mountTaskpane(rootElement);
}

void bootTaskpane().catch((error: unknown) => {
  console.error("MarkOut failed to initialize the task pane.", error);
  reportBootFailure(error);
});
