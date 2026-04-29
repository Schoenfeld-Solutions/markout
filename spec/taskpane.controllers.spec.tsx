/** @jest-environment jsdom */

import { act } from "react";
import type { ReactElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import type { ComposeMarkdownService } from "../src/lib/compose-markdown";
import type { ComposeNotificationService } from "../src/lib/compose-notifications";
import type { DiagnosticEventInput } from "../src/lib/runtime";
import {
  useAutoRenderNotificationController,
  usePreviewController,
  useSelectionStateController,
} from "../src/taskpane/controllers";
import type { PanelKey } from "../src/taskpane/types";

(
  globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true;

function deferred<T>(): {
  promise: Promise<T>;
  reject: (error: unknown) => void;
  resolve: (value: T) => void;
} {
  let resolvePromise: (value: T) => void = () => undefined;
  let rejectPromise: (error: unknown) => void = () => undefined;
  const promise = new Promise<T>((resolve, reject) => {
    resolvePromise = resolve;
    rejectPromise = reject;
  });

  return {
    promise,
    reject: rejectPromise,
    resolve: resolvePromise,
  };
}

function createComposeMarkdownService(
  overrides: Partial<ComposeMarkdownService> = {}
): ComposeMarkdownService {
  return {
    getSelection: () =>
      Promise.resolve({
        hasSelection: false,
        html: null,
        source: "body",
        text: "",
      }),
    insertRenderedMarkdown: () => Promise.resolve("inserted"),
    renderPreview: () => Promise.resolve("<p>Preview</p>"),
    renderSelection: () => Promise.resolve(),
    ...overrides,
  };
}

function createNotificationService(
  overrides: Partial<ComposeNotificationService> = {}
): ComposeNotificationService {
  return {
    clearAutoRenderDismissed: () => Promise.resolve(),
    clearAutoRenderNotification: () => Promise.resolve(),
    clearTransientNotification: () => Promise.resolve(),
    hasAutoRenderBeenDismissed: () => Promise.resolve(false),
    markAutoRenderDismissed: () => Promise.resolve(),
    onAutoRenderDismiss: () => undefined,
    showAutoRenderNotification: () => Promise.resolve("outlook"),
    showTransientNotification: () => Promise.resolve("outlook"),
    ...overrides,
  };
}

function mount(element: ReactElement): {
  container: HTMLElement;
  root: Root;
} {
  document.body.innerHTML = '<div id="root"></div>';
  const container = document.getElementById("root");

  if (container === null) {
    throw new Error("Expected a controller test container.");
  }

  const root = createRoot(container);
  act(() => {
    root.render(element);
  });

  return { container, root };
}

async function flushPromises(): Promise<void> {
  await act(async () => {
    await Promise.resolve();
    await Promise.resolve();
  });
}

describe("taskpane preview controller", () => {
  afterEach(() => {
    jest.restoreAllMocks();
  });

  it("keeps empty markdown in the empty state without rendering", async () => {
    const renderPreview = jest.fn(() => Promise.resolve("<p>unused</p>"));
    const service = createComposeMarkdownService({ renderPreview });

    function PreviewProbe(): ReactElement {
      const { previewHtml, previewState } = usePreviewController(
        service,
        "   ",
        "",
        "Preview failed.",
        () => undefined
      );

      return <div data-state={previewState}>{previewHtml}</div>;
    }

    const { container, root } = mount(<PreviewProbe />);
    try {
      await flushPromises();

      expect(renderPreview).not.toHaveBeenCalled();
      expect(container.firstElementChild?.getAttribute("data-state")).toBe(
        "empty"
      );
      expect(container.textContent).toBe("");
    } finally {
      act(() => {
        root.unmount();
      });
    }
  });

  it("renders latest markdown, records diagnostics, and ignores stale results", async () => {
    const renders: {
      markdown: string;
      resolve: (html: string) => void;
      stylesheet: string | undefined;
    }[] = [];
    const events: string[] = [];
    const service = createComposeMarkdownService({
      renderPreview: (markdown, stylesheet) =>
        new Promise<string>((resolve) => {
          renders.push({ markdown, resolve, stylesheet });
        }),
    });

    function PreviewProbe({
      markdown,
      stylesheet,
    }: {
      markdown: string;
      stylesheet: string;
    }): ReactElement {
      const { previewHtml, previewState } = usePreviewController(
        service,
        markdown,
        stylesheet,
        "Preview failed.",
        () => undefined,
        (event) => {
          events.push(event.code);
        }
      );

      return (
        <div data-state={previewState} id="preview-probe">
          {previewHtml}
        </div>
      );
    }

    const { container, root } = mount(
      <PreviewProbe markdown="# First" stylesheet=".first{}" />
    );
    try {
      await flushPromises();
      act(() => {
        root.render(
          <PreviewProbe markdown="# Second" stylesheet=".second{}" />
        );
      });
      await flushPromises();

      expect(renders).toHaveLength(2);
      expect(renders.map((render) => render.markdown)).toEqual([
        "# First",
        "# Second",
      ]);
      expect(renders.map((render) => render.stylesheet)).toEqual([
        ".first{}",
        ".second{}",
      ]);

      await act(async () => {
        renders[1]?.resolve("<p>Second</p>");
        await Promise.resolve();
      });
      expect(container.textContent).toContain("Second");
      expect(
        container.querySelector("#preview-probe")?.getAttribute("data-state")
      ).toBe("ready");

      await act(async () => {
        renders[0]?.resolve("<p>First</p>");
        await Promise.resolve();
      });
      expect(container.textContent).toContain("Second");
      expect(container.textContent).not.toContain("First");
      expect(events).toEqual([
        "preview.render.started",
        "preview.render.started",
        "preview.render.succeeded",
      ]);
    } finally {
      act(() => {
        root.unmount();
      });
    }
  });

  it("reports preview failures without leaking stale ready html", async () => {
    const consoleError = jest
      .spyOn(console, "error")
      .mockImplementation(() => undefined);
    const errors: string[] = [];
    const events: string[] = [];
    const service = createComposeMarkdownService({
      renderPreview: () => Promise.reject(new TypeError("private details")),
    });

    function PreviewProbe(): ReactElement {
      const { previewHtml, previewState } = usePreviewController(
        service,
        "# Broken",
        "",
        "Preview failed.",
        (message) => {
          errors.push(message);
        },
        (event) => {
          events.push(event.code);
        }
      );

      return <div data-state={previewState}>{previewHtml}</div>;
    }

    const { container, root } = mount(<PreviewProbe />);
    try {
      await flushPromises();

      expect(container.firstElementChild?.getAttribute("data-state")).toBe(
        "empty"
      );
      expect(container.textContent).toBe("");
      expect(errors).toEqual(["Preview failed."]);
      expect(events).toEqual([
        "preview.render.started",
        "preview.render.failed",
      ]);
      expect(consoleError).toHaveBeenCalled();
    } finally {
      act(() => {
        root.unmount();
      });
    }
  });
});

describe("taskpane selection state controller", () => {
  afterEach(() => {
    jest.restoreAllMocks();
    Object.defineProperty(document, "visibilityState", {
      configurable: true,
      value: "visible",
    });
  });

  it("polls selection only on the insert panel and skips hidden refreshes", () => {
    const getSelection = jest.fn(
      () =>
        new Promise<{
          hasSelection: boolean;
          html: null;
          source: "body";
          text: string;
        }>(() => undefined)
    );
    const service = createComposeMarkdownService({ getSelection });

    function SelectionProbe({
      activePanel,
    }: {
      activePanel: PanelKey;
    }): ReactElement {
      useSelectionStateController(service, activePanel);

      return <div />;
    }

    const { root } = mount(<SelectionProbe activePanel="settings" />);
    try {
      expect(getSelection).not.toHaveBeenCalled();

      act(() => {
        root.render(<SelectionProbe activePanel="insert" />);
      });
      expect(getSelection).toHaveBeenCalledTimes(1);

      act(() => {
        window.dispatchEvent(new Event("focus"));
      });
      expect(getSelection).toHaveBeenCalledTimes(2);

      Object.defineProperty(document, "visibilityState", {
        configurable: true,
        value: "hidden",
      });
      act(() => {
        document.dispatchEvent(new Event("visibilitychange"));
      });
      expect(getSelection).toHaveBeenCalledTimes(2);
    } finally {
      act(() => {
        root.unmount();
      });
    }
  });

  it("maps body and subject selection states through explicit refreshes", async () => {
    const selections = [
      {
        hasSelection: true,
        html: "<p>Selected</p>",
        source: "body" as const,
        text: "Selected text",
      },
      {
        hasSelection: false,
        html: null,
        source: "body" as const,
        text: "",
      },
      {
        hasSelection: true,
        html: null,
        source: "subject" as const,
        text: "Subject",
      },
    ];
    const getSelection = jest.fn(() =>
      Promise.resolve(
        selections.shift() ?? {
          hasSelection: false,
          html: null,
          source: "body" as const,
          text: "",
        }
      )
    );
    const events: DiagnosticEventInput[] = [];
    const service = createComposeMarkdownService({ getSelection });
    let refreshSelection: (() => Promise<boolean>) | null = null;

    function SelectionProbe(): ReactElement {
      const { selectionState, updateSelectionState } =
        useSelectionStateController(service, "settings", (event) => {
          events.push(event);
        });
      refreshSelection = updateSelectionState;

      return (
        <div data-state={selectionState.availability} id="selection-probe">
          {selectionState.debug?.textPreview ?? ""}
        </div>
      );
    }

    const { container, root } = mount(<SelectionProbe />);
    try {
      await act(async () => {
        await refreshSelection?.();
      });
      expect(
        container.querySelector("#selection-probe")?.getAttribute("data-state")
      ).toBe("body-selection");
      expect(container.textContent).toContain("Selected text");

      await act(async () => {
        await refreshSelection?.();
      });
      expect(
        container.querySelector("#selection-probe")?.getAttribute("data-state")
      ).toBe("body-none");

      await act(async () => {
        await refreshSelection?.();
      });
      expect(
        container.querySelector("#selection-probe")?.getAttribute("data-state")
      ).toBe("subject");
      expect(events.map((event) => event.code)).toEqual([
        "selection.read.succeeded",
        "selection.read.succeeded",
        "selection.read.succeeded",
      ]);
    } finally {
      act(() => {
        root.unmount();
      });
    }
  });

  it("preserves debug data after read failures", async () => {
    const getSelection = jest
      .fn()
      .mockResolvedValueOnce({
        hasSelection: true,
        html: null,
        source: "body",
        text: "Previous debug",
      })
      .mockRejectedValueOnce(new Error("selection read failed"));
    const events: string[] = [];
    const service = createComposeMarkdownService({ getSelection });
    let refreshSelection: (() => Promise<boolean>) | null = null;

    function SelectionProbe(): ReactElement {
      const { selectionState, updateSelectionState } =
        useSelectionStateController(service, "settings", (event) => {
          events.push(event.code);
        });
      refreshSelection = updateSelectionState;

      return (
        <div data-state={selectionState.availability} id="selection-probe">
          {selectionState.debug?.textPreview ?? ""}
        </div>
      );
    }

    const { container, root } = mount(<SelectionProbe />);
    try {
      await act(async () => {
        await refreshSelection?.();
      });
      expect(getSelection).toHaveBeenCalledTimes(1);

      await act(async () => {
        await refreshSelection?.();
      });
      expect(
        container.querySelector("#selection-probe")?.getAttribute("data-state")
      ).toBe("unknown");
      expect(container.textContent).toContain("Previous debug");
      expect(events).toEqual([
        "selection.read.succeeded",
        "selection.read.failed",
      ]);
    } finally {
      act(() => {
        root.unmount();
      });
    }
  });
});

describe("taskpane auto-render notification controller", () => {
  afterEach(() => {
    jest.restoreAllMocks();
  });

  it("does nothing when no notification service is available", async () => {
    function NotificationProbe(): ReactElement {
      const { dismissAutoRenderFallbackNotice, showAutoRenderFallbackNotice } =
        useAutoRenderNotificationController(undefined, true, "Auto render");

      return (
        <button
          id="notice"
          onClick={() => {
            void dismissAutoRenderFallbackNotice();
          }}
        >
          {showAutoRenderFallbackNotice ? "visible" : "hidden"}
        </button>
      );
    }

    const { container, root } = mount(<NotificationProbe />);
    try {
      await flushPromises();
      expect(container.textContent).toBe("hidden");

      await act(async () => {
        container.querySelector<HTMLButtonElement>("#notice")?.click();
        await Promise.resolve();
      });
      expect(container.textContent).toBe("hidden");
    } finally {
      act(() => {
        root.unmount();
      });
    }
  });

  it("shows, dismisses, and restores auto-render notification states", async () => {
    let dismissHandler: (() => void) | null = null;
    const markAutoRenderDismissed = jest.fn(() => Promise.resolve());
    const events: string[] = [];
    const service = createNotificationService({
      markAutoRenderDismissed,
      onAutoRenderDismiss: (handler) => {
        dismissHandler = handler;
      },
      showAutoRenderNotification: () => Promise.resolve("pane"),
    });

    function NotificationProbe(): ReactElement {
      const { dismissAutoRenderFallbackNotice, showAutoRenderFallbackNotice } =
        useAutoRenderNotificationController(
          service,
          true,
          "Auto render",
          (event) => {
            events.push(event.code);
          }
        );

      return (
        <button
          id="notice"
          onClick={() => {
            void dismissAutoRenderFallbackNotice();
          }}
        >
          {showAutoRenderFallbackNotice ? "visible" : "hidden"}
        </button>
      );
    }

    const { container, root } = mount(<NotificationProbe />);
    try {
      await flushPromises();
      expect(container.textContent).toBe("visible");
      expect(events).toContain("notification.autorender.fallback-pane");

      await act(async () => {
        container.querySelector<HTMLButtonElement>("#notice")?.click();
        await Promise.resolve();
      });
      expect(markAutoRenderDismissed).toHaveBeenCalled();
      expect(container.textContent).toBe("hidden");

      act(() => {
        dismissHandler?.();
      });
      expect(events).toContain("notification.autorender.dismissed");
    } finally {
      act(() => {
        root.unmount();
      });
    }
  });

  it("clears notification state when disabled and records restored dismissals", async () => {
    const clearAutoRenderDismissed = jest.fn(() => Promise.resolve());
    const clearAutoRenderNotification = jest.fn(() => Promise.resolve());
    const events: string[] = [];
    const service = createNotificationService({
      clearAutoRenderDismissed,
      clearAutoRenderNotification,
      hasAutoRenderBeenDismissed: () => Promise.resolve(true),
    });

    function NotificationProbe({
      enabled,
    }: {
      enabled: boolean;
    }): ReactElement {
      const { showAutoRenderFallbackNotice } =
        useAutoRenderNotificationController(
          service,
          enabled,
          "Auto render",
          (event) => {
            events.push(event.code);
          }
        );

      return <div>{showAutoRenderFallbackNotice ? "visible" : "hidden"}</div>;
    }

    const { container, root } = mount(<NotificationProbe enabled={false} />);
    try {
      await flushPromises();
      expect(clearAutoRenderNotification).toHaveBeenCalledTimes(1);
      expect(clearAutoRenderDismissed).toHaveBeenCalledTimes(1);
      expect(events).toContain("notification.autorender.disabled");

      act(() => {
        root.render(<NotificationProbe enabled={true} />);
      });
      await flushPromises();
      expect(clearAutoRenderDismissed).toHaveBeenCalledTimes(2);
      expect(container.textContent).toBe("hidden");
      expect(events).toContain("notification.autorender.dismissal-restored");
    } finally {
      act(() => {
        root.unmount();
      });
    }
  });

  it("ignores stale notification results after disable and records service failures", async () => {
    const showNotice = deferred<"outlook" | "pane">();
    const events: string[] = [];
    const service = createNotificationService({
      showAutoRenderNotification: () => showNotice.promise,
    });

    function NotificationProbe({
      enabled,
      notificationService,
    }: {
      enabled: boolean;
      notificationService: ComposeNotificationService;
    }): ReactElement {
      const { showAutoRenderFallbackNotice } =
        useAutoRenderNotificationController(
          notificationService,
          enabled,
          "Auto render",
          (event) => {
            events.push(event.code);
          }
        );

      return <div>{showAutoRenderFallbackNotice ? "visible" : "hidden"}</div>;
    }

    const { container, root } = mount(
      <NotificationProbe enabled={true} notificationService={service} />
    );
    try {
      await flushPromises();
      act(() => {
        root.render(
          <NotificationProbe enabled={false} notificationService={service} />
        );
      });
      await flushPromises();

      await act(async () => {
        showNotice.resolve("pane");
        await Promise.resolve();
      });
      expect(container.textContent).toBe("hidden");
      expect(events).toContain("notification.autorender.disabled");
      expect(events).not.toContain("notification.autorender.fallback-pane");

      const failingService = createNotificationService({
        hasAutoRenderBeenDismissed: () => Promise.reject(new Error("failed")),
      });
      act(() => {
        root.render(
          <NotificationProbe
            enabled={true}
            notificationService={failingService}
          />
        );
      });
      await flushPromises();
      expect(container.textContent).toBe("hidden");
      expect(events).toContain("notification.autorender.failed");
    } finally {
      act(() => {
        root.unmount();
      });
    }
  });
});
