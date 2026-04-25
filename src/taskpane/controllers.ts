import {
  useDeferredValue,
  useEffect,
  useEffectEvent,
  useRef,
  useState,
} from "react";
import type { ComposeNotificationService } from "../lib/compose-notifications";
import type { ComposeMarkdownService } from "../lib/compose-markdown";
import {
  getErrorDiagnosticMetadata,
  type DiagnosticEventInput,
} from "../lib/runtime";
import type { PanelKey, SelectionState } from "./types";

const SELECTION_REFRESH_INTERVAL_MS = 1600;

export type TaskpaneDiagnosticRecorder = (event: DiagnosticEventInput) => void;

export function usePreviewController(
  composeMarkdown: ComposeMarkdownService,
  markdownInput: string,
  stylesheet: string,
  previewFailedMessage: string,
  onPreviewError: (message: string) => void,
  recordDiagnostic?: TaskpaneDiagnosticRecorder
): {
  previewHtml: string;
  previewState: "empty" | "loading" | "ready";
} {
  const [previewHtml, setPreviewHtml] = useState("");
  const [previewState, setPreviewState] = useState<
    "empty" | "loading" | "ready"
  >("empty");
  const handlePreviewError = useEffectEvent((message: string) => {
    onPreviewError(message);
  });
  const emitDiagnostic = useEffectEvent((event: DiagnosticEventInput) => {
    recordDiagnostic?.(event);
  });
  const deferredMarkdownInput = useDeferredValue(markdownInput);
  const deferredStylesheet = useDeferredValue(stylesheet);

  useEffect(() => {
    let ignore = false;

    if (deferredMarkdownInput.trim().length === 0) {
      setPreviewHtml("");
      setPreviewState("empty");
      return;
    }

    setPreviewState("loading");
    emitDiagnostic({
      area: "preview",
      code: "preview.render.started",
      level: "debug",
      metadata: {
        inputLength: deferredMarkdownInput.length,
        stylesheetLength: deferredStylesheet.length,
      },
    });
    void composeMarkdown
      .renderPreview(deferredMarkdownInput, deferredStylesheet)
      .then((html) => {
        if (ignore) {
          return;
        }

        setPreviewHtml(html);
        setPreviewState("ready");
        emitDiagnostic({
          area: "preview",
          code: "preview.render.succeeded",
          level: "info",
          metadata: {
            outputLength: html.length,
          },
        });
      })
      .catch((error: unknown) => {
        if (ignore) {
          return;
        }

        console.error("MarkOut failed to refresh the taskpane preview.", error);
        setPreviewHtml("");
        setPreviewState("empty");
        emitDiagnostic({
          area: "preview",
          code: "preview.render.failed",
          level: "error",
          metadata: getErrorDiagnosticMetadata(error),
        });
        handlePreviewError(previewFailedMessage);
      });

    return () => {
      ignore = true;
    };
  }, [
    composeMarkdown,
    deferredMarkdownInput,
    deferredStylesheet,
    previewFailedMessage,
  ]);

  return { previewHtml, previewState };
}

export function useSelectionStateController(
  composeMarkdown: ComposeMarkdownService,
  activePanel: PanelKey,
  recordDiagnostic?: TaskpaneDiagnosticRecorder
): {
  isInspectingSelection: boolean;
  selectionState: SelectionState;
  setIsInspectingSelection: (value: boolean) => void;
  updateSelectionState: () => Promise<boolean>;
} {
  const [selectionState, setSelectionState] = useState<SelectionState>({
    availability: "unknown",
    debug: null,
  });
  const [isInspectingSelection, setIsInspectingSelection] = useState(false);
  const emitDiagnostic = useEffectEvent((event: DiagnosticEventInput) => {
    recordDiagnostic?.(event);
  });

  const updateSelectionState = useEffectEvent(async (): Promise<boolean> => {
    try {
      const selection = await composeMarkdown.getSelection();

      setSelectionState({
        availability:
          selection.source === "subject"
            ? "subject"
            : selection.hasSelection
              ? "body-selection"
              : "body-none",
        debug: {
          hasSelection: selection.hasSelection,
          source: selection.source,
          textPreview: selection.text.slice(0, 200),
        },
      });
      emitDiagnostic({
        area: "selection",
        code: "selection.read.succeeded",
        level: "info",
        metadata: {
          hasSelection: selection.hasSelection,
          inputLength: selection.text.length,
          source: selection.source,
        },
      });
      return true;
    } catch (error) {
      setSelectionState((currentState) => ({
        availability: "unknown",
        debug: currentState.debug,
      }));
      emitDiagnostic({
        area: "selection",
        code: "selection.read.failed",
        level: "warning",
        metadata: getErrorDiagnosticMetadata(error),
      });
      return false;
    }
  });

  useEffect(() => {
    if (activePanel !== "insert") {
      return;
    }

    const refreshSelection = () => {
      if (document.visibilityState === "hidden") {
        return;
      }

      void updateSelectionState();
    };

    refreshSelection();
    window.addEventListener("focus", refreshSelection);
    document.addEventListener("visibilitychange", refreshSelection);
    const intervalId = window.setInterval(
      refreshSelection,
      SELECTION_REFRESH_INTERVAL_MS
    );

    return () => {
      window.clearInterval(intervalId);
      window.removeEventListener("focus", refreshSelection);
      document.removeEventListener("visibilitychange", refreshSelection);
    };
  }, [activePanel, updateSelectionState]);

  return {
    isInspectingSelection,
    selectionState,
    setIsInspectingSelection,
    updateSelectionState,
  };
}

export function useAutoRenderNotificationController(
  notificationService: ComposeNotificationService | undefined,
  autoRenderEnabled: boolean,
  stickyMessage: string,
  recordDiagnostic?: TaskpaneDiagnosticRecorder
): {
  showAutoRenderFallbackNotice: boolean;
  dismissAutoRenderFallbackNotice: () => Promise<void>;
} {
  const [showAutoRenderFallbackNotice, setShowAutoRenderFallbackNotice] =
    useState(false);
  const previousAutoRenderRef = useRef(autoRenderEnabled);
  const emitDiagnostic = useEffectEvent((event: DiagnosticEventInput) => {
    recordDiagnostic?.(event);
  });

  useEffect(() => {
    if (notificationService === undefined) {
      return;
    }

    notificationService.onAutoRenderDismiss(() => {
      setShowAutoRenderFallbackNotice(false);
      emitDiagnostic({
        area: "notification",
        code: "notification.autorender.dismissed",
        level: "info",
      });
    });
  }, [notificationService]);

  useEffect(() => {
    if (notificationService === undefined) {
      return;
    }

    let cancelled = false;
    const isCancelled = (): boolean => cancelled;
    const wasEnabled = previousAutoRenderRef.current;
    previousAutoRenderRef.current = autoRenderEnabled;

    void (async () => {
      if (!autoRenderEnabled) {
        await notificationService.clearAutoRenderNotification();
        await notificationService.clearAutoRenderDismissed();
        if (!isCancelled()) {
          setShowAutoRenderFallbackNotice(false);
          emitDiagnostic({
            area: "notification",
            code: "notification.autorender.disabled",
            level: "debug",
          });
        }
        return;
      }

      if (!wasEnabled) {
        await notificationService.clearAutoRenderDismissed();
      }

      const dismissed = await notificationService.hasAutoRenderBeenDismissed();
      if (isCancelled()) {
        return;
      }

      if (dismissed) {
        setShowAutoRenderFallbackNotice(false);
        emitDiagnostic({
          area: "notification",
          code: "notification.autorender.dismissal-restored",
          level: "debug",
        });
        return;
      }

      const surface = await notificationService.showAutoRenderNotification({
        message: stickyMessage,
      });

      if (!isCancelled()) {
        setShowAutoRenderFallbackNotice(surface === "pane");
        emitDiagnostic({
          area: "notification",
          code:
            surface === "pane"
              ? "notification.autorender.fallback-pane"
              : "notification.autorender.shown",
          level: surface === "pane" ? "warning" : "info",
        });
      }
    })().catch((error: unknown) => {
      if (isCancelled()) {
        return;
      }

      setShowAutoRenderFallbackNotice(false);
      emitDiagnostic({
        area: "notification",
        code: "notification.autorender.failed",
        level: "error",
        metadata: getErrorDiagnosticMetadata(error),
      });
    });

    return () => {
      cancelled = true;
    };
  }, [autoRenderEnabled, notificationService, stickyMessage]);

  return {
    dismissAutoRenderFallbackNotice: async (): Promise<void> => {
      setShowAutoRenderFallbackNotice(false);
      await notificationService?.markAutoRenderDismissed();
    },
    showAutoRenderFallbackNotice,
  };
}
