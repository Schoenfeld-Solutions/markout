import {
  useDeferredValue,
  useEffect,
  useEffectEvent,
  useRef,
  useState,
} from "react";
import type { ComposeNotificationService } from "../lib/compose-notifications";
import type { ComposeMarkdownService } from "../lib/compose-markdown";
import type { PanelKey, SelectionState } from "./types";

const SELECTION_REFRESH_INTERVAL_MS = 1600;

export function usePreviewController(
  composeMarkdown: ComposeMarkdownService,
  markdownInput: string,
  stylesheet: string,
  previewFailedMessage: string,
  onPreviewError: (message: string) => void
): {
  previewHtml: string;
  previewState: "empty" | "loading" | "ready";
} {
  const [previewHtml, setPreviewHtml] = useState("");
  const [previewState, setPreviewState] = useState<
    "empty" | "loading" | "ready"
  >("empty");
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
    void composeMarkdown
      .renderPreview(deferredMarkdownInput, deferredStylesheet)
      .then((html) => {
        if (ignore) {
          return;
        }

        setPreviewHtml(html);
        setPreviewState("ready");
      })
      .catch((error: unknown) => {
        if (ignore) {
          return;
        }

        console.error("MarkOut failed to refresh the taskpane preview.", error);
        setPreviewHtml("");
        setPreviewState("empty");
        onPreviewError(previewFailedMessage);
      });

    return () => {
      ignore = true;
    };
  }, [
    composeMarkdown,
    deferredMarkdownInput,
    deferredStylesheet,
    onPreviewError,
    previewFailedMessage,
  ]);

  return { previewHtml, previewState };
}

export function useSelectionStateController(
  composeMarkdown: ComposeMarkdownService,
  activePanel: PanelKey
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
      return true;
    } catch {
      setSelectionState((currentState) => ({
        availability: "unknown",
        debug: currentState.debug,
      }));
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
  stickyMessage: string
): {
  showAutoRenderFallbackNotice: boolean;
  dismissAutoRenderFallbackNotice: () => Promise<void>;
} {
  const [showAutoRenderFallbackNotice, setShowAutoRenderFallbackNotice] =
    useState(false);
  const previousAutoRenderRef = useRef(autoRenderEnabled);

  useEffect(() => {
    if (notificationService === undefined) {
      return;
    }

    notificationService.onAutoRenderDismiss(() => {
      setShowAutoRenderFallbackNotice(false);
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
        return;
      }

      const surface = await notificationService.showAutoRenderNotification({
        message: stickyMessage,
      });

      if (!isCancelled()) {
        setShowAutoRenderFallbackNotice(surface === "pane");
      }
    })();

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
