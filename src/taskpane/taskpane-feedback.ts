import {
  EMPTY_SELECTION_MESSAGE,
  FULL_DRAFT_ALREADY_RENDERED_MESSAGE,
  RENDERED_SELECTION_BLOCKED_MESSAGE,
  SUBJECT_SELECTION_UNSUPPORTED_MESSAGE,
} from "../lib/compose-markdown";
import type { RenderItemResult } from "../lib/item";
import type { LocalizedStrings } from "./i18n";
import type { PanelMessageState } from "./types";

export interface DraftRenderFeedback {
  diagnosticArea: "render" | "restore";
  diagnosticCode:
    | "draft.render.succeeded"
    | "draft.render.unchanged"
    | "draft.restore.succeeded";
  intent: PanelMessageState["intent"];
  message: string;
}

export function getDraftRenderFeedback(
  strings: LocalizedStrings,
  result: RenderItemResult
): DraftRenderFeedback {
  if (result === "rendered") {
    return {
      diagnosticArea: "render",
      diagnosticCode: "draft.render.succeeded",
      intent: "success",
      message: strings.status.draftRendered,
    };
  }

  if (result === "restored") {
    return {
      diagnosticArea: "restore",
      diagnosticCode: "draft.restore.succeeded",
      intent: "success",
      message: strings.status.draftRestored,
    };
  }

  return {
    diagnosticArea: "render",
    diagnosticCode: "draft.render.unchanged",
    intent: "info",
    message: strings.status.draftUnchanged,
  };
}

export function localizeActionError(
  strings: LocalizedStrings,
  error: unknown
): string {
  if (
    error instanceof Error &&
    error.message === SUBJECT_SELECTION_UNSUPPORTED_MESSAGE
  ) {
    return strings.tooltips.renderSelectionSubject;
  }

  if (error instanceof Error && error.message === EMPTY_SELECTION_MESSAGE) {
    return strings.tooltips.renderSelectionNoSelection;
  }

  if (
    error instanceof Error &&
    error.message === FULL_DRAFT_ALREADY_RENDERED_MESSAGE
  ) {
    return strings.tooltips.renderEntireDraft;
  }

  if (
    error instanceof Error &&
    error.message === RENDERED_SELECTION_BLOCKED_MESSAGE
  ) {
    return strings.tooltips.renderedFragmentBlocked;
  }

  if (error instanceof Error) {
    const readMatch = /^MarkOut could not read (.+)\.$/.exec(error.message);
    if (readMatch !== null) {
      const [, fileName = "the file"] = readMatch;
      return strings.status.fileReadFailed(fileName);
    }

    const decodeMatch = /^MarkOut could not decode (.+)\.$/.exec(error.message);
    if (decodeMatch !== null) {
      const [, fileName = "the file"] = decodeMatch;
      return strings.status.fileDecodeFailed(fileName);
    }

    return error.message;
  }

  return strings.status.unexpectedActionFailure;
}
