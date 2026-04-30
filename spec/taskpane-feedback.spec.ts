import {
  EMPTY_SELECTION_MESSAGE,
  FULL_DRAFT_ALREADY_RENDERED_MESSAGE,
  RENDERED_SELECTION_BLOCKED_MESSAGE,
  SUBJECT_SELECTION_UNSUPPORTED_MESSAGE,
} from "../src/lib/compose-markdown";
import { getStrings } from "../src/taskpane/i18n";
import {
  getDraftRenderFeedback,
  localizeActionError,
} from "../src/taskpane/taskpane-feedback";

describe("taskpane feedback", () => {
  it("maps draft render results to user copy and diagnostics", () => {
    const strings = getStrings("en-US");

    expect(getDraftRenderFeedback(strings, "rendered")).toEqual({
      diagnosticArea: "render",
      diagnosticCode: "draft.render.succeeded",
      intent: "success",
      message: strings.status.draftRendered,
    });
    expect(getDraftRenderFeedback(strings, "restored")).toEqual({
      diagnosticArea: "restore",
      diagnosticCode: "draft.restore.succeeded",
      intent: "success",
      message: strings.status.draftRestored,
    });
    expect(getDraftRenderFeedback(strings, "unchanged")).toEqual({
      diagnosticArea: "render",
      diagnosticCode: "draft.render.unchanged",
      intent: "info",
      message: strings.status.draftUnchanged,
    });
  });

  it("maps known compose errors to localized taskpane messages", () => {
    const strings = getStrings("en-US");

    expect(
      localizeActionError(
        strings,
        new Error(SUBJECT_SELECTION_UNSUPPORTED_MESSAGE)
      )
    ).toBe(strings.tooltips.renderSelectionSubject);
    expect(
      localizeActionError(strings, new Error(EMPTY_SELECTION_MESSAGE))
    ).toBe(strings.tooltips.renderSelectionNoSelection);
    expect(
      localizeActionError(
        strings,
        new Error(FULL_DRAFT_ALREADY_RENDERED_MESSAGE)
      )
    ).toBe(strings.tooltips.renderEntireDraft);
    expect(
      localizeActionError(
        strings,
        new Error(RENDERED_SELECTION_BLOCKED_MESSAGE)
      )
    ).toBe(strings.tooltips.renderedFragmentBlocked);
  });

  it("maps dropped file read and decode errors to file-specific copy", () => {
    const strings = getStrings("en-US");

    expect(
      localizeActionError(
        strings,
        new Error("MarkOut could not read broken.md.")
      )
    ).toBe("broken.md could not be read.");
    expect(
      localizeActionError(
        strings,
        new Error("MarkOut could not decode broken.md.")
      )
    ).toBe("broken.md could not be decoded.");
  });

  it("falls back to safe generic or original error messages", () => {
    const strings = getStrings("en-US");

    expect(localizeActionError(strings, new Error("specific failure"))).toBe(
      "specific failure"
    );
    expect(localizeActionError(strings, "not an error")).toBe(
      strings.status.unexpectedActionFailure
    );
  });
});
