/** @jest-environment jsdom */

import { act } from "react";
import type { ReactElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import type { Diagnostic } from "@codemirror/lint";
import type { TransactionSpec } from "@codemirror/state";
import {
  findLintIssueRange,
  loadCodeMirrorModules,
  toCodeMirrorDiagnostics,
  useStylesheetEditor,
  type StylesheetEditorModules,
} from "../src/taskpane/editor";
import type { PanelKey } from "../src/taskpane/types";
import type { StylesheetLintResult } from "../src/lib/stylesheet-lint";

(
  globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true;

interface FakeState {
  doc: {
    toString: () => string;
  };
}

interface FakeViewUpdate {
  docChanged: boolean;
  state: FakeState;
}

class FakeEditorView {
  public static lineWrapping = { extension: "lineWrapping" };
  public static readonly listeners: ((update: FakeViewUpdate) => void)[] = [];
  public static readonly themes: unknown[] = [];
  public static updateListener = {
    of: (listener: (update: FakeViewUpdate) => void): unknown => {
      FakeEditorView.listeners.push(listener);
      return { extension: "updateListener" };
    },
  };

  public readonly dispatches: TransactionSpec[] = [];
  public destroyed = false;
  public state: FakeState;

  public constructor(config: { state: FakeState }) {
    this.state = config.state;
  }

  public static reset(): void {
    FakeEditorView.listeners.length = 0;
    FakeEditorView.themes.length = 0;
  }

  public static theme(spec: unknown, options?: { dark?: boolean }): unknown {
    FakeEditorView.themes.push({ options, spec });
    return { extension: "theme" };
  }

  public destroy(): void {
    this.destroyed = true;
  }

  public dispatch(transaction: TransactionSpec): void {
    this.dispatches.push(transaction);

    if (
      typeof transaction === "object" &&
      "changes" in transaction &&
      typeof transaction.changes === "object" &&
      "insert" in transaction.changes &&
      typeof transaction.changes.insert === "string"
    ) {
      const nextDocument = transaction.changes.insert;
      this.state = {
        doc: {
          toString: () => nextDocument,
        },
      };
    }
  }
}

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

function createFakeModules(
  createdViews: FakeEditorView[],
  diagnostics: Diagnostic[][] = []
): StylesheetEditorModules {
  class CapturingEditorView extends FakeEditorView {
    public constructor(config: { state: FakeState }) {
      super(config);
      createdViews.push(this);
    }
  }

  CapturingEditorView.reset();

  return {
    css: () => ({ extension: "css" }) as never,
    defaultHighlightStyle: { extension: "highlight" } as never,
    EditorState: {
      create: (config) =>
        ({
          doc: {
            toString: () => config.doc,
          },
        }) as never,
    },
    EditorView: CapturingEditorView as never,
    lineNumbers: () => ({ extension: "lineNumbers" }) as never,
    setDiagnostics: (_state, nextDiagnostics) => {
      diagnostics.push([...nextDiagnostics]);
      return { effects: [] };
    },
    syntaxHighlighting: () => ({ extension: "syntaxHighlighting" }) as never,
  };
}

function mount(element: ReactElement): {
  container: HTMLElement;
  root: Root;
} {
  document.body.innerHTML = '<div id="root"></div>';
  const container = document.getElementById("root");

  if (container === null) {
    throw new Error("Expected an editor test container.");
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

describe("stylesheet editor helpers", () => {
  it("loads the production CodeMirror module bundle", async () => {
    await expect(loadCodeMirrorModules()).resolves.toMatchObject({
      css: expect.any(Function),
      EditorState: expect.any(Function),
      EditorView: expect.any(Function),
      lineNumbers: expect.any(Function),
      setDiagnostics: expect.any(Function),
      syntaxHighlighting: expect.any(Function),
    });
  });

  it("maps lint issues to useful CodeMirror ranges and diagnostics", () => {
    const lintResult: StylesheetLintResult = {
      issues: [
        {
          code: "unsupported-selector",
          message: 'The selector ".bad" is not allowed.',
          severity: "warning",
        },
        {
          code: "invalid-rule",
          message: 'The property "colour" is not allowed.',
          severity: "error",
        },
        {
          code: "empty-stylesheet",
          message: "Stylesheet is empty.",
          severity: "warning",
        },
      ],
      validRuleCount: 0,
    };

    expect(
      findLintIssueRange(".bad { colour: red; }", lintResult.issues[0]!)
    ).toEqual({ from: 0, to: 4 });
    expect(
      findLintIssueRange(".bad { colour: red; }", lintResult.issues[1]!)
    ).toEqual({ from: 7, to: 13 });
    expect(findLintIssueRange("", lintResult.issues[2]!)).toEqual({
      from: 0,
      to: 0,
    });
    expect(
      findLintIssueRange(".safe { color: red; }", {
        code: "invalid-rule",
        message: 'The property "missing" is not allowed.',
        severity: "error",
      })
    ).toEqual({
      from: 0,
      to: ".safe { color: red; }".length,
    });
    expect(
      findLintIssueRange(".safe { color: red; }", {
        code: "invalid-rule",
        message: "The rule is not allowed.",
        severity: "error",
      })
    ).toEqual({
      from: 0,
      to: ".safe { color: red; }".length,
    });
    expect(
      toCodeMirrorDiagnostics(".bad { colour: red; }", lintResult)
    ).toEqual([
      {
        from: 0,
        message: 'The selector ".bad" is not allowed.',
        severity: "warning",
        to: 4,
      },
      {
        from: 7,
        message: 'The property "colour" is not allowed.',
        severity: "error",
        to: 13,
      },
      {
        from: 0,
        message: "Stylesheet is empty.",
        severity: "warning",
        to: 0,
      },
    ]);
    expect(toCodeMirrorDiagnostics("", null)).toEqual([]);
  });
});

describe("useStylesheetEditor", () => {
  afterEach(() => {
    jest.restoreAllMocks();
  });

  it("loads CodeMirror, reports edits, syncs external stylesheet changes, and dispatches diagnostics", async () => {
    const createdViews: FakeEditorView[] = [];
    const diagnostics: Diagnostic[][] = [];
    const changes: string[] = [];
    const modules = createFakeModules(createdViews, diagnostics);
    const loadCodeMirrorModules = () => Promise.resolve(modules);
    const onStylesheetChange = (nextStylesheet: string): void => {
      changes.push(nextStylesheet);
    };
    const onEditorLoadError = (): void => undefined;
    const lintResult: StylesheetLintResult = {
      issues: [
        {
          code: "invalid-rule",
          message: 'The property "colour" is not allowed.',
          severity: "error",
        },
      ],
      validRuleCount: 0,
    };

    function EditorProbe({
      activePanel,
      stylesheet,
      stylesheetLintResult,
    }: {
      activePanel: PanelKey;
      stylesheet: string;
      stylesheetLintResult: StylesheetLintResult | null;
    }): ReactElement {
      const { codeMirrorHostRef, isCodeMirrorLoading } = useStylesheetEditor(
        activePanel,
        stylesheet,
        stylesheetLintResult,
        "dark",
        onStylesheetChange,
        onEditorLoadError,
        "Editor failed.",
        {
          loadCodeMirrorModules,
        }
      );

      return (
        <div data-loading={String(isCodeMirrorLoading)}>
          <div id="editor-host" ref={codeMirrorHostRef} />
        </div>
      );
    }

    const { container, root } = mount(
      <EditorProbe
        activePanel="settings"
        stylesheet=".mo { color: inherit; }"
        stylesheetLintResult={null}
      />
    );
    try {
      await flushPromises();

      expect(container.firstElementChild?.getAttribute("data-loading")).toBe(
        "false"
      );
      expect(createdViews).toHaveLength(1);
      expect(createdViews[0]?.state.doc.toString()).toBe(
        ".mo { color: inherit; }"
      );

      act(() => {
        FakeEditorView.listeners[0]?.({
          docChanged: false,
          state: {
            doc: {
              toString: () => ".ignored{}",
            },
          },
        });
        FakeEditorView.listeners[0]?.({
          docChanged: true,
          state: {
            doc: {
              toString: () => ".changed{}",
            },
          },
        });
      });
      expect(changes).toEqual([".changed{}"]);

      act(() => {
        root.render(
          <EditorProbe
            activePanel="settings"
            stylesheet=".next { colour: red; }"
            stylesheetLintResult={lintResult}
          />
        );
      });
      await flushPromises();

      expect(createdViews[0]?.dispatches).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            changes: {
              from: 0,
              insert: ".next { colour: red; }",
              to: ".mo { color: inherit; }".length,
            },
          }),
        ])
      );
      expect(diagnostics.at(-1)).toEqual([
        {
          from: 8,
          message: 'The property "colour" is not allowed.',
          severity: "error",
          to: 14,
        },
      ]);

      const dispatchCount = createdViews[0]?.dispatches.length ?? 0;
      act(() => {
        root.render(
          <EditorProbe
            activePanel="settings"
            stylesheet=".next { colour: red; }"
            stylesheetLintResult={lintResult}
          />
        );
      });
      await flushPromises();
      expect(createdViews[0]?.dispatches.length).toBe(dispatchCount);

      act(() => {
        root.render(
          <EditorProbe
            activePanel="help"
            stylesheet=".next { colour: red; }"
            stylesheetLintResult={lintResult}
          />
        );
      });
      expect(createdViews[0]?.destroyed).toBe(true);
    } finally {
      act(() => {
        root.unmount();
      });
    }
  });

  it("surfaces loader failures and ignores loaders resolved after unmount", async () => {
    const consoleError = jest
      .spyOn(console, "error")
      .mockImplementation(() => undefined);
    const errors: string[] = [];
    const createdViews: FakeEditorView[] = [];
    const onStylesheetChange = (): void => undefined;
    const onEditorLoadError = (message: string): void => {
      errors.push(message);
    };

    function EditorProbe({
      loadCodeMirrorModules,
    }: {
      loadCodeMirrorModules: () => Promise<StylesheetEditorModules>;
    }): ReactElement {
      const { codeMirrorHostRef, isCodeMirrorLoading } = useStylesheetEditor(
        "settings",
        ".mo{}",
        null,
        "light",
        onStylesheetChange,
        onEditorLoadError,
        "Editor failed.",
        {
          loadCodeMirrorModules,
        }
      );

      return (
        <div data-loading={String(isCodeMirrorLoading)}>
          <div id="editor-host" ref={codeMirrorHostRef} />
        </div>
      );
    }

    const { container, root } = mount(
      <EditorProbe
        loadCodeMirrorModules={() => Promise.reject(new Error("load failed"))}
      />
    );
    try {
      await flushPromises();
      expect(container.firstElementChild?.getAttribute("data-loading")).toBe(
        "false"
      );
      expect(errors).toEqual(["Editor failed."]);
      expect(consoleError).toHaveBeenCalled();
    } finally {
      act(() => {
        root.unmount();
      });
    }

    const pendingModules = deferred<StylesheetEditorModules>();
    const mounted = mount(
      <EditorProbe loadCodeMirrorModules={() => pendingModules.promise} />
    );
    mounted.root.unmount();

    await act(async () => {
      pendingModules.resolve(createFakeModules(createdViews));
      await Promise.resolve();
    });
    expect(createdViews).toHaveLength(0);
  });
});
