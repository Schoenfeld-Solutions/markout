import type { Diagnostic } from "@codemirror/lint";
import type { HighlightStyle, LanguageSupport } from "@codemirror/language";
import type {
  EditorState as CodeMirrorEditorState,
  Extension,
  TransactionSpec,
} from "@codemirror/state";
import type {
  EditorView as CodeMirrorEditorView,
  ViewUpdate as CodeMirrorViewUpdate,
} from "@codemirror/view";
import type { StyleSpec } from "style-mod";
import { tokens } from "@fluentui/react-components";
import { useEffect, useRef, useState, type RefObject } from "react";
import type { StylesheetLintResult } from "../lib/stylesheet-lint";
import type { PanelKey } from "./types";

interface CodeMirrorEditorStateConstructor {
  create(config: {
    doc: string;
    extensions: readonly Extension[];
  }): CodeMirrorEditorState;
}

interface CodeMirrorEditorViewConstructor {
  new (config: {
    parent: Element | DocumentFragment;
    state: CodeMirrorEditorState;
  }): CodeMirrorEditorView;
  lineWrapping: Extension;
  theme(
    spec: Record<string, StyleSpec>,
    options?: { dark?: boolean }
  ): Extension;
  updateListener: {
    of(listener: (update: CodeMirrorViewUpdate) => void): Extension;
  };
}

interface CodeMirrorModules {
  css: () => LanguageSupport;
  defaultHighlightStyle: HighlightStyle;
  EditorState: CodeMirrorEditorStateConstructor;
  EditorView: CodeMirrorEditorViewConstructor;
  lineNumbers: () => Extension;
  setDiagnostics: (
    state: CodeMirrorEditorState,
    diagnostics: readonly Diagnostic[]
  ) => TransactionSpec;
  syntaxHighlighting: (
    highlighter: HighlightStyle,
    options?: { fallback: boolean }
  ) => Extension;
}

async function loadCodeMirrorModules(): Promise<CodeMirrorModules> {
  const [cssModule, languageModule, lintModule, stateModule, viewModule] =
    await Promise.all([
      import("@codemirror/lang-css"),
      import("@codemirror/language"),
      import("@codemirror/lint"),
      import("@codemirror/state"),
      import("@codemirror/view"),
    ]);

  return {
    css: cssModule.css,
    defaultHighlightStyle: languageModule.defaultHighlightStyle,
    EditorState: stateModule.EditorState,
    EditorView: viewModule.EditorView,
    lineNumbers: viewModule.lineNumbers,
    setDiagnostics: lintModule.setDiagnostics,
    syntaxHighlighting: languageModule.syntaxHighlighting,
  };
}

function findLintIssueRange(
  stylesheet: string,
  issue: StylesheetLintResult["issues"][number]
): { from: number; to: number } {
  const normalizedStylesheet = stylesheet.length > 0 ? stylesheet : " ";
  const selectorMatch = /"([^"]+)"/.exec(issue.message);

  if (selectorMatch?.[1] !== undefined) {
    const index = normalizedStylesheet.indexOf(selectorMatch[1]);

    if (index !== -1) {
      return {
        from: index,
        to: index + selectorMatch[1].length,
      };
    }
  }

  const propertyMatch = /The property "([^"]+)"/.exec(issue.message);

  if (propertyMatch?.[1] !== undefined) {
    const index = normalizedStylesheet.indexOf(propertyMatch[1]);

    if (index !== -1) {
      return {
        from: index,
        to: index + propertyMatch[1].length,
      };
    }
  }

  if (issue.code === "empty-stylesheet") {
    return { from: 0, to: 0 };
  }

  return {
    from: 0,
    to: normalizedStylesheet.length,
  };
}

function toCodeMirrorDiagnostics(
  stylesheet: string,
  lintResult: StylesheetLintResult | null
): Diagnostic[] {
  if (lintResult === null) {
    return [];
  }

  return lintResult.issues.map((issue) => {
    const { from, to } = findLintIssueRange(stylesheet, issue);

    return {
      from,
      message: issue.message,
      severity: issue.severity === "error" ? "error" : "warning",
      to,
    };
  });
}

export function useStylesheetEditor(
  activePanel: PanelKey,
  stylesheet: string,
  cssLintResult: StylesheetLintResult | null,
  resolvedColorMode: "dark" | "light",
  onStylesheetChange: (stylesheet: string) => void,
  onEditorLoadError: (message: string) => void,
  loadFailedMessage: string
): {
  codeMirrorHostRef: RefObject<HTMLDivElement | null>;
  isCodeMirrorLoading: boolean;
} {
  const [isCodeMirrorLoading, setIsCodeMirrorLoading] = useState(false);
  const codeMirrorHostRef = useRef<HTMLDivElement | null>(null);
  const codeMirrorModulesRef = useRef<CodeMirrorModules | null>(null);
  const codeMirrorViewRef = useRef<CodeMirrorEditorView | null>(null);

  useEffect(() => {
    if (activePanel !== "settings") {
      setIsCodeMirrorLoading(false);
    }
  }, [activePanel]);

  useEffect(() => {
    if (activePanel !== "settings" || codeMirrorHostRef.current === null) {
      return;
    }

    let cancelled = false;
    let editorView: CodeMirrorEditorView | null = null;

    setIsCodeMirrorLoading(true);

    void loadCodeMirrorModules()
      .then((modules) => {
        if (cancelled || codeMirrorHostRef.current === null) {
          return;
        }

        codeMirrorModulesRef.current = modules;
        const editorTheme = modules.EditorView.theme(
          {
            "&": {
              backgroundColor: "transparent",
              color: tokens.colorNeutralForeground1,
              fontFamily: tokens.fontFamilyMonospace,
              fontSize: tokens.fontSizeBase300,
              minHeight: "14rem",
            },
            ".cm-scroller": {
              fontFamily: tokens.fontFamilyMonospace,
              lineHeight: tokens.lineHeightBase300,
              minHeight: "14rem",
            },
            ".cm-content": {
              caretColor: tokens.colorNeutralForeground1,
              minHeight: "14rem",
              padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalM}`,
            },
            ".cm-gutters": {
              backgroundColor: tokens.colorNeutralBackground1,
              borderRightColor: tokens.colorNeutralStroke2,
              color: tokens.colorNeutralForeground3,
            },
            ".cm-activeLine": {
              backgroundColor:
                resolvedColorMode === "dark"
                  ? "rgba(255, 255, 255, 0.04)"
                  : "rgba(15, 108, 189, 0.06)",
            },
            ".cm-activeLineGutter": {
              backgroundColor:
                resolvedColorMode === "dark"
                  ? "rgba(255, 255, 255, 0.04)"
                  : "rgba(15, 108, 189, 0.06)",
            },
            ".cm-selectionBackground": {
              backgroundColor:
                resolvedColorMode === "dark"
                  ? "rgba(96, 165, 250, 0.28) !important"
                  : "rgba(15, 108, 189, 0.24) !important",
            },
            ".cm-diagnostic": {
              fontFamily: tokens.fontFamilyBase,
            },
          },
          { dark: resolvedColorMode === "dark" }
        );
        const updateListener = modules.EditorView.updateListener.of(
          (update: CodeMirrorViewUpdate) => {
            if (!update.docChanged) {
              return;
            }

            onStylesheetChange(update.state.doc.toString());
          }
        );

        editorView = new modules.EditorView({
          parent: codeMirrorHostRef.current,
          state: modules.EditorState.create({
            doc: stylesheet,
            extensions: [
              modules.lineNumbers(),
              modules.css(),
              modules.EditorView.lineWrapping,
              modules.syntaxHighlighting(modules.defaultHighlightStyle, {
                fallback: true,
              }),
              editorTheme,
              updateListener,
            ],
          }),
        });

        codeMirrorViewRef.current = editorView;
        setIsCodeMirrorLoading(false);
      })
      .catch((error: unknown) => {
        console.error(
          "MarkOut failed to initialize the stylesheet editor.",
          error
        );
        codeMirrorModulesRef.current = null;
        codeMirrorViewRef.current = null;
        setIsCodeMirrorLoading(false);
        if (!cancelled) {
          onEditorLoadError(loadFailedMessage);
        }
      });

    return () => {
      cancelled = true;
      editorView?.destroy();
      codeMirrorViewRef.current = null;
    };
  }, [
    activePanel,
    loadFailedMessage,
    onEditorLoadError,
    onStylesheetChange,
    resolvedColorMode,
    stylesheet,
  ]);

  useEffect(() => {
    const editorView = codeMirrorViewRef.current;

    if (editorView === null) {
      return;
    }

    const currentDocument = editorView.state.doc.toString();

    if (currentDocument === stylesheet) {
      return;
    }

    editorView.dispatch({
      changes: {
        from: 0,
        insert: stylesheet,
        to: currentDocument.length,
      },
    });
  }, [stylesheet]);

  useEffect(() => {
    const editorView = codeMirrorViewRef.current;
    const modules = codeMirrorModulesRef.current;

    if (editorView === null || modules === null) {
      return;
    }

    editorView.dispatch(
      modules.setDiagnostics(
        editorView.state,
        toCodeMirrorDiagnostics(stylesheet, cssLintResult)
      )
    );
  }, [cssLintResult, stylesheet]);

  return {
    codeMirrorHostRef,
    isCodeMirrorLoading,
  };
}
