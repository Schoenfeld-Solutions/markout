import { Component, type ErrorInfo, type ReactNode } from "react";
import { createRoot } from "react-dom/client";
import {
  createComposeMarkdownService,
  type ComposeMarkdownDependencies,
} from "../lib/compose-markdown";
import { createComposeNotificationService } from "../lib/compose-notifications";
import { createOfficeSettingsStore, type SettingsStore } from "../lib/config";
import { DefaultHtmlSanitizer } from "../lib/html-sanitizer";
import {
  createItemRenderer,
  type RenderDependencies,
  type RenderItemResult,
} from "../lib/item";
import { createLazyMarkdownRenderer } from "../lib/lazy-markdown-renderer";
import { createOfficeBodyAccessor } from "../lib/body-accessor";
import { createOfficeRenderStateStore } from "../lib/render-state-store";
import {
  resolveRuntimeChannelConfig,
  type RuntimeChannelConfig,
} from "../lib/runtime";
import { TaskpaneApp } from "./app";
import {
  getStrings,
  type LocalizedStrings,
  resolveLocale,
  resolveOfficeDisplayLanguage,
} from "./i18n";

interface TaskpaneRuntimeErrorBoundaryProps {
  children: ReactNode;
  strings: LocalizedStrings;
}

interface TaskpaneRuntimeErrorBoundaryState {
  error: Error | null;
}

export class TaskpaneRuntimeErrorBoundary extends Component<
  TaskpaneRuntimeErrorBoundaryProps,
  TaskpaneRuntimeErrorBoundaryState
> {
  public override state: TaskpaneRuntimeErrorBoundaryState = {
    error: null,
  };

  public static getDerivedStateFromError(
    error: unknown
  ): TaskpaneRuntimeErrorBoundaryState {
    return {
      error:
        error instanceof Error
          ? error
          : new Error("An unknown error occurred."),
    };
  }

  public override componentDidCatch(error: Error, errorInfo: ErrorInfo): void {
    console.error("MarkOut taskpane crashed after mount.", error, errorInfo);
  }

  public override render(): ReactNode {
    const { error } = this.state;

    if (error === null) {
      return this.props.children;
    }

    return (
      <section
        id="taskpane-runtime-error"
        style={{
          backgroundColor: "#f3f2f1",
          boxSizing: "border-box",
          color: "#201f1e",
          display: "grid",
          height: "100%",
          margin: 0,
          minHeight: "100%",
          padding: "1rem",
        }}
      >
        <div
          style={{
            alignSelf: "start",
            backgroundColor: "#ffffff",
            border: "1px solid #d1d1d1",
            borderRadius: "12px",
            boxShadow: "0 8px 24px rgba(0, 0, 0, 0.08)",
            display: "grid",
            gap: "0.75rem",
            padding: "1rem",
          }}
        >
          <p
            style={{
              color: "#0f6cbd",
              fontSize: "0.75rem",
              fontWeight: 700,
              letterSpacing: "0.08em",
              margin: 0,
              textTransform: "uppercase",
            }}
          >
            MarkOut
          </p>
          <h1
            style={{
              fontSize: "1.25rem",
              lineHeight: 1.2,
              margin: 0,
            }}
          >
            {this.props.strings.runtimeError.title}
          </h1>
          <p
            style={{
              color: "#605e5c",
              lineHeight: 1.5,
              margin: 0,
            }}
          >
            {this.props.strings.runtimeError.body}
          </p>
          <pre
            style={{
              backgroundColor: "#f3f2f1",
              borderRadius: "8px",
              fontFamily:
                'ui-monospace, "Cascadia Code", "SFMono-Regular", Consolas, monospace',
              fontSize: "0.75rem",
              margin: 0,
              overflowX: "auto",
              padding: "0.75rem",
              whiteSpace: "pre-wrap",
              wordBreak: "break-word",
            }}
          >
            {`${error.name}: ${error.message}`}
          </pre>
        </div>
      </section>
    );
  }
}

export function mountTaskpane(rootElement: HTMLElement): void {
  const root = createRoot(rootElement);
  const runtimeChannelConfig = resolveRuntimeChannelConfig();
  const settingsStore = createOfficeSettingsStore(
    undefined,
    runtimeChannelConfig
  );
  const locale = resolveLocale(
    resolveOfficeDisplayLanguage(),
    settingsStore.getLanguagePreference()
  );
  const strings = getStrings(locale);
  const services = createTaskpaneRuntimeServices(
    runtimeChannelConfig,
    settingsStore
  );

  console.info("[MarkOut] taskpane runtime mounted", {
    channel: runtimeChannelConfig.channelId,
    locale,
  });

  root.render(
    <TaskpaneRuntimeErrorBoundary strings={strings}>
      <TaskpaneApp
        locale={locale}
        notificationService={createComposeNotificationService(
          undefined,
          runtimeChannelConfig
        )}
        services={services}
        settingsStore={settingsStore}
      />
    </TaskpaneRuntimeErrorBoundary>
  );
}

function createTaskpaneRuntimeServices(
  runtimeChannelConfig: RuntimeChannelConfig,
  settingsStore: SettingsStore
): {
  composeMarkdown: ReturnType<typeof createComposeMarkdownService>;
  renderEntireDraft: () => Promise<RenderItemResult>;
} {
  const markdownRenderer = createLazyMarkdownRenderer();
  const bodyAccessor = createOfficeBodyAccessor();
  const htmlSanitizer = new DefaultHtmlSanitizer();
  const composeDependencies: ComposeMarkdownDependencies = {
    bodyAccessor,
    htmlSanitizer,
    markdownRenderer,
    settingsStore,
  };
  const renderDependencies: RenderDependencies = {
    bodyAccessor,
    htmlSanitizer,
    markdownRenderer,
    renderStateStore: createOfficeRenderStateStore(
      undefined,
      runtimeChannelConfig
    ),
    settingsStore,
  };
  const itemRenderer = createItemRenderer(renderDependencies);

  return {
    composeMarkdown: createComposeMarkdownService(composeDependencies),
    renderEntireDraft: () => itemRenderer.renderItem(),
  };
}
