/** @jest-environment jsdom */

import type { ReactElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import type { SettingsStore } from "../src/lib/config";
import {
  TaskpaneRuntimeErrorBoundary,
  mountTaskpane,
} from "../src/taskpane/runtime";
import type { TaskpaneServices } from "../src/taskpane/types";
import { FakeMailboxItem, installOfficeEnvironment } from "./helpers";

jest.mock("react-dom/client", () => ({
  createRoot: jest.fn(),
}));

interface RuntimeBoundaryElementProps {
  children: ReactElement<TaskpaneAppElementProps>;
}

interface TaskpaneAppElementProps {
  locale: string;
  runtimeBuildInfo: {
    appVersion: string;
    buildRef: string;
    buildSha: string;
  };
  runtimeChannelConfig: {
    channelId: string;
    taskpaneUrl: string;
  };
  services: TaskpaneServices;
  settingsStore: SettingsStore;
}

const BUILD_ENV_KEYS = [
  "MARKOUT_APP_VERSION",
  "MARKOUT_BUILD_REF",
  "MARKOUT_BUILD_SHA",
] as const;

type BuildEnvKey = (typeof BUILD_ENV_KEYS)[number];
type BuildEnvSnapshot = Record<BuildEnvKey, string | undefined>;

function readBuildEnvSnapshot(): BuildEnvSnapshot {
  return {
    MARKOUT_APP_VERSION: process.env.MARKOUT_APP_VERSION,
    MARKOUT_BUILD_REF: process.env.MARKOUT_BUILD_REF,
    MARKOUT_BUILD_SHA: process.env.MARKOUT_BUILD_SHA,
  };
}

function restoreBuildEnvSnapshot(snapshot: BuildEnvSnapshot): void {
  for (const key of BUILD_ENV_KEYS) {
    const value = snapshot[key];

    if (value === undefined) {
      delete process.env[key];
    } else {
      process.env[key] = value;
    }
  }
}

describe("taskpane runtime", () => {
  const createRootMock = createRoot as jest.MockedFunction<typeof createRoot>;
  let originalBuildEnvSnapshot: BuildEnvSnapshot;

  beforeEach(() => {
    originalBuildEnvSnapshot = readBuildEnvSnapshot();
    for (const key of BUILD_ENV_KEYS) {
      delete process.env[key];
    }
    createRootMock.mockReset();
    jest.spyOn(console, "info").mockImplementation(() => undefined);
  });

  afterEach(() => {
    jest.restoreAllMocks();
    restoreBuildEnvSnapshot(originalBuildEnvSnapshot);
    delete (globalThis as { Office?: typeof Office }).Office;
  });

  it("mounts the taskpane with channel-scoped runtime services and host locale", async () => {
    const render = jest.fn();
    const root = {
      render,
      unmount: jest.fn(),
    } as unknown as Root;
    const mailboxItem = new FakeMailboxItem("<div># Heading</div>");
    installOfficeEnvironment({
      displayLanguage: "de-DE",
      mailboxItem,
    });
    createRootMock.mockReturnValue(root);
    const rootElement = document.createElement("div");

    mountTaskpane(rootElement);

    expect(createRootMock).toHaveBeenCalledWith(rootElement);
    expect(render).toHaveBeenCalledTimes(1);
    const boundaryElement = render.mock
      .calls[0][0] as ReactElement<RuntimeBoundaryElementProps>;
    expect(boundaryElement.type).toBe(TaskpaneRuntimeErrorBoundary);

    const taskpaneElement = boundaryElement.props.children;
    expect(taskpaneElement.props.locale).toBe("de-DE");
    expect(taskpaneElement.props.runtimeBuildInfo).toEqual({
      appVersion: "unknown",
      buildRef: "unknown",
      buildSha: "unknown",
    });
    expect(taskpaneElement.props.runtimeChannelConfig).toEqual(
      expect.objectContaining({
        channelId: "local",
        taskpaneUrl: "https://localhost:3000/taskpane.html?channel=local",
      })
    );
    expect(taskpaneElement.props.settingsStore.getLanguagePreference()).toBe(
      "system"
    );
    expect(taskpaneElement.props.services.composeMarkdown).toEqual(
      expect.objectContaining({
        getSelection: expect.any(Function),
        insertRenderedMarkdown: expect.any(Function),
        renderPreview: expect.any(Function),
        renderSelection: expect.any(Function),
      })
    );
    await expect(
      taskpaneElement.props.services.composeMarkdown.renderPreview("# Preview")
    ).resolves.toContain("<h1>Preview</h1>");
    await expect(
      taskpaneElement.props.services.renderEntireDraft()
    ).resolves.toMatch(/rendered|unchanged|restored/u);
    expect(console.info).toHaveBeenCalledWith(
      "[MarkOut] taskpane runtime mounted",
      {
        buildRef: "unknown",
        buildSha: "unknown",
        channel: "local",
        locale: "de-DE",
        version: "unknown",
      }
    );
  });

  it("normalizes non-error runtime crashes for the fallback UI", () => {
    const state = TaskpaneRuntimeErrorBoundary.getDerivedStateFromError("boom");

    expect(state.error).toBeInstanceOf(Error);
    expect(state.error?.message).toBe("An unknown error occurred.");
  });

  it("falls back to English locale when the Office display language is unsupported", () => {
    const render = jest.fn();
    installOfficeEnvironment({
      displayLanguage: "fr-FR",
      mailboxItem: new FakeMailboxItem("<div>Plain body</div>"),
    });
    createRootMock.mockReturnValue({
      render,
      unmount: jest.fn(),
    });

    mountTaskpane(document.createElement("div"));

    const boundaryElement = render.mock
      .calls[0][0] as ReactElement<RuntimeBoundaryElementProps>;
    const taskpaneElement = boundaryElement.props.children;

    expect(taskpaneElement.props.locale).toBe("en-US");
    expect(taskpaneElement.props.settingsStore.getStylesheet()).toContain(
      ".mo"
    );
    expect(console.info).toHaveBeenCalledWith(
      "[MarkOut] taskpane runtime mounted",
      {
        buildRef: "unknown",
        buildSha: "unknown",
        channel: "local",
        locale: "en-US",
        version: "unknown",
      }
    );
  });
});
