/** @jest-environment jsdom */

import { act, type ReactElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import {
  isDarkColor,
  resolveSystemColorMode,
  useResolvedColorMode,
} from "../src/taskpane/theme";
import { installOfficeEnvironment } from "./helpers";

(
  globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true;

interface TestMediaQueryList {
  addEventListener: jest.Mock;
  dispatchChange(): void;
  matches: boolean;
  removeEventListener: jest.Mock;
}

function createMediaQueryList(matches: boolean): TestMediaQueryList {
  const listeners = new Set<() => void>();

  return {
    addEventListener: jest.fn((_eventName: string, listener: () => void) => {
      listeners.add(listener);
    }),
    dispatchChange() {
      for (const listener of listeners) {
        listener();
      }
    },
    matches,
    removeEventListener: jest.fn((_eventName: string, listener: () => void) => {
      listeners.delete(listener);
    }),
  };
}

function installMatchMedia(mediaQueryList: TestMediaQueryList): () => void {
  const originalMatchMedia = window.matchMedia;

  Object.defineProperty(window, "matchMedia", {
    configurable: true,
    value: jest.fn().mockReturnValue(mediaQueryList),
  });

  return () => {
    Object.defineProperty(window, "matchMedia", {
      configurable: true,
      value: originalMatchMedia,
    });
  };
}

function mountThemeProbe(themeMode: "dark" | "light" | "system"): {
  container: HTMLElement;
  root: Root;
} {
  document.body.innerHTML = '<div id="root"></div>';
  const container = document.getElementById("root");

  if (container === null) {
    throw new Error("Expected a theme probe container.");
  }

  function ThemeProbe(): ReactElement {
    return <div id="resolved-theme">{useResolvedColorMode(themeMode)}</div>;
  }

  const root = createRoot(container);
  act(() => {
    root.render(<ThemeProbe />);
  });

  return { container, root };
}

describe("taskpane theme", () => {
  afterEach(() => {
    jest.restoreAllMocks();
    delete (globalThis as { Office?: typeof Office }).Office;
  });

  it("resolves system color mode from Office theme, browser media, and safe fallbacks", () => {
    const darkMedia = createMediaQueryList(true);
    const restoreMatchMedia = installMatchMedia(darkMedia);

    try {
      expect(resolveSystemColorMode({ bodyBackgroundColor: "#111111" })).toBe(
        "dark"
      );
      expect(resolveSystemColorMode({ bodyBackgroundColor: "#fafafa" })).toBe(
        "light"
      );
      expect(
        resolveSystemColorMode({ bodyBackgroundColor: "not-a-color" })
      ).toBe("light");
      expect(resolveSystemColorMode(undefined)).toBe("dark");
    } finally {
      restoreMatchMedia();
    }

    const originalMatchMedia = window.matchMedia;
    Object.defineProperty(window, "matchMedia", {
      configurable: true,
      value: undefined,
    });
    expect(resolveSystemColorMode(undefined)).toBe("light");
    Object.defineProperty(window, "matchMedia", {
      configurable: true,
      value: originalMatchMedia,
    });
  });

  it("normalizes hex colors before evaluating luminance", () => {
    expect(isDarkColor("#000")).toBe(true);
    expect(isDarkColor("fff")).toBe(false);
    expect(isDarkColor(" #123456 ")).toBe(true);
    expect(isDarkColor("rgb(0, 0, 0)")).toBe(false);
  });

  it("tracks browser theme changes for system mode", () => {
    const mediaQueryList = createMediaQueryList(false);
    const restoreMatchMedia = installMatchMedia(mediaQueryList);
    const mounted = mountThemeProbe("system");

    try {
      expect(mounted.container.textContent).toBe("light");

      act(() => {
        mediaQueryList.matches = true;
        mediaQueryList.dispatchChange();
      });
      expect(mounted.container.textContent).toBe("dark");
    } finally {
      act(() => {
        mounted.root.unmount();
      });
      restoreMatchMedia();
    }
  });

  it("tracks Office theme changes for system mode", async () => {
    const mediaQueryList = createMediaQueryList(false);
    const restoreMatchMedia = installMatchMedia(mediaQueryList);
    const officeEnvironment = installOfficeEnvironment();
    const mounted = mountThemeProbe("system");

    try {
      expect(mounted.container.textContent).toBe("light");

      await act(async () => {
        await officeEnvironment.triggerOfficeThemeChange({
          bodyBackgroundColor: "#111111",
        });
      });
      expect(mounted.container.textContent).toBe("dark");
    } finally {
      act(() => {
        mounted.root.unmount();
      });
      restoreMatchMedia();
    }
  });

  it("honors explicit light and dark modes over system mode", () => {
    const mediaQueryList = createMediaQueryList(true);
    const restoreMatchMedia = installMatchMedia(mediaQueryList);
    const light = mountThemeProbe("light");
    const dark = mountThemeProbe("dark");

    try {
      expect(light.container.textContent).toBe("light");
      expect(dark.container.textContent).toBe("dark");
    } finally {
      act(() => {
        light.root.unmount();
        dark.root.unmount();
      });
      restoreMatchMedia();
    }
  });
});
