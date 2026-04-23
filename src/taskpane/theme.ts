import { useEffect, useState } from "react";
import type { ThemeMode } from "../lib/config";

export function useResolvedColorMode(themeMode: ThemeMode): "dark" | "light" {
  const [systemColorMode, setSystemColorMode] = useState<"dark" | "light">(() =>
    resolveSystemColorMode()
  );

  useEffect(() => {
    const mediaQuery =
      typeof window.matchMedia === "function"
        ? window.matchMedia("(prefers-color-scheme: dark)")
        : null;
    const mailbox =
      typeof Office === "undefined" ? null : Office.context.mailbox;
    const updateSystemColorMode = (
      officeTheme?: Partial<Office.OfficeTheme>
    ) => {
      setSystemColorMode(resolveSystemColorMode(officeTheme));
    };

    updateSystemColorMode();

    const handleMediaChange = () => {
      updateSystemColorMode();
    };

    mediaQuery?.addEventListener("change", handleMediaChange);

    if (mailbox !== null && typeof mailbox.addHandlerAsync === "function") {
      mailbox.addHandlerAsync(
        Office.EventType.OfficeThemeChanged,
        (event: Office.OfficeThemeChangedEventArgs) => {
          updateSystemColorMode(event.officeTheme);
        }
      );
    }

    return () => {
      mediaQuery?.removeEventListener("change", handleMediaChange);
    };
  }, []);

  return themeMode === "system" ? systemColorMode : themeMode;
}

export function resolveSystemColorMode(
  officeTheme: Partial<Office.OfficeTheme> | undefined = readOfficeTheme()
): "dark" | "light" {
  const officeThemeColor = officeTheme?.bodyBackgroundColor;

  if (typeof officeThemeColor === "string" && officeThemeColor.length > 0) {
    return isDarkColor(officeThemeColor) ? "dark" : "light";
  }

  if (
    typeof window.matchMedia === "function" &&
    window.matchMedia("(prefers-color-scheme: dark)").matches
  ) {
    return "dark";
  }

  return "light";
}

export function isDarkColor(color: string): boolean {
  const normalizedColor = normalizeHexColor(color);

  if (normalizedColor === null) {
    return false;
  }

  const red = Number.parseInt(normalizedColor.slice(0, 2), 16);
  const green = Number.parseInt(normalizedColor.slice(2, 4), 16);
  const blue = Number.parseInt(normalizedColor.slice(4, 6), 16);
  const luminance = (0.2126 * red + 0.7152 * green + 0.0722 * blue) / 255;

  return luminance < 0.55;
}

function readOfficeTheme(): Partial<Office.OfficeTheme> | undefined {
  if (typeof Office === "undefined") {
    return undefined;
  }

  const mailboxWithTheme = Office.context.mailbox as Office.Mailbox & {
    officeTheme?: Partial<Office.OfficeTheme>;
  };

  return mailboxWithTheme.officeTheme;
}

function normalizeHexColor(color: string): string | null {
  const trimmedColor = color.trim().replace(/^#/, "");

  if (/^[0-9a-f]{6}$/i.test(trimmedColor)) {
    return trimmedColor;
  }

  if (/^[0-9a-f]{3}$/i.test(trimmedColor)) {
    return trimmedColor
      .split("")
      .map((segment) => segment.repeat(2))
      .join("");
  }

  return null;
}
