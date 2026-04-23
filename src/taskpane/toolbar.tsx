import { useEffect, useRef, useState, type RefObject } from "react";
import { getStrings, type LocalizedStrings } from "./i18n";
import {
  CreditsIcon,
  DeveloperIcon,
  HelpIcon,
  InfoIcon,
  InsertIcon,
  SettingsIcon,
} from "./icons";
import type {
  PanelKey,
  PreferenceState,
  SelectionAvailability,
  ToolbarLayoutMode,
  ToolbarPanelDescriptor,
} from "./types";

const TOOLBAR_LABEL_MIN_WIDTH = 72;

export function buildToolbarPanels(
  preferences: PreferenceState,
  strings: LocalizedStrings
): ToolbarPanelDescriptor[] {
  const panels: ToolbarPanelDescriptor[] = [];

  if (!preferences.introDismissed) {
    panels.push({
      icon: <InfoIcon />,
      key: "intro",
      label: strings.toolbar.intro,
    });
  }

  panels.push({
    icon: <InsertIcon />,
    key: "insert",
    label: strings.toolbar.insert,
  });
  panels.push({
    icon: <SettingsIcon />,
    key: "settings",
    label: strings.toolbar.settings,
  });

  if (preferences.helpVisible) {
    panels.push({
      icon: <HelpIcon />,
      key: "help",
      label: strings.toolbar.help,
    });
  }

  if (preferences.developerToolsEnabled) {
    panels.push({
      icon: <DeveloperIcon />,
      key: "developer",
      label: strings.toolbar.developer,
    });
  }

  if (preferences.creditsVisible) {
    panels.push({
      icon: <CreditsIcon />,
      key: "credits",
      label: strings.toolbar.credits,
    });
  }

  return panels;
}

export function visibleToolbarPanelCount(preferences: PreferenceState): number {
  return buildToolbarPanels(preferences, getStrings("en-US")).length;
}

export function getPanelAfterVisibilityChange(
  activePanel: PanelKey,
  changedPanel: "credits" | "developer" | "help",
  visible: boolean
): PanelKey {
  return !visible && activePanel === changedPanel ? "settings" : activePanel;
}

export function resolveToolbarLayoutMode(
  availableWidth: number,
  itemCount: number
): ToolbarLayoutMode {
  return availableWidth / Math.max(itemCount, 1) >= TOOLBAR_LABEL_MIN_WIDTH
    ? "regular"
    : "compact";
}

export function getRenderSelectionTooltip(
  strings: LocalizedStrings,
  availability: SelectionAvailability
): string {
  switch (availability) {
    case "body-selection":
      return strings.tooltips.renderSelection;
    case "body-none":
      return strings.tooltips.renderSelectionNoSelection;
    case "subject":
      return strings.tooltips.renderSelectionSubject;
    default:
      return strings.tooltips.renderSelectionUnknown;
  }
}

export function isRenderSelectionDisabled(
  isBusy: boolean,
  availability: SelectionAvailability
): boolean {
  return isBusy || availability !== "body-selection";
}

export function isInsertRenderedMarkdownDisabled(
  isBusy: boolean,
  markdownInput: string
): boolean {
  return isBusy || markdownInput.trim().length === 0;
}

export function useToolbarLayoutMode(
  itemCount: number,
  forcedMode: ToolbarLayoutMode | undefined
): { mode: ToolbarLayoutMode; ref: RefObject<HTMLElement | null> } {
  const ref = useRef<HTMLElement | null>(null);
  const [mode, setMode] = useState<ToolbarLayoutMode>(forcedMode ?? "regular");

  useEffect(() => {
    if (forcedMode !== undefined) {
      setMode(forcedMode);
      return;
    }

    const updateMode = () => {
      const availableWidth = ref.current?.clientWidth ?? window.innerWidth;
      const nextMode = resolveToolbarLayoutMode(availableWidth, itemCount);
      setMode(nextMode);
    };

    updateMode();

    if (typeof ResizeObserver === "function" && ref.current !== null) {
      const resizeObserver = new ResizeObserver(updateMode);
      resizeObserver.observe(ref.current);

      return () => {
        resizeObserver.disconnect();
      };
    }

    window.addEventListener("resize", updateMode);

    return () => {
      window.removeEventListener("resize", updateMode);
    };
  }, [forcedMode, itemCount]);

  return { mode, ref };
}
