export type SupportedLocale = "de-DE" | "en-US";

export interface LocalizedStrings {
  appTitle: string;
  buyMeACoffeePlaceholder: string;
  credits: {
    currentMaintenanceBody: string;
    currentMaintenanceTitle: string;
    openFork: string;
    panelDescription: string;
    panelTitle: string;
    starFork: string;
    upstreamBody: string;
    upstreamTitle: string;
  };
  developer: {
    hostNotesTitle: string;
    inspectSelection: string;
    noSelectionSnapshot: string;
    panelDescription: string;
    panelTitle: string;
    resolvedTheme: string;
    ribbonHint: string;
    subjectHint: string;
    taskpaneHint: string;
  };
  editor: {
    lintErrorLabel: string;
    lintButton: string;
    lintNoIssues: string;
    lintWarningLabel: string;
    resetButton: string;
    title: string;
  };
  general: {
    hidden: string;
    off: string;
    on: string;
    shown: string;
  };
  help: {
    docsDescription: string;
    docsTitle: string;
    panelDescription: string;
    panelTitle: string;
    repoDescription: string;
    repoTitle: string;
    websiteDescription: string;
    websiteTitle: string;
  };
  intro: {
    confirm: string;
    panelDescription: string;
    panelTitle: string;
    stepOneBody: string;
    stepOneTitle: string;
    stepTwoBody: string;
    stepTwoTitle: string;
  };
  insert: {
    dropzoneCopy: string;
    dropzoneTitle: string;
    emptyPreview: string;
    insertButton: string;
    inputLabel: string;
    inputPlaceholder: string;
    panelDescription: string;
    panelTitle: string;
    previewDescription: string;
    previewLoading: string;
    previewTitle: string;
    renderEntireDraftButton: string;
    renderSelectionButton: string;
  };
  localization: {
    buyMeACoffeePlaceholderDescription: string;
    supportedLanguagesNote: string;
  };
  notifications: {
    autoRenderFallbackBody: string;
    autoRenderFallbackDismiss: string;
    autoRenderFallbackTitle: string;
    autoRenderStickyBody: string;
  };
  settings: {
    autoRenderDescription: string;
    autoRenderTitle: string;
    creditsDescription: string;
    creditsTitle: string;
    developerDescription: string;
    developerTitle: string;
    helpDescription: string;
    helpTitle: string;
    introDescription: string;
    introTitle: string;
    panelDescription: string;
    panelTitle: string;
    themeDescription: string;
    themeModeDark: string;
    themeModeLight: string;
    themeModeSystem: string;
    themeTitle: string;
  };
  status: {
    autoRenderDisabled: string;
    autoRenderEnabled: string;
    cannotLoadSelection: string;
    cssLintComplete: string;
    cssLintFailed: string;
    creditsHidden: string;
    creditsShown: string;
    developerDisabled: string;
    developerEnabled: string;
    draftRendered: string;
    draftRestored: string;
    fileDecodeFailed: (fileName: string) => string;
    fileReadFailed: (fileName: string) => string;
    dropFileInstruction: string;
    fragmentInserted: string;
    fragmentReplaced: string;
    helpHidden: string;
    helpShown: string;
    introHidden: string;
    introRestored: string;
    previewFailed: string;
    selectionInspectionFailed: string;
    selectionInspectionSuccess: string;
    selectionRendered: string;
    settingsUpdateFailed: string;
    stylesheetLoaded: (fileName: string) => string;
    stylesheetSaved: string;
    stylesheetSaveFailed: string;
    themeUpdated: (mode: string) => string;
    unsupportedFileType: string;
    unexpectedActionFailure: string;
  };
  toolbar: {
    credits: string;
    developer: string;
    help: string;
    insert: string;
    intro: string;
    settings: string;
  };
  tooltips: {
    insertRenderedMarkdown: string;
    renderEntireDraft: string;
    renderSelection: string;
    renderSelectionNoSelection: string;
    renderSelectionSubject: string;
    renderSelectionUnknown: string;
    renderedFragmentBlocked: string;
    toolbarCompactHint: (label: string) => string;
  };
}

const EN_US: LocalizedStrings = {
  appTitle: "MarkOut",
  buyMeACoffeePlaceholder: "Buy me a coffee",
  credits: {
    currentMaintenanceBody:
      "This fork is maintained by Schoenfeld Solutions, with ongoing work by Gabriel-Johannes Schönfeld.",
    currentMaintenanceTitle: "Current maintenance",
    openFork: "Open the fork",
    panelDescription:
      "MarkOut continues as an independently maintained fork while preserving visible credit to the upstream project.",
    panelTitle: "Credits",
    starFork: "Leave a star",
    upstreamBody:
      "Original product direction and implementation work came from SierraSoftworks/markout.",
    upstreamTitle: "Upstream foundation",
  },
  developer: {
    hostNotesTitle: "Host notes",
    inspectSelection: "Inspect selection",
    noSelectionSnapshot: "No selection snapshot loaded yet.",
    panelDescription:
      "Inspect host theme resolution and selection state without exposing debug noise to regular users.",
    panelTitle: "Developer tools",
    resolvedTheme: "Preference: {mode}. Effective theme: {resolvedMode}.",
    ribbonHint:
      "Outlook decides whether the command appears directly on the ribbon or in Apps.",
    subjectHint:
      "MarkOut can only update the message body. Move the cursor into the body or select text there first.",
    taskpaneHint:
      "Native Outlook context-menu commands are not the delivery path for this add-in.",
  },
  editor: {
    lintErrorLabel: "Error",
    lintButton: "Lint CSS",
    lintNoIssues: "No lint findings.",
    lintWarningLabel: "Warning",
    resetButton: "Reset default stylesheet",
    title: "Inline stylesheet",
  },
  general: {
    hidden: "Hidden",
    off: "Off",
    on: "On",
    shown: "Shown",
  },
  help: {
    docsDescription:
      "Open the GitHub Pages landing page with manifests, hosted docs, and deployment notes.",
    docsTitle: "Hosted project docs",
    panelDescription:
      "Use these links for the maintained fork, hosted project docs, and the operator website.",
    panelTitle: "Help",
    repoDescription:
      "Track issues, releases, and the maintained Schoenfeld Solutions fork.",
    repoTitle: "GitHub repository",
    websiteDescription:
      "Open the Schoenfeld Solutions website. A support link can be added here later.",
    websiteTitle: "Schoenfeld Solutions",
  },
  intro: {
    confirm: "I have read this",
    panelDescription:
      "MarkOut keeps compose work Markdown-first while staying inside Outlook's taskpane and Smart Alerts model.",
    panelTitle: "Intro",
    stepOneBody:
      "Render the whole draft or only a selected body range without leaving compose mode.",
    stepOneTitle: "Render when you are ready",
    stepTwoBody:
      "Drop a file or paste Markdown, preview the fragment, then insert it where Outlook still preserves your body selection or cursor.",
    stepTwoTitle: "Insert fragments safely",
  },
  insert: {
    dropzoneCopy:
      "MarkOut accepts .md, .markdown, and .txt files. You can also paste Markdown directly into the editor below.",
    dropzoneTitle: "Drop a Markdown file here",
    emptyPreview:
      "Paste or drop Markdown to preview the fragment that will be inserted into the draft.",
    insertButton: "Insert rendered markdown",
    inputLabel: "Markdown input",
    inputPlaceholder:
      "Paste Markdown here, or drop a Markdown file into the pane.",
    panelDescription:
      "Build a fragment in the pane, replace a selected body range, or insert rendered content at the current body cursor.",
    panelTitle: "Insert rendered Markdown",
    previewDescription:
      "Preview uses the same sanitized fragment pipeline that MarkOut inserts into the draft body.",
    previewLoading: "Rendering preview...",
    previewTitle: "Preview",
    renderEntireDraftButton: "Render entire draft",
    renderSelectionButton: "Render selection",
  },
  localization: {
    buyMeACoffeePlaceholderDescription: "Reserved for a future support link.",
    supportedLanguagesNote: "Supported UI languages: English and German.",
  },
  notifications: {
    autoRenderFallbackBody:
      "Auto-render on send is enabled for this draft. MarkOut will try to render the entire draft when Smart Alerts run.",
    autoRenderFallbackDismiss: "Dismiss",
    autoRenderFallbackTitle: "Auto-render is enabled",
    autoRenderStickyBody:
      "MarkOut will render the current draft automatically when Smart Alerts run during send.",
  },
  settings: {
    autoRenderDescription:
      "Render the entire draft before send when Smart Alerts run.",
    autoRenderTitle: "Smart Alerts auto-render",
    creditsDescription: "Show or hide the credits icon in the bottom toolbar.",
    creditsTitle: "Credits visibility",
    developerDescription:
      "Reveal the developer panel and additional host diagnostics.",
    developerTitle: "Developer tools",
    helpDescription: "Show or hide the help icon in the bottom toolbar.",
    helpTitle: "Help visibility",
    introDescription: "Restore or hide the intro icon in the bottom toolbar.",
    introTitle: "Intro visibility",
    panelDescription:
      "Control theme behavior, Smart Alerts rendering, panel visibility, and the inline stylesheet.",
    panelTitle: "Settings",
    themeDescription:
      "System follows Outlook theme when the host provides it and falls back to the browser preference otherwise.",
    themeModeDark: "Dark",
    themeModeLight: "Light",
    themeModeSystem: "System",
    themeTitle: "Theme mode",
  },
  status: {
    autoRenderDisabled: "Auto-render on send disabled.",
    autoRenderEnabled: "Auto-render on send enabled.",
    cannotLoadSelection: "Selection state could not be read from Outlook.",
    cssLintComplete: "CSS lint completed.",
    cssLintFailed: "CSS lint could not be completed.",
    creditsHidden: "Credits hidden from the toolbar.",
    creditsShown: "Credits restored to the toolbar.",
    developerDisabled: "Developer tools disabled.",
    developerEnabled: "Developer tools enabled.",
    draftRendered: "The current draft was rendered successfully.",
    draftRestored: "The original draft HTML was restored successfully.",
    fileDecodeFailed: (fileName) => `${fileName} could not be decoded.`,
    fileReadFailed: (fileName) => `${fileName} could not be read.`,
    dropFileInstruction:
      "Drop a Markdown or text file to load content into MarkOut.",
    fragmentInserted:
      "Rendered Markdown was inserted at the current body cursor.",
    fragmentReplaced: "Rendered Markdown replaced the current selection.",
    helpHidden: "Help hidden from the toolbar.",
    helpShown: "Help restored to the toolbar.",
    introHidden: "Intro hidden from the toolbar.",
    introRestored: "Intro restored to the toolbar.",
    previewFailed:
      "Preview could not be rendered with the current Markdown or stylesheet.",
    selectionInspectionFailed:
      "Selection state could not be read from Outlook.",
    selectionInspectionSuccess: "Selection state refreshed from Outlook.",
    selectionRendered: "The current body selection was rendered successfully.",
    settingsUpdateFailed: "Settings could not be updated.",
    stylesheetLoaded: (fileName) => `${fileName} loaded into the insert pane.`,
    stylesheetSaved: "Stylesheet changes saved.",
    stylesheetSaveFailed: "Stylesheet changes could not be persisted.",
    themeUpdated: (mode) => `Theme mode updated to ${mode}.`,
    unsupportedFileType:
      "Only .md, .markdown, and .txt files are supported in the insert pane.",
    unexpectedActionFailure: "MarkOut could not complete that action.",
  },
  toolbar: {
    credits: "Credits",
    developer: "Developer",
    help: "Help",
    insert: "Insert",
    intro: "Intro",
    settings: "Settings",
  },
  tooltips: {
    insertRenderedMarkdown:
      "Render the Markdown in the editor and replace the current body selection, or insert it at the current body cursor.",
    renderEntireDraft:
      "Render the entire draft body. If the draft was already rendered, MarkOut restores the original draft HTML instead.",
    renderSelection:
      "Render only the currently selected Markdown text in the message body.",
    renderSelectionNoSelection:
      "Select Markdown text in the message body before using Render selection.",
    renderSelectionSubject:
      "MarkOut can only update the message body. Move the cursor into the body or select text there first.",
    renderSelectionUnknown:
      "Selection state could not be read from Outlook. Focus the message body and try again.",
    renderedFragmentBlocked:
      "MarkOut won't replace content that already contains rendered MarkOut markup.",
    toolbarCompactHint: (label) => `Open ${label}`,
  },
};

const DE_DE: LocalizedStrings = {
  appTitle: "MarkOut",
  buyMeACoffeePlaceholder: "Buy me a coffee",
  credits: {
    currentMaintenanceBody:
      "Dieser Fork wird von Schoenfeld Solutions gepflegt, mit laufender Weiterentwicklung durch Gabriel-Johannes Schönfeld.",
    currentMaintenanceTitle: "Aktuelle Pflege",
    openFork: "Fork öffnen",
    panelDescription:
      "MarkOut wird als eigenständig gepflegter Fork fortgeführt und gibt dem Upstream weiterhin sichtbar Credit.",
    panelTitle: "Credits",
    starFork: "Stern vergeben",
    upstreamBody:
      "Die ursprüngliche Produktidee und große Teile der ersten Implementierung stammen aus SierraSoftworks/markout.",
    upstreamTitle: "Upstream-Basis",
  },
  developer: {
    hostNotesTitle: "Host-Hinweise",
    inspectSelection: "Selektion prüfen",
    noSelectionSnapshot: "Noch kein Selektions-Snapshot geladen.",
    panelDescription:
      "Host-Theme und Selektionszustand prüfen, ohne normale Nutzer mit Debug-Ausgaben zu belasten.",
    panelTitle: "Developer-Tools",
    resolvedTheme: "Präferenz: {mode}. Effektives Theme: {resolvedMode}.",
    ribbonHint:
      "Outlook entscheidet, ob der Befehl direkt im Ribbon oder nur unter Apps erscheint.",
    subjectHint:
      "MarkOut kann nur den Nachrichten-Body aktualisieren. Setze den Cursor in den Body oder markiere dort Text.",
    taskpaneHint:
      "Native Outlook-Kontextmenü-Befehle sind für dieses Add-in kein Zielpfad.",
  },
  editor: {
    lintErrorLabel: "Fehler",
    lintButton: "CSS prüfen",
    lintNoIssues: "Keine Lint-Hinweise.",
    lintWarningLabel: "Warnung",
    resetButton: "Default-Stylesheet zurücksetzen",
    title: "Inline-Stylesheet",
  },
  general: {
    hidden: "Verborgen",
    off: "Aus",
    on: "An",
    shown: "Sichtbar",
  },
  help: {
    docsDescription:
      "Öffnet die GitHub-Pages-Landingpage mit Manifests, gehosteter Doku und Deployment-Hinweisen.",
    docsTitle: "Gehostete Projektdoku",
    panelDescription:
      "Diese Links führen zum gepflegten Fork, zur gehosteten Projektdoku und zur Betreiber-Website.",
    panelTitle: "Hilfe",
    repoDescription:
      "Issues, Releases und den gepflegten Schoenfeld-Solutions-Fork ansehen.",
    repoTitle: "GitHub-Repository",
    websiteDescription:
      "Öffnet die Website von Schoenfeld Solutions. Ein Support-Link kann hier später ergänzt werden.",
    websiteTitle: "Schoenfeld Solutions",
  },
  intro: {
    confirm: "Ich habe es gelesen",
    panelDescription:
      "MarkOut hält den Compose-Flow Markdown-first und bleibt dabei vollständig in Outlook-Taskpane und Smart Alerts.",
    panelTitle: "Einführung",
    stepOneBody:
      "Den ganzen Entwurf oder nur eine selektierte Stelle rendern, ohne Compose zu verlassen.",
    stepOneTitle: "Rendern, wenn du bereit bist",
    stepTwoBody:
      "Datei droppen oder Markdown einfügen, Fragment prüfen und genau dort einsetzen, wo Outlook noch Selektion oder Cursor kennt.",
    stepTwoTitle: "Fragmente sicher einfügen",
  },
  insert: {
    dropzoneCopy:
      "MarkOut akzeptiert .md-, .markdown- und .txt-Dateien. Markdown kann außerdem direkt unten in den Editor eingefügt werden.",
    dropzoneTitle: "Markdown-Datei hier ablegen",
    emptyPreview:
      "Markdown einfügen oder droppen, um das Fragment zu sehen, das in den Entwurf eingefügt wird.",
    insertButton: "Gerendertes Markdown einfügen",
    inputLabel: "Markdown-Eingabe",
    inputPlaceholder:
      "Markdown hier einfügen oder eine Markdown-Datei in die Pane ziehen.",
    panelDescription:
      "Ein Fragment in der Pane vorbereiten, einen selektierten Body-Bereich ersetzen oder gerenderten Inhalt am aktuellen Cursor einfügen.",
    panelTitle: "Gerendertes Markdown einfügen",
    previewDescription:
      "Die Vorschau nutzt dieselbe sanitizte Fragment-Pipeline wie der spätere Draft-Insert.",
    previewLoading: "Vorschau wird gerendert...",
    previewTitle: "Vorschau",
    renderEntireDraftButton: "Gesamten Entwurf rendern",
    renderSelectionButton: "Selektion rendern",
  },
  localization: {
    buyMeACoffeePlaceholderDescription:
      "Reservierter Platz für einen späteren Support-Link.",
    supportedLanguagesNote: "Unterstützte UI-Sprachen: Englisch und Deutsch.",
  },
  notifications: {
    autoRenderFallbackBody:
      "Auto-Render beim Senden ist für diesen Entwurf aktiv. MarkOut versucht, den gesamten Entwurf zu rendern, sobald Smart Alerts laufen.",
    autoRenderFallbackDismiss: "Schließen",
    autoRenderFallbackTitle: "Auto-Render ist aktiv",
    autoRenderStickyBody:
      "MarkOut rendert den aktuellen Entwurf automatisch, wenn Smart Alerts beim Senden ausgeführt werden.",
  },
  settings: {
    autoRenderDescription:
      "Den gesamten Entwurf vor dem Senden rendern, wenn Smart Alerts laufen.",
    autoRenderTitle: "Smart Alerts Auto-Render",
    creditsDescription:
      "Credits-Symbol in der unteren Toolbar ein- oder ausblenden.",
    creditsTitle: "Credits-Sichtbarkeit",
    developerDescription:
      "Developer-Panel und zusatzliche Host-Diagnosen sichtbar machen.",
    developerTitle: "Developer-Tools",
    helpDescription:
      "Hilfe-Symbol in der unteren Toolbar ein- oder ausblenden.",
    helpTitle: "Hilfe-Sichtbarkeit",
    introDescription:
      "Intro-Symbol in der unteren Toolbar wiederherstellen oder ausblenden.",
    introTitle: "Intro-Sichtbarkeit",
    panelDescription:
      "Theme-Verhalten, Smart-Alerts-Rendering, Panel-Sichtbarkeit und das Inline-Stylesheet steuern.",
    panelTitle: "Einstellungen",
    themeDescription:
      "System folgt dem Outlook-Theme, wenn der Host es liefert, und fällt sonst auf die Browser-Präferenz zurück.",
    themeModeDark: "Dunkel",
    themeModeLight: "Hell",
    themeModeSystem: "System",
    themeTitle: "Theme-Modus",
  },
  status: {
    autoRenderDisabled: "Auto-Render beim Senden deaktiviert.",
    autoRenderEnabled: "Auto-Render beim Senden aktiviert.",
    cannotLoadSelection:
      "Der Selektionszustand konnte nicht aus Outlook gelesen werden.",
    cssLintComplete: "CSS-Prüfung abgeschlossen.",
    cssLintFailed: "CSS-Prüfung konnte nicht abgeschlossen werden.",
    creditsHidden: "Credits aus der Toolbar ausgeblendet.",
    creditsShown: "Credits in der Toolbar wiederhergestellt.",
    developerDisabled: "Developer-Tools deaktiviert.",
    developerEnabled: "Developer-Tools aktiviert.",
    draftRendered: "Der aktuelle Entwurf wurde erfolgreich gerendert.",
    draftRestored:
      "Das ursprüngliche Entwurfs-HTML wurde erfolgreich wiederhergestellt.",
    fileDecodeFailed: (fileName) =>
      `${fileName} konnte nicht dekodiert werden.`,
    fileReadFailed: (fileName) => `${fileName} konnte nicht gelesen werden.`,
    dropFileInstruction:
      "Markdown- oder Textdatei ablegen, um Inhalt in MarkOut zu laden.",
    fragmentInserted:
      "Gerendertes Markdown wurde am aktuellen Body-Cursor eingefügt.",
    fragmentReplaced:
      "Gerendertes Markdown hat die aktuelle Selektion ersetzt.",
    helpHidden: "Hilfe aus der Toolbar ausgeblendet.",
    helpShown: "Hilfe in der Toolbar wiederhergestellt.",
    introHidden: "Intro aus der Toolbar ausgeblendet.",
    introRestored: "Intro in der Toolbar wiederhergestellt.",
    previewFailed:
      "Die Vorschau konnte mit dem aktuellen Markdown oder Stylesheet nicht gerendert werden.",
    selectionInspectionFailed:
      "Der Selektionszustand konnte nicht aus Outlook gelesen werden.",
    selectionInspectionSuccess:
      "Der Selektionszustand wurde aus Outlook aktualisiert.",
    selectionRendered:
      "Die aktuelle Body-Selektion wurde erfolgreich gerendert.",
    settingsUpdateFailed:
      "Die Einstellungen konnten nicht aktualisiert werden.",
    stylesheetLoaded: (fileName) =>
      `${fileName} wurde in die Insert-Pane geladen.`,
    stylesheetSaved: "Stylesheet-Änderungen gespeichert.",
    stylesheetSaveFailed:
      "Stylesheet-Änderungen konnten nicht gespeichert werden.",
    themeUpdated: (mode) => `Theme-Modus auf ${mode} gesetzt.`,
    unsupportedFileType:
      "Nur .md-, .markdown- und .txt-Dateien werden in der Insert-Pane unterstützt.",
    unexpectedActionFailure:
      "MarkOut konnte diese Aktion nicht erfolgreich abschließen.",
  },
  toolbar: {
    credits: "Credits",
    developer: "Developer",
    help: "Hilfe",
    insert: "Einfügen",
    intro: "Intro",
    settings: "Einstellungen",
  },
  tooltips: {
    insertRenderedMarkdown:
      "Das Markdown aus dem Editor rendern und die aktuelle Body-Selektion ersetzen oder am aktuellen Cursor einfügen.",
    renderEntireDraft:
      "Den gesamten Entwurf rendern. Falls der Entwurf bereits gerendert wurde, stellt MarkOut stattdessen das ursprüngliche Entwurfs-HTML wieder her.",
    renderSelection:
      "Nur den aktuell ausgewählten Markdown-Text im Nachrichten-Body rendern.",
    renderSelectionNoSelection:
      "Markiere Markdown-Text im Nachrichten-Body, bevor du Selektion rendern verwendest.",
    renderSelectionSubject:
      "MarkOut kann nur den Nachrichten-Body aktualisieren. Bewege den Cursor in den Body oder markiere dort Text.",
    renderSelectionUnknown:
      "Der Selektionszustand konnte nicht aus Outlook gelesen werden. Fokussiere den Nachrichten-Body und versuche es erneut.",
    renderedFragmentBlocked:
      "MarkOut ersetzt keine Inhalte, die bereits gerendertes MarkOut-Markup enthalten.",
    toolbarCompactHint: (label) => `${label} öffnen`,
  },
};

const LOCALE_MAP: Record<SupportedLocale, LocalizedStrings> = {
  "de-DE": DE_DE,
  "en-US": EN_US,
};

export function resolveLocale(
  displayLanguage: string | undefined,
  navigatorLanguage: string | undefined = typeof navigator === "undefined"
    ? undefined
    : navigator.language
): SupportedLocale {
  const candidates = [displayLanguage, navigatorLanguage]
    .filter((value): value is string => typeof value === "string")
    .map((value) => value.toLowerCase());

  for (const candidate of candidates) {
    if (candidate.startsWith("de")) {
      return "de-DE";
    }

    if (candidate.startsWith("en")) {
      return "en-US";
    }
  }

  return "en-US";
}

export function getStrings(locale: SupportedLocale): LocalizedStrings {
  return LOCALE_MAP[locale];
}

export function resolveOfficeDisplayLanguage(): string | undefined {
  if (typeof Office === "undefined") {
    return undefined;
  }

  return Office.context.displayLanguage;
}
