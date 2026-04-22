import { createRoot } from "react-dom/client";
import { createComposeMarkdownService } from "../lib/compose-markdown";
import { createComposeNotificationService } from "../lib/compose-notifications";
import { createOfficeSettingsStore } from "../lib/config";
import { renderItem } from "../lib/item";
import { TaskpaneApp } from "./app";
import {
  getStrings,
  resolveLocale,
  resolveOfficeDisplayLanguage,
} from "./i18n";

export function mountTaskpane(rootElement: HTMLElement): void {
  const root = createRoot(rootElement);
  const locale = resolveLocale(resolveOfficeDisplayLanguage());

  root.render(
    <TaskpaneApp
      locale={locale}
      notificationService={createComposeNotificationService()}
      services={{
        composeMarkdown: createComposeMarkdownService(),
        renderEntireDraft: renderItem,
      }}
      settingsStore={createOfficeSettingsStore()}
      strings={getStrings(locale)}
    />
  );
}
