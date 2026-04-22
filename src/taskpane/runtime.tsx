import { createRoot } from "react-dom/client";
import { createComposeMarkdownService } from "../lib/compose-markdown";
import { createOfficeSettingsStore } from "../lib/config";
import { renderItem } from "../lib/item";
import { TaskpaneApp } from "./app";

export function mountTaskpane(rootElement: HTMLElement): void {
  const root = createRoot(rootElement);

  root.render(
    <TaskpaneApp
      services={{
        composeMarkdown: createComposeMarkdownService(),
        renderEntireDraft: renderItem,
      }}
      settingsStore={createOfficeSettingsStore()}
    />
  );
}
