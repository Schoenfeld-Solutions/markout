export const defaultStylesheet = `
.mo {
  font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Helvetica,Arial,sans-serif,"Apple Color Emoji","Segoe UI Emoji","Segoe UI Symbol";
  font-size: 14px;
  color: rgb(36,41,46);
}

code {
  font-size: 1em;
  line-height: 1.2em;
  padding: 0;
  margin: 0;
  font-family: Consolas, Inconsolata, Courier, monospace;
}

pre {
  margin: 1em !important;
  padding: 1em !important;
  border: 1px solid rgba(100, 100, 100, 0.2);
  border-radius: 3px;
}

code {
  white-space: normal;
  display: inline-block;
  color: #B21D12;
}

pre code {
  white-space: pre;
  overflow: auto;
  display: block !important;
  color: #000;
}

p {
  margin: 0 0 1.2em 0 !important;
}

table, dl, blockquote, q, ul, ol {
  margin: 1.2em 0 !important;
}

ul, ol {
  padding-left: 2em;
  margin: 2em 0;
}

li {
  margin: 0.5em 0;
}

li p {
  margin: 0.5em 0 !important;
}

ul ul, ul ol, ol ul, ol ol {
  margin: 0;
  padding-left: 1em;
}

ol ol, ul ol {
  list-style-type: lower-roman;
}

ul ul ol, ul ol ol, ol ul ol, ol ol ol {
  list-style-type: lower-alpha;
}

dl {
  padding: 0;
}

dl dt {
  font-size: 1em;
  font-weight: bold;
  font-style: italic;
}

dl dd {
  margin: 0 0 1em;
  padding: 0 1em;
}

blockquote, q {
  border-left: 4px solid #DDD;
  padding: 0 1em;
  color: #777;
  quotes: none;
}

blockquote::before, blockquote::after, q::before, q::after {
  content: none;
}

h1, h2, h3, h4, h5, h6 {
  margin: 1.3em 0 1em;
  padding: 0;
  font-weight: bold;
}

h1 {
  font-size: 1.6em;
  border-bottom: 1px solid #ddd;
}

h2 {
  font-size: 1.4em;
  border-bottom: 1px solid #eee;
}

h3 {
  font-size: 1.3em;
}

h4 {
  font-size: 1.2em;
}

h5 {
  font-size: 1em;
}

h6 {
  font-size: 1em;
  color: #777;
}

table {
  padding: 0;
  border-collapse: collapse;
  border-spacing: 0;
  font-size: 1em;
  font: inherit;
  border: 0;
}

tbody {
  margin: 0;
  padding: 0;
  border: 0;
}

table tr {
  border: 0;
  border-top: 1px solid #CCC;
  background-color: white;
  margin: 0;
  padding: 0;
}

table tr:nth-child(2n) {
  background-color: #F8F8F8;
}

table tr th, table tr td {
  font-size: 1em;
  border: 1px solid #CCC;
  margin: 0;
  padding: 0.5em 1em;
}

table tr th {
 font-weight: bold;
  background-color: #F0F0F0;
}

a {
  color: #0366d6;
  text-decoration: none;
}

.hljs {
    display: block;
    font-family: Consolas, Inconsolata, Courier, monospace;
    overflow-x: auto;
    padding: 0.5em;
    color: black
}

.hljs-variable,.hljs-template-variable,.hljs-symbol,.hljs-bullet,.hljs-section,.hljs-addition,.hljs-attribute,.hljs-link {
    color: #333
}

.hljs-string {
    color: #B21D12;
}

.hljs-comment,.hljs-quote,.hljs-meta,.hljs-deletion {
    color: #ccc
}

.hljs-keyword,.hljs-selector-tag,.hljs-section,.hljs-name,.hljs-type,.hljs-strong,.hljs-attr {
    font-weight: bold
}

.hljs-literal,.hljs-number {
    color: #409EFF;
    font-weight: bold;
}

.hljs-emphasis {
    font-style: italic
}
`;

const SETTING_STYLESHEET = "markout.stylesheet";
const SETTING_AUTORENDER = "markout.autorender";

interface RoamingSettingsLike {
  get(name: string): unknown;
  saveAsync(callback: (result: Office.AsyncResult<void>) => void): void;
  set(name: string, value: unknown): void;
}

export interface SettingsStore {
  getAutoRender(): boolean;
  getStylesheet(): string;
  save(): Promise<void>;
  setAutoRender(enabled: boolean): void;
  setStylesheet(stylesheet: string): void;
}

class InMemorySettingsStore implements SettingsStore {
  private autoRender = false;
  private stylesheet = defaultStylesheet;

  public getAutoRender(): boolean {
    return this.autoRender;
  }

  public getStylesheet(): string {
    return this.stylesheet;
  }

  public async save(): Promise<void> {
    return Promise.resolve();
  }

  public setAutoRender(enabled: boolean): void {
    this.autoRender = enabled;
  }

  public setStylesheet(stylesheet: string): void {
    this.stylesheet =
      stylesheet.trim().length > 0 ? stylesheet : defaultStylesheet;
  }
}

class OfficeSettingsStore implements SettingsStore {
  public constructor(private readonly roamingSettings: RoamingSettingsLike) {}

  public getAutoRender(): boolean {
    return this.roamingSettings.get(SETTING_AUTORENDER) === true;
  }

  public getStylesheet(): string {
    const storedStylesheet = this.roamingSettings.get(SETTING_STYLESHEET);

    if (
      typeof storedStylesheet === "string" &&
      storedStylesheet.trim().length > 0
    ) {
      return storedStylesheet;
    }

    return defaultStylesheet;
  }

  public async save(): Promise<void> {
    await new Promise<void>((resolve, reject) => {
      this.roamingSettings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          const error = new Error(result.error.message);
          error.name = result.error.name;
          reject(error);
          return;
        }

        resolve();
      });
    });
  }

  public setAutoRender(enabled: boolean): void {
    this.roamingSettings.set(SETTING_AUTORENDER, enabled);
  }

  public setStylesheet(stylesheet: string): void {
    this.roamingSettings.set(SETTING_STYLESHEET, stylesheet);
  }
}

function getDefaultRoamingSettings(): RoamingSettingsLike | undefined {
  if (typeof Office === "undefined") {
    return undefined;
  }

  return Office.context.roamingSettings;
}

function isRoamingSettingsLike(
  roamingSettings: RoamingSettingsLike | null | undefined
): roamingSettings is RoamingSettingsLike {
  return (
    roamingSettings !== undefined &&
    roamingSettings !== null &&
    typeof roamingSettings.get === "function" &&
    typeof roamingSettings.set === "function" &&
    typeof roamingSettings.saveAsync === "function"
  );
}

export function createOfficeSettingsStore(
  roamingSettings:
    | RoamingSettingsLike
    | null
    | undefined = getDefaultRoamingSettings()
): SettingsStore {
  if (!isRoamingSettingsLike(roamingSettings)) {
    return new InMemorySettingsStore();
  }

  return new OfficeSettingsStore(roamingSettings);
}

export function getStylesheet(): string {
  return createOfficeSettingsStore().getStylesheet();
}

export function setStylesheet(stylesheet: string): void {
  createOfficeSettingsStore().setStylesheet(stylesheet);
}

export async function saveStylesheet(stylesheet?: string): Promise<string> {
  const settingsStore = createOfficeSettingsStore();

  if (stylesheet !== undefined) {
    settingsStore.setStylesheet(stylesheet);
  }

  await settingsStore.save();
  return settingsStore.getStylesheet();
}

export function getAutoRender(): boolean {
  return createOfficeSettingsStore().getAutoRender();
}

export async function setAutoRender(enabled: boolean): Promise<boolean> {
  const settingsStore = createOfficeSettingsStore();
  settingsStore.setAutoRender(enabled);
  await settingsStore.save();
  return enabled;
}
