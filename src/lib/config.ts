export const defaultStylesheet = `
.mo {
  line-height: 1.5;
}

code {
  font-size: 0.95em;
  margin: 0;
  padding: 0.08em 0.3em;
  background-color: rgba(127, 127, 127, 0.08);
  border-radius: 4px;
  white-space: normal;
}

.hljs,
pre {
  background-color: rgba(127, 127, 127, 0.08);
  border: 1px solid rgba(127, 127, 127, 0.22);
  border-radius: 6px;
  display: block;
  margin: 1em 0 !important;
  overflow-x: auto;
  padding: 0.75em !important;
  white-space: pre-wrap;
  word-break: break-word;
}

pre code {
  background-color: transparent;
  border-radius: 0;
  display: block !important;
  padding: 0;
  white-space: pre-wrap;
  word-break: break-word;
}

p {
  margin: 0 0 1em 0 !important;
}

table, dl, blockquote, q, ul, ol {
  margin: 1em 0 !important;
}

ul, ol {
  margin: 1em 0;
  padding-left: 1.5em;
}

li {
  margin: 0.35em 0;
}

li p {
  margin: 0.35em 0 !important;
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
  font-weight: bold;
  font-style: italic;
}

dl dd {
  margin: 0 0 1em 0;
  padding: 0 1em;
}

blockquote, q {
  border-left: 4px solid rgba(127, 127, 127, 0.35);
  padding: 0 1em;
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
  border-bottom: 1px solid rgba(127, 127, 127, 0.3);
  padding-bottom: 0.12em;
}

h2 {
  font-size: 1.4em;
  border-bottom: 1px solid rgba(127, 127, 127, 0.24);
  padding-bottom: 0.1em;
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
}

table {
  border-spacing: 0;
  border-collapse: collapse;
  margin: 0;
  padding: 0;
  border: 0;
  width: 100%;
}

table tr {
  border: 0;
  border-top: 1px solid rgba(127, 127, 127, 0.28);
  background-color: transparent;
  margin: 0;
  padding: 0;
}

table tr:nth-child(2n) {
  background-color: rgba(127, 127, 127, 0.06);
}

table tr th, table tr td {
  border: 1px solid rgba(127, 127, 127, 0.28);
  margin: 0;
  padding: 0.5em 1em;
}

table tr th {
  background-color: rgba(127, 127, 127, 0.1);
  font-weight: bold;
}

a {
  text-decoration: underline;
}

.hljs-keyword,.hljs-selector-tag,.hljs-section,.hljs-name,.hljs-type,.hljs-strong,.hljs-attr {
  font-weight: bold;
}

.hljs-literal,.hljs-number {
  font-weight: bold;
}

.hljs-emphasis {
  font-style: italic;
}
`;

const SETTING_AUTORENDER = "markout.autorender";
const SETTING_CREDITS_VISIBLE = "markout.creditsVisible";
const SETTING_DEVELOPER_TOOLS = "markout.developerToolsEnabled";
const SETTING_HELP_VISIBLE = "markout.helpVisible";
const SETTING_INTRO_DISMISSED = "markout.introDismissed";
const SETTING_LANGUAGE_PREFERENCE = "markout.languagePreference";
const SETTING_STYLESHEET = "markout.stylesheet";
const SETTING_THEME_MODE = "markout.themeMode";

interface RoamingSettingsLike {
  get(name: string): unknown;
  saveAsync(callback: (result: Office.AsyncResult<void>) => void): void;
  set(name: string, value: unknown): void;
}

export type ThemeMode = "dark" | "light" | "system";
export type LanguagePreference = "de-DE" | "en-US" | "system";

export interface SettingsStore {
  getAutoRender(): boolean;
  getCreditsVisible(): boolean;
  getDeveloperToolsEnabled(): boolean;
  getHelpVisible(): boolean;
  getIntroDismissed(): boolean;
  getLanguagePreference(): LanguagePreference;
  getStylesheet(): string;
  getThemeMode(): ThemeMode;
  save(): Promise<void>;
  setAutoRender(enabled: boolean): void;
  setCreditsVisible(visible: boolean): void;
  setDeveloperToolsEnabled(enabled: boolean): void;
  setHelpVisible(visible: boolean): void;
  setIntroDismissed(dismissed: boolean): void;
  setLanguagePreference(preference: LanguagePreference): void;
  setStylesheet(stylesheet: string): void;
  setThemeMode(mode: ThemeMode): void;
}

function isThemeMode(value: unknown): value is ThemeMode {
  return value === "dark" || value === "light" || value === "system";
}

function isLanguagePreference(value: unknown): value is LanguagePreference {
  return value === "de-DE" || value === "en-US" || value === "system";
}

function normalizeStylesheet(stylesheet: string): string {
  return stylesheet.trim().length > 0 ? stylesheet : defaultStylesheet;
}

class InMemorySettingsStore implements SettingsStore {
  private autoRender = false;
  private creditsVisible = true;
  private developerToolsEnabled = false;
  private helpVisible = true;
  private introDismissed = false;
  private languagePreference: LanguagePreference = "system";
  private stylesheet = defaultStylesheet;
  private themeMode: ThemeMode = "system";

  public getAutoRender(): boolean {
    return this.autoRender;
  }

  public getCreditsVisible(): boolean {
    return this.creditsVisible;
  }

  public getDeveloperToolsEnabled(): boolean {
    return this.developerToolsEnabled;
  }

  public getHelpVisible(): boolean {
    return this.helpVisible;
  }

  public getIntroDismissed(): boolean {
    return this.introDismissed;
  }

  public getLanguagePreference(): LanguagePreference {
    return this.languagePreference;
  }

  public getStylesheet(): string {
    return this.stylesheet;
  }

  public getThemeMode(): ThemeMode {
    return this.themeMode;
  }

  public async save(): Promise<void> {
    return Promise.resolve();
  }

  public setAutoRender(enabled: boolean): void {
    this.autoRender = enabled;
  }

  public setCreditsVisible(visible: boolean): void {
    this.creditsVisible = visible;
  }

  public setDeveloperToolsEnabled(enabled: boolean): void {
    this.developerToolsEnabled = enabled;
  }

  public setHelpVisible(visible: boolean): void {
    this.helpVisible = visible;
  }

  public setIntroDismissed(dismissed: boolean): void {
    this.introDismissed = dismissed;
  }

  public setLanguagePreference(preference: LanguagePreference): void {
    this.languagePreference = preference;
  }

  public setStylesheet(stylesheet: string): void {
    this.stylesheet = normalizeStylesheet(stylesheet);
  }

  public setThemeMode(mode: ThemeMode): void {
    this.themeMode = mode;
  }
}

class OfficeSettingsStore implements SettingsStore {
  public constructor(private readonly roamingSettings: RoamingSettingsLike) {}

  public getAutoRender(): boolean {
    return this.roamingSettings.get(SETTING_AUTORENDER) === true;
  }

  public getCreditsVisible(): boolean {
    const storedValue = this.roamingSettings.get(SETTING_CREDITS_VISIBLE);
    return typeof storedValue === "boolean" ? storedValue : true;
  }

  public getDeveloperToolsEnabled(): boolean {
    return this.roamingSettings.get(SETTING_DEVELOPER_TOOLS) === true;
  }

  public getHelpVisible(): boolean {
    const storedValue = this.roamingSettings.get(SETTING_HELP_VISIBLE);
    return typeof storedValue === "boolean" ? storedValue : true;
  }

  public getIntroDismissed(): boolean {
    return this.roamingSettings.get(SETTING_INTRO_DISMISSED) === true;
  }

  public getLanguagePreference(): LanguagePreference {
    const storedPreference = this.roamingSettings.get(
      SETTING_LANGUAGE_PREFERENCE
    );

    return isLanguagePreference(storedPreference) ? storedPreference : "system";
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

  public getThemeMode(): ThemeMode {
    const storedThemeMode = this.roamingSettings.get(SETTING_THEME_MODE);

    return isThemeMode(storedThemeMode) ? storedThemeMode : "system";
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

  public setCreditsVisible(visible: boolean): void {
    this.roamingSettings.set(SETTING_CREDITS_VISIBLE, visible);
  }

  public setDeveloperToolsEnabled(enabled: boolean): void {
    this.roamingSettings.set(SETTING_DEVELOPER_TOOLS, enabled);
  }

  public setHelpVisible(visible: boolean): void {
    this.roamingSettings.set(SETTING_HELP_VISIBLE, visible);
  }

  public setIntroDismissed(dismissed: boolean): void {
    this.roamingSettings.set(SETTING_INTRO_DISMISSED, dismissed);
  }

  public setLanguagePreference(preference: LanguagePreference): void {
    this.roamingSettings.set(SETTING_LANGUAGE_PREFERENCE, preference);
  }

  public setStylesheet(stylesheet: string): void {
    this.roamingSettings.set(
      SETTING_STYLESHEET,
      normalizeStylesheet(stylesheet)
    );
  }

  public setThemeMode(mode: ThemeMode): void {
    this.roamingSettings.set(SETTING_THEME_MODE, mode);
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

export function getAutoRender(): boolean {
  return createOfficeSettingsStore().getAutoRender();
}

export function getStylesheet(): string {
  return createOfficeSettingsStore().getStylesheet();
}

export async function saveStylesheet(stylesheet?: string): Promise<string> {
  const settingsStore = createOfficeSettingsStore();

  if (stylesheet !== undefined) {
    settingsStore.setStylesheet(stylesheet);
  }

  await settingsStore.save();
  return settingsStore.getStylesheet();
}

export async function setAutoRender(enabled: boolean): Promise<boolean> {
  const settingsStore = createOfficeSettingsStore();
  settingsStore.setAutoRender(enabled);
  await settingsStore.save();
  return enabled;
}
