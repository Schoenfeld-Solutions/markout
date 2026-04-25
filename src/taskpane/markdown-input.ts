const LEADING_PASTE_SPACE_PATTERN = /(^|\r\n|\n|\r)([ \t\u00a0\u202f\u2007]+)/g;
const PASTE_SPACE_PATTERN = /[\u00a0\u202f\u2007]/g;

export function normalizeMarkdownInput(value: string): string {
  return value.replace(
    LEADING_PASTE_SPACE_PATTERN,
    (_match, lineStart: string, indentation: string) =>
      `${lineStart}${indentation.replace(PASTE_SPACE_PATTERN, " ")}`
  );
}
