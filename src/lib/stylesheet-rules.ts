export interface ParsedStyleDeclaration {
  propertyName: string;
  propertyValue: string;
  priority: string;
}

export interface ParsedStyleRule {
  declarationText: string;
  selectorText: string;
}

export function parseStyleRules(stylesheet: string): ParsedStyleRule[] {
  return stylesheet
    .replaceAll(/\/\*[\s\S]*?\*\//g, "")
    .split("}")
    .flatMap((ruleFragment) => {
      const separatorIndex = ruleFragment.indexOf("{");

      if (separatorIndex === -1) {
        return [];
      }

      const selectorText = ruleFragment.slice(0, separatorIndex).trim();
      const declarationText = ruleFragment.slice(separatorIndex + 1).trim();

      if (selectorText.length === 0 || declarationText.length === 0) {
        return [];
      }

      return [{ declarationText, selectorText }];
    });
}

export function isInlineableSelector(selectorText: string): boolean {
  return selectorText
    .split(",")
    .map((selector) => selector.trim())
    .every((selector) => selector.length > 0 && !selector.includes(":"));
}

export function parseDeclarationText(
  declarationText: string
): ParsedStyleDeclaration[] {
  return declarationText
    .split(";")
    .map((fragment) => fragment.trim())
    .filter((fragment) => fragment.length > 0)
    .flatMap((fragment) => {
      const separatorIndex = fragment.indexOf(":");

      if (separatorIndex === -1) {
        return [];
      }

      const propertyName = fragment.slice(0, separatorIndex).trim();
      const rawValue = fragment.slice(separatorIndex + 1).trim();

      if (propertyName.length === 0 || rawValue.length === 0) {
        return [];
      }

      const [propertyValue = rawValue, priorityToken] = rawValue.split(
        /\s+!/,
        2
      );

      return [
        {
          priority:
            priorityToken?.trim().toLowerCase() === "important"
              ? "important"
              : "",
          propertyName,
          propertyValue:
            priorityToken === undefined ? rawValue : propertyValue.trim(),
        },
      ];
    });
}
