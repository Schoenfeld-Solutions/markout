import { isSafeStyleProperty } from "./html-sanitizer";
import {
  isInlineableSelector,
  parseDeclarationText,
  parseStyleRules,
} from "./stylesheet-rules";

export type StylesheetLintSeverity = "error" | "warning";

export interface StylesheetLintIssue {
  code:
    | "empty-stylesheet"
    | "empty-rule"
    | "invalid-rule"
    | "pseudo-selector"
    | "sanitizer-unsafe-property"
    | "unsupported-selector";
  message: string;
  severity: StylesheetLintSeverity;
}

export interface StylesheetLintResult {
  issues: StylesheetLintIssue[];
  validRuleCount: number;
}

const UNSUPPORTED_SELECTOR_PATTERNS = [
  {
    code: "unsupported-selector" as const,
    message:
      "Child and sibling combinators are not supported by the MarkOut inliner.",
    pattern: /[>+~]/,
  },
  {
    code: "unsupported-selector" as const,
    message: "Attribute selectors are not supported by the MarkOut inliner.",
    pattern: /\[[^\]]+\]/,
  },
];

export function lintStylesheet(stylesheet: string): StylesheetLintResult {
  const trimmedStylesheet = stylesheet.trim();

  if (trimmedStylesheet.length === 0) {
    return {
      issues: [
        {
          code: "empty-stylesheet",
          message: "The stylesheet is empty.",
          severity: "warning",
        },
      ],
      validRuleCount: 0,
    };
  }

  const issues: StylesheetLintIssue[] = [];
  const parsedRules = parseStyleRules(trimmedStylesheet);
  const unmatchedBraces =
    countOccurrences(trimmedStylesheet, "{") !==
    countOccurrences(trimmedStylesheet, "}");

  if (unmatchedBraces) {
    issues.push({
      code: "invalid-rule",
      message: "At least one CSS rule is missing an opening or closing brace.",
      severity: "error",
    });
  }

  const ruleFragments = trimmedStylesheet
    .replaceAll(/\/\*[\s\S]*?\*\//g, "")
    .split("}")
    .map((fragment) => fragment.trim())
    .filter((fragment) => fragment.length > 0);

  for (const fragment of ruleFragments) {
    if (!fragment.includes("{")) {
      issues.push({
        code: "invalid-rule",
        message: `The rule "${fragment}" is missing a declaration block.`,
        severity: "error",
      });
    }
  }

  for (const rule of parsedRules) {
    if (rule.declarationText.trim().length === 0) {
      issues.push({
        code: "empty-rule",
        message: `The selector "${rule.selectorText}" does not define any declarations.`,
        severity: "warning",
      });
      continue;
    }

    if (rule.selectorText.includes(":")) {
      issues.push({
        code: "pseudo-selector",
        message: `The selector "${rule.selectorText}" contains a pseudo selector and will be ignored by the inline renderer.`,
        severity: "warning",
      });
    }

    if (!isInlineableSelector(rule.selectorText)) {
      issues.push({
        code: "unsupported-selector",
        message: `The selector "${rule.selectorText}" is not fully supported by the MarkOut inline renderer.`,
        severity: "warning",
      });
    }

    for (const pattern of UNSUPPORTED_SELECTOR_PATTERNS) {
      if (pattern.pattern.test(rule.selectorText)) {
        issues.push({
          code: pattern.code,
          message: `The selector "${rule.selectorText}" uses an unsupported selector form. ${pattern.message}`,
          severity: "warning",
        });
      }
    }

    const declarations = parseDeclarationText(rule.declarationText);

    if (declarations.length === 0) {
      issues.push({
        code: "invalid-rule",
        message: `The selector "${rule.selectorText}" does not contain a valid property declaration.`,
        severity: "error",
      });
      continue;
    }

    for (const declaration of declarations) {
      if (!isSafeStyleProperty(declaration.propertyName)) {
        issues.push({
          code: "sanitizer-unsafe-property",
          message: `The property "${declaration.propertyName}" on "${rule.selectorText}" is stripped by the MarkOut sanitizer.`,
          severity: "warning",
        });
      }
    }
  }

  return {
    issues,
    validRuleCount: parsedRules.length,
  };
}

function countOccurrences(value: string, token: string): number {
  return value.split(token).length - 1;
}
