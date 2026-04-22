const js = require("@eslint/js");
const eslintConfigPrettier = require("eslint-config-prettier");
const tsParser = require("@typescript-eslint/parser");
const tsPlugin = require("@typescript-eslint/eslint-plugin");

const browserGlobals = {
  DOMParser: "readonly",
  FileReader: "readonly",
  Office: "readonly",
  clearTimeout: "readonly",
  console: "readonly",
  document: "readonly",
  localStorage: "readonly",
  matchMedia: "readonly",
  navigator: "readonly",
  self: "readonly",
  setTimeout: "readonly",
  window: "readonly",
};

const testGlobals = {
  afterEach: "readonly",
  beforeEach: "readonly",
  __dirname: "readonly",
  describe: "readonly",
  expect: "readonly",
  it: "readonly",
  jest: "readonly",
};

module.exports = [
  {
    ignores: [
      "dist/**",
      "coverage/**",
      "node_modules/**",
      "assets/**",
      "manifest*.xml",
      "eslint.config.cjs",
      "jest.config.cjs",
      "src/shims/**/*.js",
      "webpack.config.js",
    ],
  },
  js.configs.recommended,
  {
    files: ["**/*.ts", "**/*.tsx"],
    languageOptions: {
      parser: tsParser,
      parserOptions: {
        ecmaVersion: "latest",
        project: "./tsconfig.json",
        sourceType: "module",
        tsconfigRootDir: __dirname,
      },
      globals: browserGlobals,
    },
    plugins: {
      "@typescript-eslint": tsPlugin,
    },
    rules: {
      ...tsPlugin.configs["recommended-type-checked"].rules,
      ...tsPlugin.configs["stylistic-type-checked"].rules,
      "no-undef": "off",
      "@typescript-eslint/consistent-type-imports": "error",
      "@typescript-eslint/no-confusing-void-expression": [
        "error",
        { ignoreArrowShorthand: true },
      ],
      "@typescript-eslint/no-explicit-any": "error",
      "@typescript-eslint/no-floating-promises": "error",
      "@typescript-eslint/no-misused-promises": [
        "error",
        { checksVoidReturn: false },
      ],
      "@typescript-eslint/no-unnecessary-condition": "error",
    },
  },
  {
    files: ["spec/**/*.ts", "spec/**/*.tsx"],
    languageOptions: {
      globals: testGlobals,
    },
    rules: {
      "@typescript-eslint/no-unsafe-assignment": "off",
      "@typescript-eslint/no-unsafe-call": "off",
      "@typescript-eslint/no-unsafe-member-access": "off",
      "@typescript-eslint/unbound-method": "off",
    },
  },
  eslintConfigPrettier,
];
