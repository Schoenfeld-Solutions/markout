module.exports = {
  collectCoverageFrom: ["src/**/*.{ts,tsx}", "!src/**/*.d.ts"],
  coverageThreshold: {
    global: {
      // Jest applies "global" after subtracting files with explicit thresholds.
      // Keep this as a ratcheting floor for the remaining repo surface.
      branches: 54,
      functions: 65,
      lines: 66,
      statements: 66,
    },
    "./src/lib/body-accessor.ts": {
      branches: 85,
      functions: 100,
      lines: 95,
      statements: 95,
    },
    "./src/lib/compose-markdown.ts": {
      branches: 82,
      functions: 92,
      lines: 93,
      statements: 93,
    },
    "./src/lib/render-state-store.ts": {
      branches: 73,
      functions: 100,
      lines: 90,
      statements: 90,
    },
    "./src/lib/runtime.ts": {
      branches: 88,
      functions: 100,
      lines: 94,
      statements: 94,
    },
    "./src/taskpane/panels.tsx": {
      branches: 47,
      functions: 68,
      lines: 70,
      statements: 70,
    },
    "./src/taskpane/preferences.ts": {
      branches: 100,
      functions: 100,
      lines: 100,
      statements: 100,
    },
  },
  maxWorkers: 1,
  preset: "ts-jest",
  roots: ["<rootDir>/spec"],
  testEnvironment: "node",
  testMatch: ["**/*.spec.ts", "**/*.spec.tsx"],
};
