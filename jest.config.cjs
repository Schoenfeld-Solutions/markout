module.exports = {
  collectCoverageFrom: ["src/**/*.{ts,tsx}", "!src/**/*.d.ts"],
  maxWorkers: 1,
  preset: "ts-jest",
  roots: ["<rootDir>/spec"],
  testEnvironment: "node",
  testMatch: ["**/*.spec.ts", "**/*.spec.tsx"],
};
