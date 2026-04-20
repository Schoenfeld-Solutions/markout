module.exports = {
  collectCoverageFrom: ["src/**/*.ts", "!src/**/*.d.ts"],
  maxWorkers: 1,
  preset: "ts-jest",
  roots: ["<rootDir>/spec"],
  testEnvironment: "node",
  testMatch: ["**/*.spec.ts"],
};
