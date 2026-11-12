import { defineConfig } from "vitest/config"

export default defineConfig({
  test: {
    globals: true,
    environment: "node",
    // Tool output formats dates with toLocaleDateString(); pin the zone so specs read the
    // same on every machine and in CI.
    env: { TZ: "UTC" },
    coverage: {
      provider: "v8",
      reporter: ["text", "json", "html"],
      include: ["src/**/*.ts"],
      exclude: ["src/**/*.spec.ts", "src/**/*.test.ts", "src/test-support/**", "**/*.d.ts", "**/node_modules/**"],
    },
  },
})
