import { defineConfig } from "vitest/config"

export default defineConfig({
  test: {
    globals: true,
    environment: "node",
    // Specs chdir into temp directories, which worker threads do not allow.
    pool: "forks",
    // Blocks the real network and moves HOME before any spec file is imported.
    setupFiles: ["./src/test-support/setup.ts"],
    // Tool output formats dates with toLocaleString()/toLocaleDateString(); pin the zone and
    // the locale so specs read the same on every machine and in CI.
    env: { TZ: "UTC", LC_ALL: "en_US.UTF-8", LANG: "en_US.UTF-8" },
    coverage: {
      provider: "v8",
      reporter: ["text", "json", "html"],
      include: ["src/**/*.ts"],
      exclude: ["src/**/*.spec.ts", "src/**/*.test.ts", "src/test-support/**", "**/*.d.ts", "**/node_modules/**"],
    },
  },
})
