import { defineConfig } from "tsdown"

const isDev = process.env.NODE_ENV !== "production"

// This package ships CLI/MCP binaries, not a single library entry. The `bin`,
// `main`, and `files` fields plus CI all reference `dist/*.js`, so every entry
// builds to `dist/` in both dev and prod (no lib/dist split).
export default defineConfig({
  entry: [
    "src/todo-index.ts",
    "src/cli.ts",
    "src/create-mcp-config.ts",
    "src/auth-server.ts",
    "src/setup.ts",
    "src/token-manager.ts",
  ],
  outDir: "dist",
  format: ["esm"],
  platform: "node",
  // Emit `.js` (valid ESM under "type": "module") so the `bin`/`main` paths and
  // source `./*.js` imports keep resolving — tsdown would otherwise emit `.mjs`.
  outExtensions: () => ({ js: ".js" }),
  dts: true,
  clean: true,
  splitting: false,
  sourcemap: isDev,
  minify: !isDev,
})
