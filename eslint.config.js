import js from "@eslint/js"
import prettierRecommended from "eslint-plugin-prettier/recommended"
import simpleImportSort from "eslint-plugin-simple-import-sort"
import globals from "globals"
import tseslint from "typescript-eslint"

export default [
  {
    ignores: ["dist/**", "lib/**", "node_modules/**", "coverage/**", "test-api-exploration.js"],
  },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  prettierRecommended,
  {
    plugins: {
      "simple-import-sort": simpleImportSort,
    },
    languageOptions: {
      globals: {
        ...globals.node,
        ...globals.es2021,
      },
      ecmaVersion: 2020,
      sourceType: "module",
    },
    rules: {
      "prettier/prettier": ["error", {}, { usePrettierrc: true }],
      "@typescript-eslint/no-unused-vars": "off",
      "@typescript-eslint/explicit-function-return-type": "off",
      // Graph API responses are loosely typed; allow `any` as a warning rather than blocking.
      "@typescript-eslint/no-explicit-any": "warn",
      "simple-import-sort/imports": "error",
      "simple-import-sort/exports": "error",
    },
  },
]
