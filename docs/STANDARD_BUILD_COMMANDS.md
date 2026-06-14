# Standard Build Commands for TypeScript/JavaScript Projects

## Overview

This document defines a standardized set of npm/pnpm scripts that should be used across all repositories for consistency and developer experience.

## Core Commands

### 1. **`check`** - Quick validation (no build)

```json
"check": "pnpm run format:check && pnpm run typecheck"
```

- Validates code formatting
- Checks TypeScript types
- Fast feedback loop for developers

### 2. **`check:fix`** - Auto-fix issues

```json
"check:fix": "pnpm run format && pnpm run typecheck"
```

- Automatically fixes formatting issues
- Still validates types (can't auto-fix)

### 3. **`ci`** - Complete CI validation

```json
"ci": "pnpm run check && pnpm run build && pnpm run test"
```

- Full validation suite for CI/CD pipelines
- Ensures code is ready for production

### 4. **`precommit`** - Pre-commit hook

```json
"precommit": "pnpm run check:fix && pnpm run build"
```

- Auto-fixes what it can
- Prevents bad commits

## Individual Commands

### Formatting

```json
"format": "prettier --write \"src/**/*.{ts,js,json}\" \"*.{json,md}\"",
"format:check": "prettier --check \"src/**/*.{ts,js,json}\" \"*.{json,md}\""
```

### Linting (if using ESLint)

```json
"lint": "eslint . --ext .ts,.tsx,.js,.jsx",
"lint:fix": "eslint . --ext .ts,.tsx,.js,.jsx --fix"
```

### Type Checking

```json
"typecheck": "tsc --noEmit"
```

### Building

```json
"build": "tsup",  // or your build tool
"build:watch": "tsup --watch"
```

### Testing

```json
"test": "vitest",
"test:watch": "vitest --watch",
"test:coverage": "vitest --coverage"
```

## Example Complete package.json Scripts Section

```json
{
  "scripts": {
    // Development
    "dev": "tsx watch src/index.ts",
    "build": "tsup",
    "build:watch": "tsup --watch",

    // Code Quality
    "format": "prettier --write \"src/**/*.{ts,js,json}\" \"*.{json,md}\"",
    "format:check": "prettier --check \"src/**/*.{ts,js,json}\" \"*.{json,md}\"",
    "lint": "eslint . --ext .ts,.tsx,.js,.jsx",
    "lint:fix": "eslint . --ext .ts,.tsx,.js,.jsx --fix",
    "typecheck": "tsc --noEmit",

    // Testing
    "test": "vitest",
    "test:watch": "vitest --watch",
    "test:coverage": "vitest --coverage",

    // Composite Commands
    "check": "pnpm run format:check && pnpm run lint && pnpm run typecheck",
    "check:fix": "pnpm run format && pnpm run lint:fix && pnpm run typecheck",
    "ci": "pnpm run check && pnpm run build && pnpm run test",
    "precommit": "pnpm run check:fix && pnpm run build",

    // Production
    "start": "node dist/index.js",
    "clean": "rm -rf dist coverage .turbo"
  }
}
```

## Usage Patterns

### For Developers

1. **During development**: Use `pnpm run dev` for hot reloading
2. **Before committing**: Run `pnpm run check:fix` to auto-fix issues
3. **To validate changes**: Run `pnpm run check` for quick validation

### For CI/CD

1. **GitHub Actions**: Use `pnpm run ci` in your workflow
2. **Pre-commit hooks**: Configure husky to run `pnpm run precommit`
3. **PR checks**: Require `pnpm run ci` to pass

### For Different Project Types

#### Library/Package

```json
{
  "scripts": {
    "prepublishOnly": "pnpm run ci",
    "release": "pnpm run ci && changeset publish"
  }
}
```

#### Application

```json
{
  "scripts": {
    "docker:build": "docker build -t app .",
    "deploy": "pnpm run ci && pnpm run docker:build && pnpm run docker:push"
  }
}
```

## Benefits

1. **Consistency**: Same commands work across all repos
2. **Discoverability**: New developers know what to run
3. **CI/CD Integration**: Standard `ci` command for all pipelines
4. **Progressive Enhancement**: Start with basics, add more as needed
5. **Tool Agnostic**: Works with any build tool (tsup, esbuild, webpack, etc.)

## Migration Guide

To migrate an existing project:

1. Add the core commands (`check`, `check:fix`, `ci`, `precommit`)
2. Ensure individual commands exist (format, typecheck, etc.)
3. Update CI/CD to use `pnpm run ci`
4. Add pre-commit hooks with husky:
   ```bash
   npx husky add .husky/pre-commit "pnpm run precommit"
   ```

## Tool Recommendations

- **Build**: tsup (fastest, zero-config)
- **Format**: Prettier
- **Lint**: ESLint with typescript-eslint
- **Test**: Vitest (fastest, Jest-compatible)
- **Type Check**: TypeScript strict mode
- **Package Manager**: pnpm (fastest, efficient)

## Notes

- Always include `typecheck` even if your build tool checks types (for speed)
- Use `--no-emit` for typecheck to avoid generating files
- Consider adding `clean` script to remove build artifacts
- Add tool-specific config files (.prettierrc, tsconfig.json, etc.)
