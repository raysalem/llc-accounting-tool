# Code Quality (The Output Rules)

## Modern JS Only
- Use `const`, `let` (no `var`).
- Arrow functions for callbacks.
- Modules (ESM or CommonJS depending on project config).
- `async`/`await` for asynchronous operations.
- Promises where applicable.

## Clear, Consistent Naming
- **Variables/Functions**: Descriptive `camelCase` (e.g., `calculateTotal`, `vendorList`).
- **Classes**: `PascalCase` (e.g., `TransactionProcessor`).
- Avoid single-letter variables except in standard loop contexts (`i`, `j`).

## Modularity & Single Responsibility
- Request small, focused functions and components.
- Break large functions into helper methods.

## Strict Typing/Linting
- Adhere to ESLint rules.
- Check types defensively (e.g., `typeof val === 'number'`).

## Error Handling
- explicit `try...catch` blocks for all I/O and external calls.
- `.catch()` handling for Promises.
- Throw descriptive Error objects, do not just return null silently for critical failures.

# Verification (The Safety Net)

## Test Everything
- LLMs hallucinate dependencies and logic; run unit tests.
- Verify outputs against expected values.

## Debug & Inspect
- Step through code to find phantom variables or incorrect logic.
- Use `console.log` or debug flags to trace execution paths.

## Reduce Dependencies
- Ask for vanilla JS solutions where possible.
- Only use trusted, popular libraries (e.g., `exceljs` is approved).
