# Developer Quality Checklist

## Code Quality Standards
- [ ] **No "Toy" Scopes**: Critical variables must be explicitly defined in local scope or passed as arguments. Do not rely on valid-state from 500 lines above.
- [ ] **No "Slop" Logs**: Remove all `console.log` debugging lines before "finishing" a task. Use the built-in `showDebug` flag buffer if tracing is needed.
- [ ] **Strict Typing**: When parsing amounts, immediately handle `NaN`, `null`, and `undefined` safely.

## Testing Standards
- [ ] **Golden Master Test**: Use `tests/Full_Accounting_Test_Case.xlsx` (or create one) for regression testing instead of generating weak file structures via script.
- [ ] **Parity Check**: If you add a column/calculation to Excel output, YOU MUST add it to PDF output.
- [ ] **Path Safety**: Always handle UNC paths (`\\server\share`) and Spaces in paths (`"My Documents"`).

## Architecture
- [ ] **Setup Parser**: Ensure Setup sheet changes preserve "Table" structure or valid header row detection (7+ keywords wins).
- [ ] **Asset Polarity**: Asset categories generally behave inversely to Expenses regarding Bank polarity (Spending = Increase). Ensure new features respect this.
