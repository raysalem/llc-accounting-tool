# Reflection: Refine Polarity Logic

## Review
- **Prompt**:
    1.  Resolve polarity issue (don't blindly abs). -> **Done** (Conditional check `val < 0`).
    2.  Run command and confirm. -> **Done** (Executed `report.js` with args).
- **Workflow**: 4 MDs created.
- **Rules**:
    - **Followed Prompt**: Yes.
    - **JS**: Yes.
    - **Robustness**: Handled positive/negative logic explicitly.

## Quality
- **Impact**: Prevents financial reporting errors (reporting refunds as expenses).

## Next
- Commit.
