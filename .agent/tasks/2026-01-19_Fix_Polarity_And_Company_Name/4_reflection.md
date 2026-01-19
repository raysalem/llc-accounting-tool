# Reflection: Fix Polarity and Company Name

## Review
- **Prompt**: "logic polarity wrong (no abs)", "business name is setuptag".
- **Workflow**: Plan -> Implement -> Verify.
- **Rules**:
    - **Execution**: Verified.
    - **Constraints**: Followed user preference for negative values.

## Quality
- **Robustness**: The key normalization (`replace(/[^a-z0-9]/g, '')`) ensures `Company Name` and `Business Name` are treated identically, solving the Payer Info issue permanently.
- **Correctness**: The robust 1099 logic now correctly handles checking polarity (for threshold) while preserving sign (for reporting).

## Next
- Commit.
