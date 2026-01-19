# Implementation: 1099 CSV Generation

## Details
1.  **Split Columns**: `report.js` now looks for `1099 type` and `1099 required` headers in Setup, prioritizing them over the single `1099` column.
    - Logic respects user input: If `required` is "No", it ignores the vendor. If `type` is present, it uses it. Default is "NEC".
2.  **CSV Output**:
    - Creates `[BusinessName]-1099.csv` in the same directory as the accounting file.
    - Flattens Payer Info (from Setup vertical table) and Recipient Info (from external `vendor.csv`/Setup) into a single row per vendor.
    - Header includes Payer/Recipient/Amount/Type fields.
3.  **Validation**:
    - Logic for "only positive amounts" and "thresholds" checks are applied before adding to CSV rows.

## Robustness
- **Filenames**: Sanitizes the business name to ensure safe filenames (`safeName = payerName.replace...`).
- **Flexible Keys**: Reads Payer Info using loose key matching but maps to standard CSV headers (Payer Name, Payer TIN, etc.).
- **Fallback**: Defaults to "Unknown_Payer-1099.csv" if no company info is found.

## Considerations
- **Override**: If the CSV file already exists, it is overwritten (standard behavior for generated reports).
