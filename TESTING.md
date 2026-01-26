# Testing Guide

This document outlines how to verify the correctness of the Accounting Tool's logic, particularly for complex scenarios like Transfer double-counting and Liability polarity.

## Test Case 1: Credit Card Double Counting
**Scenario:** A payment is sent from the Bank to pay off a Credit Card.
*   **Bank Sheet:** Shows a negative transaction (e.g., -100). Category: `AX CC` (Liability).
*   **CC Sheet:** Shows a positive transaction (e.g., +100). Category: `Transfer AX CC` (Transfer).

**Expected Behavior:**
1.  **Bank Balance:** Decreases by 100. (Correct).
2.  **Liability Balance (AX CC):**
    *   Decreases by 100 due to Bank Payment.
    *   **IGNORES** the CC side transaction (because it is a Transfer).
    *   **Net Effect:** Liability reduces by 100.
    *   *If failed:* Liability reduces by 200 (Double Counted).

**Verification:**
1.  Run `node report.js ... --debug`.
2.  Check the PDF or Console output for "AX CC" Liability.
3.  Ensure "Subtractions" column matches the **Bank Payment Only** (+ Refunds).
4.  Ensure "Ending Balance" reflects only one payment.

## Test Case 2: Refunds vs. Payments
**Scenario:** A refund for a purchase vs. a payment for the bill.
*   **Refund ($50)**: Return of goods.
    *   **Category:** `Office Supplies` (Expense).
    *   **Action:** Reduces Expense sum. Reduces Liability (Debt goes down).
    *   **Result:** Should be INCLUDED in the report.
*   **Payment ($100)**: Money transfer.
    *   **Category:** `Transfer AX CC`.
    *   **Action:** No effect on P&L. No double-effect on Liability.
    *   **Result:** Should be EXCLUDED from Linkage logic.

**How to verify:**
*   Check the "Subtractions" column in the Liability Detail.
*   It should equal: `Total Bank Payments` + `Total Refunds`.
*   It should NOT include `Total CC Payments Received`.

## Test Case 3: Transfer Account Validation
**Scenario:** Using a Transfer category on the wrong sheet.
*   Setup: `Transfer AX CC` maps to `Transfer Account: AX CC`.
*   Action: Use `Transfer AX CC` on the `Bank` sheet.
*   **Expected:** The tool should issue a CRITICAL WARNING in the console, because the Bank sheet is not the `AX CC` account.

