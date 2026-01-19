# Found It: Transaction Loop Crash

## The Blind Spot
I was focusing on the **Ledger Loop** (Lines 750+) because I was looking there earlier.
**BUT** the stack trace `report.js:555` is the **Transaction Loop**.
There is **ANOTHER** 1099 check inside the Transaction Loop (Lines 630-632 in view above).

```javascript
630:                     const type = vendor1099Map.get(vLower);
631:                     vendor1099Stats[type][displayVendor] = (vendor1099Stats[type][displayVendor] || 0) + amount;
```

## The Bug
- `vendor1099Map.get(vLower)` returns an **OBJECT** `{ type: 'NEC', req: '' }`.
- Line 631 tries to use this object as the key for `vendor1099Stats`:
    - `vendor1099Stats[object]` -> `vendor1099Stats["[object Object]"]`.
- `vendor1099Stats` is `{ NEC: {}, INT: {} }`.
- "object Object" key is undefined.
- So `vendor1099Stats["[object Object]"]` is `undefined`.
- Then `undefined[displayVendor]` throws **TypeError: Cannot set properties of undefined**.
- **THIS IS THE CRASH.**

## Solution
- I must applying the logic update (checking `is1099.type`) to **BOTH** loops (Ledger Loop AND Transaction Loop).
- I missed the Transaction Loop in my previous multi-replace.

## Plan
1.  Update Transaction Loop (Line ~630) to handle the object structure safely.
2.  (Optional) Remove `try/catch` from Ledger Loop if desired, but nice to keep.
3.  Re-enable debug logging if needed (but probably not needed now that cause is found).
