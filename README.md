# GST Reconciliation Studio

A clean, reconciliation-only Windows/Streamlit application for matching GSTR-2B
against the inward purchase register.

## Package contents

- `app.py` — reconciliation interface
- `reconciliation.py` — reconciliation + checking engine
- `sample_data/Sample_GSTR_2B.xlsx` — ready-to-test sample
- `sample_data/Sample_Inward_Register.xlsx` — ready-to-test sample
- `sample_data/EXPECTED_SAMPLE_RESULTS.txt` — what the samples should produce
- `requirements.txt` — required Python packages
- `run_windows.bat` — Windows launcher

## Windows

1. Extract the ZIP.
2. Double-click `run_windows.bat`.
3. The browser opens Streamlit.
4. Use the included sample files to test.
5. Replace them with your actual GSTR-2B and purchase register.

## Columns

Recommended columns on both files:

GSTIN, **Supplier Name**, Invoice No, Invoice Date, 2B Month / Return Period,
Taxable Value, IGST, CGST, SGST, Cess, Invoice Value.

Header names can vary — common aliases are recognized automatically
(`Supplier GSTIN`, `Trade Name`, `Party Name`, `Document No`, `Return Period`,
`Taxable Amount`, and so on).

**Supplier name is optional.** Where a line has no supplier name, it is filled in
from the GSTIN using a supplier directory built from whichever file does carry
names (normally GSTR-2B). Every row records whether the name came from the file,
was derived from the GSTIN, or is still unknown.

## Matching, strongest first

1. Exact GSTIN + invoice number.
2. Same GSTIN + same invoice serial, ignoring format
   (`INV/1006/26-27` matches `1006`).
3. Same PAN + exact invoice number — catches a purchase booked against the
   supplier's registration in another state.
4. Same PAN + same invoice serial.
5. Split invoices: several register lines summed to one 2B invoice, **and**
   several 2B lines summed to one register line (rate-wise splits).
6. Weighted fuzzy scoring over invoice number, invoice serial, GSTIN, PAN,
   supplier name, tax values and invoice date. A pair is only accepted with a
   score of 55+ **and** at least two independent anchors, so stronger matching
   never becomes loose guessing.
7. Near-miss suggestions: anything still unmatched is scored against the other
   side's leftovers and the three closest candidates are reported *without*
   being auto-matched, so a reviewer gets a starting point rather than a dead
   end.

Each tier is resolved across the whole file before the next one runs, and fuzzy
pairs are assigned in descending score order, so a strong match can never be
taken by a weaker one that simply appeared earlier in the file.

## Statuses

`Matched`, `Value Mismatch`, `GSTIN Mismatch`, `Same PAN - Different GSTIN`,
`Tax Head Mismatch (IGST vs CGST/SGST)`, `Likely Match - Invoice Format`,
`Likely Match - Register Invoice Error`, `In 2B Only`, `In Register Only`.

Every matched row also carries a match basis, a confidence score, the value
difference, a date warning and a supplier-name warning.

## Analysis on top of the match

Every reconciled line also carries:

- **ITC action** — claim, hold, correct the register, book the invoice, or chase
  the supplier — plus the **ITC amount at risk** on that line.
- **Difference diagnosis** — not just the delta but why: taxable value differs,
  tax rate differs, same total tax under different heads, cess only, a possible
  digit transposition (difference divisible by 9), or a rounding difference that
  a higher tolerance would absorb.
- **Rate comparison** — implied GST rate on each side, with a warning when they
  disagree.
- **Cross-period filing** — the return period the supplier filed in against the
  period the ITC was booked in, with the gap in months.
- **Confidence band** — auto / good / low / unmatched, so high-confidence lines
  can be signed off in bulk.
- **Aging** — days outstanding and a bucket for every open item.

Report-level analysis:

- **ITC Statement** — ITC as per books vs GSTR-2B, head by head, split into
  matched, under review, not in 2B and not in books.
- **Action Plan** — supplier-wise follow-up list, highest ITC exposure first.
- **Supplier Risk** — a 0-100 risk score per supplier from match rate, ITC
  exposure and invoices missing from 2B, banded Low / Medium / High / Critical.
- **Aging Analysis** — open exposure by age bucket and status.
- **Suggestions** — near-miss candidates for unmatched items.

## Built-in checks

Run on both files, listed on the `Data Quality Checks` sheet by severity:

- GSTIN missing, malformed, or failing the check-digit test
- Duplicate invoice (same GSTIN + invoice number + same values) — double ITC risk
- Repeated invoice number with different values — confirm rate-wise split
- IGST together with CGST/SGST on one line
- CGST not equal to SGST
- Taxable + tax + cess not equal to the invoice value
- Implied tax rate that is not a notified GST rate
- Missing, unreadable or future invoice dates
- Section 16(4) ITC time limit likely expired
- Negative values (credit / debit notes)
- Supplier name missing

## Excel report

Summary, ITC Statement, Action Plan, Supplier Summary (with risk score), Review
Queue, In Register Only, In 2B Only, Suggestions, Aging Analysis, Matched
Details, Monthly Analysis, Data Quality Checks, Supplier Directory, both
uploaded files as received, and the full reconciliation — with
filters, frozen headers, money formatting, exception highlighting and clickable
links back to the matching GSTR-2B row.
