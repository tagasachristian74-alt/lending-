# lending-

Loan-book management system for **ONE US MONEY** (Gloria, Oriental Mindoro), built as an
Excel workbook. Borrower register, loan origination, collections, expenses and daily cash
reconciliation, all driven by ID keys.

| File | What it is |
|---|---|
| `Lending_Business_System_V9.xlsx` | **Current working file.** Use this. |
| `Lending_Business_System_V8 recreated.xlsx` | Previous version, kept unchanged as a backup. |
| `build_v9.py` | The script that generated V9 from V8. Re-run with `python3 build_v9.py`. |

---

## Sheet map

| Sheet | Purpose |
|---|---|
| `Start Here` | Instructions, colour legend, flag reference, suggested next IDs. |
| `Dashboard` | Management summary — KPIs, portfolio health, cash activity, two charts. |
| `Master List` | Borrower register. Key: **Borrower ID** (`BR-001`). |
| `Loan Release` | Loan origination. Key: **Loan Reference** (`LN-001`). |
| `Collection` | One row per payment, linked to a Loan Reference. |
| `Expenses` | One row per expense. |
| `Daily Cash Report` | One row per business day. Cash reconciliation. |
| `Control Center` | Every data-quality flag counted in one place. |
| `Lists` | Source for every dropdown. Add names and categories here. |
| `COLLECTION SHEET` | Printable daily collection form (standalone). |
| `Daily Cash Report Format` | Printable daily cash report (standalone). |

## Data flow

```
Master List (Borrower ID)
      └── Loan Release (Borrower ID → name)          key: Loan Reference
                └── Collection (Loan Reference → borrower)
Expenses ──────────────┐
Loan Release ──────────┼── Daily Cash Report   (Cash channel only)
Collection ────────────┘
                              └── Dashboard / Control Center
```

Only **Cash**-channel transactions reach the Daily Cash Report. GCash, Bank Transfer and
Cheque are deliberately excluded from cash on hand.

## Colour code

| | Meaning |
|---|---|
| White cell, blue text | You type here. The only editable cells. |
| Grey cell | Calculated. Do not type. |
| Green cell | Looked up from another sheet. Do not type. |
| Red flag | Duplicate ID, broken link or cash mismatch. |
| Yellow flag | Row is incomplete. |
| Green flag | Row passed its checks. |

## Loan maths

- Interest is **flat on original principal**: `Principal × Rate × ROUNDUP(Term / period)`,
  where period is 1 / 7 / 30 / 365 days for Daily / Weekly / Monthly / Yearly.
- Every started interest period counts as a full period.
- `Total Receivable = Principal + Interest`. The processing fee is **not** added to the
  receivable — it is income collected on release.
- `Installment = Total Receivable ÷ ROUNDUP(Term / repayment period)`. Repayment frequency
  is chosen separately from the interest basis.
- `Maturity = First Due Date + Term − 1`.
- `Cash: Net = Collections + Processing Fees − Principal Released − Expenses`;
  `Ending Cash = Beginning Cash + Net`.

## Rules

- Never delete or rename an ID once it has been used.
- Never overwrite an old transaction — add a new row.
- Dropdowns are enforced. A value not on `Lists` is rejected on entry.
- `Control Center` should read 0 across the board before closing the books.

---

## What changed from V8 to V9

### Bugs fixed

| # | V8 problem | V9 |
|---|---|---|
| 1 | `Loan Release`!A8 header literally read `"Column1"`, in both the cell *and* the table definition | Renamed to **Loan Reference** in both places |
| 2 | The loan key cell was blank, so `Outstanding` evaluated to `#VALUE!`, which cascaded to `Master List` and `Dashboard` as **0** instead of ₱4,450 | Loan `LN-001` assigned; `Outstanding` is now error-proof against text and returns **₱4,450** |
| 3 | The ₱650 payment had no Loan Reference, so it was linked to nothing | Linked to `LN-001`; loan balance now updates |
| 4 | `Dashboard` "Outstanding Receivable" pointed at a **text label**, displaying the words `Outstanding (PHP)` | Points at the value |
| 5 | `Dashboard` "Processing Fees" and "Active Loans" both pointed at an **empty cell** → 0, understating cash in by ₱500 and skewing Expense-to-Inflow to 2.31 | Point at real totals; ratio is now **1.30** |
| 6 | `Master List` columns *Credit Limit, Total Released, Total Outstanding, Notes* had no formulas at all — dead columns | Total Released / Total Outstanding now calculate per borrower; Credit Limit is a live input |
| 7 | A stray `ACTIVE` sat on an empty row, making "Active Borrowers" read 1 for the wrong reason | Removed; `Status` is now derived per borrower |
| 8 | `Collection` "Total Collected" summed only principal + interest, **silently dropping Penalty/Other** | Penalty is now included, with its own summary total |
| 9 | `Collection` column A had no header at all | Now **Collection ID**, auto-generated |
| 10 | Sheet titles said V8 while `Control Center` and `Dashboard` said V4 | Consistently V9 |

### Added

- **30 input validations** across the five data sheets — dropdowns for Borrower ID, Loan
  Reference, Interest Basis, Repayment Frequency, Channel, Category, Collector, Officer and
  Approver, plus range checks on every amount and every date. Bad input is rejected on
  entry with an explanatory message. V8 had **zero** validations despite `Lists` claiming to
  power the dropdowns.
- **`Start Here` sheet** — workflow, colour legend, suggested next IDs, and a table of what
  each flag means and what to do about it.
- **`Days Overdue` column** on Loan Release, with a heat-map scale for collection follow-up.
- Auto-generated **Collection ID** and **Expense ID** for the audit trail.
- Second dashboard chart (cash in vs cash out) and a rebuilt line chart.
- New `Control Center` checks: Loans Over 30 Days Past Due, and a total that no longer
  double-counts.
- Frozen header panes, filter dropdowns, tab colours, consistent currency and date formats,
  gridlines off on presentation sheets.
- `Start Here`, being reference-only, is sheet-protected (no password — Review ▸ Unprotect
  Sheet to lift it).

### Your data

The single V8 transaction set was carried over exactly as recorded. Two keys that did not
exist in V8 had to be created for the relational chain to work at all: Loan Reference
`LN-001` and the link from the payment to it. **No other value was invented.** Three fields
that were blank in V8 are still blank, so three flags are still showing:

| Flag | Where | To clear it |
|---|---|---|
| INCOMPLETE | `Collection` row 9 | Pick a **Channel** (the payment is not in the cash report until you do) |
| INCOMPLETE | `Expenses` row 9 | Enter **Paid To** for the ₱1,500 passbook expense |
| ENTER ACTUAL CASH | `Daily Cash Report` row 9 | Count the till and enter **Actual Cash** |

The loan also correctly reads **PAST DUE** — maturity was 06-May-2026, 188 days ago.

Delete row 9 on each transaction sheet to start from an empty book.

---

## Verification

`build_v9.py` refuses to save a workbook that fails any of these gates:

- a formula with unbalanced parentheses or unbalanced quotes;
- an Excel table whose column names do not match its header row (that mismatch is what
  produced `Column1` in V8);
- a dropdown range containing a blank cell, or a list that lost its header.

V9 was then recalculated independently with the `formulas` engine:

- **35/35** expected values matched — loan maths, balances, lookups, auto-generated IDs,
  cash reconciliation, dashboard KPIs and control counts.
- **0** Excel errors (`#VALUE!`, `#REF!`, `#DIV/0!`) across all 37,527 computed cells.
- 18,646 formulas parse; all parentheses balanced.
- All 5 table definitions match their headers; 12 defined names resolve; all 10 dropdown
  lists fully populated; 30 validations present.
- Both print forms preserved (54 and 27 cells, merges intact); no columns left hidden.

Reproduce with:

```bash
pip install openpyxl formulas
python3 build_v9.py
```
