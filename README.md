# BHF purchase order parser

Python take-home (optional) for the Software Engineer, Full Stack & Automation role at BHF.

`parse_po.py` reads a vendor purchase order printed as fixed-width text and writes one spreadsheet row per SKU,
in the column layout of `output_template_blank.xlsx`. It also checks the arithmetic on every line.

- Input: [`data/purchase_order_sample.txt`](data/purchase_order_sample.txt)
- Output (generated, committed): [`output/purchase_order_sample.xlsx`](output/purchase_order_sample.xlsx)

## Setup and run

Requires Python 3.11+.

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

python parse_po.py data/purchase_order_sample.txt   # writes output/purchase_order_sample.xlsx
pytest -q                                           # 7 tests
```

Output on the sample:

```
Wrote 12 rows to output/purchase_order_sample.xlsx
```

Any mismatches are printed under that line, and they are also written into the spreadsheet.

## Approach

The PO is a printed report, so the script reads it line by line with a small number of regular expressions.
There is no general-purpose parser.

1. **Header fields** (`parse_header`): each label gets one regex, e.g. `SHIP DATE:\s*(\S+)`.
   `PO#` needs a negative lookbehind so it doesn't match `REF MASTER PO#`.
   The VENDOR block is read from its title down to the next blank line.
   SHIP TO and BILL TO are printed side by side. I split each line at the column where the `BILL TO` heading starts.
2. **Line items** (`parse_items`): one regex matches an item row. The free-text description is matched lazily,
   and the numeric columns are anchored to the end of the line, so descriptions containing spaces or `/` work.
   The **UPC is on the line directly under each item**. When an item row matches, the next line must be a
   12-14 digit UPC. If it isn't, the script raises an error with the line number instead of writing a row with a blank UPC.
3. **Validation** (`validate_item`, `validate_totals`): values are parsed as `Decimal`, so the comparisons
   are exact and not affected by float rounding.
   - `EXT QTY` is checked against `CTNS × CSPK`.
   - `EXT COST` is checked against `COST × EXT QTY`.
   - A mismatched cell is filled red and gets an Excel comment with the expected value, e.g. `Expected 136 (CTNS x CSPK)`.
     The same message is printed to the console with the SKU.
   - The script also sums the items and compares them with the printed `TOTALS:` line (ext cost, cartons, ext qty, kilograms).
     Problems found there are printed to the console.
4. **Output** (`write_xlsx`): the script loads the blank template and appends rows to it, so the sheet has
   exactly the template's 23 columns and nothing else.

On the sample, all 12 rows and the TOTALS line pass validation, so nothing is highlighted. Two tests change one
printed quantity to show the flag works: `test_mismatch_is_flagged` checks the messages, and
`test_mismatch_is_highlighted_in_output` checks the fill and comment in the written spreadsheet.

## Assumptions

- **Total Cost** is the PO total (`Total Invoice Value: 447,455.12`), repeated on every row like the other header fields.
  Per-line cost is already in `EXT Cost`.
- **VENDOR, SHIP TO, BILL TO** are each one cell, with the address lines joined by `", "`.
  Leading codes are kept (`50201`, `331`). PHONE and FAX lines are dropped from all three blocks.
- **BUYER** is kept as printed, including the buyer number (`610 JORDAN MILLER`).
- **SHIP DATE** is kept as printed (`5/02/26`) and not converted to an Excel date.
  The year is ambiguous (two digits), and the downstream format isn't specified.
- **SKU, UPC, Vendor Part #, Dept** are stored as text so leading zeros (`04212`, `0091120032129`) survive.
  Prices, quantities, cube and kilograms are stored as numbers.
- **FILENAME** is the input file's name.
- **Validation flags** are cell formatting (fill and comment), not extra columns. The brief says the output must
  match it exactly, and a downstream import would likely reject unexpected columns.
- Each check uses the _printed_ values of the other columns: EXT COST is checked against COST × printed EXT QTY.
  This way a typo in one column flags only that column, not both.
- Exact equality is correct for this document: costs have 3 decimals and quantities are integers,
  so `COST × EXT QTY` is exact at 3 decimals.
- The document is a single page. The page header and footer are not repeated.

## What I'd change with more time

- **Multi-page POs**: skip repeated page headers and footers and continue items across pages.
  I'd want a real multi-page sample before guessing at the layout.
- **More samples** from other vendors. The regexes are tuned to this one layout. A second vendor would show
  which parts are actually fixed and which vary.
- **Rounding tolerance** (e.g. to the cent) if other vendors print rounded ext costs.
- **Header-level checks**: compare the TOTALS line with `Total Invoice Value`, and check UPC check digits.
- **Batch mode**: process a folder of POs into one sheet (the FILENAME column already supports this).

## AI assistance

I used Claude Code (Anthropic) for this project. Its instruction file is [`AGENTS.md`](AGENTS.md), committed as-is.
I read the brief and the data first and decided the plan and the output assumptions above, including the layout,
the Total Cost interpretation, how addresses are joined, and where the validation flags go. Claude wrote the first
draft of the code, tests and README. I reviewed every line, ran the tests and checked the output spreadsheet by hand.

## Time spent

About 2 hours.
