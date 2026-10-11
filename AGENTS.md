# AGENTS.md

Context for AI coding assistants working in this repo.

## What this is

A small Python script that turns one vendor purchase order (semi-structured text) into a spreadsheet.
It is an optional take-home assessment (2-3 hour budget), so the goal is a small, clear script the author
can explain line by line. Prefer the simplest thing that works. Do not build a generic PO-parsing framework.

Required behaviour:
- Read `data/purchase_order_sample.txt`.
- Extract header fields (buyer, ship terms, PO#, ref master PO#, ship date, vendor, ship to, bill to) and
  every line item, including the UPC that sits on the line under each item row.
- Write one row per SKU into a copy of `data/output_template_blank.xlsx`, header fields repeated on every row.
- Validate: recompute EXT QTY (CTNS x CSPK) and EXT COST (COST x EXT QTY) and flag mismatches.

## Stack (do not change)

- Python 3.11+, standard library plus `openpyxl`. Tests use `pytest`.
- No pandas, no classes unless clearly needed, no new dependencies without asking first.

## Conventions

- One module, `parse_po.py`, built from plain functions: parse header, parse items, validate, write xlsx, main.
- Use `Decimal` for money and measurements so validation comparisons are exact.
- Keep SKU, UPC and vendor part # as strings (leading zeros matter).
- The output sheet must match the template exactly: same 23 columns, no extra columns or sheets.
  Flag validation mismatches by highlighting the cell (red fill + comment with the expected value).
- Fail loudly (raise `ValueError` with the line number) when the document doesn't look as expected,
  rather than silently writing a wrong row.

## Layout

```
parse_po.py          parsing, validation, xlsx output, CLI
test_parse_po.py     pytest tests against the sample PO
data/                inputs (sample PO, blank template)
output/              generated spreadsheet (committed)
```

## Workflow

- Run: `python parse_po.py data/purchase_order_sample.txt`
- Test: `pytest -q`
- Do not run git commits; the author commits.
- Record assumptions in README.md as you make them.
