"""Parse a vendor purchase order (fixed-width text) into the BHF output spreadsheet.

Usage:
    python parse_po.py data/purchase_order_sample.txt [--out output/purchase_order_sample.xlsx]
"""

import argparse
import re
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill

TEMPLATE_PATH = Path(__file__).parent / "data" / "output_template_blank.xlsx"

# Mismatched cells are highlighted rather than flagged in extra columns, so the sheet keeps the template's exact columns.
MISMATCH_FILL = PatternFill(fill_type="solid", start_color="FFC7CE", end_color="FFC7CE")

# One line item row, e.g.
# 210 71097188      04212         NORDPEAK ASHER T CMFST GRY            149.99   49.400   6718.400    68     2      136  0.0991       7.72
# The description can contain spaces, so it is matched lazily and the numeric columns are anchored to the end.
ITEM_RE = re.compile(
    r"^(?P<dept>\d+)\s+(?P<sku>\d+)\s+(?P<vendor_part>\S+)\s+(?P<description>.+?)\s+"
    r"(?P<retail>[\d.]+)\s+(?P<cost>[\d.]+)\s+(?P<ext_cost>[\d.]+)\s+"
    r"(?P<ctns>\d+)\s+(?P<cspk>\d+)\s+(?P<ext_qty>\d+)\s+"
    r"(?P<cube>[\d.]+)\s+(?P<kilograms>[\d.]+)\s*$"
)

# The UPC sits alone, indented, on the line directly under each item row.
UPC_RE = re.compile(r"^\s+(\d{12,14})\s*$")

# TOTALS: <ext cost> <ctns> <ext qty> <kilograms>
TOTALS_RE = re.compile(r"^TOTALS:\s+([\d.]+)\s+(\d+)\s+(\d+)\s+([\d.]+)\s*$")


def find(pattern, text, field):
    """Return the first capture group of `pattern` in `text`, or raise if the field is missing."""
    match = re.search(pattern, text, re.MULTILINE)
    if not match:
        raise ValueError(f"Could not find {field} in purchase order")
    return match.group(1).strip()


def read_block(lines, title):
    """Return the lines of an address block that starts with a `title` line, a dashed underline,
    and runs until the next blank line."""
    start = next(i for i, line in enumerate(lines) if line.startswith(title))
    block = []
    for line in lines[start + 2:]:  # skip the title and the dashed underline
        if not line.strip():
            break
        block.append(line)
    return block


def join_address(lines):
    """Join address lines into one cell, dropping PHONE / FAX lines."""
    kept = [line.strip() for line in lines if line.strip() and not line.strip().startswith(("PHONE", "FAX"))]
    return ", ".join(kept)


def parse_header(text):
    lines = text.splitlines()

    # SHIP TO and BILL TO are printed side by side; split each line at the column where "BILL TO" starts.
    ship_bill_lines = read_block(lines, "SHIP TO")
    bill_to_column = next(line for line in lines if line.startswith("SHIP TO")).index("BILL TO")
    ship_to = [line[:bill_to_column] for line in ship_bill_lines]
    bill_to = [line[bill_to_column:] for line in ship_bill_lines]

    return {
        "buyer": find(r"BUYER:\s*(.+?)\s{2,}", text, "BUYER"),
        "ship_terms": find(r"SHIP TERMS:\s*(.+?)\s{2,}", text, "SHIP TERMS"),
        # The lookbehind stops "REF MASTER PO#:" from matching as the PO#.
        "po_number": find(r"(?<!MASTER )PO#:\s*(\S+)", text, "PO#"),
        "ref_master_po": find(r"REF MASTER PO#:\s*(\S+)", text, "REF MASTER PO#"),
        "ship_date": find(r"SHIP DATE:\s*(\S+)", text, "SHIP DATE"),
        "vendor": join_address(read_block(lines, "VENDOR")),
        "ship_to": join_address(ship_to),
        "bill_to": join_address(bill_to),
        "total_cost": Decimal(find(r"Total Invoice Value:\s*([\d,.]+)", text, "Total Invoice Value").replace(",", "")),
    }


def parse_items(text):
    lines = text.splitlines()
    items = []
    for i, line in enumerate(lines):
        match = ITEM_RE.match(line)
        if not match:
            continue

        next_line = lines[i + 1] if i + 1 < len(lines) else ""
        upc_match = UPC_RE.match(next_line)
        if not upc_match:
            raise ValueError(f"Line {i + 2}: expected a UPC under SKU {match['sku']}, got {next_line!r}")

        items.append({
            "dept": match["dept"],
            "sku": match["sku"],
            "upc": upc_match.group(1),
            "vendor_part": match["vendor_part"],
            "description": match["description"].strip(),
            "retail": Decimal(match["retail"]),
            "cost": Decimal(match["cost"]),
            "ext_cost": Decimal(match["ext_cost"]),
            "ctns": int(match["ctns"]),
            "cspk": int(match["cspk"]),
            "ext_qty": int(match["ext_qty"]),
            "cube": Decimal(match["cube"]),
            "kilograms": Decimal(match["kilograms"]),
        })

    if not items:
        raise ValueError("No line items found in purchase order")
    return items


def parse_totals(text):
    match = re.search(TOTALS_RE.pattern, text, re.MULTILINE)
    if not match:
        raise ValueError("Could not find TOTALS line in purchase order")
    return {
        "ext_cost": Decimal(match.group(1)),
        "ctns": int(match.group(2)),
        "ext_qty": int(match.group(3)),
        "kilograms": Decimal(match.group(4)),
    }


def validate_item(item):
    """Recompute EXT QTY and EXT COST. Returns {template column: message} for each value that doesn't match.

    Each check uses the printed values of the other columns, so a typo in one column flags only that column.
    """
    problems = {}
    expected_qty = item["ctns"] * item["cspk"]
    if expected_qty != item["ext_qty"]:
        problems["EXT QTY"] = f"Expected {expected_qty} (CTNS x CSPK)"
    expected_cost = item["cost"] * item["ext_qty"]
    if expected_cost != item["ext_cost"]:
        problems["EXT Cost"] = f"Expected {expected_cost} (Cost x EXT QTY)"
    return problems


def validate_totals(items, totals):
    """Compare the sums of the parsed items against the printed TOTALS line. Returns a list of problems."""
    sums = {
        "ext_cost": sum(item["ext_cost"] for item in items),
        "ctns": sum(item["ctns"] for item in items),
        "ext_qty": sum(item["ext_qty"] for item in items),
        "kilograms": sum(item["kilograms"] for item in items),
    }
    return [
        f"TOTALS {name}: printed {totals[name]}, items sum to {sums[name]}"
        for name in sums
        if sums[name] != totals[name]
    ]


def build_rows(filename, header, items):
    """One row per SKU, in the template's column order."""
    rows = []
    for item in items:
        rows.append([
            filename,
            header["buyer"],
            header["ship_terms"],
            header["po_number"],
            header["ref_master_po"],
            header["ship_date"],
            header["vendor"],
            header["ship_to"],
            header["bill_to"],
            item["dept"],
            item["sku"],
            item["upc"],
            item["vendor_part"],
            item["description"],
            float(item["retail"]),
            float(item["cost"]),
            float(item["ext_cost"]),
            item["ctns"],
            item["cspk"],
            item["ext_qty"],
            float(item["cube"]),
            float(item["kilograms"]),
            float(header["total_cost"]),
        ])
    return rows


def write_xlsx(rows, problems, out_path, template_path=TEMPLATE_PATH):
    """Fill a copy of the template so the header row always matches it exactly.

    `problems` holds one validate_item() result per row. Each mismatched cell gets a red fill and a comment
    with the expected value.
    """
    workbook = load_workbook(template_path)
    sheet = workbook.active
    column_of = {cell.value: cell.column for cell in sheet[1]}
    for row, row_problems in zip(rows, problems):
        sheet.append(row)
        for column_name, message in row_problems.items():
            cell = sheet.cell(row=sheet.max_row, column=column_of[column_name])
            cell.fill = MISMATCH_FILL
            cell.comment = Comment(message, "parse_po")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(out_path)


def main():
    parser = argparse.ArgumentParser(description="Convert a vendor purchase order text file to the BHF spreadsheet.")
    parser.add_argument("input", type=Path, help="purchase order .txt file")
    parser.add_argument("--out", type=Path, help="output .xlsx (default: output/<input name>.xlsx)")
    args = parser.parse_args()

    out_path = args.out or Path("output") / f"{args.input.stem}.xlsx"
    text = args.input.read_text()

    header = parse_header(text)
    items = parse_items(text)
    rows = build_rows(args.input.name, header, items)
    problems = [validate_item(item) for item in items]
    write_xlsx(rows, problems, out_path)

    print(f"Wrote {len(rows)} rows to {out_path}")
    for item, item_problems in zip(items, problems):
        for column_name, message in item_problems.items():
            print(f"  SKU {item['sku']} {column_name}: {message}")
    for problem in validate_totals(items, parse_totals(text)):
        print(f"  {problem}")


if __name__ == "__main__":
    main()
