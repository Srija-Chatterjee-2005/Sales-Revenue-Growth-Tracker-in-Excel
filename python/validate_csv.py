from __future__ import annotations

import csv
import sys
from pathlib import Path

EXPECTED_HEADERS = [
    "Order ID","Date","Month","Quarter","Region","City","Sales Executive",
    "Product Category","Product Name","Units Sold","Selling Price","Revenue",
    "Cost","Profit","Profit Margin %","Customer Type","Sales Channel",
    "Target Sales","Achievement %","Growth %","Order Status","Discount %","Rating"
]

def validate_csv(path: Path) -> int:
    if not path.exists():
        print(f"ERROR: File not found: {path}")
        return 1

    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        try:
            headers = next(reader)
        except StopIteration:
            print("ERROR: CSV is empty.")
            return 1

        if headers != EXPECTED_HEADERS:
            print("ERROR: Header structure does not match SalesIQ.")
            print("Expected:")
            print(EXPECTED_HEADERS)
            print("Found:")
            print(headers)
            return 1

        row_count = 0
        width_errors = 0

        for line_no, row in enumerate(reader, start=2):
            row_count += 1
            if len(row) != len(EXPECTED_HEADERS):
                width_errors += 1
                print(f"WARNING: row {line_no} has {len(row)} columns.")

        if row_count == 0:
            print("ERROR: No data rows were found.")
            return 1

        print(f"VALID: {row_count:,} rows found.")
        if width_errors:
            print(f"WARNING: {width_errors} malformed rows found.")
            return 2

    return 0

if __name__ == "__main__":
    default = Path(__file__).resolve().parents[1] / "data" / "sample_sales_data.csv"
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else default
    raise SystemExit(validate_csv(target))
