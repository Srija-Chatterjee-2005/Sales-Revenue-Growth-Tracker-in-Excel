from __future__ import annotations

import csv
from pathlib import Path

def clean_sales_csv(input_path: Path, output_path: Path) -> None:
    with input_path.open("r", encoding="utf-8-sig", newline="") as src:
        reader = csv.DictReader(src)
        if reader.fieldnames is None:
            raise ValueError("CSV contains no headers.")

        seen = set()
        cleaned = []

        for row in reader:
            order_id = (row.get("Order ID") or "").strip()
            if not order_id or order_id in seen:
                continue
            seen.add(order_id)

            for key, value in list(row.items()):
                if isinstance(value, str):
                    row[key] = value.strip()

            cleaned.append(row)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8", newline="") as dst:
        writer = csv.DictWriter(dst, fieldnames=reader.fieldnames)
        writer.writeheader()
        writer.writerows(cleaned)

    print(f"Cleaned rows: {len(cleaned):,}")
    print(f"Saved: {output_path}")

if __name__ == "__main__":
    root = Path(__file__).resolve().parents[1]
    clean_sales_csv(
        root / "data" / "sample_sales_data.csv",
        root / "data" / "cleaned_sales_data.csv",
    )
