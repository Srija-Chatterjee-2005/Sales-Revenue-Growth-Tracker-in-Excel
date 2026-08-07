from __future__ import annotations

import csv
from collections import defaultdict
from pathlib import Path


MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]


def monthly_revenue(csv_path: Path) -> list[float]:
    totals = defaultdict(float)

    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            month = (row.get("Month") or "").strip()
            try:
                revenue = float(row.get("Revenue") or 0)
            except ValueError:
                revenue = 0
            totals[month] += revenue

    return [totals[m] for m in MONTHS]


def next_month_forecast(values: list[float]) -> float:
    """
    Lightweight transparent forecast used for portfolio/demo purposes.
    Uses the last 3 months' average and their average month-over-month growth.
    """
    if len(values) < 3:
        return 0.0

    last3 = values[-3:]
    growth_rates = []

    for previous, current in zip(last3, last3[1:]):
        if previous:
            growth_rates.append((current - previous) / previous)

    average_growth = sum(growth_rates) / len(growth_rates) if growth_rates else 0.0
    baseline = sum(last3) / len(last3)

    return baseline * (1 + average_growth)


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    csv_path = root / "data" / "sample_sales_data.csv"

    values = monthly_revenue(csv_path)
    forecast = next_month_forecast(values)

    print("Monthly revenue:")
    for month, value in zip(MONTHS, values):
        print(f"{month}: ₹{value:,.0f}")

    print(f"\nNext month forecast: ₹{forecast:,.0f}")


if __name__ == "__main__":
    main()
