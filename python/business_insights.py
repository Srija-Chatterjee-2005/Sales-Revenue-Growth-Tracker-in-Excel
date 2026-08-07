from __future__ import annotations

import csv
from collections import defaultdict
from pathlib import Path


def load_rows(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def safe_float(value: str | None) -> float:
    try:
        return float(value or 0)
    except ValueError:
        return 0.0


def generate_insights(rows: list[dict[str, str]]) -> dict[str, object]:
    region_revenue = defaultdict(float)
    category_profit = defaultdict(float)
    executive_revenue = defaultdict(float)

    total_revenue = 0.0
    total_profit = 0.0
    total_target = 0.0
    completed = 0

    for row in rows:
        revenue = safe_float(row.get("Revenue"))
        profit = safe_float(row.get("Profit"))
        target = safe_float(row.get("Target Sales"))

        total_revenue += revenue
        total_profit += profit
        total_target += target

        region_revenue[row.get("Region","Unknown")] += revenue
        category_profit[row.get("Product Category","Unknown")] += profit
        executive_revenue[row.get("Sales Executive","Unknown")] += revenue

        if row.get("Order Status") == "Completed":
            completed += 1

    top_region = max(region_revenue, key=region_revenue.get) if region_revenue else "N/A"
    top_category = max(category_profit, key=category_profit.get) if category_profit else "N/A"
    top_executive = max(executive_revenue, key=executive_revenue.get) if executive_revenue else "N/A"

    achievement = total_revenue / total_target if total_target else 0.0
    margin = total_profit / total_revenue if total_revenue else 0.0
    completion_rate = completed / len(rows) if rows else 0.0

    return {
        "total_revenue": total_revenue,
        "total_profit": total_profit,
        "profit_margin": margin,
        "achievement": achievement,
        "completion_rate": completion_rate,
        "top_region": top_region,
        "top_category": top_category,
        "top_executive": top_executive,
    }


def print_insights(insights: dict[str, object]) -> None:
    print("SalesIQ Business Insights")
    print("-" * 30)
    print(f"Revenue: ₹{insights['total_revenue']:,.0f}")
    print(f"Profit: ₹{insights['total_profit']:,.0f}")
    print(f"Profit margin: {insights['profit_margin']:.1%}")
    print(f"Target achievement: {insights['achievement']:.1%}")
    print(f"Completion rate: {insights['completion_rate']:.1%}")
    print(f"Top region: {insights['top_region']}")
    print(f"Top category: {insights['top_category']}")
    print(f"Top executive: {insights['top_executive']}")

    if insights["achievement"] < 1:
        print("Action: Focus on weaker regions/categories to close the target gap.")
    else:
        print("Action: Protect top-performing segments and scale winning practices.")


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    rows = load_rows(root / "data" / "sample_sales_data.csv")
    print_insights(generate_insights(rows))


if __name__ == "__main__":
    main()
