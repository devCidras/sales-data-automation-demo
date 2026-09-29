"""Generates a realistic, messy sales export to use as the "before" demo file.

Simulates the kind of CSV export a small business would actually pull from a
POS or invoicing system: inconsistent casing, stray whitespace, mixed date
formats, currency symbols in numeric fields, duplicate rows and missing
values.
"""

import csv
import random
from datetime import date, timedelta
from pathlib import Path

random.seed(42)

PRODUCTS = [
    ("Office Chair", "Furniture", 89.90),
    ("Adjustable Desk", "Furniture", 249.00),
    ("24in Monitor", "Electronics", 139.50),
    ("Mechanical Keyboard", "Electronics", 59.90),
    ("Wireless Mouse", "Electronics", 24.90),
    ("LED Desk Lamp", "Lighting", 34.50),
    ("Modular Shelving Unit", "Furniture", 119.00),
    ("Bluetooth Headphones", "Electronics", 79.00),
    ("XL Mouse Pad", "Accessories", 14.90),
    ("HD Webcam", "Electronics", 44.90),
]

REGIONS = ["North", "South", "East", "West", "Central"]
CUSTOMERS = [f"Customer {i:03d}" for i in range(1, 41)]

DATE_FORMATS = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]


def messy_product_name(name: str) -> str:
    variant = random.choice([str.upper, str.lower, str.title, lambda s: s])
    padded = random.choice([True, False])
    out = variant(name)
    return f"  {out}  " if padded else out


def messy_price(price: float) -> str:
    variant = random.random()
    if variant < 0.4:
        return f"{price:.2f} EUR"
    if variant < 0.6:
        return f"€{price:.2f}"
    return f"{price:.2f}"


def random_date() -> str:
    day_offset = random.randint(0, 179)  # ~6 months of data
    d = date(2026, 3, 1) + timedelta(days=day_offset)
    fmt = random.choice(DATE_FORMATS)
    return d.strftime(fmt)


def build_rows(n: int) -> list[dict]:
    rows = []
    for _ in range(n):
        product, category, unit_price = random.choice(PRODUCTS)
        qty = random.randint(1, 6)
        is_return = random.random() < 0.05
        if is_return:
            qty = -abs(qty)

        row = {
            "Date": random_date(),
            "Product": messy_product_name(product),
            "Category": category,
            "Quantity": qty,
            "Unit Price": messy_price(unit_price),
            "Region": random.choice(REGIONS),
            "Customer": random.choice(CUSTOMERS),
        }

        # Randomly blank out a non-critical field to simulate missing data.
        if random.random() < 0.03:
            row["Region"] = ""
        if random.random() < 0.02:
            row["Customer"] = ""

        rows.append(row)
    return rows


def main() -> None:
    rows = build_rows(480)

    # Inject some exact duplicate rows, as often happens with re-exports.
    duplicates = random.sample(rows, 15)
    rows.extend(duplicates)

    # Inject a few rows with a missing/garbage critical field (unusable).
    for _ in range(6):
        bad_row = dict(random.choice(rows))
        bad_row["Unit Price"] = ""
        rows.append(bad_row)

    random.shuffle(rows)

    out_path = Path(__file__).parent / "data" / "raw_sales.csv"
    with out_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)

    print(f"Sample file generated: {out_path} ({len(rows)} rows)")


if __name__ == "__main__":
    main()
