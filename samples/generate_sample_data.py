"""Generates mock Excel spreadsheets for testing the reporting pipeline."""

from datetime import datetime, timedelta
from pathlib import Path
import random
import pandas as pd


def generate_sample_workbook(region_name: str, output_path: Path, num_records: int = 25):
    categories = ["Hardware", "Software", "Cloud Services", "Consulting", "Support"]
    statuses = ["Completed", "Completed", "Completed", "Shipped", "In Transit", "Pending", "Cancelled"]
    reps_by_region = {
        "North": ["Alice Morgan", "David Vance", "Sarah Chen"],
        "South": ["Carlos Mendez", "Elena Rostova", "Marcus Brody"],
        "West": ["Jordan Taylor", "Keiko Tanaka", "Liam O'Connor"],
        "East": ["Noah Patel", "Olivia Wright", "Zoe Fischer"],
    }
    reps = reps_by_region.get(region_name, ["Agent A", "Agent B"])

    start_date = datetime.now() - timedelta(days=30)

    records = []
    for i in range(num_records):
        tx_date = start_date + timedelta(days=random.randint(0, 30))
        cat = random.choice(categories)
        units = random.randint(1, 45)
        unit_price = {
            "Hardware": 350.0,
            "Software": 1200.0,
            "Cloud Services": 2500.0,
            "Consulting": 180.0,
            "Support": 75.0,
        }.get(cat, 100.0)

        revenue = round(units * unit_price * random.uniform(0.85, 1.15), 2)
        status = random.choice(statuses)

        records.append({
            "Transaction ID": f"TX-{region_name[:2].upper()}-{1000 + i}",
            "Date": tx_date.strftime("%Y-%m-%d"),
            "Region": region_name,
            "Sales Rep": random.choice(reps),
            "Product Category": cat,
            "Units Sold": units,
            "Revenue": revenue,
            "Status": status,
        })

    df = pd.DataFrame(records)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Sales_Data", index=False)
    print(f"Created sample workbook: {output_path} ({len(df)} rows)")


def generate_all_samples(target_dir: str = "samples/data"):
    out_dir = Path(target_dir)
    regions = ["North", "South", "West"]
    for reg in regions:
        generate_sample_workbook(reg, out_dir / f"Regional_Sales_{reg}.xlsx", num_records=20)


if __name__ == "__main__":
    generate_all_samples()
