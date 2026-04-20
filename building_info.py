r"""
Fill building info into Filleddata_output.xlsx and Global_assumptions.csv.

Usage:
  python building_info.py

Reads building info from:
  D:\JARVIS back up 16092025\JARVIS backup\Code - Full set\building info.xlsx

Updates:
  1. Filleddata_output.xlsx (Datadata2 sheet, cols AE-AF) — A-B pairs only
  2. Global_assumptions.csv — E-F pairs (Seismic status, Fck, Fy, Slab thickness)

Run node_coordinate_calculator.py first to create Filleddata_output.xlsx.
"""

import csv
from pathlib import Path
from openpyxl import load_workbook

BUILDING_INFO_PATH = Path(r"D:\JARVIS back up 16092025\JARVIS backup\Code - Full set\building info.xlsx")
FILLED_PATH = Path(r"D:\JARVIS back up 16092025\JARVIS backup\STD ANL model\Filleddata.xlsx")
GLOBAL_CSV_PATH = Path(r"D:\JARVIS back up 16092025\JARVIS backup\Code - Full set\Global_assumptions.csv")


def read_building_info():
    """Read building info. Returns (ab_pairs, ef_pairs) separately."""
    wb = load_workbook(BUILDING_INFO_PATH, data_only=True)
    ws = wb["building info"]
    ab_pairs = []
    ef_pairs = []
    for row in ws.iter_rows(min_row=1, values_only=True):
        if row[0] is not None:
            ab_pairs.append((str(row[0]).strip(), row[1]))
        if len(row) > 4 and row[4] is not None:
            ef_pairs.append((str(row[4]).strip(), row[5]))
    wb.close()
    return ab_pairs, ef_pairs


def update_global_csv(ef_pairs):
    """Update matching keys in Global_assumptions.csv with E-F values."""
    ef_dict = {k: v for k, v in ef_pairs}

    rows = []
    with open(GLOBAL_CSV_PATH, "r", newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        for row in reader:
            if row and row[0] in ef_dict:
                row[1] = ef_dict[row[0]]
            rows.append(row)

    with open(GLOBAL_CSV_PATH, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerows(rows)


def main():
    if not FILLED_PATH.exists():
        print(f"Error: {FILLED_PATH} not found. Run node_coordinate_calculator.py first.")
        return

    print(f"Reading building info from: {BUILDING_INFO_PATH}")
    ab_pairs, ef_pairs = read_building_info()

    # 1. Write A-B pairs to AE-AF in Filleddata
    wb = load_workbook(FILLED_PATH)
    ws = wb["Datadata2"]
    for i, (key, value) in enumerate(ab_pairs, 2):
        ws.cell(i, 31, key)     # AE: parameter name
        ws.cell(i, 32, value)   # AF: parameter value
    wb.save(FILLED_PATH)
    print(f"Written {len(ab_pairs)} params to AE-AF in Filleddata")

    # 2. Write E-F pairs to Global_assumptions.csv
    update_global_csv(ef_pairs)
    print(f"Updated {len(ef_pairs)} params in Global_assumptions.csv:")
    for k, v in ef_pairs:
        print(f"  {k} = {v}")

    print(f"\nDone!")


if __name__ == "__main__":
    main()
