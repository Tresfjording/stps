import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook

# Parse repeated horizontal blocks (kuponger)

FILE = "stps_tolk.xlsx"
SHEET = "kuponger"

# Konfig: vertikal og horisontal offset
START_ROW = 1            # A1-basis
WEEK_STEP = 23           # rader per uke (vertikal forskyvning)
BLOCKS = 9               # antall horisontale blokker
BLOCK_START_COL = 14     # N = kolonne 14 (1-basert)
BLOCK_SPACING = 10       # start->start mellom blokker
BLOCK_WIDTH = 8          # bredde på hver blokk
ROWS_PER_BLOCK = 12      # rader i hver blokk (f.eks. 5..16)


def read_blocks(file=FILE, sheet=SHEET):
    wb = load_workbook(file, data_only=True)
    ws = wb[sheet] if sheet in wb.sheetnames else wb.active

    weeks = []

    for week_idx in range(100):
        row_start = START_ROW + week_idx * WEEK_STEP
        r0 = row_start + 4
        r1 = r0 + ROWS_PER_BLOCK - 1

        # stopp hvis ukens område er tomt i kolonne A (ingen data)
        colA_vals = [ws.cell(row=r, column=1).value for r in range(r0, r1+1)]
        if all(v in (None, "") for v in colA_vals):
            break

        week = {"week_index": week_idx, "blocks": []}

        for block_idx in range(BLOCKS):
            start_col = BLOCK_START_COL + block_idx * BLOCK_SPACING
            block_vals = []
            for r in range(r0, r1+1):
                row_vals = [ws.cell(row=r, column=c).value for c in range(start_col, start_col + BLOCK_WIDTH)]
                block_vals.append(row_vals)

            # sjekk om blokken inneholder noe
            has_data = any(any(cell not in (None, "") for cell in row) for row in block_vals)
            week["blocks"].append({
                "block_idx": block_idx,
                "start_col": start_col,
                "has_data": has_data,
                "values": block_vals
            })

        weeks.append(week)

    return weeks


def print_summary(weeks):
    for w in weeks:
        print(f"Week {w['week_index']}:")
        for b in w["blocks"]:
            print(f"  Block {b['block_idx']} start_col={b['start_col']} has_data={b['has_data']}")


if __name__ == '__main__':
    weeks = read_blocks()
    print(f"Found {len(weeks)} weeks with data")
    print_summary(weeks)


