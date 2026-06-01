
from calendar import week

import pandas as pd
import matplotlib.pyplot as plt   # ✅ denne manglet

file = "stps_tolk.xlsx"


xls = pd.ExcelFile(file, engine="openpyxl")


print(xls.sheet_names)



# les hele arket
df = pd.read_excel(file, sheet_name="kuponger", header=None, engine="openpyxl")

# ----------------------
# ⚙️ KONFIGURASJON
# ----------------------

step = 22  # rader per uke (juster hvis 22/23)
start_row = 1  # rad for første dato (B2 ≈ index 1)

# tipper offsets (juster hvis flere)
players = {
    "AHH": {
        "rows_offset": 15,
        "cols_play": [13,14,15],   # N:P
        "cols_points": [17,18,19], # R:T
        "col_correct": 20          # U
    }
}

player_name = "AHH"   # <-- velg tipper her
current_week = 0

player = players[player_name]

# ----------------------
# 🎨 PLOTT
# ----------------------

fig, ax = plt.subplots(figsize=(10,6))

def draw(week):
    ax.clear()
print("TYPE week:", type(week))
row_start = start_row + week * step
data_row = row_start + player["rows_offset"]

print("Dato rad:", row_start)
print("Data rad:", data_row)

print(df.iloc[data_row-2:data_row+5, 10:25])  # debug-visning

print("Dato rad:", row_start)
print("Data rad:", data_row)
print(df.iloc[data_row-2:data_row+5, 10:25])

...

    # dato
date = df.iloc[row_start, 1]

    # rad med data
data_row = row_start + player["rows_offset"]

    # hent rekker
plays = [int(x) for x in df.iloc[data_row, player["cols_play"]]]

    # hent poeng
points = [int(x) for x in df.iloc[data_row, player["cols_points"]]]

    # hent riktige
correct = int(df.iloc[data_row, player["col_correct"]])

    # ----------------------
    # VISUALISER
    # ----------------------

text = f"{player_name} – Uke {week+1}\nDato: {date}\n\n"

text += "Rekker:\n"
for p in plays:
        text += str(p) + "\n"

text += "\nPoeng:\n"
for p in points:
        text += str(p) + "\n"

text += f"\n✅ Rette: {correct}"

ax.text(0.1, 0.5, text, fontsize=12, va='center')

ax.set_title("Kupong‑viewer")
ax.axis("off")


def on_key(event):
    global current_week

    if event.key == "right":
        current_week += 1
    elif event.key == "left":
        current_week -= 1

    # loop rundt
    max_weeks = 40
    current_week = current_week % max_weeks

    draw(current_week)
    fig.canvas.draw()


fig.canvas.mpl_connect("key_press_event", on_key)

draw(current_week)

plt.show()


