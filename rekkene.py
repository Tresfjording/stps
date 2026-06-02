import pandas as pd
import matplotlib.pyplot as plt

# ----------------------
# 📄 LAST EXCEL
# ----------------------
file = "stps_tolk.xlsx"

xls = pd.ExcelFile(file, engine="openpyxl")
print(xls.sheet_names)

df = pd.read_excel(file, sheet_name="kuponger", header=None, engine="openpyxl")

# ----------------------
# ⚙️ KONFIGURASJON
# ----------------------
step = 22
start_row = 1

players = {
    "AHH": {
        "rows_offset": 15,
        "cols_play": [13, 14, 15],   # N:P
        "cols_points": [17, 18, 19], # R:T
        "col_correct": 20            # U
    }
}

player_name = "AHH"
current_week = 0

player = players[player_name]

# ----------------------
# 🎨 PLOTT
# ----------------------
fig, ax = plt.subplots(figsize=(10, 6))


# ----------------------
# 🧠 DRAW FUNKSJON
# ----------------------
def draw(week_index):
    ax.clear()
    ax.set_axis_off()

    # beregn ukeoffset
    row_start = start_row + week_index * step

    r0 = row_start + 4   # tilsvarer rad 5
    r1 = row_start + 13  # tilsvarer rad 13



for i in range(len(venstre)):
    # ✅ hent venstre del (A:H)
    venstre = df.iloc[r0:r1, 0:8].values

    # ✅ hent tipper (N:T)
    hoyre = df.iloc[r0:r1, 13:20].values
    
    # venstre
    v = venstre[i]
    venstre_str = "\t".join(
        str(x) if pd.notna(x) else "" for x in v
    )

    # høyre
    h = hoyre[i]
    hoyre_str = "\t".join(
        str(x) if pd.notna(x) else "" for x in h
    )

    text += venstre_str + "\t|\t" + hoyre_str + "\n"


# ----------------------
# ⌨️ KEY NAVIGASJON
# ----------------------
def on_key(event):
    global current_week

    if event.key == "right":
        current_week += 1
    elif event.key == "left":
        current_week -= 1

    max_weeks = 40
    current_week = current_week % max_weeks

    draw(current_week)


# ----------------------
# 🚀 START PROGRAM
# ----------------------
fig.canvas.mpl_connect("key_press_event", on_key)

draw(0)

plt.show()
