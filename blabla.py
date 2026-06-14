import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
import mpld3


file = "stps_tolk.xlsx"

df = pd.read_excel(file, sheet_name="Hovedtabell", engine="openpyxl")

df = df.rename(columns={df.columns[0]: "Navn"})
df = df.dropna(subset=["Navn"])

columns = df.columns[1:]  # alle kolonner unntatt Navn
current = 0

fig, ax = plt.subplots(figsize=(10, 6))


def draw(col_index):
    ax.clear()

    col = columns[col_index]

    temp = df[["Navn", col]].dropna()
    temp[col] = pd.to_numeric(temp[col], errors="coerce")
    temp = temp.dropna()

    leaderboard = temp.set_index("Navn")[col].sort_values(ascending=True)

    ax.barh(leaderboard.index, leaderboard.values, color="steelblue")

    # legg på tall
    for i, v in enumerate(leaderboard):
        ax.text(v + 1, i, str(int(v)), va='center')

    ax.set_title(f"{col}")
    ax.set_xlabel("Poeng")

    # fjern ramme
    for spine in ax.spines.values():
        spine.set_visible(False)

    ax.tick_params(left=False)
    ax.grid(axis="x")


def on_key(event):
    global current

    if event.key == "right":
        current += 1
    elif event.key == "left":
        current -= 1

    current = current % len(columns)  # loop rundt

    draw(current)
    fig.canvas.draw()


fig.canvas.mpl_connect("key_press_event", on_key)


draw(current)
plt.show()

# Create your plot layout
fig, ax = plt.subplots()
#ax.plot([1, 2, 3, 4], [1, 4, 9, 16])
#ax.set_title(f"{col}")
ax.plot("Poeng")

# Export the figure to an interactive HTML document
mpld3.save_html(fig, "interactive_report.html")