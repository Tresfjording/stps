import pandas as pd
import matplotlib.pyplot as plt

plt.close('all')

file = "stps_tolk.xlsx"

df = pd.read_excel(file, sheet_name="Hovedtabell", engine="openpyxl")

df = df.rename(columns={
    df.columns[0]: "Navn",
    df.columns[1]: "Poeng"
})

df = df.dropna(subset=["Navn"])
df["Poeng"] = pd.to_numeric(df["Poeng"], errors="coerce")
df = df.dropna()

leaderboard = df.set_index("Navn")["Poeng"]

# ✅ riktig sortering for horisontal graf
leaderboard = leaderboard.sort_values(ascending=True)

# ✅ topp 3 highlight (bottom = best i barh)
colors = ["steelblue"] * len(leaderboard)
colors[-1] = "gold"
colors[-2] = "silver"
colors[-3] = "#cd7f32"

plt.figure(figsize=(10,6))

ax = leaderboard.plot(kind="barh", color=colors)

# Fjern ramme
for spine in ax.spines.values():
    spine.set_visible(False)

# Fjern små ticks
ax.tick_params(left=False)

# Tekst
leader = leaderboard.max()
for i, v in enumerate(leaderboard):
    gap = leader - v
    label = f"{int(v)} 👑" if gap == 0 else f"{int(v)} (-{int(gap)})"
    ax.text(v + 50, i, label, va='center')

plt.title("Hovedtabell")
plt.xlabel("Poeng")
plt.ylabel("Spiller")
plt.grid(axis="x")

plt.show()


ax = leaderboard.plot(kind="barh", color=colors)

# ✅ legg på tall
leader = leaderboard.max()


leader = leaderboard.max()


# Fjern kantlinjer
for spine in ax.spines.values():
    spine.set_visible(False)

ax.tick_params(left=False)   # fjerner små streker på y-aksen

for i, v in enumerate(leaderboard):
    gap = leader - v
    label = f"{int(v)} (-{int(gap)})" if gap > 0 else f"{int(v)} (0)"
    
    ax.text(v + 50, i, label, va='center')


for i, (name, v) in enumerate(leaderboard.items()):
    crown = " 👑" if i == len(leaderboard)-1 else ""
    gap = leader - v
    ax.text(v + 50, i, f"{int(v)} (-{int(gap)}){crown}", va='center')

import datetime
import pandas as pd
import matplotlib.pyplot as plt

file = "stps_tolk.xlsx"

# Les arket
df = pd.read_excel(file, sheet_name="Hovedtabell", engine="openpyxl")

# Første kolonne = Navn
df = df.rename(columns={df.columns[0]: "Navn"})

df = df.dropna(subset=["Navn"])

# Gå gjennom alle andre kolonner
for col in df.columns[1:]:

    print(f"\n📊 ANALYSE: {col}")

    # Hent kolonne
    temp = df[["Navn", col]].dropna()

    # gjør om til tall hvis mulig
    temp[col] = pd.to_numeric(temp[col], errors="coerce")
    temp = temp.dropna()

    # lag leaderboard
    leaderboard = temp.set_index("Navn")[col].sort_values(ascending=True)

    print(leaderboard)

    # -------------------
    # 📈 GRAF
    # -------------------
    plt.figure(figsize=(8,5))

    ax = leaderboard.plot(kind="barh", color="steelblue")

    # legg på tall
    for i, v in enumerate(leaderboard):
        ax.text(v + 1, i, str(int(v)), va='center')

    plt.title(f"{col}")
    plt.xlabel("Verdi")
    plt.ylabel("Spiller")

    # fjern ramme (clean look)
    for spine in ax.spines.values():
        spine.set_visible(False)

    ax.tick_params(left=False)

    plt.grid(axis="x")
    plt.tight_layout()

    plt.show()


plt.title(f"Hovedtabell – {datetime.date.today()}")


plt.title("Hovedtabell")
plt.xlabel("Poeng")
plt.ylabel("Spiller")
plt.grid(axis="x")

plt.show()
plt.savefig(f"{col}.png", dpi=150)

