import pandas as pd
import matplotlib.pyplot as plt

print("🚀 Skorgen Tippelag analyse\n")

# ------------------------------
# 📥 LES DATA
# ------------------------------
df = pd.read_excel("SkorgenTippelag.xlsm", sheet_name="til py")

# Rens kolonner
df.columns = df.columns.str.strip().str.lower()

# Gjør dato lesbar
df["dato"] = pd.to_datetime(df["dato"], origin="1899-12-30", unit="D")

# ------------------------------
# 🔄 WIDE → LONG
# ------------------------------
df_long = df.melt(id_vars="dato", var_name="navn", value_name="poeng")

df_long = df_long.dropna()
df_long = df_long[~df_long["navn"].isin(["utg", "lev"])]
df_long = df_long.sort_values("dato")

# ------------------------------
# 🏆 LEADERBOARD
# ------------------------------
leaderboard = df_long.groupby("navn")["poeng"].sum().sort_values(ascending=False)

print("🏆 Leaderboard:")
print(leaderboard, "\n")

# ------------------------------
# 🔥 FORM (siste 3 runder)
# ------------------------------
last_dates = sorted(df_long["dato"].unique())[-3:]

form = df_long[df_long["dato"].isin(last_dates)] \
    .groupby("navn")["poeng"].mean() \
    .sort_values(ascending=False)

print("🔥 Form:")
print(form, "\n")

# ------------------------------
# 🥇 VINNERE PER RUNDE
# ------------------------------
winners = df_long.loc[df_long.groupby("dato")["poeng"].idxmax()]

print("🥇 Vinnere:")
print(winners[["dato", "navn", "poeng"]], "\n")

# ------------------------------
# 📊 STABILITET
# ------------------------------
stability = df_long.groupby("navn")["poeng"].std().fillna(0).sort_values()

print("📊 Stabilitet:")
print(stability, "\n")

# ------------------------------
# 📈 AKKUMULERT LIGA
# ------------------------------
pivot = df_long.pivot(index="dato", columns="navn", values="poeng").fillna(0)
cum = pivot.cumsum()

# sorter etter leaderboard
cum = cum[leaderboard.index]

cum.plot(marker="o")

plt.title("Liga utvikling")
plt.xlabel("Dato")
plt.ylabel("Poeng")
plt.grid()

plt.legend(title="Navn", bbox_to_anchor=(1.05, 1))
plt.tight_layout()
plt.show()

# ------------------------------
# 🏁 RANKING OVER TID
# ------------------------------
rank = cum.rank(axis=1, ascending=False)

rank.plot()

plt.title("Plassering over tid")
plt.xlabel("Dato")
plt.ylabel("Plassering")
plt.gca().invert_yaxis()
plt.grid()

plt.tight_layout()
plt.show()

# ------------------------------
# 🧠 AUTOMATISK KOMMENTAR
# ------------------------------
print("🧠 Analyse:")

print(f"🏆 {leaderboard.idxmax()} leder ligaen")
print(f"🔥 {form.idxmax()} er i best form")
print(f"🥶 {leaderboard.idxmin()} ligger sist")
