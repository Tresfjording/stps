import pandas as pd
import matplotlib.pyplot as plt

print("🚀 Skorgen Tippelag analyse\n")

# ✅ Les arket du laget
df = pd.read_excel("SkorgenTippelag.xlsm", sheet_name="til py")

# ------------------------------
# ✅ Rens kolonnenavn
# ------------------------------
df.columns = df.columns.str.strip().str.lower()

# ------------------------------
# 🔄 Gjør wide → long format
# ------------------------------
df_long = df.melt(id_vars="dato", var_name="navn", value_name="poeng")

# Fjern tomt
df_long = df_long.dropna()

# Fjern system
df_long = df_long[~df_long["navn"].isin(["utg", "lev"])]

# ------------------------------
# 🏆 LEADERBOARD
# ------------------------------
leaderboard = df_long.groupby("navn")["poeng"].sum().sort_values(ascending=False)

print("🏆 Leaderboard:")
print(leaderboard, "\n")

# ------------------------------
# 🔥 FORM (siste 3 runder)
# ------------------------------
last_3 = df_long["dato"].unique()[-3:]

form = df_long[df_long["dato"].isin(last_3)] \
    .groupby("navn")["poeng"].mean() \
    .sort_values(ascending=False)

print("🔥 Form (siste 3):")
print(form, "\n")

# ------------------------------
# 🥇 VINNERE PER RUNDE
# ------------------------------
winners = df_long.loc[df_long.groupby("dato")["poeng"].idxmax()]

print("🥇 Vinnere:")
print(winners, "\n")

# ------------------------------
# 📊 STABILITET
# ------------------------------
stability = df_long.groupby("navn")["poeng"].std().fillna(0).sort_values()

print("📊 Stabilitet:")
print(stability, "\n")

# ------------------------------
# 📈 GRAF (akkumulert liga)
# ------------------------------
pivot = df_long.pivot(index="dato", columns="navn", values="poeng").fillna(0)
cum = pivot.cumsum()

cum.plot(marker="o")

plt.title("Liga utvikling")
plt.xlabel("Runde")
plt.ylabel("Poeng")
plt.grid()

plt.legend(title="Navn", bbox_to_anchor=(1.05, 1))
plt.tight_layout()
plt.show()
