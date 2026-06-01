import pandas as pd
import matplotlib.pyplot as plt

print("🚀 Skorgen Tippelag analyse (FIXED)\n")

# ✅ Les tabellen som starter i kolonne P
df = pd.read_excel(
    "SkorgenTippelag.xlsm",
    engine="openpyxl",
    sheet_name="Statistikk",
    usecols="P:Z",   # 👉 juster hvis nødvendig
    header=0
)

# Fjern tomme rader
df = df.dropna(how="all")

print("Kolonner funnet:")
print(df.columns, "\n")

# ------------------------------
# ✅ FINN riktig kolonnenavn
# ------------------------------
# Ofte noe sånt:
# 'Navn', 'Poeng', 'Poengsnitt'

df = df.rename(columns={
    df.columns[0]: "Navn",
    df.columns[1]: "Poeng"
})

# Fjern tomme navn
df = df.dropna(subset=["Navn"])

# Fjern systemrekker
df = df[~df["Navn"].isin(["UTG", "LEV"])]

# ✅ VIKTIG: Konverter Poeng riktig
df["Poeng"] = pd.to_numeric(df["Poeng"], errors="coerce")

df = df.dropna()

# ------------------------------
# 🏆 LEADERBOARD
# ------------------------------
leaderboard = df.set_index("Navn")["Poeng"].sort_values(ascending=False)

print("🏆 Leaderboard:")
print(leaderboard, "\n")

# ------------------------------
# 📈 GRAF
# ------------------------------
leaderboard.plot(kind="bar")

plt.title("Total poeng per spiller")
plt.ylabel("Poeng")
plt.xlabel("Navn")
plt.grid(axis="y")

plt.tight_layout()
plt.show()

from openpyxl import load_workbook
import pandas as pd

wb = load_workbook("SkorgenTippelag.xlsm", data_only=True)
ws = wb["Statistikk"]

table = ws.tables["t_tp"]

data = ws[table.ref]

df = pd.DataFrame([[cell.value for cell in row] for row in data])

df.columns = df.iloc[0]
df = df[1:]
