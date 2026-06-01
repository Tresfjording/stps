print("🚀 Test ac Autmate Git 01.06.2026  - 10:17:58")
import matplotlib.pyplot as plt
import pandas as pd
df_data = pd.DataFrame({
    "Uke": [1, 1, 2],
    "Navn": ["A", "B", "A"],
    "Poeng": [10, 20, 15]
})

# ✅ Pivot (poeng per uke)
pivot = df_data.pivot(index="Uke", columns="Navn", values="Poeng")

# ✅ Kumulativ (league race)
pivot.cumsum().plot(figsize=(10,6))
plt.title("Liga utvikling")
plt.ylabel("Poeng")
plt.xlabel("Uke")
plt.grid()
plt.show()

# ✅ Leaderboard
print("\n🏆 Leaderboard:")
leaderboard = (
    df_data.groupby("Navn")["Poeng"]
    .sum()
    .sort_values(ascending=False)
)
print(leaderboard)

# ✅ Form siste 5
print("\n🔥 Form:")
form = (
    df_data.groupby("Navn")
    .tail(5)
    .groupby("Navn")["Poeng"]
    .mean()
    .sort_values(ascending=False)
)
print(form)

# ✅ Ukesvinnere
print("\n🥇 Vinnere per uke:")
winners = df_data.loc[df_data.groupby("Uke")["Poeng"].idxmax()]
print(winners)

# ✅ Ukens taper 😄
print("\n🥴 Dårligste per uke:")
losers = df_data.loc[df_data.groupby("Uke")["Poeng"].idxmin()]
print(losers)

# ✅ Stabilitet
print("\n📊 Stabilitet (lav = jevn):")
consistency = df_data.groupby("Navn")["Poeng"].std().sort_values()
print(consistency)
