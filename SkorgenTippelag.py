import pandas as pd

file = "SkorgenTippelag.xlsm"

# Les hele arket uten header (fordi Excel er visuelt)
df = pd.read_excel(file, sheet_name="Kuponger", header=None)

# Spillere vi vet finnes
players = ["AHH", "RB", "EB", "KAF", "ØG", "JH", "TOH", "UTG", "LEV"]

data = []

current_week = None

for i, row in df.iterrows():
    row_values = row.astype(str).tolist()

    # Finn uke (linje med dato og ukenummer)
    if "lørdag" in row_values:
        # Ukenummer står ofte litt over – vi bruker index
        current_week = f"Uke_{i}"

    # Finn linje med spillernavn (AHH, RB osv)
    for p in players:
        if p in row_values:

            # Les poeng fra samme rad (eller rett under – må justeres litt)
            try:
                idx = row_values.index(p)
                score = row_values[idx + 1]

                if score.isdigit():
                    data.append([current_week, p, int(score)])

            except:
                continue

# Lag DataFrame
df_data = pd.DataFrame(data, columns=["Uke", "Navn", "Poeng"])

print(df_data.head())