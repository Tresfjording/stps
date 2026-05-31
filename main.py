print("JEG ER HER!!!")
import sqlite3
import os
from tkinter import ON

print("\nDEBUG INFO:")
print("Working dir:", os.getcwd())

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
print("BASE_DIR:", BASE_DIR)

db_path = os.path.join(BASE_DIR, "data", "stps.db")
schema_path = os.path.join(BASE_DIR, "database", "schema.sql")

print("DB path:", db_path)
print("Schema path:", schema_path)


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

db_path = os.path.join(BASE_DIR, "data", "stps.db")
schema_path = os.path.join(BASE_DIR, "database", "schema.sql")

print("DB path:", db_path)
print("Schema path:", schema_path)

conn = sqlite3.connect(db_path)
cursor = conn.cursor()

with open(schema_path, "r") as f:
    sql_script = f.read()

#---cursor.executescript(sql_script)---

conn.commit()   # ✅ VIKTIG

print("\n--- SQL SCRIPT ---")
print(sql_script)

cursor.executescript(sql_script)
# --- Paths ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

db_path = os.path.join(BASE_DIR, "data", "stps.db")
schema_path = os.path.join(BASE_DIR, "database", "schema.sql")

# Sørg for at data-mappen finnes
os.makedirs(os.path.join(BASE_DIR, "data"), exist_ok=True)

# --- Database connection ---
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

# --- Opprett tabeller ---
with open(schema_path, "r") as f:
    cursor.executescript(f.read())

conn.commit()

print("✅ Database opprettet!")


# --- RESET DATA (kun under utvikling!) ---




# --- LEGG INN UKE ---
cursor.execute(
    "INSERT OR IGNORE INTO weeks (week_number, date) VALUES (?, ?)",
    (1, "2026-08-01")
)

conn.commit()

print("Uke lagt inn!")

# --- LEGG INN UKE 2 ---
cursor.execute(
    "INSERT OR IGNORE INTO weeks (week_number, date) VALUES (?, ?)",
    (2, "2026-08-08")
)

conn.commit()

# --- FINN week_id ---
cursor.execute("SELECT id FROM weeks WHERE week_number = 2")
week_id = cursor.fetchone()[0]


# --- LEGG INN KAMPER ---
for i in range(1, 13):
    cursor.execute(
        "INSERT OR IGNORE INTO matches (week_id, match_number, result) VALUES (?, ?, ?)",
        (week_id, i, None)
    )

  # --- LEGG INN SPILLERE ---
players = ["AHH", "RB", "EB", "KAF", "ØG", "JH", "TOH", "UTG", "LEV"]

for p in players:
    cursor.execute("INSERT OR IGNORE INTO players (name) VALUES (?)", (p,))

conn.commit()

print("Spillere lagt inn!")


# --- LEGG INN TIPS ---
for name in players:
    cursor.execute("SELECT id FROM players WHERE name = ?", (name,))
    result = cursor.fetchone()

    if result is None:
        raise ValueError(f"Fant ikke spiller: {name}")

    player_id = result[0]

    # 🔥 BESTEM STRATEGI FØR SQL
    if name == "AHH":
        h, u, b = 50, 30, 20
    elif name == "RB":
        h, u, b = 70, 20, 10
    elif name == "EB":
        h, u, b = 30, 40, 30
    elif name == "KAF":
        h, u, b = 60, 20, 20
    elif name == "ØG":
        h, u, b = 40, 40, 20
    elif name == "JH":
        h, u, b = 55, 25, 20
    elif name == "TOH":
        h, u, b = 45, 35, 20
    elif name == "UTG":
        h, u, b = 50, 30, 20
    elif name == "LEV":
        h, u, b = 50, 30, 20
    else:
        h, u, b = 50, 30, 20

    # ✅ SQL KUN HER
    for match_number in range(1, 5):
        cursor.execute("""
            INSERT OR IGNORE INTO tips
            (player_id, week_id, match_number, h_percent, u_percent, b_percent)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (player_id, week_id, match_number, h, u, b))

conn.commit()

print("Tips lagt inn!")

print("Spillere lagt inn!")

print("Kamper lagt inn!")

# --- SETT RESULTATER ---
for i in range(1, 5):
    cursor.execute("""
        UPDATE matches
        SET result = 'H'
        WHERE week_id = ? AND match_number = ?
    """, (week_id, i))

conn.commit()

# --- TEST TIPS ---
cursor.execute("SELECT id FROM players WHERE name = ?", ("AHH",))
player_id = cursor.fetchone()[0]


# --- HENT RIKTIG WEEK_ID ---
#---cursor.execute("SELECT id FROM weeks WHERE week_number = ?", (2,))---
#---week_id = cursor.fetchone()[0]---


# Legg inn tips for 4 kamper
for match_number in range(1, 5):
    cursor.execute("""
        INSERT OR IGNORE INTO tips (player_id, week_id, match_number, h_percent, u_percent, b_percent)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (player_id, week_id, match_number, 50, 30, 20))

conn.commit()

print("Test-tips lagt inn!")

# --- HENT RESULTAT ---
# Midlertidig test: sett kamp 1 = H
# Sett resultat på kamp 1–4


# --- SETT RESULTATER FØRST ---
for i in range(1, 5):
    cursor.execute("""

UPDATE matches
SET result = 'H'
WHERE week_id = 1 AND match_number = ?
    """, (i,))

conn.commit()


# --- SÅ henter du tips + resultater ---
JOIN matches m
ON t.week_id = m.week_id
AND t.match_number = m.match_number

rows = cursor.fetchall()

totals = {}
correct = {}

for row in rows:
    print("DEBUG row:", row)   # 🔥 debug her

    player_id, h, u, b, result = row

    # --- POENG ---
    if result == "H":
        percent = h
        max_points = 65
    elif result == "U":
        percent = u
        max_points = 74
    else:
        percent = b
        max_points = 65


points = min(percent, max_points)

print(f"Player {player_id} fikk {points} poeng")

    # --- SUM ---
if player_id not in totals:
        totals[player_id] = 0
totals[player_id] += points

    # --- RETTE ---
if player_id not in correct:
        correct[player_id] = 0

if result == "H" and h > 0:
        correct[player_id] += 1
elif result == "U" and u > 0:
        correct[player_id] += 1
elif result == "B" and b > 0:
        correct[player_id] += 1


    # --- SUM ---
if player_id not in totals:
        totals[player_id] = 0
totals[player_id] += points

    # --- RETTE ---
if player_id not in correct:
        correct[player_id] = 0

if result == "H" and h > 0:
        correct[player_id] += 1
elif result == "U" and u > 0:
        correct[player_id] += 1
elif result == "B" and b > 0:
        correct[player_id] += 1


points = min(percent, max_points)

print(f"Player {player_id} fikk {points} poeng")

for player_id, total in totals.items():
    correct_count = correct[player_id]

    bonus = 0
    if correct_count == 12:
        bonus = 1070
    elif correct_count == 11:
        bonus = 161
    elif correct_count == 10:
        bonus = 58

    cursor.execute("""
        INSERT OR IGNORE INTO weekly_results
        (player_id, week_id, total_points, correct, bonus)
        VALUES (?, ?, ?, ?, ?)
    """, (player_id, week_id, total, correct_count, bonus))

conn.commit()

# --- SUM TOTAL ---
if player_id not in totals:
        totals[player_id] = 0

totals[player_id] += points

    # --- TELLE RETTE ---
if player_id not in correct:
        correct[player_id] = 0

if result == "H" and h > 0:
        correct[player_id] += 1
elif result == "U" and u > 0:
        correct[player_id] += 1
elif result == "B" and b > 0:
        correct[player_id] += 1


# --- DEBUG ---
print("\nDEBUG:")
print(f"Antall rader i rows: {len(rows)}")
print("Totals:", totals)
print("Correct:", correct)

# --- TOTAL ---
print("\nTOTAL POENG:")
for player_id, total in totals.items():
    print(f"Player {player_id}: {total} poeng")

# --- BONUS ---
print("\nBONUS:")
for player_id, count in correct.items():
    bonus = 0

    if count == 12:
        bonus = 1070

    

# --- LEADERBOARD ---
print("\nLEADERBOARD:")

cursor.execute("SELECT id, name FROM players")
player_names = dict(cursor.fetchall())

sorted_players = sorted(
    totals.items(),
    key=lambda x: (x[1], correct[x[0]]),
    reverse=True
)


for i, (player_id, total) in enumerate(sorted_players, start=1):
    name = player_names[player_id]
    print(f"{i}. {name}: {total} poeng")

winner_id, winner_points = sorted_players[0]
winner_name = player_names[winner_id]

print(f"\n🏆 Vinner uke {week_id}: {winner_name} med {winner_points} poeng!")

cursor.execute("""
SELECT 
    p.name,
    SUM(r.total_points) as total,
    SUM(r.correct) as total_correct
FROM weekly_results r
JOIN players p ON r.player_id = p.id
GROUP BY p.name
ORDER BY total DESC, total_correct DESC
""")

for row in cursor.fetchall():
    print(row)




cursor.execute("""
SELECT 
    p.name,
    SUM(r.total_points) as total,
    SUM(r.correct) as total_correct
FROM weekly_results r
JOIN players p ON r.player_id = p.id
GROUP BY p.name
ORDER BY total DESC, total_correct DESC
""")





for i, row in enumerate(cursor.fetchall(), start=1):
    name, total, correct = row
    print(f"{i}. {name}: {total} poeng ({correct} rette)")

import pandas as pd

rows = cursor.fetchall()


# hent data én gang
cursor.execute("""
SELECT 
    p.name,
    SUM(r.total_points) as total,
    SUM(r.correct) as total_correct
FROM weekly_results r
JOIN players p ON r.player_id = p.id
GROUP BY p.name
ORDER BY total DESC, total_correct DESC
""")
rows = cursor.fetchall()

print("\nSAMMENLAGT:")

for i, (name, total, correct) in enumerate(rows, start=1):
    print(f"{i}. {name}: {total} poeng ({correct} rette)")

import pandas as pd

df = pd.DataFrame(rows, columns=["Navn", "Poeng", "Rette"])

df.to_excel("tippelag.xlsx", sheet_name="Sammenlagt", index=False)

unique_rows = set(rows)
print("Unike rader:", len(unique_rows))

# --- HISTORIKK ---
with pd.ExcelWriter("tippelag.xlsx", engine="openpyxl") as writer:

    # --- SAMMENLAGT ---
    df.to_excel(writer, sheet_name="Sammenlagt", index=False)

    # --- HISTORIKK ---
    cursor.execute("""
    SELECT 
        p.name,
        w.week_number,
        r.total_points,
        r.correct
    FROM weekly_results r
    JOIN players p ON r.player_id = p.id
    JOIN weeks w ON r.week_id = w.id
    ORDER BY p.name, w.week_number
    """)

    rows = cursor.fetchall()
    df_hist = pd.DataFrame(rows, columns=["Navn", "Uke", "Poeng", "Rette"])
    df_hist.to_excel(writer, sheet_name="Historikk", index=False)

    # --- ÉN FANE PER SPILLER ---
    for player in df_hist["Navn"].unique():
        df_player = df_hist[df_hist["Navn"] == player]
        df_player.to_excel(writer, sheet_name=player, index=False)

print("✅ Excel med flere ark laget!")
