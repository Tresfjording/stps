import pandas as pd
import streamlit as st

st.title("🏆 Skorgen Tippelag Dashboard")

# ------------------------------
# LES DATA
# ------------------------------
df = pd.read_excel("SkorgenTippelag.xlsm", sheet_name="til py")
df.columns = df.columns.str.lower()

df["dato"] = pd.to_datetime(df["dato"], origin="1899-12-30", unit="D")

# LONG FORMAT
df_long = df.melt(id_vars="dato", var_name="navn", value_name="poeng")
df_long = df_long.dropna()
df_long = df_long[~df_long["navn"].isin(["utg", "lev"])]

# ------------------------------
# LEADERBOARD
# ------------------------------
leaderboard = df_long.groupby("navn")["poeng"].sum().sort_values(ascending=False)

st.header("🏆 Leaderboard")
st.dataframe(leaderboard)

# ------------------------------
# GRAF
# ------------------------------
pivot = df_long.pivot(index="dato", columns="navn", values="poeng").fillna(0)
cum = pivot.cumsum()

st.header("Skorgen Tippelag utvikling")
st.line_chart(pivot)

# ------------------------------
# FORM
# ------------------------------
last_dates = sorted(df_long["dato"].unique())[-3:]

form = df_long[df_long["dato"].isin(last_dates)] \
    .groupby("navn")["poeng"].mean() \
    .sort_values(ascending=False)

st.header("🔥 Form")
st.dataframe(form)

# ------------------------------
# KOMMENTAR
# ------------------------------
st.header("🧠 Analyse")

st.write(f"🏆 {leaderboard.idxmax()} leder ligaen")
st.write(f"🔥 {form.idxmax()} er i best form")
st.write(f"🥶 {leaderboard.idxmin()} ligger sist")