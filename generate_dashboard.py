from blabla import build_tipper_dashboard_sheet

# Generate dashboard from stps_2026.xlsm (has cached formula values)
# Save to Dashboard.xlsx (plain xlsx = no macro issues, opens in any Excel)
df = build_tipper_dashboard_sheet('stps_2026.xlsm', sheet_name='Dashboard')
print(f'Dashboard.xlsx regenerated with {len(df)} data rows')
print(f'   Tippers: {sorted(df["Tipper"].unique())}')
