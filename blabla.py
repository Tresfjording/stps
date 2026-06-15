import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
import mpld3
import io
import base64

file = "stps_tolk.xlsx"

df = pd.read_excel(file, sheet_name="Hovedtabell", engine="openpyxl")

df = df.rename(columns={df.columns[0]: "Navn"})
df = df.rename(columns={df.columns[5]: "tp"})

# 2. Generate your chart with matplotlib.pyplot
fig, ax = plt.subplots(figsize=(7, 4))
ax.bar(df.iloc[:, 0], df.iloc[:, 5], color="#ff0000")
ax.set_title("Årsresultat")
ax.set_xlabel(df.columns[0])
ax.set_ylabel(df.columns[5])
plt.tight_layout()

# 3. Save the plot to an in-memory buffer and encode to base64
img_buffer = io.BytesIO()
plt.savefig(img_buffer, format="png")
img_buffer.seek(0)
img_base64 = base64.b64encode(img_buffer.getvalue()).decode("utf-8")
plt.close(fig)  # Free up system memory

# 4. Convert the pandas DataFrame directly to an HTML table string
html_table = df.to_html(classes="styled-table", index=False )

# 5. Combine everything into a single, clean HTML document
html_content = f"""<!DOCTYPE html>
<html>
<head>
    <title>STPS Data Rapport</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 30px; background-color: #f9f9f9; }}
        .container {{ max-width: 100%; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
        h2 {{ color: #333; border-bottom: 2px solid #3498db; padding-bottom: 5px; }}
        .chart-container {{ text-align: center; margin: 30px 0; }}
        .styled-table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
        .styled-table th {{ background-color: #3498db; color: white; padding: 10px; text-align: left; }}
        .styled-table td {{ padding: 10px; border-bottom: 1px solid #ddd; }}
        .styled-table tr:nth-child(even) {{ background-color: #f2f2f2; }}
    </style>
</head>
<body style="width: 100%">
    <div class="container" style="width: 100%">
        <h2>Skorgen Tippelag 2025/26</h2>
        <div class="chart-container">
            <!-- Embedding the matplotlib figure dynamically using base64 -->
            <img src="data:image/png;base64,{img_base64}" alt="Matplotlib Chart">
        </div>
        
        <h2>Medlemmenenes prestasjoner</h2>
        {html_table}
    </div>
</body>
</html>
"""

# 6. Write the final string content out to a production HTML file
with open("report.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print("HTML report successfully generated as 'report.html'!")
