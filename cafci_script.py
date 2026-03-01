import requests
import pandas as pd
from datetime import datetime
import os
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import pagesizes

# ==========================================================
# CONFIG
# ==========================================================

URL = "https://api.pub.cafci.org.ar/pb_get?d=1772397842680"
hoy = datetime.now()
hoy_str = hoy.strftime("%Y-%m-%d")

archivo = f"CAFCI_{hoy_str}.xlsx"

# ==========================================================
# DESCARGA
# ==========================================================

print("Descargando planilla...")
r = requests.get(URL)
with open(archivo, "wb") as f:
    f.write(r.content)

# ==========================================================
# LECTURA Y LIMPIEZA
# ==========================================================

df = pd.read_excel(archivo)
df = df.dropna(how="all")
df.columns = df.columns.str.strip()
df["Fecha"] = hoy_str

# Normalizar rendimiento diario
if "Rendimiento Diario %" in df.columns:
    df["Rendimiento_Diario_%"] = (
        df["Rendimiento Diario %"]
        .astype(str)
        .str.replace(",", ".")
        .str.replace("%", "")
        .astype(float)
    )
else:
    df["Rendimiento_Diario_%"] = 0

# ==========================================================
# HISTORICO ACUMULATIVO
# ==========================================================

hist_file = "CAFCI_Historico.xlsx"

if os.path.exists(hist_file):
    df_hist = pd.read_excel(hist_file)
    df_total = pd.concat([df_hist, df])
else:
    df_total = df

df_total = df_total.drop_duplicates()
df_total.to_excel(hist_file, index=False)

# ==========================================================
# 1️⃣ RESUMEN DÍA / MES / AÑO
# ==========================================================

rend_dia = df["Rendimiento_Diario_%"].mean()

# Mes
df_mes = df_total[
    pd.to_datetime(df_total["Fecha"]).dt.month == hoy.month
]
rend_mes = df_mes["Rendimiento_Diario_%"].mean()

# Año
df_anio = df_total[
    pd.to_datetime(df_total["Fecha"]).dt.year == hoy.year
]
rend_anio = df_anio["Rendimiento_Diario_%"].mean()

df_resumen = pd.DataFrame([{
    "Fecha": hoy_str,
    "Rendimiento_Del_Dia_%": round(rend_dia, 3),
    "Rendimiento_Del_Mes_%": round(rend_mes, 3),
    "Rendimiento_Del_Anio_%": round(rend_anio, 3),
}])

df_resumen.to_excel("Resumen_Rendimientos.xlsx", index=False)

# ==========================================================
# 2️⃣ CSV LIMPIO POWER BI
# ==========================================================

columnas_utiles = [
    col for col in df.columns
    if "Fondo" in col or "Rendimiento" in col or "Patrimonio" in col
]

df_limpio = df[columnas_utiles + ["Fecha"]]
df_limpio.to_csv("FCI_Limpio.csv", index=False)

# ==========================================================
# 3️⃣ PDF EJECUTIVO MONEY MARKET T+0
# ==========================================================

df_mm = df[
    df.astype(str).apply(lambda row: row.str.contains("Money", case=False).any(), axis=1)
]

top10 = df_mm.sort_values("Rendimiento_Diario_%", ascending=False).head(10)

pdf = SimpleDocTemplate(
    "Reporte_MoneyMarket_T0.pdf",
    pagesize=pagesizes.A4
)

elements = []
styles = getSampleStyleSheet()

elements.append(Paragraph(f"Reporte Money Market T+0 - {hoy_str}", styles["Title"]))
elements.append(Spacer(1, 12))

data = [["Fondo", "Rend. Diario %"]]

for _, row in top10.iterrows():
    data.append([
        str(row.get("Fondo", "")),
        round(row.get("Rendimiento_Diario_%", 0), 3)
    ])

table = Table(data)
elements.append(table)

pdf.build(elements)

print("Proceso completo.")
