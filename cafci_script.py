import requests
import pandas as pd
from datetime import datetime
import os
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import pagesizes

# ==========================================================
# CONFIGURACIÓN
# ==========================================================

URL = "https://api.pub.cafci.org.ar/pb_get?d=1772397842680"

hoy = datetime.now()
hoy_str = hoy.strftime("%Y-%m-%d")

archivo = f"CAFCI_{hoy_str}.xlsx"

# ==========================================================
# DESCARGA
# ==========================================================

print("Descargando planilla CAFCI...")
r = requests.get(URL)
with open(archivo, "wb") as f:
    f.write(r.content)

# ==========================================================
# LECTURA
# ==========================================================

df = pd.read_excel(archivo)
df = df.dropna(how="all")
df.columns = df.columns.str.strip()
df["Fecha"] = hoy_str

# ==========================================================
# FUNCIÓN LIMPIEZA %
# ==========================================================

def limpiar_pct(col):
    return (
        col.astype(str)
        .str.replace(",", ".", regex=False)
        .str.replace("%", "", regex=False)
        .astype(float)
    )

# ==========================================================
# DETECCIÓN DINÁMICA DE COLUMNAS
# ==========================================================

col_fondo = next(c for c in df.columns if "Fondo" in c)
col_moneda = next(c for c in df.columns if "Moneda" in c)
col_plazo = next(c for c in df.columns if "Plazo" in c or "Liquid" in c)

col_dia = next(c for c in df.columns if "Diario" in c)
col_mes = next(c for c in df.columns if "Mes" in c)
col_anio = next(c for c in df.columns if "Año" in c or "Anio" in c)

# ==========================================================
# NORMALIZAR RENDIMIENTOS
# ==========================================================

df["Rendimiento_Del_Dia_%"] = limpiar_pct(df[col_dia])
df["Rendimiento_Del_Mes_%"] = limpiar_pct(df[col_mes])
df["Rendimiento_Del_Anio_%"] = limpiar_pct(df[col_anio])

# ==========================================================
# HISTÓRICO ACUMULATIVO
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

rend_dia = df["Rendimiento_Del_Dia_%"].mean()

df_mes = df_total[
    pd.to_datetime(df_total["Fecha"]).dt.month == hoy.month
]
rend_mes = df_mes["Rendimiento_Del_Dia_%"].mean()

df_anio = df_total[
    pd.to_datetime(df_total["Fecha"]).dt.year == hoy.year
]
rend_anio = df_anio["Rendimiento_Del_Dia_%"].mean()

df_resumen = pd.DataFrame([{
    "Fecha": hoy_str,
    "Rendimiento_Del_Dia_%": round(rend_dia, 3),
    "Rendimiento_Del_Mes_%": round(rend_mes, 3),
    "Rendimiento_Del_Anio_%": round(rend_anio, 3),
}])

df_resumen.to_excel("Resumen_Rendimientos.xlsx", index=False)

# ==========================================================
# 2️⃣ CSV LIMPIO PROFESIONAL POWER BI
# ==========================================================

df_powerbi = pd.DataFrame({
    "Nombre_Fondo": df[col_fondo],
    "Moneda": df[col_moneda],
    "Fecha": hoy_str,
    "Plazo_Liquidacion": df[col_plazo],
    "Rendimiento_Del_Dia_%": df["Rendimiento_Del_Dia_%"],
    "Rendimiento_Del_Mes_%": df["Rendimiento_Del_Mes_%"],
    "Rendimiento_Del_Anio_%": df["Rendimiento_Del_Anio_%"],
})

df_powerbi.to_csv("FCI_Limpio.csv", index=False)

# ==========================================================
# 3️⃣ PDF EJECUTIVO MONEY MARKET T+0
# ==========================================================

df_mm = df[
    df[col_plazo].astype(str).str.contains("0", na=False)
]

top10 = df_mm.sort_values("Rendimiento_Del_Dia_%", ascending=False).head(10)

pdf = SimpleDocTemplate(
    "Reporte_MoneyMarket_T0.pdf",
    pagesize=pagesizes.A4
)

elements = []
styles = getSampleStyleSheet()

elements.append(Paragraph(f"Reporte Money Market T+0 - {hoy_str}", styles["Title"]))
elements.append(Spacer(1, 12))

data = [["Fondo", "Moneda", "Rend. Día %"]]

for _, row in top10.iterrows():
    data.append([
        str(row[col_fondo]),
        str(row[col_moneda]),
        round(row["Rendimiento_Del_Dia_%"], 3)
    ])

table = Table(data)
elements.append(table)

pdf.build(elements)

print("Proceso completo. Archivos generados correctamente.")
