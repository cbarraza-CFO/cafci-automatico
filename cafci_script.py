import requests
import pandas as pd
from datetime import datetime
import os
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table
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

# Normalizamos nombres
df.columns = df.columns.str.lower()

df["fecha"] = hoy_str

print("Columnas detectadas:")
print(df.columns.tolist())

# ==========================================================
# FUNCIÓN BUSCAR COLUMNA SEGURA
# ==========================================================

def buscar_columna(palabras):
    for col in df.columns:
        for palabra in palabras:
            if palabra in col:
                return col
    return None

# ==========================================================
# DETECTAR COLUMNAS
# ==========================================================

col_fondo = buscar_columna(["fondo", "denominacion", "nombre"])
col_moneda = buscar_columna(["moneda"])
col_plazo = buscar_columna(["plazo", "liquid"])

col_dia = buscar_columna(["diario"])
col_mes = buscar_columna(["mes"])
col_anio = buscar_columna(["año", "anio"])

if not col_fondo:
    raise Exception("No se encontró columna de nombre de fondo")

# ==========================================================
# LIMPIEZA %
# ==========================================================

def limpiar_pct(col):
    return (
        col.astype(str)
        .str.replace(",", ".", regex=False)
        .str.replace("%", "", regex=False)
        .astype(float)
    )

df["rend_dia"] = limpiar_pct(df[col_dia])
df["rend_mes"] = limpiar_pct(df[col_mes])
df["rend_anio"] = limpiar_pct(df[col_anio])

# ==========================================================
# HISTÓRICO
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
# RESUMEN
# ==========================================================

rend_dia = df["rend_dia"].mean()

df_total["fecha"] = pd.to_datetime(df_total["fecha"])
rend_mes = df_total[df_total["fecha"].dt.month == hoy.month]["rend_dia"].mean()
rend_anio = df_total[df_total["fecha"].dt.year == hoy.year]["rend_dia"].mean()

df_resumen = pd.DataFrame([{
    "fecha": hoy_str,
    "rendimiento_dia_%": round(rend_dia,3),
    "rendimiento_mes_%": round(rend_mes,3),
    "rendimiento_anio_%": round(rend_anio,3)
}])

df_resumen.to_excel("Resumen_Rendimientos.xlsx", index=False)

# ==========================================================
# CSV LIMPIO POWER BI
# ==========================================================

df_powerbi = pd.DataFrame({
    "Nombre_Fondo": df[col_fondo],
    "Moneda": df[col_moneda] if col_moneda else "",
    "Fecha": hoy_str,
    "Plazo_Liquidacion": df[col_plazo] if col_plazo else "",
    "Rendimiento_Del_Dia_%": df["rend_dia"],
    "Rendimiento_Del_Mes_%": df["rend_mes"],
    "Rendimiento_Del_Anio_%": df["rend_anio"],
})

df_powerbi.to_csv("FCI_Limpio.csv", index=False)

# ==========================================================
# PDF MONEY MARKET T+0
# ==========================================================

if col_plazo:
    df_mm = df[df[col_plazo].astype(str).str.contains("0", na=False)]
else:
    df_mm = df.copy()

top10 = df_mm.sort_values("rend_dia", ascending=False).head(10)

pdf = SimpleDocTemplate("Reporte_MoneyMarket_T0.pdf", pagesize=pagesizes.A4)
elements = []
styles = getSampleStyleSheet()

elements.append(Paragraph(f"Reporte Money Market T+0 - {hoy_str}", styles["Title"]))
elements.append(Spacer(1, 12))

data = [["Fondo", "Rend Día %"]]

for _, row in top10.iterrows():
    data.append([
        str(row[col_fondo]),
        round(row["rend_dia"],3)
    ])

table = Table(data)
elements.append(table)

pdf.build(elements)

print("Proceso finalizado correctamente.")
