import requests
import pandas as pd
from datetime import datetime
import os
import re
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
# LECTURA HEADER MULTINIVEL (FILAS 8 Y 9)
# ==========================================================

df = pd.read_excel(archivo, header=[7, 8])
df.columns = [f"{a} {b}".strip().lower() for a, b in df.columns]
df = df.dropna(how="all")

df["fecha"] = hoy_str

# ==========================================================
# USAR PRIMERA COLUMNA COMO FONDO
# ==========================================================

col_fondo = df.columns[0]

# Rellenar filas vacías (si hay filas de continuación)
df[col_fondo] = df[col_fondo].ffill()

# ==========================================================
# IDENTIFICAR COLUMNAS VARIACION %
# ==========================================================

cols_variacion = [c for c in df.columns if "variacion cuotaparte %" in c]

def extraer_fecha(col):
    match = re.search(r"\d{2}/\d{2}/\d{2}", col)
    if match:
        return datetime.strptime(match.group(), "%d/%m/%y")
    return None

cols_fechas = [(c, extraer_fecha(c)) for c in cols_variacion]
cols_fechas = [x for x in cols_fechas if x[1] is not None]
cols_fechas.sort(key=lambda x: x[1], reverse=True)

col_dia = cols_fechas[0][0] if len(cols_fechas) > 0 else None
col_mes = cols_fechas[1][0] if len(cols_fechas) > 1 else None
col_anio = cols_fechas[2][0] if len(cols_fechas) > 2 else None

# ==========================================================
# LIMPIAR %
# ==========================================================

def limpiar_pct(col):
    return (
        col.astype(str)
        .str.replace(",", ".", regex=False)
        .str.replace("%", "", regex=False)
        .replace("nan", None)
        .astype(float)
    )

df["rend_dia"] = limpiar_pct(df[col_dia]) if col_dia else 0
df["rend_mes"] = limpiar_pct(df[col_mes]) if col_mes else 0
df["rend_anio"] = limpiar_pct(df[col_anio]) if col_anio else 0

# ==========================================================
# BUSCAR OTRAS COLUMNAS
# ==========================================================

def buscar_columna(texto):
    for col in df.columns:
        if texto in col:
            return col
    return None

col_moneda = buscar_columna("moneda fondo")
col_plazo = buscar_columna("plazo liq")
col_valor = buscar_columna("valor (mil cuotapartes) actual")

# ==========================================================
# LIMPIAR FILAS CATEGORIA (SIN USAR TIPO_RENTA)
# ==========================================================

df_final = df[
    ~df[col_fondo].str.lower().str.contains("renta variable|renta fija", na=False)
].copy()

df_final = df_final[df_final[col_fondo].notna()]

# ==========================================================
# HISTORICO ACUMULATIVO
# ==========================================================

hist_file = "CAFCI_Historico.xlsx"

if os.path.exists(hist_file):
    df_hist = pd.read_excel(hist_file)
    df_total = pd.concat([df_hist, df_final])
else:
    df_total = df_final

df_total = df_total.drop_duplicates()
df_total.to_excel(hist_file, index=False)

# ==========================================================
# CSV POWER BI
# ==========================================================

df_powerbi = pd.DataFrame({
    "Nombre_Fondo": df_final[col_fondo],
    "Moneda": df_final[col_moneda] if col_moneda else "",
    "Fecha": hoy_str,
    "Plazo_Liquidacion": df_final[col_plazo] if col_plazo else "",
    "Valor_Cuotaparte_Actual": df_final[col_valor] if col_valor else "",
    "Rendimiento_Del_Dia_%": df_final["rend_dia"],
    "Rendimiento_Del_Mes_%": df_final["rend_mes"],
    "Rendimiento_Del_Anio_%": df_final["rend_anio"],
})

df_powerbi.to_csv("FCI_Limpio.csv", index=False)

# ==========================================================
# PDF MONEY MARKET T+0
# ==========================================================

if col_plazo:
    df_mm = df_powerbi[df_powerbi["Plazo_Liquidacion"].astype(str).str.contains("0", na=False)]
else:
    df_mm = df_powerbi.copy()

top10 = df_mm.sort_values("Rendimiento_Del_Dia_%", ascending=False).head(10)

pdf = SimpleDocTemplate("Reporte_MoneyMarket_T0.pdf", pagesize=pagesizes.A4)
elements = []
styles = getSampleStyleSheet()

elements.append(Paragraph(f"Reporte Money Market T+0 - {hoy_str}", styles["Title"]))
elements.append(Spacer(1, 12))

data = [["Fondo", "Rend Día %"]]

for _, row in top10.iterrows():
    data.append([
        str(row["Nombre_Fondo"]),
        round(row["Rendimiento_Del_Dia_%"], 3)
    ])

elements.append(Table(data))
pdf.build(elements)

print("Proceso finalizado correctamente.")
