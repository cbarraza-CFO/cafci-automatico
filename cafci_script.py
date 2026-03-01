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
# USAR PRIMERA COLUMNA COMO FONDO (CLAVE)
# ==========================================================

col_fondo = df.columns[0]

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
# CREAR TIPO_RENTA POR BLOQUE (USANDO PRIMERA COLUMNA)
# ==========================================================

df["Tipo_Renta"] = None
categoria_actual = None

def normalizar(txt):
    return " ".join(str(txt).lower().strip().split())

for i in range(len(df)):
    nombre = normalizar(df[col_fondo].iloc[i])

    if nombre == normalizar("Renta Variable Peso Argentina"):
        categoria_actual = "Renta Variable Peso Argentina"
        continue

    if nombre == normalizar("Renta Fija Peso Argentina"):
        categoria_actual = "Renta Fija Peso Argentina"
        continue

    df.at[i, "Tipo_Renta"] = categoria_actual

# Eliminar filas categoría
df_final = df[
    ~df[col_fondo].apply(
        lambda x: normalizar(x) in [
            normalizar("Renta Variable Peso Argentina"),
            normalizar("Renta Fija Peso Argentina"),
        ]
    )
].copy()

df_final = df_final[df_final[col_fondo].notna()]

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
# HISTORICO
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
    "Tipo_Renta": df_final["Tipo_Renta"],
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

print("Proceso finalizado correctamente.")
