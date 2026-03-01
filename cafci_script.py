import requests
import pandas as pd
from datetime import datetime
import os

URL = "https://api.pub.cafci.org.ar/pb_get?d=1772397842680"

hoy = datetime.now().strftime("%Y-%m-%d")
archivo = f"CAFCI_{hoy}.xlsx"

print("Descargando planilla...")

r = requests.get(URL)
with open(archivo, "wb") as f:
    f.write(r.content)

print("Procesando archivo...")

df = pd.read_excel(archivo)
df = df.dropna(how="all")
df["Fecha"] = hoy

hist_file = "CAFCI_Historico.xlsx"

if os.path.exists(hist_file):
    df_hist = pd.read_excel(hist_file)
    df_total = pd.concat([df_hist, df])
else:
    df_total = df

df_total = df_total.drop_duplicates()
df_total.to_excel(hist_file, index=False)

print("Histórico actualizado correctamente.")
