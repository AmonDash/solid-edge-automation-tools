import pandas as pd
from datetime import datetime

# === Cargar archivos Excel ===
archivo_anterior = "/content/Report old.xlsx"
archivo_nuevo = "/content/Report new.xlsx"

# Asegúrate de que los encabezados de tus archivos tienen estas columnas:
COLUMNAS = ["File Name", "Quantity", "Modified"]

df_old = pd.read_excel(archivo_anterior, usecols=COLUMNAS)
df_new = pd.read_excel(archivo_nuevo, usecols=COLUMNAS)

# Renombrar columnas para uniformidad
df_old.columns = ["File_Name", "Quantity", "Modified"]
df_new.columns = ["File_Name", "Quantity", "Modified"]

# Convertir columna Modified a datetime si no lo está
df_old["Modified"] = pd.to_datetime(df_old["Modified"], errors='coerce')
df_new["Modified"] = pd.to_datetime(df_new["Modified"], errors='coerce')

# Comparar por nombre de archivo
df_comp = pd.merge(df_old, df_new, on="File_Name", how="outer", suffixes=("_old", "_new"))

# Clasificación de cambio
def clasificar_cambio(row):
    if pd.isna(row['Quantity_old']):
        return "Nueva"
    elif pd.isna(row['Quantity_new']):
        return "Eliminada"
    elif row['Quantity_old'] != row['Quantity_new']:
        return "Cantidad modificada"
    else:
        return "Sin cambio"

def clasificar_modificado(row):
    mod_old = row["Modified_old"]
    mod_new = row["Modified_new"]
    if pd.isna(mod_old) or pd.isna(mod_new):
        return "No"
    return "Sí" if mod_old != mod_new else "No"

def tipo_modificacion(row):
    if row["Cambio"] == "Sin cambio" and row["Modificado"] == "Sí":
        return "Solo fecha modificada"
    elif row["Cambio"] != "Sin cambio" and row["Modificado"] == "Sí":
        return "Cantidad y fecha modificadas"
    elif row["Cambio"] != "Sin cambio":
        return row["Cambio"]
    else:
        return "Sin cambio"

# Aplicar clasificación
df_comp["Cambio"] = df_comp.apply(clasificar_cambio, axis=1)
df_comp["Modificado"] = df_comp.apply(clasificar_modificado, axis=1)
df_comp["Tipo_Modificacion"] = df_comp.apply(tipo_modificacion, axis=1)

# Reorganizar columnas
df_resultado = df_comp[[
    "File_Name", "Quantity_old", "Modified_old",
    "Quantity_new", "Modified_new", "Cambio", "Modificado", "Tipo_Modificacion"
]]

# Filtrar solo las piezas con algún cambio
df_filtrado = df_resultado[
    (df_resultado["Cambio"] != "Sin cambio") | (df_resultado["Modificado"] == "Sí")
]

# === Guardar Excel ===
with pd.ExcelWriter("/content/Comparacion_Cambios_Solo.xlsx", engine='openpyxl') as writer:
    df_filtrado.to_excel(writer, index=False, sheet_name="Cambios Detectados")
    pd.DataFrame({
        "Archivo Comparado (anterior)": [archivo_anterior],
        "Archivo Comparado (nuevo)": [archivo_nuevo],
    }).to_excel(writer, index=False, sheet_name="Info")

print("✅ Comparación finalizada. Archivo generado: Comparacion_Cambios_Solo.xlsx")


