import pandas as pd
import numpy as np
import re
import os
from datetime import datetime

# ==========================================
# 1. CONFIGURACIÓN
# ==========================================
FECHA_CORTE = pd.to_datetime("2026-03-05") 

CONFIG = {
    "ACTIVOS": "Activos_Inactivos.xlsx",
    "AGR": "Agrupador 10.xlsx",
    "PERM": "PermisosAsignados.xlsx",
    "BAJAS": "Bajas.xlsx",
    "AUT": "00 - FTE AUTORIZADO.xlsx"
}

OUTPUT_NAME = f"HC_FTE_COMPLETO_{FECHA_CORTE.strftime('%Y-%m-%d')}.xlsx"

# ==========================================
# 2. FUNCIONES DE APOYO
# ==========================================
def clean_rut(x):
    if pd.isna(x): return ""
    # Elimina todo lo que no sea número o letra K
    s = str(x).upper()
    s = re.sub(r'[^0-9K]', '', s)
    return s.strip()

def norm_text(x):
    if pd.isna(x): return ""
    s = str(x).upper().replace("\t", " ").strip()
    return re.sub(r"\s+", " ", s)

def find_column(df, keywords):
    for col in df.columns:
        if any(key in str(col).upper() for key in keywords):
            return col
    return None

# ==========================================
# 3. CARGA DE DATOS MAESTROS
# ==========================================
print("--- Iniciando Proceso ---")

# A. AGRUPADOR (Mapeo de Cargos - Versión Definitiva)
print("Buscando tabla en Agrupador...")

df_tabla_agr = None

# Lista de hojas donde podría estar la tabla
hojas_posibles = ["Hoja1", "AGRUPADOR"]
# Intentamos obtener las hojas reales del archivo para no fallar por nombres
try:
    xl_temp = pd.ExcelFile(CONFIG["AGR"])
    hojas_reales = xl_temp.sheet_names
    for h in hojas_reales:
        if h not in hojas_posibles:
            hojas_posibles.append(h)
except:
    pass

for nombre_hoja in hojas_posibles:
    try:
        print(f"Probando en hoja: {nombre_hoja}...")
        df_temp = pd.read_excel(CONFIG["AGR"], sheet_name=nombre_hoja, header=None)
        
        # Buscamos la fila que tiene los encabezados
        for i, row in df_temp.iterrows():
            # Convertimos toda la fila a texto para buscar las palabras clave
            row_vals = [str(v).upper().strip() for v in row.values]
            if "CARGO" in row_vals and "AGRUPA_CARGO_2" in row_vals:
                print(f"¡Tabla encontrada en {nombre_hoja}, fila {i+1}!")
                df_tabla_agr = pd.read_excel(CONFIG["AGR"], sheet_name=nombre_hoja, header=i)
                break
        if df_tabla_agr is not None:
            break
    except Exception as e:
        continue

if df_tabla_agr is None:
    raise ValueError(f"CRÍTICO: No se encontró la tabla con 'CARGO' y 'AGRUPA_CARGO_2' en ninguna hoja de {CONFIG['AGR']}.")

# Limpiar nombres de columnas
df_tabla_agr.columns = [str(c).strip().upper() for c in df_tabla_agr.columns]

# Extraer columnas por nombre
col_raw_name = [c for c in df_tabla_agr.columns if "CARGO" == c or "CARGO_RAW" in c][0]
col_agp_name = [c for c in df_tabla_agr.columns if "AGRUPA_CARGO_2" in c][0]
# El FTE suele ser la última columna del bloque
col_fte_candidates = [c for c in df_tabla_agr.columns if "FTE" in c]
col_fte_name = col_fte_candidates[-1] 

# Crear dataframe de mapeo limpio
df_mapeo = df_tabla_agr[[col_raw_name, col_agp_name, col_fte_name]].copy()
df_mapeo.columns = ["CARGO_RAW", "AGRUPA_CARGO_2", "FTE_TEO"]

# Limpiar datos y convertir FTE a número
df_mapeo["CARGO_RAW_KEY"] = df_mapeo["CARGO_RAW"].apply(norm_text)
df_mapeo["AGRUPA2_KEY"]   = df_mapeo["AGRUPA_CARGO_2"].apply(norm_text)
df_mapeo["FTE_TEO"]       = pd.to_numeric(df_mapeo["FTE_TEO"], errors="coerce").fillna(0)

# Mapas finales
map_cargo_to_agrupa2 = df_mapeo.drop_duplicates("CARGO_RAW_KEY").set_index("CARGO_RAW_KEY")["AGRUPA_CARGO_2"]
map_agrupa2_to_fte   = df_mapeo.drop_duplicates("AGRUPA2_KEY").set_index("AGRUPA2_KEY")["FTE_TEO"]

print("Tabla de Agrupador cargada exitosamente.")

# B. INCLUSIONES
df_incl = pd.read_excel(CONFIG["AGR"], sheet_name="INCLS")
col_rut_incl = find_column(df_incl, ["RUT", "IDENTIFICADOR"])
# Buscamos la columna de FTE (suele ser la 9na columna, índice 8)
map_incl = df_incl.drop_duplicates(subset=[col_rut_incl])
map_incl["RUT_KEY"] = map_incl[col_rut_incl].apply(clean_rut)
map_incl = map_incl.set_index("RUT_KEY").iloc[:, 7] # Columna I aproximadamente

# C. FILTROS (Activos y Bajas)
df_act = pd.read_excel(CONFIG["ACTIVOS"])
c_rut_act = find_column(df_act, ["RUT", "IDENTIFICADOR"])
c_est_act = find_column(df_act, ["ESTADO", "STATUS"])
df_act["RUT_KEY"] = df_act[c_rut_act].apply(clean_rut)
ruts_activos = set(df_act[df_act[c_est_act].astype(str).str.upper().str.contains("ACTIVO", na=False)]["RUT_KEY"])

df_bajas = pd.read_excel(CONFIG["BAJAS"])
c_rut_baj = find_column(df_bajas, ["RUT", "USER", "ID"])
c_fec_baj = find_column(df_bajas, ["TERMINATION", "FECHA", "BAJA"])
df_bajas["RUT_KEY"] = df_bajas[c_rut_baj].apply(clean_rut)
df_bajas["FECHA_DT"] = pd.to_datetime(df_bajas[c_fec_baj], errors='coerce')
ruts_baja = set(df_bajas[df_bajas["FECHA_DT"] <= FECHA_CORTE]["RUT_KEY"])

# D. LICENCIAS
df_perm = pd.read_excel(CONFIG["PERM"])
c_rut_per = find_column(df_perm, ["RUT", "IDENTIFICADOR"])
df_perm["RUT_KEY"] = df_perm[c_rut_per].apply(clean_rut)
df_perm["F_INI"] = pd.to_datetime(df_perm.iloc[:, 1], dayfirst=True, errors='coerce')
df_perm["F_FIN"] = pd.to_datetime(df_perm.iloc[:, 2], dayfirst=True, errors='coerce')
df_perm["DURACION"] = (df_perm["F_FIN"] - df_perm["F_INI"]).dt.days + 1
lic_cero = set(df_perm[(df_perm.iloc[:, 3].astype(str).str.contains("LICENCIA", case=False)) & (df_perm["DURACION"] > 15) & (df_perm["F_INI"] <= FECHA_CORTE) & (df_perm["F_FIN"] >= FECHA_CORTE)]["RUT_KEY"])

# ==========================================
# 4. PROCESO BASE (GESTIÓN ASISTENCIA)
# ==========================================
archivo_ga = next((f for f in os.listdir(".") if "GESTION" in f.upper().replace("Ó","O") and f.endswith(".xlsx")), None)
df_ga = pd.read_excel(archivo_ga, header=1)

# Identificar columnas en GA
c_rut_ga = find_column(df_ga, ["IDENTIFICADOR", "RUT"])
c_car_ga = find_column(df_ga, ["CARGO"])
c_cec_ga = find_column(df_ga, ["GRUPO", "CECO"])

df_ga["RUT_KEY"] = df_ga[c_rut_ga].apply(clean_rut)
df_ga["CARGO_KEY"] = df_ga[c_car_ga].apply(norm_text)
df_ga["CECO_KEY"] = df_ga[c_cec_ga].apply(clean_rut)

# APLICAR FILTROS (Aquí estaba el error de vacío)
# Solo si el RUT está en activos y NO está en bajas
df_final = df_ga[df_ga["RUT_KEY"].isin(ruts_activos)].copy()
df_final = df_final[~df_final["RUT_KEY"].isin(ruts_baja)]

print(f"Personas procesadas tras filtros: {len(df_final)}")

# Mapeos
df_final["AGRUPA_CARGO_2"] = df_final["CARGO_KEY"].map(map_cargo_to_agrupa2)
df_final["FTE_TEORICO"] = df_final["AGRUPA_CARGO_2"].map(map_agrupa2_to_fte).fillna(0)
df_final["FTE_REAL"] = df_final["RUT_KEY"].map(map_incl).fillna(df_final["FTE_TEORICO"])
df_final.loc[df_final["RUT_KEY"].isin(lic_cero), "FTE_REAL"] = 0

# ==========================================
# 5. RESUMEN TIENDAS
# ==========================================
archivo_aut = next((f for f in os.listdir(".") if "AUTORIZADO" in f.upper() and f.endswith(".xlsx")), CONFIG["AUT"])
df_fa_raw = pd.read_excel(archivo_aut, sheet_name="FEB_26", header=6)
df_fa_raw["CECO_KEY"] = df_fa_raw["CECO"].apply(clean_rut)

resumen_tiendas = df_final.groupby("CECO_KEY").agg({"RUT_KEY": "nunique", "FTE_REAL": "sum"}).rename(columns={"RUT_KEY": "HC_REAL", "FTE_REAL": "FTE_EJECUTADO"})
resumen_tiendas["FTE_AUT_META"] = resumen_tiendas.index.map(df_fa_raw.set_index("CECO_KEY")["FTE AUT"])
resumen_tiendas["CUMPLIMIENTO_%"] = (resumen_tiendas["FTE_EJECUTADO"] / resumen_tiendas["FTE_AUT_META"]) * 100

# ==========================================
# 6. SALIDA
# ==========================================
with pd.ExcelWriter(OUTPUT_NAME) as writer:
    df_final.to_excel(writer, sheet_name="DETALLE_PERSONAS", index=False)
    resumen_tiendas.reset_index().to_excel(writer, sheet_name="RESUMEN_TIENDAS", index=False)
    pd.DataFrame({"Métrica": ["HC Total", "FTE Total"], "Valor": [len(df_final), df_final["FTE_REAL"].sum()]}).to_excel(writer, sheet_name="RESUMEN_EJECUTIVO", index=False)

print("¡Hecho! Revisa el archivo generado.")