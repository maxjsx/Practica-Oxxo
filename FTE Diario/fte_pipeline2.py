import pandas as pd
import numpy as np
import re
from datetime import datetime

# ==========================================
# 1. CONFIGURACIÓN Y RUTAS DE ARCHIVOS
# ==========================================
FECHA_CORTE = pd.to_datetime("2026-03-05") # Cambiar a la fecha de hoy

PATH_GA    = "GestionAsistencia.xlsx"
PATH_PERM  = "PermisosAsignados.xlsx"
PATH_AGR   = "Agrupador 10.xlsx"
PATH_FTE   = "FTE_Autorizado.xlsx"
PATH_BAJAS = "Bajas.xlsx"

OUTPUT_NAME = f"HC_FTE_FINAL_{FECHA_CORTE.strftime('%Y-%m-%d')}.xlsx"

# ==========================================
# 2. FUNCIONES DE LIMPIEZA
# ==========================================
def clean_rut(x):
    if pd.isna(x): return ""
    return str(x).upper().replace(".", "").replace("-", "").replace(" ", "").strip()

def norm_text(x):
    if pd.isna(x): return ""
    s = str(x).upper().replace("\t", " ").strip()
    return re.sub(r"\s+", " ", s)

# ==========================================
# 3. CARGA Y PROCESAMIENTO DE DATOS
# ==========================================

# A. AGRUPADOR (Lógica de tu jefe: Col 15 -> 17 -> 19)
# Leemos el rango O5:S68 (Hoja1)
df_agr_raw = pd.read_excel(PATH_AGR, sheet_name="Hoja1", header=None)
df_tabla_agr = df_agr_raw.iloc[4:68, 14:19].copy() # Index 14=O, 18=S
df_tabla_agr.columns = ["CARGO_RAW", "X1", "AGRUPA_CARGO_2", "X2", "FTE_TEO"]

df_tabla_agr["CARGO_RAW_KEY"] = df_tabla_agr["CARGO_RAW"].apply(norm_text)
df_tabla_agr["AGRUPA2_KEY"]   = df_tabla_agr["AGRUPA_CARGO_2"].apply(norm_text)

# Mapas de búsqueda
map_cargo_to_agrupa2 = df_tabla_agr.drop_duplicates("CARGO_RAW_KEY").set_index("CARGO_RAW_KEY")["AGRUPA_CARGO_2"]
map_agrupa2_to_fte   = df_tabla_agr.drop_duplicates("AGRUPA2_KEY").set_index("AGRUPA2_KEY")["FTE_TEO"]

# B. INCLUSIONES (Hoja INCLS: D=RUT, I=FTE)
df_incl = pd.read_excel(PATH_AGR, sheet_name="INCLS")
df_incl["RUT_KEY"] = df_incl.iloc[:, 3].apply(clean_rut) # Col D
map_incl = df_incl.set_index("RUT_KEY").iloc[:, 7] # Col I (FTE Incl)

# C. BAJAS (Filtro previo)
df_bajas = pd.read_excel(PATH_BAJAS)
# Identificamos columnas por nombre común (ajustar si varían)
col_rut_baja = [c for c in df_bajas.columns if "USER" in str(c).upper() or "ID" in str(c).upper()][0]
col_fecha_baja = [c for c in df_bajas.columns if "TERMINATION" in str(c).upper() or "FECHA" in str(c).upper()][0]

df_bajas["RUT_KEY"] = df_bajas[col_rut_baja].apply(clean_rut)
df_bajas[col_fecha_baja] = pd.to_datetime(df_bajas[col_fecha_baja])
ruts_baja = df_bajas[df_bajas[col_fecha_baja] <= FECHA_CORTE]["RUT_KEY"].unique()

# D. PERMISOS (Regla estricta de Licencia Médica > 15 días)
df_perm = pd.read_excel(PATH_PERM)
df_perm["RUT_KEY"] = df_perm.iloc[:, 0].apply(clean_rut) # Asumiendo Col A es RUT
df_perm["F_INI"] = pd.to_datetime(df_perm["DESDE"]) # Ajustar nombres
df_perm["F_FIN"] = pd.to_datetime(df_perm["HASTA"])
df_perm["TIPO"]  = df_perm["TIPO"].fillna("").str.upper()

# Duración: Fin - Inicio + 1
df_perm["DURACION"] = (df_perm["F_FIN"] - df_perm["F_INI"]).dt.days + 1

# Filtramos solo Licencias Médicas activas hoy con Duración > 15
licencias_largas = df_perm[
    (df_perm["TIPO"].str.contains("LICENCIA MÉDICA|LICENCIA MEDICA")) &
    (df_perm["F_INI"] <= FECHA_CORTE) & (df_perm["F_FIN"] >= FECHA_CORTE) &
    (df_perm["DURACION"] > 15)
]["RUT_KEY"].unique()

# E. GESTIÓN ASISTENCIA (Base Principal)
df_ga = pd.read_excel(PATH_GA, header=1) # Header suele estar en fila 2
df_ga["RUT_KEY"] = df_ga["IDENTIFICADOR"].apply(clean_rut)
df_ga["CARGO_KEY"] = df_ga["CARGO"].apply(norm_text)

# ==========================================
# 4. APLICACIÓN DE REGLAS DE NEGOCIO
# ==========================================

# 1. Eliminar Bajas
df_final = df_ga[~df_ga["RUT_KEY"].isin(ruts_baja)].copy()

# 2. Asignar Agrupa_Cargo_2 y FTE Teórico
df_final["AGRUPA_CARGO_2"] = df_final["CARGO_KEY"].map(map_cargo_to_agrupa2)
df_final["FTE_TEORICO"]    = df_final["AGRUPA_CARGO_2"].map(map_agrupa2_to_fte).fillna(0)

# 3. Prioridad Inclusión
df_final["FTE_REAL"] = df_final["RUT_KEY"].map(map_incl).fillna(df_final["FTE_TEORICO"])

# 4. Regla de Licencia Médica (Si es >15 días, FTE Real = 0)
df_final.loc[df_final["RUT_KEY"].isin(licencias_largas), "FTE_REAL"] = 0

# ==========================================
# 5. SALIDA A EXCEL
# ==========================================
with pd.ExcelWriter(OUTPUT_NAME) as writer:
    df_final.to_excel(writer, sheet_name="DETALLE_FTE", index=False)
    
    # Resumen rápido
    resumen = pd.DataFrame({
        "Métrica": ["HC Total", "FTE Real Total", "Licencias > 15 días"],
        "Valor": [df_final["RUT_KEY"].nunique(), df_final["FTE_REAL"].sum(), len(licencias_largas)]
    })
    resumen.to_excel(writer, sheet_name="RESUMEN", index=False)

print(f"Proceso terminado. Archivo guardado como: {OUTPUT_NAME}")