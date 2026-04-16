import pandas as pd
import numpy as np
import re
import os

# ==========================================
# 1. CONFIGURACIÓN Y FECHA
# ==========================================
FECHA_CORTE = pd.to_datetime("2026-03-05")

def limpiar_rut(x):
    if pd.isna(x): return ""
    return re.sub(r'[^0-9Kk]', '', str(x).upper())

# ==========================================
# 2. CARGA DEL AGRUPADOR (Hoja: AGRUPADOR)
# ==========================================
print("Procesando Agrupador...")
# Según tu archivo, la tabla real empieza en la fila 5 (header=4)
df_agr = pd.read_excel("Agrupador 10.xlsx", sheet_name="AGRUPADOR", header=4)
df_agr.columns = [str(c).strip() for c in df_agr.columns]

map_cargo_to_agrupa2 = df_agr.drop_duplicates("CARGO").set_index("CARGO")["AGRUPA CARGO_2"]
map_agrupa2_to_fte = df_agr.drop_duplicates("AGRUPA CARGO_2").set_index("AGRUPA CARGO_2")["FTE TEORICO"]

# Inclusiones (Hoja INCLS, header en fila 4)
df_incl = pd.read_excel("Agrupador 10.xlsx", sheet_name="INCLS", header=3)
df_incl["RUT_KEY"] = df_incl["Identificador"].apply(limpiar_rut)
map_incl = df_incl.drop_duplicates("RUT_KEY").set_index("RUT_KEY")["FTE INCLS"]

# ==========================================
# 3. CARGA DE FILTROS (ACTIVOS, BAJAS, PERMISOS)
# ==========================================
print("Cargando filtros de personal...")

# Activos
df_act = pd.read_excel("Activos_Inactivos.xlsx")
col_rut_per = "Chile RUN - Rol Único Nacional National ID Information"
ruts_activos = set(df_act[df_act["Employee Status"] == "Active"][col_rut_per].apply(limpiar_rut))

# Bajas
df_bajas = pd.read_excel("Bajas.xlsx")
df_bajas["FECHA_BAJA"] = pd.to_datetime(df_bajas["Employment Details Termination Date"], errors='coerce')
ruts_baja = set(df_bajas[df_bajas["FECHA_BAJA"] <= FECHA_CORTE][col_rut_per].apply(limpiar_rut))

# Permisos (Licencias > 15 días)
df_perm = pd.read_excel("PermisosAsignados.xlsx")
df_perm["F_INI"] = pd.to_datetime(df_perm["Fecha Inicio"], dayfirst=True)
df_perm["F_FIN"] = pd.to_datetime(df_perm["Fecha Fin"], dayfirst=True)
df_perm["DURACION"] = (df_perm["F_FIN"] - df_perm["F_INI"]).dt.days + 1
ruts_licencia_larga = set(df_perm[
    (df_perm["Tipo Permiso"].str.contains("Licencia", na=False)) & 
    (df_perm["DURACION"] > 15) & 
    (df_perm["F_INI"] <= FECHA_CORTE) & (df_perm["F_FIN"] >= FECHA_CORTE)
]["Rut"].apply(limpiar_rut))

# ==========================================
# 4. BASE PRINCIPAL (GESTIÓN ASISTENCIA)
# ==========================================
print("Procesando Asistencia...")
df_ga = pd.read_excel("GestiondeAsistencia.xlsx", header=1)
df_ga["RUT_KEY"] = df_ga["Identificador"].apply(limpiar_rut)
df_ga["CECO_KEY"] = df_ga["Grupo"].apply(limpiar_rut) # Usamos Grupo como CECO

# Filtro de Seguridad
df_final = df_ga[df_ga["RUT_KEY"].isin(ruts_activos)].copy()
df_final = df_final[~df_final["RUT_KEY"].isin(ruts_baja)]

# Mapeo de FTE
df_final["AGRUPA_CARGO_2"] = df_final["Cargo"].map(map_cargo_to_agrupa2)
df_final["FTE_TEORICO"] = df_final["AGRUPA_CARGO_2"].map(map_agrupa2_to_fte).fillna(0)
df_final["FTE_REAL"] = df_final["RUT_KEY"].map(map_incl).fillna(df_final["FTE_TEORICO"])
df_final.loc[df_final["RUT_KEY"].isin(ruts_licencia_larga), "FTE_REAL"] = 0

# ==========================================
# 5. FTE AUTORIZADO (RESUMEN POR TIENDA)
# ==========================================
print("Cruzando con FTE Autorizado...")
# En tu archivo, la hoja es "MARZO_26 " (con un espacio al final) y el header en fila 7
df_aut = pd.read_excel("00 - FTE AUTORIZADO.xlsx", sheet_name="MARZO_26 ", header=6)
df_aut.columns = [str(c).strip() for c in df_aut.columns]
df_aut = df_aut[df_aut["ESTADO RH"] == "Abierta"] 

map_autorizado = df_aut.set_index("CECO")["FTE AUT"].to_dict()
map_nombres = df_aut.set_index("CECO")["NOMBRE MAESTRA"].to_dict()

resumen = df_final.groupby("CECO_KEY").agg({
    "RUT_KEY": "nunique",
    "FTE_REAL": "sum"
}).rename(columns={"RUT_KEY": "HC_REAL", "FTE_REAL": "FTE_EJECUTADO"})

resumen["NOMBRE_TIENDA"] = resumen.index.map(map_nombres)
resumen["FTE_AUTORIZADO"] = resumen.index.map(map_autorizado)
resumen["DIFERENCIA"] = resumen["FTE_EJECUTADO"] - resumen["FTE_AUTORIZADO"]

# ==========================================
# 6. EXPORTAR
# ==========================================
output_file = f"REPORTE_FTE_FINAL_{FECHA_CORTE.date()}.xlsx"
with pd.ExcelWriter(output_file) as writer:
    df_final.to_excel(writer, sheet_name="Detalle", index=False)
    resumen.reset_index().to_excel(writer, sheet_name="Resumen_Tiendas", index=False)

print(f"¡Proceso completado! Archivo creado: {output_file}")