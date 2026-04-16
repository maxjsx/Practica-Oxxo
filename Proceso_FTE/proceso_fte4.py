import pandas as pd
import numpy as np
import re
import os

# ==========================================
# 1. CONFIGURACIÓN Y FUNCIONES
# ==========================================
FECHA_CORTE = pd.to_datetime("2026-03-11")

def limpiar_rut(x):
    if pd.isna(x): return ""
    return re.sub(r'[^0-9Kk]', '', str(x).upper())

def clean_text(x):
    if pd.isna(x): return ""
    return str(x).strip().upper()

print("🚀 Iniciando conciliación final de FTE...")

# ==========================================
# 2. CARGA DE MAESTRA DE TIENDAS (AUTORIZADO)
# ==========================================
print("Cargando matriz de tiendas...")
# Usamos Marzo_26, header en fila 7 (index 6)
df_aut = pd.read_excel("00 - FTE AUTORIZADO.xlsx", sheet_name="MARZO_26 ", header=6)
df_aut.columns = [str(c).strip().upper() for c in df_aut.columns]
df_aut["NOMBRE_KEY"] = df_aut["NOMBRE MAESTRA"].apply(clean_text)

# Diccionarios de Tienda
dict_ceco = dict(zip(df_aut["NOMBRE_KEY"], df_aut["CECO"]))
dict_fte_aut = dict(zip(df_aut["CECO"], pd.to_numeric(df_aut["FTE AUT"], errors='coerce')))
dict_jefe = dict(zip(df_aut["CECO"], df_aut["JEFE OPERACIONES"]))
dict_asesor = dict(zip(df_aut["CECO"], df_aut["ASESOR TIENDA"]))
dict_reclutador = dict(zip(df_aut["CECO"], df_aut["RECLUTADOR"]))
dict_tienda_nom = dict(zip(df_aut["CECO"], df_aut["NOMBRE MAESTRA"]))

# ==========================================
# 3. CARGA DE AGRUPADOR E INCLUSIONES
# ==========================================
print("Cargando reglas del Agrupador...")
df_agr = pd.read_excel("Agrupador 10.xlsx", sheet_name="AGRUPADOR", header=4)
map_agrupa = dict(zip(df_agr["CARGO"].apply(clean_text), df_agr["AGRUPA CARGO_2"]))
map_fte_teo = dict(zip(df_agr["AGRUPA CARGO_2"].apply(clean_text), pd.to_numeric(df_agr["FTE TEORICO"], errors='coerce')))

# Inclusiones: Columna I es FTE INCLS (index 8)
df_incl = pd.read_excel("Agrupador 10.xlsx", sheet_name="INCLS", header=3)
df_incl["RUT_KEY"] = df_incl["Identificador"].apply(limpiar_rut)
df_incl = df_incl.drop_duplicates("RUT_KEY")

dict_fte_incl = dict(zip(df_incl["RUT_KEY"], pd.to_numeric(df_incl["FTE INCLS"], errors='coerce')))
dict_is_incl = dict(zip(df_incl["RUT_KEY"], df_incl["INCLS"]))

# ==========================================
# 4. CARGA DE FILTROS DE PERSONAL
# ==========================================
print("Procesando Filtros de Personal...")
# Activos (solo para info, no para filtrar)
df_act = pd.read_excel("Activos_Inactivos.xlsx")
col_rut_act = "Chile RUN - Rol Único Nacional National ID Information"
df_act["RUT_KEY"] = df_act[col_rut_act].apply(limpiar_rut)
dict_ingreso = dict(zip(df_act["RUT_KEY"], df_act["Employment Details Hire Date"]))

# Bajas (Único filtro de salida permitido)
df_bajas = pd.read_excel("Bajas.xlsx")
df_bajas["RUT_KEY"] = df_bajas["Chile RUN - Rol Único Nacional National ID Information"].apply(limpiar_rut)
df_bajas["FECHA_BAJA"] = pd.to_datetime(df_bajas["Employment Details Termination Date"], errors='coerce')
ruts_baja = set(df_bajas[df_bajas["FECHA_BAJA"] <= FECHA_CORTE]["RUT_KEY"])

# Permisos (Licencias > 15 días)
df_perm = pd.read_excel("PermisosAsignados.xlsx")
df_perm["RUT_KEY"] = df_perm["Rut"].apply(limpiar_rut)
df_perm["F_INI"] = pd.to_datetime(df_perm["Fecha Inicio"], dayfirst=True)
df_perm["F_FIN"] = pd.to_datetime(df_perm["Fecha Fin"], dayfirst=True)
df_perm["DURACION"] = (df_perm["F_FIN"] - df_perm["F_INI"]).dt.days + 1
ruts_licencia_larga = set(df_perm[
    (df_perm["Tipo Permiso"].str.contains("Licencia", case=False, na=False)) & 
    (df_perm["DURACION"] > 15) & (df_perm["F_INI"] <= FECHA_CORTE) & (df_perm["F_FIN"] >= FECHA_CORTE)
]["RUT_KEY"])

# ==========================================
# 5. CONSTRUCCIÓN DE LA SÁBANA (Sheet1)
# ==========================================
print("Construyendo detalle de personas...")
df_ga = pd.read_excel("GestiondeAsistencia.xlsx", header=1)
df_ga["RUT_KEY"] = df_ga["Identificador"].apply(limpiar_rut)

# FILTRO: Solo eliminamos si es una BAJA confirmada.
df_final = df_ga[~df_ga["RUT_KEY"].isin(ruts_baja)].copy()

# Inyección de columnas
df_final.insert(2, "Nombre completo", df_final["Nombre"] + " " + df_final["Apellidos"])
df_final.insert(7, "Ceco", df_final["Grupo"].apply(clean_text).map(dict_ceco))
df_final.insert(9, "Fecha ingreso", df_final["RUT_KEY"].map(dict_ingreso).dt.date)

# Cálculos de FTE
df_final["Agrupador"] = df_final["Cargo"].apply(clean_text).map(map_agrupa)
df_final["FTE TEORICO"] = df_final["Agrupador"].fillna("").astype(str).apply(clean_text).map(map_fte_teo).fillna(0)

# Lógica de FTE Real: Inclusión > Licencia > Teórico
df_final["fte incls"] = df_final["RUT_KEY"].map(dict_fte_incl)
df_final["FTE REAL"] = df_final["fte incls"].fillna(df_final["FTE TEORICO"])
df_final.loc[df_final["RUT_KEY"].isin(ruts_licencia_larga), "FTE REAL"] = 0

# ==========================================
# 6. RESUMEN EJECUTIVO (REPLICANDO HOJA 2)
# ==========================================
print("Generando Resumen de Tiendas...")
resumen = df_final.groupby("Ceco").agg(
    HC=("Identificador", "nunique"),
    FTE_R=("FTE REAL", "sum")
).reset_index()

resumen["FTE AUTORIZADO"] = resumen["Ceco"].map(dict_fte_aut).fillna(0)
resumen["NECESIDAD"] = resumen["FTE_R"] - resumen["FTE AUTORIZADO"]
resumen["EC%"] = np.where(resumen["FTE AUTORIZADO"] == 0, 0, resumen["FTE_R"] / resumen["FTE AUTORIZADO"])

def clasificar_ec(ec):
    ec = round(ec, 4)
    if ec >= 1: return "Completa"
    elif ec >= 0.8: return "Incompleta"
    else: return "Crítica"

resumen["TIPO EC"] = resumen["EC%"].apply(clasificar_ec)
resumen["FECHA"] = FECHA_CORTE.date()
resumen["TIENDA"] = resumen["Ceco"].map(dict_tienda_nom)
resumen["JEFE DISTRITO"] = resumen["Ceco"].map(dict_jefe)
resumen["ASESOR TIENDA"] = resumen["Ceco"].map(dict_asesor)
resumen["ENCARGADO RECLUTAMIENTO"] = resumen["Ceco"].map(dict_reclutador)

# ==========================================
# 7. EXPORTACIÓN
# ==========================================
archivo_salida = f"FTE_CONCILIADO_{FECHA_CORTE.date()}.xlsx"
with pd.ExcelWriter(archivo_salida) as writer:
    df_final.to_excel(writer, sheet_name="Sheet1", index=False)
    resumen.to_excel(writer, sheet_name="Resumen", index=False)

ruta = os.path.abspath(archivo_salida)
print(f"\n✅ CONCILIACIÓN COMPLETADA. Archivo en:\n{ruta}")
print(f"HC Final Detectado: {len(df_final)}")