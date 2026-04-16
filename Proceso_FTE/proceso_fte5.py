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

# Función a prueba de balas para convertir "0,68" o "0.68" en números reales
def to_float_safe(val):
    if pd.isna(val) or str(val).strip() == "": 
        return np.nan
    val_str = str(val).strip().replace(',', '.')
    try:
        return float(val_str)
    except:
        return np.nan

print("🚀 Iniciando conciliación final de FTE (V5.0 - Ajuste Fino)...")

# ==========================================
# 2. CARGA DE MAESTRA DE TIENDAS (AUTORIZADO)
# ==========================================
print("Cargando matriz de tiendas...")
df_aut = pd.read_excel("00 - FTE AUTORIZADO.xlsx", sheet_name="MARZO_26 ", header=6)
df_aut.columns = [str(c).strip().upper() for c in df_aut.columns]
df_aut["NOMBRE_KEY"] = df_aut["NOMBRE MAESTRA"].apply(clean_text)

dict_ceco = dict(zip(df_aut["NOMBRE_KEY"], df_aut["CECO"]))
dict_fte_aut = dict(zip(df_aut["CECO"], df_aut["FTE AUT"].apply(to_float_safe)))
dict_jefe = dict(zip(df_aut["CECO"], df_aut["JEFE OPERACIONES"]))
dict_asesor = dict(zip(df_aut["CECO"], df_aut["ASESOR TIENDA"]))
dict_reclutador = dict(zip(df_aut["CECO"], df_aut["RECLUTADOR"]))
dict_tienda_nom = dict(zip(df_aut["CECO"], df_aut["NOMBRE MAESTRA"]))

# ==========================================
# 3. CARGA DE AGRUPADOR E INCLUSIONES (DOBLE MAPEO)
# ==========================================
print("Cargando reglas del Agrupador y diccionario de variaciones...")

# A) Mapeo de Variaciones (Hoja1) -> Traduce "Cajero Uber..." a "CAJERO PT 30"
try:
    df_map_raw = pd.read_excel("Agrupador 10.xlsx", sheet_name="Hoja1")
    df_map_raw.columns = [str(c).strip().upper() for c in df_map_raw.columns]
    map_cargo_to_agrupa2 = dict(zip(df_map_raw["CARGO"].apply(clean_text), df_map_raw["AGRUPA CARGO_2"].apply(clean_text)))
except:
    map_cargo_to_agrupa2 = {}

# B) Mapeo de FTE Real (Hoja: AGRUPADOR) -> Asigna el valor numérico
df_agr = pd.read_excel("Agrupador 10.xlsx", sheet_name="AGRUPADOR", header=4)
df_agr.columns = [str(c).strip().upper() for c in df_agr.columns]
# Agregamos también los cargos directos por si no estaban en la Hoja1
map_cargo_to_agrupa2.update(dict(zip(df_agr["CARGO"].apply(clean_text), df_agr["AGRUPA CARGO_2"].apply(clean_text))))
map_agrupa2_to_fte = dict(zip(df_agr["AGRUPA CARGO_2"].apply(clean_text), df_agr["FTE TEORICO"].apply(to_float_safe)))

# C) Inclusiones
df_incl = pd.read_excel("Agrupador 10.xlsx", sheet_name="INCLS", header=3)
df_incl.columns = [str(c).strip().upper() for c in df_incl.columns]
df_incl["RUT_KEY"] = df_incl["IDENTIFICADOR"].apply(limpiar_rut)
df_incl = df_incl.drop_duplicates("RUT_KEY")

dict_fte_incl = dict(zip(df_incl["RUT_KEY"], df_incl["FTE INCLS"].apply(to_float_safe)))
dict_is_incl = dict(zip(df_incl["RUT_KEY"], df_incl["INCLS"]))

# ==========================================
# 4. CARGA DE FILTROS DE PERSONAL
# ==========================================
print("Procesando Filtros de Personal...")
df_act = pd.read_excel("Activos_Inactivos.xlsx")
col_rut_act = "Chile RUN - Rol Único Nacional National ID Information"
df_act["RUT_KEY"] = df_act[col_rut_act].apply(limpiar_rut)
dict_ingreso = dict(zip(df_act["RUT_KEY"], df_act["Employment Details Hire Date"]))

df_bajas = pd.read_excel("Bajas.xlsx")
df_bajas["RUT_KEY"] = df_bajas["Chile RUN - Rol Único Nacional National ID Information"].apply(limpiar_rut)
df_bajas["FECHA_BAJA"] = pd.to_datetime(df_bajas["Employment Details Termination Date"], errors='coerce')
ruts_baja = set(df_bajas[df_bajas["FECHA_BAJA"] <= FECHA_CORTE]["RUT_KEY"])

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

# FILTRO: Solo eliminamos si es una BAJA
df_final = df_ga[~df_ga["RUT_KEY"].isin(ruts_baja)].copy()

df_final.insert(2, "Nombre completo", df_final["Nombre"] + " " + df_final["Apellidos"])
df_final.insert(7, "Ceco", df_final["Grupo"].apply(clean_text).map(dict_ceco))
df_final.insert(9, "Fecha ingreso", df_final["RUT_KEY"].map(dict_ingreso).dt.date)

# TRADUCCIÓN DOBLE DE FTE
# 1. Traduce nombre raro a Agrupador Base
df_final["Agrupador"] = df_final["Cargo"].apply(clean_text).map(map_cargo_to_agrupa2)
# Si no lo encuentra en el diccionario, usa el mismo nombre del cargo
df_final["Agrupador"] = df_final["Agrupador"].fillna(df_final["Cargo"].apply(clean_text))
# 2. Asigna el número teórico
df_final["FTE TEORICO"] = df_final["Agrupador"].map(map_agrupa2_to_fte).fillna(0)

# Lógica de FTE Real
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

resumen = resumen.rename(columns={"Ceco": "CECO", "FTE_R": "FTE R"})
cols_res = ["CECO", "HC", "FTE R", "FTE AUTORIZADO", "NECESIDAD", "EC%", "TIPO EC", "FECHA", "TIENDA", "JEFE DISTRITO", "ASESOR TIENDA", "ENCARGADO RECLUTAMIENTO"]
resumen = resumen[[c for c in cols_res if c in resumen.columns]]

# ==========================================
# 7. EXPORTACIÓN
# ==========================================
print("Guardando archivo final...")
archivo_salida = f"FTE_CONCILIADO_V5_{FECHA_CORTE.date()}.xlsx"
with pd.ExcelWriter(archivo_salida) as writer:
    df_final.to_excel(writer, sheet_name="Sheet1", index=False)
    resumen.to_excel(writer, sheet_name="Resumen", index=False)

ruta = os.path.abspath(archivo_salida)
print(f"\n✅ CONCILIACIÓN COMPLETADA. Archivo en:\n{ruta}")
print(f"HC Final Detectado: {len(df_final)}")
print(f"FTE Real Calculado: {df_final['FTE REAL'].sum():.2f} (Debería ser ~1447.91)")