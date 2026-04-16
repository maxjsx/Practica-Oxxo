import pandas as pd
import numpy as np
import re
import os

# ==========================================
# 1. CONFIGURACIÓN Y FUNCIONES
# ==========================================
FECHA_CORTE = pd.to_datetime("2026-03-11") # Ajustado a la fecha del reporte

def limpiar_rut(x):
    if pd.isna(x): return ""
    return re.sub(r'[^0-9Kk]', '', str(x).upper())

def clean_text(x):
    if pd.isna(x): return ""
    return str(x).strip().upper()

print("Iniciando generación de Reporte Maestro de FTE...")

# ==========================================
# 2. CARGA DE AUTORIZADO Y DICCIONARIOS
# ==========================================
print("Cargando matriz de tiendas (Autorizado)...")
df_aut = pd.read_excel("00 - FTE AUTORIZADO.xlsx", sheet_name="MARZO_26 ", header=6)
df_aut.columns = [str(c).strip().upper() for c in df_aut.columns]
df_aut["NOMBRE_KEY"] = df_aut["NOMBRE MAESTRA"].apply(clean_text)

dict_ceco = dict(zip(df_aut["NOMBRE_KEY"], df_aut["CECO"]))
dict_fte_aut = dict(zip(df_aut["CECO"], df_aut["FTE AUT"]))
dict_jefe = dict(zip(df_aut["CECO"], df_aut["JEFE OPERACIONES"]))
dict_asesor = dict(zip(df_aut["CECO"], df_aut["ASESOR TIENDA"]))
dict_estado_tienda = dict(zip(df_aut["CECO"], df_aut["ESTADO RH"]))
dict_reclutador = dict(zip(df_aut["CECO"], df_aut["RECLUTADOR"]))
dict_tienda_nom = dict(zip(df_aut["CECO"], df_aut["NOMBRE MAESTRA"]))

# ==========================================
# 3. CARGA DE AGRUPADOR E INCLUSIONES
# ==========================================
print("Cargando Agrupador e Inclusiones...")
df_agr = pd.read_excel("Agrupador 10.xlsx", sheet_name="AGRUPADOR", header=4)
map_agrupa = dict(zip(df_agr["CARGO"].apply(clean_text), df_agr["AGRUPA CARGO_2"]))
map_fte_teo = dict(zip(df_agr["AGRUPA CARGO_2"].apply(clean_text), pd.to_numeric(df_agr["FTE TEORICO"], errors='coerce')))

df_incl = pd.read_excel("Agrupador 10.xlsx", sheet_name="INCLS", header=3)
df_incl["RUT_KEY"] = df_incl["Identificador"].apply(limpiar_rut)
df_incl = df_incl.drop_duplicates("RUT_KEY")

dict_fte_incl = dict(zip(df_incl["RUT_KEY"], df_incl["FTE INCLS"]))
dict_is_incl = dict(zip(df_incl["RUT_KEY"], df_incl["INCLS"]))
dict_art_22 = dict(zip(df_incl["RUT_KEY"], df_incl["ART 22"]))

# ==========================================
# 4. CARGA DE ACTIVOS, BAJAS Y PERMISOS
# ==========================================
print("Procesando histórico de personas (Activos, Bajas, Licencias)...")
df_act = pd.read_excel("Activos_Inactivos.xlsx")
col_rut_act = "Chile RUN - Rol Único Nacional National ID Information"
df_act["RUT_KEY"] = df_act[col_rut_act].apply(limpiar_rut)

ruts_activos = set(df_act[df_act["Employee Status"] == "Active"]["RUT_KEY"])
dict_ingreso = dict(zip(df_act["RUT_KEY"], df_act["Employment Details Hire Date"]))
dict_sindicato = dict(zip(df_act["RUT_KEY"], df_act.get("Sindicato", "0"))) 

df_bajas = pd.read_excel("Bajas.xlsx")
col_rut_baj = [c for c in df_bajas.columns if "RUN" in str(c)][0]
df_bajas["RUT_KEY"] = df_bajas[col_rut_baj].apply(limpiar_rut)
df_bajas["FECHA_BAJA"] = pd.to_datetime(df_bajas["Employment Details Termination Date"], errors='coerce')
ruts_baja = set(df_bajas[df_bajas["FECHA_BAJA"] <= FECHA_CORTE]["RUT_KEY"])
dict_egreso = dict(zip(df_bajas["RUT_KEY"], df_bajas["FECHA_BAJA"]))

df_perm = pd.read_excel("PermisosAsignados.xlsx")
df_perm["RUT_KEY"] = df_perm["Rut"].apply(limpiar_rut)
df_perm["F_INI"] = pd.to_datetime(df_perm["Fecha Inicio"], dayfirst=True)
df_perm["F_FIN"] = pd.to_datetime(df_perm["Fecha Fin"], dayfirst=True)
df_perm["DURACION"] = (df_perm["F_FIN"] - df_perm["F_INI"]).dt.days + 1

df_perm_act = df_perm[(df_perm["F_INI"] <= FECHA_CORTE) & (df_perm["F_FIN"] >= FECHA_CORTE)].drop_duplicates("RUT_KEY")
dict_perm_tipo = dict(zip(df_perm_act["RUT_KEY"], df_perm_act["Tipo Permiso"]))
dict_perm_ini = dict(zip(df_perm_act["RUT_KEY"], df_perm_act["F_INI"].dt.date))
dict_perm_fin = dict(zip(df_perm_act["RUT_KEY"], df_perm_act["F_FIN"].dt.date))
dict_perm_dias = dict(zip(df_perm_act["RUT_KEY"], df_perm_act["DURACION"]))

ruts_licencia_larga = set(df_perm_act[
    (df_perm_act["Tipo Permiso"].str.contains("Licencia", case=False, na=False)) & 
    (df_perm_act["DURACION"] > 15)
]["RUT_KEY"])

# ==========================================
# 5. BASE ASISTENCIA (CONSTRUCCIÓN HOJA 1)
# ==========================================
print("Construyendo Sabana de Datos (Sheet1)...")
df_ga = pd.read_excel("GestiondeAsistencia.xlsx", header=1)
df_ga["RUT_KEY"] = df_ga["Identificador"].apply(limpiar_rut)

# Filtro
df_final = df_ga[df_ga["RUT_KEY"].isin(ruts_activos)].copy()
df_final = df_final[~df_final["RUT_KEY"].isin(ruts_baja)]

# Inyectar columnas
df_final.insert(2, "Nombre completo", df_final["Nombre"] + " " + df_final["Apellidos"])

def format_rut_sap(r):
    return f"{r[:-1]}-{r[-1]}" if len(r)>1 else r

df_final.insert(4, "RUT SAP", df_final["RUT_KEY"].apply(format_rut_sap))
df_final.insert(5, "RUT TALANA", df_final["RUT_KEY"])

# Mapear CECO 
df_final.insert(7, "Ceco", df_final["Grupo"].apply(clean_text).map(dict_ceco))
df_final.insert(9, "Fecha ingreso", df_final["RUT_KEY"].map(dict_ingreso).dt.date)
df_final.insert(10, "Fecha egreso", df_final["RUT_KEY"].map(dict_egreso).dt.date)

# Mapeos de cálculo
df_final["Agrupador"] = df_final["Cargo"].apply(clean_text).map(map_agrupa)
df_final["FTE TEORICO"] = df_final["Agrupador"].fillna("").astype(str).apply(clean_text).map(map_fte_teo).fillna(0)

df_final["Tipo Permiso Ext"] = df_final["RUT_KEY"].map(dict_perm_tipo)
df_final["inicio"] = df_final["RUT_KEY"].map(dict_perm_ini)
df_final["fin"] = df_final["RUT_KEY"].map(dict_perm_fin)
df_final["dias"] = df_final["RUT_KEY"].map(dict_perm_dias)
df_final["fte incls"] = df_final["RUT_KEY"].map(dict_fte_incl)

# Calculo Maestro
df_final["FTE REAL"] = df_final["fte incls"].fillna(df_final["FTE TEORICO"])
df_final.loc[df_final["RUT_KEY"].isin(ruts_licencia_larga), "FTE REAL"] = 0

df_final["INCLS"] = df_final["RUT_KEY"].map(dict_is_incl)
df_final["ART 22"] = df_final["RUT_KEY"].map(dict_art_22)
df_final["RRLL"] = df_final["RUT_KEY"].map(dict_sindicato).fillna(0)
df_final["RRLL2"] = 0

# Limpiar
df_final = df_final.drop(columns=["RUT_KEY"])

# ==========================================
# 6. CONSTRUCCIÓN RESUMEN (HOJA 2)
# ==========================================
print("Generando Resumen Ejecutivo (Hoja2)...")
resumen = df_final.groupby("Ceco").agg(
    HC=("Identificador", "nunique"),
    FTE_R=("FTE REAL", "sum")
).reset_index()

# Obligamos a numérico para evitar el error de resta
resumen["FTE_R"] = pd.to_numeric(resumen["FTE_R"], errors='coerce').fillna(0)
resumen["FTE AUTORIZADO"] = pd.to_numeric(resumen["Ceco"].map(dict_fte_aut), errors='coerce').fillna(0)

resumen["NECESIDAD"] = resumen["FTE_R"] - resumen["FTE AUTORIZADO"]
resumen["EC%"] = np.where(resumen["FTE AUTORIZADO"] == 0, 0, resumen["FTE_R"] / resumen["FTE AUTORIZADO"])

def clasificar_ec(ec):
    if pd.isna(ec): return ""
    ec = round(ec, 4)
    if ec == 1: return "Completa"
    elif ec > 1: return "Sobredotación"
    elif ec >= 0.8: return "Incompleta"
    else: return "Crítica"

resumen["TIPO EC"] = resumen["EC%"].apply(clasificar_ec)
resumen["FECHA"] = FECHA_CORTE.date()
resumen["TIENDA"] = resumen["Ceco"].map(dict_tienda_nom)
resumen["JEFE DISTRITO"] = resumen["Ceco"].map(dict_jefe)
resumen["ASESOR TIENDA"] = resumen["Ceco"].map(dict_asesor)
resumen["ESTADO TIENDA"] = resumen["Ceco"].map(dict_estado_tienda)
resumen["ENCARGADO RECLUTAMIENTO"] = resumen["Ceco"].map(dict_reclutador)

resumen = resumen.rename(columns={"Ceco": "CECO", "FTE_R": "FTE R"})

# Ordenar columnas
cols_res = ["CECO", "HC", "FTE R", "FTE AUTORIZADO", "NECESIDAD", "EC%", "TIPO EC", "FECHA", "TIENDA", "JEFE DISTRITO", "ASESOR TIENDA", "ESTADO TIENDA", "ENCARGADO RECLUTAMIENTO"]
resumen = resumen[[c for c in cols_res if c in resumen.columns]]

# ==========================================
# 7. EXPORTACIÓN PROFESIONAL
# ==========================================
print("Guardando archivo...")
archivo_salida = f"FTE_AUTOMATIZADO_{FECHA_CORTE.date()}.xlsx"

with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
    df_final.to_excel(writer, sheet_name="Sheet1", index=False)
    resumen.to_excel(writer, sheet_name="Resumen", index=False)

# Mostrar ruta exacta donde se guardó
ruta_absoluta = os.path.abspath(archivo_salida)
print(f"\n✅ ¡Éxito total! Archivo generado y guardado en:\n{ruta_absoluta}")