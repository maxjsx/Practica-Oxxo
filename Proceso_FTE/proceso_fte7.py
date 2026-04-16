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

def to_float_safe(val):
    if pd.isna(val) or str(val).strip() == "": 
        return np.nan
    val_str = str(val).strip().replace(',', '.')
    try:
        return float(val_str)
    except:
        return np.nan

print("🚀 Iniciando conciliación de FTE y formateo gráfico...")

# ==========================================
# 2. CARGA DE MAESTRA DE TIENDAS (AUTORIZADO)
# ==========================================
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
# 3. CARGA DE AGRUPADOR E INCLUSIONES
# ==========================================
map_cargo_to_agrupa = {}
map_cargo_to_fte = {}

df_agr = pd.read_excel("Agrupador 10.xlsx", sheet_name="AGRUPADOR", header=4)
df_agr.columns = [str(c).strip().upper() for c in df_agr.columns]
map_cargo_to_agrupa.update(dict(zip(df_agr["CARGO"].apply(clean_text), df_agr["AGRUPA CARGO_2"].apply(clean_text))))
map_cargo_to_fte.update(dict(zip(df_agr["CARGO"].apply(clean_text), df_agr["FTE TEORICO"].apply(to_float_safe))))

try:
    df_map_raw = pd.read_excel("Agrupador 10.xlsx", sheet_name="Hoja1")
    df_map_raw.columns = [str(c).strip().upper() for c in df_map_raw.columns]
    map_cargo_to_agrupa.update(dict(zip(df_map_raw["CARGO"].apply(clean_text), df_map_raw.iloc[:, 2].apply(clean_text))))
    map_cargo_to_fte.update(dict(zip(df_map_raw["CARGO"].apply(clean_text), df_map_raw.iloc[:, 4].apply(to_float_safe))))
except: pass

df_incl = pd.read_excel("Agrupador 10.xlsx", sheet_name="INCLS", header=3)
df_incl.columns = [str(c).strip().upper() for c in df_incl.columns]
df_incl["RUT_KEY"] = df_incl["IDENTIFICADOR"].apply(limpiar_rut)
df_incl = df_incl.drop_duplicates("RUT_KEY")

dict_fte_incl = dict(zip(df_incl["RUT_KEY"], df_incl["FTE INCLS"].apply(to_float_safe)))

# ==========================================
# 4. CARGA DE FILTROS DE PERSONAL
# ==========================================
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
df_ga = pd.read_excel("GestiondeAsistencia.xlsx", header=1)
df_ga["RUT_KEY"] = df_ga["Identificador"].apply(limpiar_rut)
df_final = df_ga[~df_ga["RUT_KEY"].isin(ruts_baja)].copy()

df_final.insert(2, "Nombre completo", df_final["Nombre"] + " " + df_final["Apellidos"])
df_final.insert(7, "Ceco", df_final["Grupo"].apply(clean_text).map(dict_ceco))
df_final.insert(9, "Fecha ingreso", df_final["RUT_KEY"].map(dict_ingreso).dt.date)

df_final["Agrupador"] = df_final["Cargo"].apply(clean_text).map(map_cargo_to_agrupa)
df_final["Agrupador"] = df_final["Agrupador"].fillna(df_final["Cargo"].apply(clean_text))
df_final["FTE TEORICO"] = df_final["Cargo"].apply(clean_text).map(map_cargo_to_fte).fillna(0)

df_final["fte incls"] = df_final["RUT_KEY"].map(dict_fte_incl)
df_final["fte incls"] = df_final["fte incls"].replace(0, np.nan)
df_final["FTE REAL"] = df_final["fte incls"].fillna(df_final["FTE TEORICO"])
df_final.loc[df_final["RUT_KEY"].isin(ruts_licencia_larga), "FTE REAL"] = 0

# ==========================================
# 6. RESUMEN EJECUTIVO (MATEMÁTICAS)
# ==========================================
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
# 7. EXPORTACIÓN CON DISEÑO PROFESIONAL (XLSXWRITER)
# ==========================================
print("Pintando Excel y ajustando formato...")
archivo_salida = f"FTE_REPORTE_GERENCIAL_{FECHA_CORTE.date()}.xlsx"

# Abrimos Excel usando xlsxwriter para inyectar diseño
writer = pd.ExcelWriter(archivo_salida, engine='xlsxwriter')
df_final.to_excel(writer, sheet_name="Detalle Asistencia", index=False)
resumen.to_excel(writer, sheet_name="Resumen Tiendas", index=False)

workbook = writer.book
ws_resumen = writer.sheets['Resumen Tiendas']

# --- DEFINICIÓN DE ESTILOS ---
formato_header = workbook.add_format({
    'bold': True, 'font_color': 'white', 'bg_color': '#1F4E78', # Azul Corporativo
    'border': 1, 'align': 'center', 'valign': 'vcenter'
})
formato_porcentaje = workbook.add_format({'num_format': '0.00%', 'align': 'center'})
formato_decimal = workbook.add_format({'num_format': '#,##0.00', 'align': 'center'})
formato_centrado = workbook.add_format({'align': 'center'})

# Colores Condicionales
fmt_critica = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) # Rojo
fmt_incompleta = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'}) # Amarillo
fmt_completa = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'}) # Verde
fmt_sobre = workbook.add_format({'bg_color': '#DDEBF7', 'font_color': '#203764'}) # Celeste

# --- APLICAR ESTILOS A LA HOJA RESUMEN ---
# Formatear encabezados
for col_num, value in enumerate(resumen.columns.values):
    ws_resumen.write(0, col_num, value, formato_header)

# Ajustar anchos y formatos de columnas específicos
ws_resumen.set_column('A:A', 10, formato_centrado) # CECO
ws_resumen.set_column('B:B', 8, formato_centrado)  # HC
ws_resumen.set_column('C:E', 15, formato_decimal)  # FTE R, AUTORIZADO, NECESIDAD
ws_resumen.set_column('F:F', 12, formato_porcentaje) # EC%
ws_resumen.set_column('G:G', 16, formato_centrado) # TIPO EC
ws_resumen.set_column('H:H', 12, formato_centrado) # FECHA
ws_resumen.set_column('I:I', 28)                   # TIENDA
ws_resumen.set_column('J:L', 25)                   # NOMBRES DE JEFES

# Agregar Autofiltrado
ws_resumen.autofilter(0, 0, len(resumen), len(resumen.columns) - 1)

# Aplicar Formato Condicional a la columna "TIPO EC" (Columna G, index 6)
ws_resumen.conditional_format(1, 6, len(resumen), 6,
    {'type': 'cell', 'criteria': '==', 'value': '"Crítica"', 'format': fmt_critica})
ws_resumen.conditional_format(1, 6, len(resumen), 6,
    {'type': 'cell', 'criteria': '==', 'value': '"Incompleta"', 'format': fmt_incompleta})
ws_resumen.conditional_format(1, 6, len(resumen), 6,
    {'type': 'cell', 'criteria': '==', 'value': '"Completa"', 'format': fmt_completa})
ws_resumen.conditional_format(1, 6, len(resumen), 6,
    {'type': 'cell', 'criteria': '==', 'value': '"Sobredotación"', 'format': fmt_sobre})

writer.close()

ruta = os.path.abspath(archivo_salida)
print(f"\n✅ REPORTE GERENCIAL CREADO.\nEl archivo está bellamente formateado y guardado en:\n{ruta}")