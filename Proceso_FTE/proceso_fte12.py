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
        return 0.0
    val_str = str(val).strip().replace(',', '.')
    try:
        return float(val_str)
    except:
        return 0.0

print("🚀 Iniciando FTE V11.0 (Fallback Inteligente SAP -> Talana)...")

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
# 3. CARGA DE AGRUPADOR E INCLUSIONES
# ==========================================
print("Cargando Agrupador y Diccionarios...")
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

# Guardamos un set con todos los cargos válidos conocidos para saber si SAP "tira error"
cargos_validos = set(map_cargo_to_agrupa.keys()).union(set(map_cargo_to_fte.keys()))

# ==========================================
# 4. CARGA DE FILTROS Y CARGOS DE SAP
# ==========================================
print("Procesando histórico de personas y Cargos SAP...")
df_act = pd.read_excel("Activos_Inactivos.xlsx")
col_rut_act = "Chile RUN - Rol Único Nacional National ID Information"
df_act["RUT_KEY"] = df_act[col_rut_act].apply(limpiar_rut)

dict_ingreso = dict(zip(df_act["RUT_KEY"], df_act["Employment Details Hire Date"]))
dict_cargo_sap = dict(zip(df_act["RUT_KEY"], df_act["Position Title"]))

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
# 5. CONSTRUCCIÓN DE LA SÁBANA Y FALLBACK
# ==========================================
print("Aplicando lógica de Fallback de Cargos...")
df_ga = pd.read_excel("GestiondeAsistencia.xlsx", header=1)
df_ga["RUT_KEY"] = df_ga["Identificador"].apply(limpiar_rut)
df_final = df_ga[~df_ga["RUT_KEY"].isin(ruts_baja)].copy()

df_final.insert(2, "Nombre completo", df_final["Nombre"] + " " + df_final["Apellidos"])
df_final.insert(7, "Ceco", df_final["Grupo"].apply(clean_text).map(dict_ceco))
df_final.insert(9, "Fecha ingreso", df_final["RUT_KEY"].map(dict_ingreso).dt.date)

df_final["CARGO_SAP"] = df_final["RUT_KEY"].map(dict_cargo_sap)

# --- LÓGICA INTELIGENTE DE FALLBACK ---
def definir_cargo_inteligente(row):
    tal = str(row["Cargo"]).strip().upper()
    sap = str(row["CARGO_SAP"]).strip().upper()

    if pd.isna(row["CARGO_SAP"]) or sap == "" or sap == "NAN":
        return tal

    # 1. Conservar los decimales: Si Talana dice "PT" y SAP no, priorizamos Talana.
    tal_is_pt = "PT" in tal or "PART TIME" in tal or "HRS" in tal
    sap_is_pt = "PT" in sap or "PART TIME" in sap or "HRS" in sap
    if tal_is_pt and not sap_is_pt:
        return tal

    # 2. Si SAP está en nuestro diccionario, lo usamos como fuente primaria.
    if sap in cargos_validos:
        return sap

    # 3. Fallback final: Si SAP "tira error" (no está en el diccionario), volvemos a Talana.
    return tal

df_final["CARGO OFICIAL"] = df_final.apply(definir_cargo_inteligente, axis=1)
# --------------------------------------

# Mapeo usando CARGO OFICIAL
df_final["Agrupador"] = df_final["CARGO OFICIAL"].apply(clean_text).map(map_cargo_to_agrupa)
df_final["Agrupador"] = df_final["Agrupador"].fillna(df_final["CARGO OFICIAL"].apply(clean_text))
df_final["FTE TEORICO"] = df_final["CARGO OFICIAL"].apply(clean_text).map(map_cargo_to_fte).fillna(0)

df_final["fte incls"] = df_final["RUT_KEY"].map(dict_fte_incl)
df_final["fte incls"] = df_final["fte incls"].replace(0, np.nan)
df_final["FTE REAL"] = df_final["fte incls"].fillna(df_final["FTE TEORICO"])
df_final.loc[df_final["RUT_KEY"].isin(ruts_licencia_larga), "FTE REAL"] = 0

# ==========================================
# 6. RESUMEN EJECUTIVO Y ANÁLISIS DE BRECHAS
# ==========================================
print("Generando Tablas Dinámicas y Brechas...")
resumen = df_final.groupby("Ceco").agg(HC=("Identificador", "nunique"), FTE_R=("FTE REAL", "sum")).reset_index()
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

total_hc = resumen["HC"].sum()
total_fte = resumen["FTE R"].sum()
total_tiendas = len(resumen)

resumen_general = pd.DataFrame({"MÉTRICA": ["Total HC (Filtrado)", "Total FTE Real", "Total Tiendas"], "VALOR": [total_hc, total_fte, total_tiendas]})
resumen_estado = resumen.groupby("TIPO EC").agg(CANTIDAD=("CECO", "count")).reset_index()
resumen_estado["% DEL TOTAL"] = resumen_estado["CANTIDAD"] / total_tiendas
resumen_estado = resumen_estado.sort_values(by="CANTIDAD", ascending=False)

# Brechas
map_cols_brecha = {
    "LIDER": "LIDER", "JEFE DE SALA": "JEFE DE SALA", "CAJERO": "CAJERO",
    "CAJERO PT 30": "CAJERO PT 30", "PT25": "CAJERO PT 25", "CAJERP PT 20": "CAJERO PT 20", "PT15": "CAJERO PT 15"
}

real_cargos = df_final.groupby(['Ceco', 'Agrupador'])['Identificador'].nunique().unstack(fill_value=0)
brechas_list = []
df_aut_filtrado = df_aut[df_aut["CECO"].isin(resumen["CECO"])].copy()

for index, row in df_aut_filtrado.iterrows():
    ceco = row["CECO"]
    tienda = row["NOMBRE MAESTRA"]
    fila_brecha = {"CECO": ceco, "TIENDA": tienda}
    for col_meta, nombre_agrupador in map_cols_brecha.items():
        meta = to_float_safe(row.get(col_meta, 0))
        real = real_cargos.loc[ceco, nombre_agrupador] if (ceco in real_cargos.index and nombre_agrupador in real_cargos.columns) else 0
        fila_brecha[f"Meta {nombre_agrupador}"] = meta
        fila_brecha[f"Real {nombre_agrupador}"] = real
        fila_brecha[f"VACANTES {nombre_agrupador}"] = meta - real
    brechas_list.append(fila_brecha)

df_brechas = pd.DataFrame(brechas_list)

# ==========================================
# 7. EXPORTACIÓN CON DISEÑO PROFESIONAL
# ==========================================
print("Pintando Excel y generando Tablas Oficiales...")
archivo_salida = f"FTE_DASHBOARD_GERENCIAL_{FECHA_CORTE.date()}.xlsx"

try:
    writer = pd.ExcelWriter(archivo_salida, engine='xlsxwriter')
    
    df_final.drop(columns=["RUT_KEY"]).to_excel(writer, sheet_name="Detalle Asistencia", index=False)

    workbook = writer.book
    
    ws_resumen = workbook.add_worksheet('Resumen Tiendas')
    writer.sheets['Resumen Tiendas'] = ws_resumen

    resumen.to_excel(writer, sheet_name="Resumen Tiendas", index=False, header=False, startrow=1, startcol=0)
    ws_resumen.add_table(0, 0, len(resumen), len(resumen.columns) - 1, {'columns': [{'header': c} for c in resumen.columns], 'name': 'TablaPrincipal', 'style': 'Table Style Medium 2'})

    resumen_general.to_excel(writer, sheet_name="Resumen Tiendas", index=False, header=False, startrow=2, startcol=14)
    ws_resumen.add_table(1, 14, 1 + len(resumen_general), 15, {'columns': [{'header': c} for c in resumen_general.columns], 'name': 'TablaGeneral', 'style': 'Table Style Dark 11'})

    resumen_estado.to_excel(writer, sheet_name="Resumen Tiendas", index=False, header=False, startrow=7, startcol=14)
    ws_resumen.add_table(6, 14, 6 + len(resumen_estado), 16, {'columns': [{'header': c} for c in resumen_estado.columns], 'name': 'TablaEstados', 'style': 'Table Style Medium 3'})

    ws_brechas = workbook.add_worksheet('Brechas por Cargo')
    writer.sheets['Brechas por Cargo'] = ws_brechas
    
    df_brechas.to_excel(writer, sheet_name="Brechas por Cargo", index=False, header=False, startrow=1, startcol=0)
    ws_brechas.add_table(0, 0, len(df_brechas), len(df_brechas.columns) - 1, {
        'columns': [{'header': c} for c in df_brechas.columns], 
        'name': 'TablaBrechas', 
        'style': 'Table Style Medium 2'
    })

    formato_porcentaje = workbook.add_format({'num_format': '0.00%', 'align': 'center'})
    formato_decimal = workbook.add_format({'num_format': '#,##0.00', 'align': 'center'})
    formato_entero = workbook.add_format({'num_format': '#,##0', 'align': 'center'})
    formato_centrado = workbook.add_format({'align': 'center'})

    fmt_critica = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    fmt_incompleta = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})
    fmt_completa = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    fmt_sobre = workbook.add_format({'bg_color': '#DDEBF7', 'font_color': '#203764'})

    ws_resumen.set_column('A:A', 10, formato_centrado) 
    ws_resumen.set_column('B:B', 8, formato_centrado)  
    ws_resumen.set_column('C:E', 15, formato_decimal)  
    ws_resumen.set_column('F:F', 12, formato_porcentaje) 
    ws_resumen.set_column('G:G', 16, formato_centrado) 
    ws_resumen.set_column('H:H', 12, formato_centrado) 
    ws_resumen.set_column('I:I', 28)                   
    ws_resumen.set_column('J:L', 25)                   
    ws_resumen.set_column('O:O', 25) 
    ws_resumen.set_column('P:P', 14, formato_entero) 
    ws_resumen.set_column('Q:Q', 14, formato_porcentaje) 
    ws_resumen.write_number(3, 15, total_fte, formato_decimal)

    ws_resumen.conditional_format(1, 6, len(resumen), 6, {'type': 'cell', 'criteria': '==', 'value': '"Crítica"', 'format': fmt_critica})
    ws_resumen.conditional_format(1, 6, len(resumen), 6, {'type': 'cell', 'criteria': '==', 'value': '"Incompleta"', 'format': fmt_incompleta})
    ws_resumen.conditional_format(1, 6, len(resumen), 6, {'type': 'cell', 'criteria': '==', 'value': '"Completa"', 'format': fmt_completa})
    ws_resumen.conditional_format(1, 6, len(resumen), 6, {'type': 'cell', 'criteria': '==', 'value': '"Sobredotación"', 'format': fmt_sobre})

    ws_resumen.conditional_format(7, 14, 7+len(resumen_estado), 14, {'type': 'cell', 'criteria': '==', 'value': '"Crítica"', 'format': fmt_critica})
    ws_resumen.conditional_format(7, 14, 7+len(resumen_estado), 14, {'type': 'cell', 'criteria': '==', 'value': '"Incompleta"', 'format': fmt_incompleta})
    ws_resumen.conditional_format(7, 14, 7+len(resumen_estado), 14, {'type': 'cell', 'criteria': '==', 'value': '"Completa"', 'format': fmt_completa})
    ws_resumen.conditional_format(7, 14, 7+len(resumen_estado), 14, {'type': 'cell', 'criteria': '==', 'value': '"Sobredotación"', 'format': fmt_sobre})

    ws_brechas.set_column('A:A', 10, formato_centrado)
    ws_brechas.set_column('B:B', 30)
    ws_brechas.set_column('C:W', 18, formato_entero) 
    
    for col_idx in range(4, len(df_brechas.columns), 3):
        ws_brechas.conditional_format(1, col_idx, len(df_brechas), col_idx, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_critica})
        ws_brechas.conditional_format(1, col_idx, len(df_brechas), col_idx, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt_sobre})
        ws_brechas.conditional_format(1, col_idx, len(df_brechas), col_idx, {'type': 'cell', 'criteria': '==', 'value': 0, 'format': fmt_completa})

    writer.close()
    
    ruta = os.path.abspath(archivo_salida)
    print(f"\n✅ REPORTE EXACTO CREADO CON ÉXITO.\nEl archivo está en:\n{ruta}")

except PermissionError:
    print(f"\n❌ ERROR: El archivo '{archivo_salida}' está abierto en Excel. Ciérralo y vuelve a intentar.")