import pandas as pd
import numpy as np
import re
import os

# ==========================================
# 1. CONFIGURACIÓN Y FUNCIONES (Clonación de Power Query)
# ==========================================
FECHA_CORTE = pd.to_datetime("2026-04-30") 

# El PQ limpia el RUT de Talana reemplazando el punto por nada
def clean_rut_talana(x):
    if pd.isna(x): return ""
    return str(x).replace('.', '').strip().upper()

# El PQ limpia el RUT de SAP y Asistencia rellenando con ceros
def clean_rut_sap(x):
    if pd.isna(x): return ""
    return str(x).zfill(12).strip().upper()

def clean_text(x):
    if pd.isna(x): return ""
    return str(x).strip().upper() # El PQ original usa Text.Upper y Text.Trim

def to_float_safe(val):
    if pd.isna(val) or str(val).strip() == "": 
        return np.nan
    val_str = str(val).strip().replace(',', '.')
    try:
        return float(val_str)
    except:
        return np.nan

print("🚀 Iniciando FTE V19.0 (Modo: Clon de Power Query ABRIL 2026)...")

# ==========================================
# 2. CARGA DE MAESTRA DE TIENDAS (AUTORIZADO)
# ==========================================
print("Cargando matriz de tiendas...")
df_aut = pd.read_excel("00 - FTE AUTORIZADO.xlsx", sheet_name="ABRIL_26", header=6)
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
map_cargo_to_fte.update(dict(zip(df_agr["CARGO"].apply(clean_text), df_agr["FTE TEORICO2"].apply(to_float_safe))))

try:
    df_map_raw = pd.read_excel("Agrupador 10.xlsx", sheet_name="Hoja1")
    df_map_raw.columns = [str(c).strip().upper() for c in df_map_raw.columns]
    map_cargo_to_agrupa.update(dict(zip(df_map_raw["CARGO"].apply(clean_text), df_map_raw.iloc[:, 2].apply(clean_text))))
    map_cargo_to_fte.update(dict(zip(df_map_raw["CARGO"].apply(clean_text), df_map_raw.iloc[:, 4].apply(to_float_safe))))
except: pass

df_incl = pd.read_excel("Agrupador 10.xlsx", sheet_name="INCLS", header=3)
df_incl.columns = [str(c).strip().upper() for c in df_incl.columns]
df_incl["RUT_KEY"] = df_incl["IDENTIFICADOR"].apply(lambda x: clean_rut_talana(str(x)))
df_incl = df_incl.drop_duplicates("RUT_KEY")

# Respetamos el número que viene en INCLS (incluso si es 0 literal)
dict_fte_incl = dict(zip(df_incl["RUT_KEY"], df_incl["FTE INCLS"].apply(to_float_safe)))

# ==========================================
# 4. CARGA DE FILTROS Y PERMISOS (Como lo hace el PQ)
# ==========================================
print("Procesando histórico, Bajas y Permisos...")
df_act = pd.read_excel("Copia_de_Copia_de_Activos_inactivos_OXXO_Chile_2__Copy_2026_04_07_12_35_17.xlsx")
col_rut_act = "Chile RUN - Rol Único Nacional National ID Information"
df_act["RUT_SAP"] = df_act[col_rut_act].apply(clean_rut_sap)
dict_cargo_sap = dict(zip(df_act["RUT_SAP"], df_act["Position Title"]))

df_talana_emp = pd.read_excel("Lista Empleados de oxxo (3).xlsx")
df_talana_emp["RUT_TALANA"] = df_talana_emp["RUT"].apply(clean_rut_talana)
dict_cargo_talana_emp = dict(zip(df_talana_emp["RUT_TALANA"], df_talana_emp["Cargo"]))

df_bajas = pd.read_excel("Copia_de_Copia_de_Bajas_Tienda_OXXO_Chile_Copy_2026_04_07_12_37_34.xlsx")
df_bajas["RUT_SAP"] = df_bajas["Chile RUN - Rol Único Nacional National ID Information"].apply(clean_rut_sap)
df_bajas["FECHA_BAJA"] = pd.to_datetime(df_bajas["Employment Details Termination Date"], errors='coerce')
ruts_baja = set(df_bajas[df_bajas["FECHA_BAJA"] <= FECHA_CORTE]["RUT_SAP"])

df_perm = pd.read_excel("Permisosasignados202604070841_216b6d90-b885-4e1b-953b-0289a1acf2b7.xlsx")
df_perm["RUT_TALANA"] = df_perm["Rut"].apply(clean_rut_talana)
df_perm["F_INI"] = pd.to_datetime(df_perm["Fecha Inicio"], dayfirst=True)
df_perm["F_FIN"] = pd.to_datetime(df_perm["Fecha Fin"], dayfirst=True)
# Clonamos el cálculo exacto de días del PQ: Duration.Days([Fecha Fin] - [Fecha Inicio]) + 1
df_perm["DIAS"] = (df_perm["F_FIN"] - df_perm["F_INI"]).dt.days + 1

# Cualquier permiso mayor a 15 días (sin importar el tipo)
dict_dias_permiso = dict(zip(df_perm["RUT_TALANA"], df_perm["DIAS"]))

# ==========================================
# 5. CONSTRUCCIÓN DE LA SÁBANA (Lógica exacta PQ)
# ==========================================
print("Cruzando con lógica de Power Query...")
df_ga = pd.read_excel("GestióndeAsistencia202604070840_08ba6579-40d4-4bf5-bf3f-bf4f61b24d1c.xlsx", header=1)

# Creamos las dos llaves maestras que usa el Power Query
df_ga["RUT_SAP"] = df_ga["Identificador"].apply(clean_rut_sap)
df_ga["RUT_TALANA"] = df_ga["Identificador"].apply(clean_rut_talana)

# Filtramos bajas
df_final = df_ga[~df_ga["RUT_SAP"].isin(ruts_baja)].copy()

df_final.insert(2, "NOMBRE COMPLETO", df_final["Nombre"] + " " + df_final["Apellidos"])
df_final.insert(7, "CECO_1", df_final["Grupo"].apply(clean_text).map(dict_ceco))

# --- REGLA CARGO FINAL (Clon PQ) ---
# "each if [SAP ACTIVOS INACTIVOS.Position Title] = null then [TALANA LISTA EMPLEADOS.Cargo] else [SAP ACTIVOS INACTIVOS.Position Title]"
def definir_cargo_pq(row):
    sap_title = dict_cargo_sap.get(row["RUT_SAP"])
    tal_title = dict_cargo_talana_emp.get(row["RUT_TALANA"])
    if pd.isna(sap_title) or sap_title is None or str(sap_title).strip() == "":
        return tal_title
    return sap_title

df_final["Cargo final"] = df_final.apply(definir_cargo_pq, axis=1)
# --------------------------------------

# Texto a Mayúsculas y Mapeo al Agrupador
df_final["CARGO_CLEAN"] = df_final["Cargo final"].apply(clean_text)
df_final["AGRUPADOR"] = df_final["CARGO_CLEAN"].map(map_cargo_to_agrupa)
df_final["AGRUPADOR"] = df_final["AGRUPADOR"].fillna(df_final["CARGO_CLEAN"])
df_final["FTE TEORICO"] = df_final["CARGO_CLEAN"].map(map_cargo_to_fte).fillna(0.0)

# Mapeamos Días de Permiso e Inclusiones
df_final["Dias 2"] = df_final["RUT_TALANA"].map(dict_dias_permiso).fillna(0)
df_final["FTE INCLS"] = df_final["RUT_TALANA"].map(dict_fte_incl)

# --- REGLA FTE REAL (Clon PQ) ---
# "if [Dias 2] > 15 then 0 else if [AGRUPA INCLS.FTE INCLS] = 0 then 0 else [AGRUPA_CARGO.FTE TEORICO]"
def calcular_fte_real(row):
    if row["Dias 2"] > 15:
        return 0.0
    elif row["FTE INCLS"] == 0:
        return 0.0
    elif pd.notna(row["FTE INCLS"]):
        return float(row["FTE INCLS"]) # Si existe INCLS y no es 0, lo usamos (comportamiento implícito esperado de rescate)
    else:
        return row["FTE TEORICO"]

df_final["FTE REAL"] = df_final.apply(calcular_fte_real, axis=1)

# ==========================================
# 6. RESUMEN EJECUTIVO Y ANÁLISIS DE BRECHAS
# ==========================================
print("Generando Tablas Dinámicas y Brechas...")
resumen = df_final.groupby("CECO_1").agg(HC=("Identificador", "nunique"), FTE_R=("FTE REAL", "sum")).reset_index()
resumen["FTE AUTORIZADO"] = resumen["CECO_1"].map(dict_fte_aut).fillna(0)
resumen["NECESIDAD"] = resumen["FTE_R"] - resumen["FTE AUTORIZADO"]
resumen["EC%"] = np.where(resumen["FTE AUTORIZADO"] == 0, 0, resumen["FTE_R"] / resumen["FTE AUTORIZADO"])

def clasificar_ec(ec):
    ec = round(ec, 4)
    if ec >= 1: return "Completa"
    elif ec >= 0.8: return "Incompleta"
    else: return "Crítica"

resumen["TIPO EC"] = resumen["EC%"].apply(clasificar_ec)
resumen["FECHA"] = FECHA_CORTE.date()
resumen["TIENDA"] = resumen["CECO_1"].map(dict_tienda_nom)
resumen["JEFE DISTRITO"] = resumen["CECO_1"].map(dict_jefe)
resumen["ASESOR TIENDA"] = resumen["CECO_1"].map(dict_asesor)
resumen["ENCARGADO RECLUTAMIENTO"] = resumen["CECO_1"].map(dict_reclutador)

resumen = resumen.rename(columns={"CECO_1": "CECO", "FTE_R": "FTE R"})
cols_res = ["CECO", "HC", "FTE R", "FTE AUTORIZADO", "NECESIDAD", "EC%", "TIPO EC", "FECHA", "TIENDA", "JEFE DISTRITO", "ASESOR TIENDA", "ENCARGADO RECLUTAMIENTO"]
resumen = resumen[[c for c in cols_res if c in resumen.columns]]

total_hc = resumen["HC"].sum()
total_fte = resumen["FTE R"].sum()
total_tiendas = len(resumen)

resumen_general = pd.DataFrame({"MÉTRICA": ["Total HC (Filtrado)", "Total FTE Real", "Total Tiendas"], "VALOR": [total_hc, total_fte, total_tiendas]})
resumen_estado = resumen.groupby("TIPO EC").agg(CANTIDAD=("CECO", "count")).reset_index()
resumen_estado["% DEL TOTAL"] = resumen_estado["CANTIDAD"] / total_tiendas
resumen_estado = resumen_estado.sort_values(by="CANTIDAD", ascending=False)

map_cols_brecha = {
    "LIDER": ("LIDER DE TIENDA", "LIDER"), 
    "JEFE DE SALA": ("JEFE DE SALA", "JEFE DE SALA"), 
    "CAJERO": ("CAJERO", "CAJERO"),
    "CAJERO PT 30": ("CAJERO PT 30", "CAJERO PT 30"), 
    "PT25": ("CAJERO PT 25", "CAJERO PT 25"), 
    "CAJERP PT 20": ("CAJERO PT 20", "CAJERO PT 20"), 
    "PT15": ("CAJERO PT 15", "CAJERO PT 15")
}

real_cargos = df_final.groupby(['CECO_1', 'AGRUPADOR'])['Identificador'].nunique().unstack(fill_value=0)
brechas_list = []
df_aut_filtrado = df_aut[df_aut["CECO"].isin(resumen["CECO"])].copy()

for index, row in df_aut_filtrado.iterrows():
    ceco = row["CECO"]
    tienda = row["NOMBRE MAESTRA"]
    fila_brecha = {"CECO": ceco, "TIENDA": tienda}
    for col_meta, (nombre_agrupador, display_name) in map_cols_brecha.items():
        meta = to_float_safe(row.get(col_meta, 0))
        real = real_cargos.loc[ceco, nombre_agrupador] if (ceco in real_cargos.index and nombre_agrupador in real_cargos.columns) else 0
        fila_brecha[f"Meta {display_name}"] = meta
        fila_brecha[f"Real {display_name}"] = real
        fila_brecha[f"VACANTES {display_name}"] = meta - real
    brechas_list.append(fila_brecha)

df_brechas = pd.DataFrame(brechas_list)

# ==========================================
# 7. EXPORTACIÓN CON DISEÑO PROFESIONAL
# ==========================================
print("Pintando Excel y generando Tablas Oficiales...")
archivo_salida = f"FTE_DASHBOARD_GERENCIAL_{FECHA_CORTE.date()}.xlsx"

try:
    writer = pd.ExcelWriter(archivo_salida, engine='xlsxwriter')
    
    df_export = df_final.drop(columns=["RUT_SAP", "RUT_TALANA", "CARGO_CLEAN"])
    df_export.to_excel(writer, sheet_name="Detalle Asistencia", index=False)

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
    ws_brechas.add_table(0, 0, len(df_brechas), len(df_brechas.columns) - 1, {'columns': [{'header': c} for c in df_brechas.columns], 'name': 'TablaBrechas', 'style': 'Table Style Medium 2'})

    formato_porcentaje = workbook.add_format({'num_format': '0.00%', 'align': 'center'})
    formato_decimal = workbook.add_format({'num_format': '#,##0.00', 'align': 'center'})
    formato_entero = workbook.add_format({'num_format': '#,##0', 'align': 'center'})
    formato_centrado = workbook.add_format({'align': 'center'})

    fmt_critica = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    fmt_incompleta = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})
    fmt_completa = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    fmt_sobre = workbook.add_format({'bg_color': '#DDEBF7', 'font_color': '#203764'})

    for ws in [ws_resumen, ws_brechas]:
        ws.set_column('A:A', 10, formato_centrado) 

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

    for criteria, fmt in [('"Crítica"', fmt_critica), ('"Incompleta"', fmt_incompleta), ('"Completa"', fmt_completa), ('"Sobredotación"', fmt_sobre)]:
        ws_resumen.conditional_format(1, 6, len(resumen), 6, {'type': 'cell', 'criteria': '==', 'value': criteria, 'format': fmt})
        ws_resumen.conditional_format(7, 14, 7+len(resumen_estado), 14, {'type': 'cell', 'criteria': '==', 'value': criteria, 'format': fmt})

    ws_brechas.set_column('B:B', 30)
    ws_brechas.set_column('C:W', 18, formato_entero) 
    
    for col_idx in range(4, len(df_brechas.columns), 3):
        ws_brechas.conditional_format(1, col_idx, len(df_brechas), col_idx, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_critica})
        ws_brechas.conditional_format(1, col_idx, len(df_brechas), col_idx, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt_sobre})
        ws_brechas.conditional_format(1, col_idx, len(df_brechas), col_idx, {'type': 'cell', 'criteria': '==', 'value': 0, 'format': fmt_completa})

    writer.close()
    
    ruta = os.path.abspath(archivo_salida)
    print(f"\n✅ REPORTE EXACTO (CLON PQ) CREADO CON ÉXITO.\nEl archivo está en:\n{ruta}")

except PermissionError:
    print(f"\n❌ ERROR: El archivo '{archivo_salida}' está abierto en Excel. Ciérralo y vuelve a intentar.")