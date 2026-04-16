import pandas as pd
import numpy as np
import re
import os

# ==========================================
# 1. CONFIGURACIÓN Y FUNCIONES
# ==========================================
FECHA_CORTE = pd.to_datetime("2026-04-30")

def clean_rut_talana(x):
    if pd.isna(x): return ""
    return str(x).replace('.', '').strip().upper()

def clean_rut_sap(x):
    if pd.isna(x): return ""
    return str(x).zfill(12).strip().upper()

def clean_text(x):
    if pd.isna(x): return ""
    return re.sub(r'\s+', ' ', str(x).strip().upper())

def to_float_safe(val):
    if pd.isna(val) or str(val).strip() == "": 
        return np.nan
    val_str = str(val).strip().replace(',', '.')
    try:
        return float(val_str)
    except:
        return np.nan

print("🚀 Iniciando FTE V-FINAL (Asignación Directa Columna E - ABRIL 2026)...")

# ==========================================
# 2. CARGA DE MAESTRA DE TIENDAS
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
# 3. CARGA DE AGRUPADOR
# ==========================================
print("Cargando Agrupador...")
df_agr = pd.read_excel("Agrupador 10.xlsx", sheet_name="AGRUPADOR", header=4)

cargo_col = df_agr.iloc[:, 14].apply(clean_text)
agrupador_col = df_agr.iloc[:, 15].apply(clean_text)
fte_col = df_agr.iloc[:, 18].apply(to_float_safe)

dict_cargo_agrup = dict(zip(cargo_col, agrupador_col))

df_agr_temp = pd.DataFrame({"AGRUPADOR": agrupador_col, "FTE": fte_col})
df_agr_unique = df_agr_temp.drop_duplicates(subset=["AGRUPADOR"])
dict_agrup_fte = dict(zip(df_agr_unique["AGRUPADOR"], df_agr_unique["FTE"]))

# ==========================================
# 4. CARGA DE SAP, TALANA EMPLEADOS Y PERMISOS
# ==========================================
print("Procesando SAP, Bajas y Permisos...")
df_act = pd.read_excel("Copia_de_Copia_de_Activos_inactivos_OXXO_Chile_2__Copy_2026_04_07_12_35_17.xlsx")
df_act["RUT_SAP"] = df_act["Chile RUN - Rol Único Nacional National ID Information"].apply(clean_rut_sap)
dict_sap_term = dict(zip(df_act["RUT_SAP"], pd.to_datetime(df_act["Employment Details Termination Date"], errors='coerce')))
dict_cargo_sap = dict(zip(df_act["RUT_SAP"], df_act["Position Title"]))

df_talana_emp = pd.read_excel("Lista Empleados de oxxo (3).xlsx")
df_talana_emp["RUT_TALANA"] = df_talana_emp["RUT"].apply(clean_rut_talana)
dict_tal_ingreso = dict(zip(df_talana_emp["RUT_TALANA"], pd.to_datetime(df_talana_emp["Fecha de Ingreso"], errors='coerce', dayfirst=True)))
dict_cargo_talana_emp = dict(zip(df_talana_emp["RUT_TALANA"], df_talana_emp["Cargo"]))

df_bajas = pd.read_excel("Copia_de_Copia_de_Bajas_Tienda_OXXO_Chile_Copy_2026_04_07_12_37_34.xlsx")
df_bajas["RUT_SAP"] = df_bajas["Chile RUN - Rol Único Nacional National ID Information"].apply(clean_rut_sap)
dict_sap_term.update(dict(zip(df_bajas["RUT_SAP"], pd.to_datetime(df_bajas["Employment Details Termination Date"], errors='coerce'))))

# PERMISOS: Se excluyen Vacaciones para que no sumen a los 15 días
df_perm = pd.read_excel("Permisosasignados202604070841_216b6d90-b885-4e1b-953b-0289a1acf2b7.xlsx")
df_perm["RUT_TALANA"] = df_perm["Rut"].apply(clean_rut_talana)
df_perm = df_perm[~df_perm["Tipo Permiso"].str.contains("Vacaciones", case=False, na=False)].copy()

df_perm["F_INI"] = pd.to_datetime(df_perm["Fecha Inicio"], dayfirst=True)
df_perm["F_FIN"] = pd.to_datetime(df_perm["Fecha Fin"], dayfirst=True)
df_perm["DIAS"] = (df_perm["F_FIN"] - df_perm["F_INI"]).dt.days + 1

df_perm_filtrado = df_perm.sort_values(by="DIAS", ascending=False)
df_perm_unique = df_perm_filtrado.drop_duplicates(subset=["RUT_TALANA"])
dict_dias_permiso = dict(zip(df_perm_unique["RUT_TALANA"], df_perm_unique["DIAS"]))

# ==========================================
# 5. CONSTRUCCIÓN DE LA SÁBANA
# ==========================================
print("Ejecutando cruces...")
df_ga = pd.read_excel("GestióndeAsistencia202604070840_08ba6579-40d4-4bf5-bf3f-bf4f61b24d1c.xlsx", header=1)

df_ga["RUT_SAP"] = df_ga["Identificador"].apply(clean_rut_sap)
df_ga["RUT_TALANA"] = df_ga["Identificador"].apply(clean_rut_talana)
df_ga.insert(2, "NOMBRE COMPLETO", df_ga["Nombre"] + " " + df_ga["Apellidos"])
df_ga.insert(7, "CECO_1", df_ga["Grupo"].apply(clean_text).map(dict_ceco))

df_final = df_ga[df_ga["CECO_1"].notna()].copy()

def parse_fecha(x):
    try: return pd.to_datetime(str(x).split(' ')[-1], format='%d-%m-%Y')
    except: return pd.NaT

df_final["FECHA_PARSEADA"] = df_final["Fecha"].apply(parse_fecha)

def is_mantener(row):
    term = dict_sap_term.get(row["RUT_SAP"])
    fecha = row["FECHA_PARSEADA"]
    ingreso = dict_tal_ingreso.get(row["RUT_TALANA"])
    
    if pd.isna(term) or term.year == 1899 or (pd.notna(fecha) and term > fecha): return True
    if pd.notna(ingreso) and pd.notna(term) and ingreso > term: return True
    return False

df_final["Mantener"] = df_final.apply(is_mantener, axis=1)
df_final = df_final[df_final["Mantener"]].copy()

def cargo_final_pq(row):
    sap = dict_cargo_sap.get(row["RUT_SAP"])
    if pd.isna(sap) or sap is None or str(sap).strip() == "":
        return dict_cargo_talana_emp.get(row["RUT_TALANA"])
    return sap

df_final["Cargo final"] = df_final.apply(cargo_final_pq, axis=1)

df_final["CARGO_UPPER"] = df_final["Cargo final"].apply(clean_text)
df_final["AGRUPA_CARGO.AGRUPADOR"] = df_final["CARGO_UPPER"].map(dict_cargo_agrup)
df_final["AGRUPADOR_RESCATE"] = df_final["AGRUPA_CARGO.AGRUPADOR"].fillna(df_final["CARGO_UPPER"])
df_final["AGRUPA_CARGO.FTE TEORICO"] = df_final["AGRUPADOR_RESCATE"].map(dict_agrup_fte).fillna(0.0)

df_final["Dias 2"] = df_final["RUT_TALANA"].map(dict_dias_permiso).fillna(0)

def calcular_fte_real(row):
    # Se eliminó la regla que castigaba con el "0" de la hoja INCLS.
    # Ahora, si no superan los 15 días de licencia, reciben directamente su valor de la Columna E.
    if row["Dias 2"] > 15:
        return 0.0
    else:
        return row["AGRUPA_CARGO.FTE TEORICO"]

df_final["FTE REAL"] = df_final.apply(calcular_fte_real, axis=1)
df_final = df_final.drop_duplicates(subset=["Identificador"])

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

real_cargos = df_final.groupby(['CECO_1', 'AGRUPADOR_RESCATE'])['Identificador'].nunique().unstack(fill_value=0)
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
    
    df_export = df_final.drop(columns=["RUT_SAP", "RUT_TALANA", "CARGO_UPPER", "Mantener", "FECHA_PARSEADA", "CECO_1", "AGRUPA_CARGO.AGRUPADOR", "AGRUPADOR_RESCATE"])
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
    print(f"\n✅ REPORTE EXACTO CREADO CON ÉXITO.\nEl archivo está en:\n{ruta}")

except PermissionError:
    print(f"\n❌ ERROR: El archivo '{archivo_salida}' está abierto en Excel. Ciérralo y vuelve a intentar.")