from __future__ import annotations

import re
import unicodedata
from pathlib import Path
import pandas as pd

# =============================
# Config
# =============================
print(">>> Estoy ejecutando este archivo:", __file__)

ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "input"
OUTPUT = ROOT / "output"
OUTPUT.mkdir(exist_ok=True)

TODAY = pd.Timestamp.today().normalize()  # puedes fijarlo: pd.Timestamp("2026-02-05")

print("INPUT =", INPUT)
print("=== Archivos reales en input ===")
for f in INPUT.iterdir():
    print(" -", f.name)
print("================================")


def pick_one(pattern: str) -> Path:
    matches = list(INPUT.glob(pattern))
    if not matches:
        raise FileNotFoundError(f"No encontré '{pattern}' dentro de {INPUT}")
    matches.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return matches[0]


FILES = {
    "gestion": pick_one("GestionAsistencia*.xls*"),
    "bajas": pick_one("Copia_de_bajas*.xls*"),
    "activos": pick_one("Activos_inactivos*.xls*"),
    "talana": pick_one("Lista*Empleados_Talana*.xls*"),
    "permisos": pick_one("PermisosAsignados*.xls*"),
    "fte_aut": pick_one("00 - FTE AUTORIZADO*.xls*"),
    "agrupador": pick_one("Agrupador 5*.xls*"),
}

print("=== Archivos seleccionados ===")
for k, v in FILES.items():
    print(f"{k}: {v.name}")
print("==============================\n")


# =============================
# Helpers
# =============================
def normalize_rut(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip().upper()
    s = re.sub(r"[^0-9K]", "", s)
    return s


def strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def normalize_text(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).upper().replace("\xa0", " ").strip()
    s = re.sub(r"\s+", " ", s)
    return s


def store_key(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip().upper().replace("\xa0", " ")
    s = strip_accents(s)
    s = s.replace("–", "-").replace("—", "-").replace("−", "-")

    for pref in ("OKM ", "OXXO ", "TIENDA ", "LOCAL "):
        if s.startswith(pref):
            s = s[len(pref):].lstrip()

    s = re.sub(r"[^A-Z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def clean_display_name(x) -> str:
    if pd.isna(x):
        return ""
    s = normalize_text(x)
    for pref in ("OKM ", "OXXO "):
        if s.startswith(pref):
            s = s[len(pref):].lstrip()
    return s


def find_col(cols, must_contain_any):
    ucols = [str(c).upper() for c in cols]
    for token in must_contain_any:
        for i, c in enumerate(ucols):
            if token in c:
                return cols[i]
    return None


# =============================
# Lectura FTE Autorizado (tiendas)
# =============================
def pick_latest_month_sheet(sheet_names: list[str]) -> str:
    months = {
        "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
        "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12
    }
    best = None  # (year, month, name)
    for name in sheet_names:
        up = str(name).upper()
        if "_" in up:
            parts = up.split("_")
            if len(parts) >= 2 and parts[0] in months and re.fullmatch(r"\d{2}", parts[1]):
                month = months[parts[0]]
                year = 2000 + int(parts[1])
                cand = (year, month, name)
                if best is None or cand > best:
                    best = cand
    if best:
        return best[2]
    return sheet_names[0]


def read_fte_aut_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None)

    header_row = None
    for i in range(min(200, len(raw))):
        row_vals = [("" if pd.isna(v) else str(v).strip().upper()) for v in raw.iloc[i].tolist()]
        has_ceco = any(v == "CECO" for v in row_vals)
        has_maestra = any("NOMBRE MAESTRA" in v for v in row_vals)
        if has_ceco and has_maestra:
            header_row = i
            break

    if header_row is None:
        raise ValueError(f"No encontré header (CECO / NOMBRE MAESTRA) en la hoja: {sheet_name}")

    df = pd.read_excel(path, sheet_name=sheet_name, header=header_row)
    df.columns = [str(c).strip().upper() for c in df.columns]

    needed = ["CECO", "NOMBRE MAESTRA", "FTE AUT"]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas {missing} en {sheet_name}. Columnas: {list(df.columns)}")

    df = df[needed].copy()
    df["CECO"] = df["CECO"].astype(str).str.strip().str.upper()
    df["NOMBRE MAESTRA"] = df["NOMBRE MAESTRA"].map(normalize_text)
    df["FTE AUT"] = pd.to_numeric(df["FTE AUT"], errors="coerce")

    df = df[df["CECO"].str.match(r"^[A-Z0-9]{4,6}$", na=False)]
    df = df[df["NOMBRE MAESTRA"].astype(str).str.len() > 0]
    df = df[df["FTE AUT"].notna()]
    df = df[df["FTE AUT"].between(0, 500)]

    df["STORE_KEY"] = df["NOMBRE MAESTRA"].map(store_key)
    df["NOMBRE_DISPLAY"] = df["NOMBRE MAESTRA"].map(clean_display_name)

    df = df.drop_duplicates(subset=["CECO"], keep="last").copy()

    # si quieres limitar a tiendas reales:
    df = df[df["CECO"].str.match(r"^50[A-Z0-9]{3}$", na=False)].copy()

    return df


# =============================
# Agrupador 5: CARGO -> AGRUPA_CARGO y FTE base (fallback)
# =============================
def read_agrupador_cargo_map(path: Path) -> pd.DataFrame:
    # IMPORTANTE: header=4 puede variar. Si te falla, lo ajustamos.
    df = pd.read_excel(path, sheet_name="AGRUPADOR", header=4)
    df.columns = [str(c).strip().upper().replace("\xa0", " ") for c in df.columns]

    cargo_col  = "CARGO" if "CARGO" in df.columns else find_col(df.columns, ["CARGO"])
    agrupa_col = "AGRUPA CARGO_2" if "AGRUPA CARGO_2" in df.columns else find_col(df.columns, ["AGRUPA CARGO"])
    fte_col    = "FTE TEORICO" if "FTE TEORICO" in df.columns else find_col(df.columns, ["FTE TEORICO", "FTE"])

    if not cargo_col or not agrupa_col or not fte_col:
        raise ValueError(f"Faltan columnas en AGRUPADOR. Detectado: CARGO={cargo_col}, AGRUPA={agrupa_col}, FTE={fte_col}. Columnas: {list(df.columns)}")

    out = df[[cargo_col, agrupa_col, fte_col]].copy()
    out = out.rename(columns={
        cargo_col: "CARGO",
        agrupa_col: "AGRUPA_CARGO",
        fte_col: "FTE_AGRUPADOR",
    })

    out["CARGO"] = out["CARGO"].map(normalize_text)
    out["AGRUPA_CARGO"] = out["AGRUPA_CARGO"].map(normalize_text)

    out["FTE_AGRUPADOR"] = (
        out["FTE_AGRUPADOR"]
        .astype(str)
        .str.replace(",", ".", regex=False)
    )
    out["FTE_AGRUPADOR"] = pd.to_numeric(out["FTE_AGRUPADOR"], errors="coerce")

    out = out.dropna(subset=["CARGO", "AGRUPA_CARGO", "FTE_AGRUPADOR"])
    out = out[out["FTE_AGRUPADOR"].between(0, 1.2)]
    out = out.drop_duplicates(subset=["CARGO"], keep="last")
    return out


# =============================
# COMPOSICION FTE (en 00 - FTE AUTORIZADO): AGRUPA_CARGO -> FTE_TEORICO_PERSONA
# =============================
def read_composicion_fte(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name="COMPOSICION FTE", header=None)

    header_idx = None
    for i in range(min(80, len(raw))):
        v = raw.iloc[i, 1] if raw.shape[1] > 1 else None
        if isinstance(v, str) and v.strip().upper() == "CARGO":
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("No encontré encabezado 'CARGO' en sheet COMPOSICION FTE")

    df = raw.iloc[header_idx + 1:, [0, 1]].copy()
    df.columns = ["FTE_FACTOR", "AGRUPA_CARGO"]

    df["AGRUPA_CARGO"] = df["AGRUPA_CARGO"].astype(str).map(normalize_text)
    df["FTE_FACTOR"] = pd.to_numeric(df["FTE_FACTOR"], errors="coerce")

    df = df.dropna(subset=["AGRUPA_CARGO", "FTE_FACTOR"])
    df = df[~df["AGRUPA_CARGO"].str.contains(r"^FTE AUTORIZADO$|^CARGO$", regex=True, na=False)]
    df = df.drop_duplicates(subset=["AGRUPA_CARGO"], keep="last")

    df = df.rename(columns={"FTE_FACTOR": "FTE_TEORICO_PERSONA"})
    return df[["AGRUPA_CARGO", "FTE_TEORICO_PERSONA"]]


# =============================
# Lectura Gestión
# =============================
def read_gestion(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Sheet1", header=1)
    df.columns = [str(c).strip().upper() for c in df.columns]

    if "IDENTIFICADOR" not in df.columns or "GRUPO" not in df.columns:
        raise ValueError(f"No encuentro IDENTIFICADOR o GRUPO. Columnas: {list(df.columns)}")

    df = df.rename(columns={"IDENTIFICADOR": "RUT"})
    df["RUT"] = df["RUT"].map(normalize_rut)
    df["GRUPO"] = df["GRUPO"].map(normalize_text)
    df["STORE_KEY"] = df["GRUPO"].map(store_key)

    # CARGO si viene en Gestión (si no, lo traeremos desde Talana igual)
    cargo_col = "CARGO" if "CARGO" in df.columns else find_col(df.columns, ["CARGO", "PUESTO", "POSICION", "POSITION", "JOB"])
    if cargo_col and cargo_col != "CARGO":
        df = df.rename(columns={cargo_col: "CARGO"})

    if "CARGO" in df.columns:
        df["CARGO"] = df["CARGO"].map(normalize_text)
    else:
        df["CARGO"] = ""

    return df


# =============================
# Lectura Activos / Bajas (fechas)
# =============================
def read_bajas(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    df.columns = [str(c).strip().upper() for c in df.columns]

    rut_col = "CHILE RUN - ROL ÚNICO NACIONAL NATIONAL ID INFORMATION"
    egreso_col = "EMPLOYMENT DETAILS TERMINATION DATE"

    if rut_col not in df.columns or egreso_col not in df.columns:
        print("Aviso: Bajas no tiene columnas esperadas. Columnas:", list(df.columns))
        return pd.DataFrame(columns=["RUT", "FECHA EGRESO"])

    out = df[[rut_col, egreso_col]].copy()
    out = out.rename(columns={rut_col: "RUT", egreso_col: "FECHA EGRESO"})
    out["RUT"] = out["RUT"].map(normalize_rut)
    out["FECHA EGRESO"] = pd.to_datetime(out["FECHA EGRESO"], errors="coerce")
    return out.dropna(subset=["RUT"]).drop_duplicates(subset=["RUT"], keep="last")


def read_activos(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    df.columns = [str(c).strip().upper() for c in df.columns]

    rut_col = "CHILE RUN - ROL ÚNICO NACIONAL NATIONAL ID INFORMATION"
    ingreso_col = "EMPLOYMENT DETAILS HIRE DATE"

    if rut_col not in df.columns or ingreso_col not in df.columns:
        print("Aviso: Activos/Inactivos no tiene columnas esperadas. Columnas:", list(df.columns))
        return pd.DataFrame(columns=["RUT", "FECHA INGRESO"])

    out = df[[rut_col, ingreso_col]].copy()
    out = out.rename(columns={rut_col: "RUT", ingreso_col: "FECHA INGRESO"})
    out["RUT"] = out["RUT"].map(normalize_rut)
    out["FECHA INGRESO"] = pd.to_datetime(out["FECHA INGRESO"], errors="coerce")
    return out.dropna(subset=["RUT"]).drop_duplicates(subset=["RUT"], keep="last")


# =============================
# Talana: Inclusión y Cargo
# =============================
def read_talana_inclusion(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    df.columns = [str(c).strip().upper().replace("\xa0", " ") for c in df.columns]

    rut_col = find_col(df.columns, ["RUT", "RUN", "NATIONAL ID", "IDENTIFICADOR"])
    if not rut_col:
        return pd.DataFrame(columns=["RUT", "FTE_INCLUSION"])

    incl_col = find_col(df.columns, ["FTE INCLUSION", "INCLUSION FTE", "INCLUSION"])
    if not incl_col:
        return pd.DataFrame(columns=["RUT", "FTE_INCLUSION"])

    out = df[[rut_col, incl_col]].copy()
    out = out.rename(columns={rut_col: "RUT", incl_col: "FTE_INCLUSION"})
    out["RUT"] = out["RUT"].map(normalize_rut)
    out["FTE_INCLUSION"] = pd.to_numeric(out["FTE_INCLUSION"], errors="coerce")
    out = out.dropna(subset=["RUT", "FTE_INCLUSION"]).drop_duplicates(subset=["RUT"], keep="last")
    return out[["RUT", "FTE_INCLUSION"]]


def read_talana_cargos(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    df.columns = [str(c).strip().upper().replace("\xa0", " ") for c in df.columns]

    rut_col = find_col(df.columns, ["RUT", "RUN", "NATIONAL ID", "IDENTIFICADOR"])
    cargo_col = find_col(df.columns, ["CARGO", "PUESTO", "POSICION", "POSITION", "JOB"])

    if not rut_col or not cargo_col:
        raise ValueError(f"No encontré columnas RUT/CARGO en Talana. Columnas: {list(df.columns)}")

    out = df[[rut_col, cargo_col]].copy()
    out = out.rename(columns={rut_col: "RUT", cargo_col: "CARGO"})
    out["RUT"] = out["RUT"].map(normalize_rut)
    out["CARGO"] = out["CARGO"].map(normalize_text)
    out = out.dropna(subset=["RUT"]).drop_duplicates(subset=["RUT"], keep="last")
    return out


# =============================
# Permisos
# =============================
def read_permisos(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    df.columns = [str(c).strip().upper() for c in df.columns]

    rut_col = "RUT" if "RUT" in df.columns else find_col(df.columns, ["RUT", "RUN", "NATIONAL ID", "IDENTIFICADOR"])
    tipo_col = "TIPO PERMISO" if "TIPO PERMISO" in df.columns else find_col(df.columns, ["TIPO", "MOTIVO", "PERMISO", "AUSENCIA", "CLASE"])
    ini_col = "FECHA INICIO" if "FECHA INICIO" in df.columns else find_col(df.columns, ["INICIO", "DESDE", "START"])
    fin_col = "FECHA FIN" if "FECHA FIN" in df.columns else find_col(df.columns, ["TERMINO", "HASTA", "FIN", "END"])

    if not rut_col or not tipo_col or not ini_col or not fin_col:
        print("Aviso: PermisosAsignados no calza (me faltan columnas). Columnas:", list(df.columns))
        return pd.DataFrame(columns=["RUT", "TIPO", "FECHA_INICIO", "FECHA_FIN", "DURACION_INICIAL"])

    out = df[[rut_col, tipo_col, ini_col, fin_col]].copy()
    out = out.rename(columns={
        rut_col: "RUT",
        tipo_col: "TIPO",
        ini_col: "FECHA_INICIO",
        fin_col: "FECHA_FIN",
    })

    out["RUT"] = out["RUT"].map(normalize_rut)
    out["TIPO"] = out["TIPO"].map(normalize_text)

    out["FECHA_INICIO"] = pd.to_datetime(out["FECHA_INICIO"], errors="coerce", dayfirst=True).dt.normalize()
    out["FECHA_FIN"] = pd.to_datetime(out["FECHA_FIN"], errors="coerce", dayfirst=True).dt.normalize()

    dur = (out["FECHA_FIN"] - out["FECHA_INICIO"]).dt.days + 1
    out["DURACION_INICIAL"] = dur

    out = out.dropna(subset=["RUT", "FECHA_INICIO"])
    return out[["RUT", "TIPO", "FECHA_INICIO", "FECHA_FIN", "DURACION_INICIAL"]]


def permisos_vigentes_hoy(permisos: pd.DataFrame) -> pd.DataFrame:
    if permisos.empty:
        return pd.DataFrame(columns=["RUT", "TIPO", "FECHA_INICIO", "FECHA_FIN", "DURACION_INICIAL", "VIGENTE_HOY"])

    p = permisos.copy()
    p["VIGENTE_HOY"] = (p["FECHA_INICIO"].notna()) & (p["FECHA_INICIO"] <= TODAY) & (
        p["FECHA_FIN"].isna() | (p["FECHA_FIN"] >= TODAY)
    )

    p["_vig"] = p["VIGENTE_HOY"].astype(int)
    p = p.sort_values(by=["RUT", "_vig", "DURACION_INICIAL", "FECHA_INICIO"], ascending=[True, False, False, False])
    p = p.drop_duplicates(subset=["RUT"], keep="first").drop(columns=["_vig"])
    return p


# =============================
# Reglas FTE por persona
# =============================
def aplicar_reglas_fte_persona(base: pd.DataFrame,
                              permisos: pd.DataFrame,
                              inclusion: pd.DataFrame) -> pd.DataFrame:
    out = base.copy()

    p = permisos_vigentes_hoy(permisos)
    out = out.merge(p[["RUT", "TIPO", "VIGENTE_HOY", "DURACION_INICIAL"]], on="RUT", how="left")

    if not inclusion.empty:
        out = out.merge(inclusion, on="RUT", how="left")
    else:
        out["FTE_INCLUSION"] = pd.NA

    out["FTE_TEORICO_PERSONA"] = pd.to_numeric(out["FTE_TEORICO_PERSONA"], errors="coerce").fillna(1.0)

    out["ES_INCLUSION"] = out["FTE_INCLUSION"].notna()

    t = out.get("TIPO", pd.Series([""] * len(out))).fillna("").astype(str)
    out["ES_LICENCIA"] = t.str.contains("LICEN", case=False, na=False)
    out["ES_VACACIONES"] = t.str.contains("VAC", case=False, na=False)

    out["FTE_REAL_PERSONA"] = out["FTE_TEORICO_PERSONA"]

    cond_lic_0 = (out["ES_LICENCIA"]) & (out["VIGENTE_HOY"] == True) & (pd.to_numeric(out["DURACION_INICIAL"], errors="coerce") > 15)
    out.loc[cond_lic_0, "FTE_REAL_PERSONA"] = 0.0

    out.loc[out["ES_INCLUSION"], "FTE_REAL_PERSONA"] = pd.to_numeric(out.loc[out["ES_INCLUSION"], "FTE_INCLUSION"], errors="coerce")

    out["FTE_REAL_PERSONA"] = pd.to_numeric(out["FTE_REAL_PERSONA"], errors="coerce").fillna(0.0)

    return out


# =============================
# Resumen por tienda
# =============================
def resumen_por_tienda(base_persona: pd.DataFrame, universo: pd.DataFrame) -> pd.DataFrame:
    tmp = base_persona.dropna(subset=["CECO"]).copy()

    agg = tmp.groupby("CECO", dropna=False).agg(
        DOTACION_REAL=("RUT", "nunique"),
        FTE_REAL=("FTE_REAL_PERSONA", "sum"),
    ).reset_index()

    res = universo.merge(agg, on="CECO", how="inner")
    res = res[res["DOTACION_REAL"] > 0].copy()

    res = res.rename(columns={
        "FTE_REAL": "FTE REAL",
        "FTE TEORICO": "FTE TEORICO"
    })

    res["BRECHA (REAL-TEORICO)"] = res["FTE REAL"] - res["FTE TEORICO"]

    cols = ["CECO", "GRUPO", "DOTACION_REAL", "FTE TEORICO", "FTE REAL", "BRECHA (REAL-TEORICO)"]
    res = res[cols].sort_values(["CECO"])
    return res


# =============================
# Main
# =============================
def main():
    gestion   = read_gestion(FILES["gestion"])
    activos   = read_activos(FILES["activos"])
    bajas     = read_bajas(FILES["bajas"])
    permisos  = read_permisos(FILES["permisos"])
    inclusion = read_talana_inclusion(FILES["talana"])

    # FTE Autorizado (tiendas)
    xls = pd.ExcelFile(FILES["fte_aut"])
    month_sheet = pick_latest_month_sheet(xls.sheet_names)
    fte_aut = read_fte_aut_sheet(FILES["fte_aut"], sheet_name=month_sheet)

    # Base Gestión -> CECO
    base = gestion.merge(
        fte_aut[["CECO", "STORE_KEY", "FTE AUT"]],
        on="STORE_KEY",
        how="left"
    )

    base = base.merge(activos, on="RUT", how="left")
    base = base.merge(bajas, on="RUT", how="left")

    base = base.rename(columns={
        "FECHA INGRESO": "FECHA DE INGRESO",
        "FECHA EGRESO": "FECHA DE EGRESO",
        "FTE AUT": "FTE TEORICO TIENDA",
    })

    # Cargo desde Talana (para asegurar)
    tal_cargos = read_talana_cargos(FILES["talana"])
    base = base.merge(tal_cargos, on="RUT", how="left", suffixes=("", "_TAL"))

    # Si venía CARGO de gestión y Talana también, priorizamos Talana cuando exista
    if "CARGO_TAL" in base.columns:
        base["CARGO"] = base["CARGO_TAL"].fillna(base["CARGO"])
        base = base.drop(columns=["CARGO_TAL"])

    base["CARGO"] = base["CARGO"].fillna("").map(normalize_text)

    # Mapeos de FTE teórico persona
    agr_map  = read_agrupador_cargo_map(FILES["agrupador"])  # CARGO -> AGRUPA_CARGO + FTE_AGRUPADOR
    comp_map = read_composicion_fte(FILES["fte_aut"])        # AGRUPA_CARGO -> FTE_TEORICO_PERSONA

    base = base.merge(agr_map, on="CARGO", how="left")

    sin_agr = base.loc[base["AGRUPA_CARGO"].isna() & (base["CARGO"] != ""), "CARGO"].value_counts().head(20)
    print("\n=== TOP 20 CARGOS SIN MATCH EN AGRUPADOR ===")
    print(sin_agr.to_string())
    print("===========================================\n", flush=True)

    base = base.merge(comp_map, on="AGRUPA_CARGO", how="left")

    # Fallback: si no hay en composición, usar FTE del agrupador; si tampoco, 1.0
    if "FTE_AGRUPADOR" in base.columns:
        base["FTE_TEORICO_PERSONA"] = base["FTE_TEORICO_PERSONA"].fillna(base["FTE_AGRUPADOR"])

    base["FTE_TEORICO_PERSONA"] = pd.to_numeric(base["FTE_TEORICO_PERSONA"], errors="coerce").fillna(1.0)

    # Universo tiendas
    universo = fte_aut.rename(columns={
        "NOMBRE_DISPLAY": "GRUPO",
        "FTE AUT": "FTE TEORICO",
    })[["CECO", "GRUPO", "FTE TEORICO"]].copy()

    # Reglas por persona
    base_persona = aplicar_reglas_fte_persona(base, permisos, inclusion)

    # Resumen por tienda
    resumen = resumen_por_tienda(base_persona, universo)

    pivot = resumen.pivot_table(
        index=["CECO", "GRUPO"],
        values=["FTE TEORICO", "FTE REAL", "BRECHA (REAL-TEORICO)"],
        aggfunc="sum"
    ).reset_index()

    # Export
    out_path = OUTPUT / f"FTE_resultado_{TODAY.strftime('%Y-%m-%d_%H%M%S')}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        base_persona.to_excel(writer, index=False, sheet_name="BASE_PERSONA")
        resumen.to_excel(writer, index=False, sheet_name="RESUMEN_TIENDAS")
        pivot.to_excel(writer, index=False, sheet_name="PIVOT")

    print(f"OK -> {out_path}")


if __name__ == "__main__":
    main()