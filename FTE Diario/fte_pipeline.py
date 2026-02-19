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

TODAY = pd.Timestamp.today().normalize()  # si quieres fijarlo: pd.Timestamp("2026-02-05")

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
    "agrupador": pick_one("Agrupador 9*.xls*"),

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
    """Devuelve la primera columna cuyo nombre contenga alguno de los tokens."""
    ucols = [c.upper() for c in cols]
    for token in must_contain_any:
        for i, c in enumerate(ucols):
            if token in c:
                return cols[i]
    return None


# =============================
# Lectura FTE Autorizado
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
    for name in sheet_names:
        if "_" in str(name):
            return name
    return sheet_names[0]


def read_fte_aut_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None)

    header_row = None
    for i in range(min(200, len(raw))):
        row_vals = []
        for v in raw.iloc[i].tolist():
            row_vals.append("" if pd.isna(v) else str(v).strip().upper())

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

    # Limpieza: CECO válido + FTE válido
    df = df[df["CECO"].str.match(r"^[A-Z0-9]{4,6}$", na=False)]
    df = df[df["NOMBRE MAESTRA"].astype(str).str.len() > 0]
    df = df[df["FTE AUT"].notna()]
    df = df[df["FTE AUT"].between(0, 500)]

    df["STORE_KEY"] = df["NOMBRE MAESTRA"].map(store_key)
    df["NOMBRE_DISPLAY"] = df["NOMBRE MAESTRA"].map(clean_display_name)

    # dedup por CECO (por si viene repetido)
    df = df.drop_duplicates(subset=["CECO"], keep="last").copy()

    # IMPORTANTE: si quieres limitar a “tiendas reales”:
    df = df[df["CECO"].str.match(r"^50[A-Z0-9]{3}$", na=False)].copy()

    return df

def read_agrupador_cargos(path: Path) -> pd.DataFrame:
    # Leemos sin header para encontrar la fila donde realmente empieza la tabla
    raw = pd.read_excel(path, sheet_name="AGRUPADOR", header=None)

    header_row = None
    for i in range(min(200, len(raw))):
        row = [("" if pd.isna(v) else str(v).strip().upper()) for v in raw.iloc[i].tolist()]
        if "CARGO" in row and any("FTE" in v for v in row):
            header_row = i
            break

    if header_row is None:
        raise ValueError("No encontré la fila header de la tabla de cargos (busqué 'CARGO' y algo con 'FTE').")

    df = pd.read_excel(path, sheet_name="AGRUPADOR", header=header_row)
    df.columns = [str(c).strip().upper().replace("\xa0", " ") for c in df.columns]

    # Columnas que necesitamos (en tu archivo suele venir como 'AGRUPA CARGO_2')
    cargo_col = "CARGO" if "CARGO" in df.columns else find_col(df.columns, ["CARGO"])
    agrupa_col = find_col(df.columns, ["AGRUPA CARGO"])
    fte_col = find_col(df.columns, ["FTE TEORICO", "FTE TEORI", "FTE"])

    if not cargo_col or not agrupa_col or not fte_col:
        raise ValueError(f"No encontré columnas esperadas en la tabla de cargos. Columnas: {list(df.columns)}")

    out = df[[cargo_col, agrupa_col, fte_col]].copy()
    out = out.rename(columns={
        cargo_col: "CARGO",
        agrupa_col: "AGRUPA CARGO",
        fte_col: "FTE_TEORICO_PERSONA"
    })

    out["CARGO"] = out["CARGO"].map(normalize_text)
    out["AGRUPA CARGO"] = out["AGRUPA CARGO"].map(normalize_text)
    out["FTE_TEORICO_PERSONA"] = pd.to_numeric(out["FTE_TEORICO_PERSONA"], errors="coerce")

    out = out.dropna(subset=["CARGO", "FTE_TEORICO_PERSONA"])
    out = out[out["FTE_TEORICO_PERSONA"].between(0, 1.2)]  # limpieza
    out = out.drop_duplicates(subset=["CARGO"], keep="last")

    return out


def read_agrupador_inclusion(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="INCLS", header=3)
    df.columns = [str(c).strip().upper().replace("\xa0", " ") for c in df.columns]

    # Columnas reales en tu archivo: Identificador, AGRUPA CARGO, FTE INCLS, etc.
    rut_col = "IDENTIFICADOR" if "IDENTIFICADOR" in df.columns else find_col(df.columns, ["IDENTIFICADOR", "RUT", "RUN"])
    fte_col = "FTE INCLS" if "FTE INCLS" in df.columns else find_col(df.columns, ["FTE INCLS", "INCLS", "FTE"])

    if not rut_col or not fte_col:
        raise ValueError(f"No encontré Identificador / FTE INCLS en INCLS. Columnas: {list(df.columns)}")

    out = df[[rut_col, fte_col]].copy()
    out = out.rename(columns={rut_col: "RUT", fte_col: "FTE_INCLUSION"})
    out["RUT"] = out["RUT"].map(normalize_rut)
    out["FTE_INCLUSION"] = pd.to_numeric(out["FTE_INCLUSION"], errors="coerce")

    out = out.dropna(subset=["RUT", "FTE_INCLUSION"])
    return out.drop_duplicates(subset=["RUT"], keep="last")


def fte_teorico_desde_cargo(cargo) -> float:
    """
    Regla:
    - Si CARGO contiene PT + horas (20/25/30/etc) => horas/44
    - Si NO contiene PT => 1.0
    """
    c = normalize_text(cargo)  # ya deja MAYUS + espacios normalizados

    # 1) PT30 / PT 30 / PT-30 / PT 30 HRS / PT30HRS
    m = re.search(r"\bPT\s*[-]?\s*(\d{1,2})\b", c)
    if not m:
        m = re.search(r"\bPT(\d{1,2})\b", c)
    if not m:
        m = re.search(r"\bPT\s*[-]?\s*(\d{1,2})\s*(HRS|HORAS)?\b", c)

    # 2) PART TIME 30
    if not m:
        m = re.search(r"\bPART\s*TIME\s*(\d{1,2})\b", c)

    if m:
        horas = int(m.group(1))
        # por seguridad: solo horas razonables
        if 1 <= horas <= 44:
            return round(horas / 44.0, 6)

    # si no aparece PT => full time
    return 1.0

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

    # CARGO (importante para el merge con agrupador)
    # dentro de read_gestion(...)
    cargo_col = "CARGO" if "CARGO" in df.columns else find_col(df.columns, ["CARGO", "PUESTO", "POSICION", "POSITION", "JOB"])
    if cargo_col and cargo_col != "CARGO":
    df = df.rename(columns={cargo_col: "CARGO"})

    df["CARGO"] = df["CARGO"].map(normalize_text) if "CARGO" in df.columns else ""

    # NO fijar FTE acá
    # df["FTE_TEORICO_PERSONA"] = 1.0  <-- QUITAR
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
    return out


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
    return out


# =============================
# Lectura Talana (Inclusión)
# =============================
def read_talana_inclusion(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    df.columns = [str(c).strip().upper() for c in df.columns]

    rut_col = find_col(df.columns, ["RUT", "RUN", "NATIONAL ID"])
    if not rut_col:
        print("Aviso: No encontré columna RUT en Talana. Columnas:", list(df.columns))
        return pd.DataFrame(columns=["RUT", "FTE_INCLUSION"])

    # Busca una columna que diga INCLUSION y/o FTE
    incl_col = find_col(df.columns, ["FTE INCLUSION", "INCLUSION FTE", "INCLUSION"])
    if not incl_col:
        # si no existe, devuelve vacío (no rompe)
        return pd.DataFrame(columns=["RUT", "FTE_INCLUSION"])

    out = df[[rut_col, incl_col]].copy()
    out = out.rename(columns={rut_col: "RUT", incl_col: "FTE_INCLUSION"})
    out["RUT"] = out["RUT"].map(normalize_rut)
    out["FTE_INCLUSION"] = pd.to_numeric(out["FTE_INCLUSION"], errors="coerce")
    out = out.dropna(subset=["RUT"])
    out = out.dropna(subset=["FTE_INCLUSION"])
    return out[["RUT", "FTE_INCLUSION"]].drop_duplicates(subset=["RUT"], keep="last")


# =============================
# Lectura Permisos (Vacaciones / Licencias)
# =============================
def read_permisos(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    df.columns = [str(c).strip().upper() for c in df.columns]

    # En tu archivo existen exactamente estas:
    # RUT, TIPO PERMISO, FECHA INICIO, FECHA FIN
    rut_col = "RUT" if "RUT" in df.columns else find_col(df.columns, ["RUT", "RUN", "NATIONAL ID", "IDENTIFICADOR"])
    tipo_col = "TIPO PERMISO" if "TIPO PERMISO" in df.columns else find_col(df.columns, ["TIPO", "MOTIVO", "PERMISO", "AUSENCIA", "CLASE"])
    ini_col = "FECHA INICIO" if "FECHA INICIO" in df.columns else find_col(df.columns, ["INICIO", "DESDE", "START"])
    fin_col = "FECHA FIN" if "FECHA FIN" in df.columns else find_col(df.columns, ["TERMINO", "HASTA", "FIN", "END"])

    if not rut_col or not tipo_col or not ini_col or not fin_col:
        print("Aviso: PermisosAsignados no calza (me faltan columnas). Columnas:", list(df.columns))
        print("Detectado:", [("RUT", rut_col), ("TIPO", tipo_col), ("INICIO", ini_col), ("FIN", fin_col)])
        # devolvemos esquema estándar vacío (para que no reviente)
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


    # Calcula duración inicial en días (inclusive)
    # Si FECHA_FIN viene vacía, dejamos NaN (no podemos calcular duración)
    dur = (out["FECHA_FIN"] - out["FECHA_INICIO"]).dt.days + 1
    out["DURACION_INICIAL"] = dur

    out = out.dropna(subset=["RUT", "FECHA_INICIO"])
    return out[["RUT", "TIPO", "FECHA_INICIO", "FECHA_FIN", "DURACION_INICIAL"]]


def permisos_vigentes_hoy(permisos: pd.DataFrame) -> pd.DataFrame:
    """devuelve 1 fila por RUT con el 'permiso vigente' (si hay varios, toma el más largo o el más reciente)."""
    if permisos.empty:
        return pd.DataFrame(columns=["RUT", "TIPO", "FECHA_INICIO", "FECHA_FIN", "DURACION_INICIAL", "VIGENTE_HOY"])

    p = permisos.copy()

    # vigente: inicio <= hoy y (fin es NaT o fin >= hoy)
    p["VIGENTE_HOY"] = (p["FECHA_INICIO"].notna()) & (p["FECHA_INICIO"] <= TODAY) & (
        p["FECHA_FIN"].isna() | (p["FECHA_FIN"] >= TODAY)
    )

    # si fin < hoy => no vigente
    # Nos quedamos con todos (porque la regla dice: si ya terminó, se comporta normal)
    # pero para evaluar licencia>15 vigente, usamos VIGENTE_HOY.

    # Si hay varias filas por RUT, priorizamos:
    # 1) Vigente hoy primero
    # 2) Mayor duración inicial
    # 3) Fecha inicio más reciente
    p["_vig"] = p["VIGENTE_HOY"].astype(int)
    p = p.sort_values(by=["RUT", "_vig", "DURACION_INICIAL", "FECHA_INICIO"], ascending=[True, False, False, False])
    p = p.drop_duplicates(subset=["RUT"], keep="first").drop(columns=["_vig"])
    return p


# =============================
# Cálculo FTE por persona (reglas)
# =============================
def aplicar_reglas_fte_persona(base: pd.DataFrame,
                              permisos: pd.DataFrame,
                              inclusion: pd.DataFrame) -> pd.DataFrame:
    out = base.copy()

    # 1) merge permisos (1 por RUT)
    p = permisos_vigentes_hoy(permisos)
    out = out.merge(p[["RUT", "TIPO", "VIGENTE_HOY", "DURACION_INICIAL"]], on="RUT", how="left")

    # 2) merge inclusion
    if not inclusion.empty:
        out = out.merge(inclusion, on="RUT", how="left")
    else:
        out["FTE_INCLUSION"] = pd.NA

    # Normalizaciones
    out["FTE_TEORICO_PERSONA"] = pd.to_numeric(out["FTE_TEORICO_PERSONA"], errors="coerce").fillna(1.0)

    # Flags
    out["ES_INCLUSION"] = out["FTE_INCLUSION"].notna()

    # Detecta licencia / vacaciones por texto
    t = out.get("TIPO", pd.Series([""] * len(out))).fillna("").astype(str)
    out["ES_LICENCIA"] = t.str.contains("LICEN", case=False, na=False)
    out["ES_VACACIONES"] = t.str.contains("VAC", case=False, na=False)

    # 3) Regla base: teorico
    out["FTE_REAL_PERSONA"] = out["FTE_TEORICO_PERSONA"]

    # 4) Licencia: si vigente hoy y duración inicial > 15 => 0
    cond_lic_0 = (out["ES_LICENCIA"]) & (out["VIGENTE_HOY"] == True) & (pd.to_numeric(out["DURACION_INICIAL"], errors="coerce") > 15)
    out.loc[cond_lic_0, "FTE_REAL_PERSONA"] = 0.0

    # 5) Inclusión sobre-escribe todo
    out.loc[out["ES_INCLUSION"], "FTE_REAL_PERSONA"] = pd.to_numeric(out.loc[out["ES_INCLUSION"], "FTE_INCLUSION"], errors="coerce")

    # Seguridad: NaN -> 0
    out["FTE_REAL_PERSONA"] = pd.to_numeric(out["FTE_REAL_PERSONA"], errors="coerce").fillna(0.0)

    return out


# =============================
# Resumen por tienda
# =============================
def resumen_por_tienda(base_persona: pd.DataFrame, universo: pd.DataFrame) -> pd.DataFrame:
    """
    Devuelve 1 fila por CECO con:
      - GRUPO (nombre desde universo)
      - DOTACION_REAL (nunique RUT)
      - FTE REAL (suma FTE_REAL_PERSONA)
      - FTE TEORICO (desde universo)
      - BRECHA
    """
    tmp = base_persona.dropna(subset=["CECO"]).copy()

    agg = tmp.groupby("CECO", dropna=False).agg(
        DOTACION_REAL=("RUT", "nunique"),
        FTE_REAL=("FTE_REAL_PERSONA", "sum"),
    ).reset_index()

    res = universo.merge(agg, on="CECO", how="inner")

    # Solo tiendas con dotación hoy
    res = res[res["DOTACION_REAL"] > 0].copy()

    # Renombres finales (para que quede como tu Excel)
    res = res.rename(columns={
        "FTE_REAL": "FTE REAL",
        "FTE TEORICO": "FTE TEORICO"
    })

    res["BRECHA (REAL-TEORICO)"] = res["FTE REAL"] - res["FTE TEORICO"]

    # Orden de columnas
    cols = ["CECO", "GRUPO", "DOTACION_REAL", "FTE TEORICO", "FTE REAL", "BRECHA (REAL-TEORICO)"]
    res = res[cols].sort_values(["CECO"])

    return res




# =============================
# Main
# =============================
def main():
    gestion = read_gestion(FILES["gestion"])
    activos = read_activos(FILES["activos"])
    bajas = read_bajas(FILES["bajas"])
    permisos = read_permisos(FILES["permisos"])
    inclusion = read_talana_inclusion(FILES["talana"])

    # FTE Autorizado
    xls = pd.ExcelFile(FILES["fte_aut"])
    month_sheet = pick_latest_month_sheet(xls.sheet_names)
    fte_aut = read_fte_aut_sheet(FILES["fte_aut"], sheet_name=month_sheet)

    # Base: Gestión -> CECO por STORE_KEY
    base = gestion.merge(
        fte_aut[["CECO", "STORE_KEY", "FTE AUT"]],
        on="STORE_KEY",
        how="left"
    )

    # fechas por RUT
    base = base.merge(activos, on="RUT", how="left")
    base = base.merge(bajas, on="RUT", how="left")
    base = base.rename(columns={
        "FECHA INGRESO": "FECHA DE INGRESO",
        "FECHA EGRESO": "FECHA DE EGRESO",
        "FTE AUT": "FTE TEORICO TIENDA",
    })
    agr_cargos = read_agrupador_cargos(FILES["agrupador"])

    # Traer AGRUPA CARGO + FTE_TEORICO_PERSONA desde el Agrupador
    base = base.merge(agr_cargos, on="CARGO", how="left")

    # Si algún cargo no matchea, por seguridad:
    base["FTE_TEORICO_PERSONA"] = pd.to_numeric(base["FTE_TEORICO_PERSONA"], errors="coerce").fillna(1.0)
    base["AGRUPA CARGO"] = base["AGRUPA CARGO"].fillna(base["CARGO"])

    # Debug útil (déjalo al principio hasta que todo calce)
    no_match = base[base["AGRUPA CARGO"].isna() | base["FTE_TEORICO_PERSONA"].isna()]
    print("Cargos sin match en Agrupador:", no_match["CARGO"].nunique())
    print(no_match["CARGO"].value_counts().head(25))

    # FTE teórico por persona desde CARGO (PTxx -> xx/44, sino 1.0)
    # FTE teórico por persona desde CARGO
    base["FTE_TEORICO_PERSONA"] = base["CARGO"].apply(fte_teorico_desde_cargo)
    base["FTE_TEORICO_PERSONA"] = pd.to_numeric(base["FTE_TEORICO_PERSONA"], errors="coerce").fillna(1.0)

    print(base["FTE_TEORICO_PERSONA"].describe())
    print(base[["CARGO","FTE_TEORICO_PERSONA"]].drop_duplicates().head(15).to_string(index=False))


    # (opcional) debug rápido
    # print("Cargos sin match:", base.loc[base["FTE_TEORICO_PERSONA"].isna(), "CARGO"].value_counts().head(20))

    # Universo tiendas
    universo = fte_aut.rename(columns={
        "NOMBRE_DISPLAY": "GRUPO",
        "FTE AUT": "FTE TEORICO",
    })[["CECO", "GRUPO", "FTE TEORICO"]].copy()

      


    # Aplica reglas por persona (usa FTE_TEORICO_PERSONA como base)
    base_persona = aplicar_reglas_fte_persona(base, permisos, inclusion)

    # Resumen final
    resumen = resumen_por_tienda(base_persona, universo)

    pivot = resumen.pivot_table(
        index=["CECO", "GRUPO"],
        values=["FTE TEORICO", "FTE REAL", "BRECHA (REAL-TEORICO)"],
        aggfunc="sum"
    ).reset_index()

    out_path = OUTPUT / "FTE_resultado.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        base_persona.to_excel(writer, index=False, sheet_name="BASE_PERSONA")
        resumen.to_excel(writer, index=False, sheet_name="RESUMEN_TIENDAS")
        pivot.to_excel(writer, index=False, sheet_name="PIVOT")

    print(f"OK -> {out_path}")



if __name__ == "__main__":
    main()
