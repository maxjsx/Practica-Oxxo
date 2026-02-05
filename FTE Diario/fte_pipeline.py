from __future__ import annotations

import re
import unicodedata
from pathlib import Path
import pandas as pd

print(">>> Estoy ejecutando este archivo:", __file__)

# -----------------------------
# Config
# -----------------------------
ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "input"
OUTPUT = ROOT / "output"
OUTPUT.mkdir(exist_ok=True)

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


# -----------------------------
# Helpers
# -----------------------------
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
    """
    Llave robusta para cruzar nombres de tienda entre Gestión y FTE Autorizado.
    - Mayúsculas
    - Sin tildes
    - Quita prefijos tipo OKM / OXXO
    - Quita puntuación (guiones, etc.) -> espacios
    - Solo A-Z0-9 y espacios
    """
    if pd.isna(x):
        return ""
    s = str(x).strip().upper().replace("\xa0", " ")
    s = strip_accents(s)

    # normaliza distintos guiones
    s = s.replace("–", "-").replace("—", "-").replace("−", "-")

    # quita prefijos comunes
    for pref in ("OKM ", "OXXO ", "TIENDA ", "LOCAL "):
        if s.startswith(pref):
            s = s[len(pref):].lstrip()

    # deja alfanumérico, resto a espacio
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def clean_display_name(x) -> str:
    # nombre “bonito” (sin OKM, sin dobles espacios, etc.)
    if pd.isna(x):
        return ""
    s = normalize_text(x)
    for pref in ("OKM ", "OXXO "):
        if s.startswith(pref):
            s = s[len(pref):].lstrip()
    return s


# -----------------------------
# FTE autorizado
# -----------------------------
def read_fte_aut_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None)

    header_row = None
    for i in range(min(150, len(raw))):
        row_vals = []
        for v in raw.iloc[i].tolist():
            if pd.isna(v):
                row_vals.append("")
            else:
                row_vals.append(str(v).strip().upper())

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

    # --- LIMPIEZA CRÍTICA: eliminar filas basura ---
    # CECO tipo "50XRJ" (4-6 alfanum)
    df["CECO_OK"] = df["CECO"].str.match(r"^[A-Z0-9]{4,6}$", na=False)
    df = df[df["CECO_OK"]].drop(columns=["CECO_OK"])

    df = df[df["NOMBRE MAESTRA"].astype(str).str.len() > 0]
    df = df[df["FTE AUT"].notna()]

    # (opcional) corta outliers absurdos
    df = df[df["FTE AUT"].between(0, 500)]

    # Llaves
    df["STORE_KEY"] = df["NOMBRE MAESTRA"].map(store_key)
    df["NOMBRE_DISPLAY"] = df["NOMBRE MAESTRA"].map(clean_display_name)

    return df


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


# -----------------------------
# Lecturas fuentes
# -----------------------------
def read_agrupador(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None)

    header_row = None

    # 1) Detectar fila header buscando CR + (CENTRO DE COSTO / TIENDA / NOMBRE / etc.)
    for i in range(min(200, len(raw))):
        row_vals = []
        for v in raw.iloc[i].tolist():
            if pd.isna(v):
                row_vals.append("")
            else:
                row_vals.append(str(v).strip().upper().replace("\xa0", " "))

        joined = " | ".join(row_vals)

        has_code = (
            ("CECO" in joined) or
            ("COST CENTER" in joined) or
            ("CENTRO DE COSTO" in joined) or
            ("COST CENTER CODE" in joined) or
            ("CR" in row_vals)  # <- CLAVE para tu archivo
        )

        has_name = (
            ("CENTRO DE COSTO" in joined) or
            ("TIENDA" in joined) or
            ("NOMBRE" in joined) or
            ("MAESTRA" in joined) or
            ("GRUPO" in joined) or
            ("LOCAL" in joined)
        )

        # Para tu caso: CR + CENTRO DE COSTO es suficiente
        if ("CR" in row_vals and "CENTRO DE COSTO" in joined) or (has_code and has_name):
            header_row = i
            break

    if header_row is None:
        sample = raw.iloc[:20, :12].fillna("").astype(str).to_string(index=True, header=False)
        raise ValueError(
            "No encontré la fila de encabezados en Agrupador 5.\n"
            "Muestra primeras filas:\n" + sample
        )

    df = pd.read_excel(path, header=header_row)
    df.columns = [str(c).strip().upper().replace("\xa0", " ") for c in df.columns]

    # 2) Encontrar columna CECO/código (en tu archivo normalmente es CR)
    ceco_col = None
    # prioridad: CR exacto
    if "CR" in df.columns:
        ceco_col = "CR"
    else:
        # fallback: cualquier cosa que parezca código de centro de costo
        ceco_candidates = [c for c in df.columns if any(k in c for k in ["CECO", "COST CENTER CODE", "CENTRO DE COSTO CODE"])]
        if ceco_candidates:
            ceco_col = ceco_candidates[0]

    if ceco_col is None:
        raise ValueError(f"No encontré columna código (CR/CECO) en Agrupador. Columnas: {list(df.columns)}")

    # 3) Columna nombre (en tu archivo es CENTRO DE COSTO)
    name_col = None
    if "CENTRO DE COSTO" in df.columns:
        name_col = "CENTRO DE COSTO"
    else:
        name_candidates = [c for c in df.columns if any(k in c for k in ["TIENDA", "NOMBRE", "MAESTRA", "GRUPO", "LOCAL"])]
        if name_candidates:
            name_col = name_candidates[0]

    out = df[[ceco_col] + ([name_col] if name_col else [])].copy()
    out = out.rename(columns={ceco_col: "CECO"})

    out["CECO"] = out["CECO"].astype(str).str.strip().str.upper()

    if name_col:
        out = out.rename(columns={name_col: "GRUPO"})
        out["GRUPO"] = out["GRUPO"].map(normalize_text)
    else:
        out["GRUPO"] = ""

    # limpiar vacíos + únicos
    out = out[out["CECO"].astype(str).str.len() > 0]
    out = out.drop_duplicates(subset=["CECO"], keep="last")

    return out[["CECO", "GRUPO"]]



def read_gestion(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Sheet1", header=1)
    df.columns = [str(c).strip().upper() for c in df.columns]

    if "IDENTIFICADOR" not in df.columns or "GRUPO" not in df.columns:
        raise ValueError(f"No encuentro IDENTIFICADOR o GRUPO. Columnas: {list(df.columns)}")

    df = df.rename(columns={"IDENTIFICADOR": "RUT"})
    df["RUT"] = df["RUT"].map(normalize_rut)
    df["GRUPO"] = df["GRUPO"].map(normalize_text)
    df["STORE_KEY"] = df["GRUPO"].map(store_key)
    return df


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


def read_talana(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    df.columns = [str(c).strip().upper() for c in df.columns]
    if "RUT" in df.columns:
        df["RUT"] = df["RUT"].map(normalize_rut)
    if "GRUPO" in df.columns:
        df["GRUPO"] = df["GRUPO"].map(normalize_text)
    return df


# -----------------------------
# Cálculos FTE
# -----------------------------
def compute_fte_real(base: pd.DataFrame) -> pd.DataFrame:
    """
    1 fila por CECO (tienda real). El nombre oficial vendrá desde fte_aut/universo.
    """
    tmp = base.dropna(subset=["CECO"]).copy()

    grp = tmp.groupby("CECO", dropna=False).agg(
        DOTACION_REAL=("RUT", "nunique")
    ).reset_index()

    grp["FTE REAL"] = grp["DOTACION_REAL"].astype(float)
    return grp



# -----------------------------
# Pipeline
# -----------------------------
def main():
    gestion = read_gestion(FILES["gestion"])
    bajas = read_bajas(FILES["bajas"])
    activos = read_activos(FILES["activos"])
    _talana = read_talana(FILES["talana"])
    agrupador = read_agrupador(FILES["agrupador"])

    # FTE Autorizado (hoja más reciente)
    xls = pd.ExcelFile(FILES["fte_aut"])
    month_sheet = pick_latest_month_sheet(xls.sheet_names)
    fte_aut = read_fte_aut_sheet(FILES["fte_aut"], sheet_name=month_sheet)

    # dedup por CECO
    fte_aut = fte_aut.dropna(subset=["CECO"]).copy()
    fte_aut["CECO"] = fte_aut["CECO"].astype(str).str.strip().str.upper()
    fte_aut = fte_aut.drop_duplicates(subset=["CECO"], keep="last")

    # filtrar a universo (234) según agrupador
    agrupador["CECO"] = agrupador["CECO"].astype(str).str.strip().str.upper()
    fte_aut = fte_aut[fte_aut["CECO"].isin(agrupador["CECO"])].copy()

    # BASE del día: cruzar Gestión -> FTE autorizado por STORE_KEY (robusto)
    base = gestion.merge(
        fte_aut[["CECO", "STORE_KEY", "FTE AUT"]],
        on="STORE_KEY",
        how="left"
    )

    # traer fechas por RUT
    base = base.merge(activos, on="RUT", how="left")
    base = base.merge(bajas, on="RUT", how="left")

    base = base.rename(columns={
        "FECHA INGRESO": "FECHA DE INGRESO",
        "FECHA EGRESO": "FECHA DE EGRESO",
        "FTE AUT": "FTE TEORICO",
    })

    

    # FTE REAL desde el día (solo presentes)
    fte_real = compute_fte_real(base)  # ahora devuelve CECO, DOTACION_REAL, FTE REAL

    # Universo tiendas (desde FTE autorizado limpio)
    universo = fte_aut.rename(columns={
        "NOMBRE_DISPLAY": "GRUPO",
        "FTE AUT": "FTE TEORICO",
    })[["CECO", "GRUPO", "FTE TEORICO"]].copy()

    # Resumen SOLO TIENDAS CON DOTACIÓN HOY (inner + filtro)
    resumen = universo.merge(fte_real, on="CECO", how="inner")
    resumen = resumen[resumen["DOTACION_REAL"] > 0].copy()

    resumen["BRECHA (REAL-TEORICO)"] = resumen["FTE REAL"] - resumen["FTE TEORICO"]
    # FTE REAL desde el día (solo donde hay CECO)
    fte_real = compute_fte_real(base)  # devuelve CECO, DOTACION_REAL, FTE REAL

    # Universo tiendas (desde FTE autorizado limpio)
    universo = fte_aut.rename(columns={
        "NOMBRE_DISPLAY": "GRUPO",
        "FTE AUT": "FTE TEORICO",
    })[["CECO", "GRUPO", "FTE TEORICO"]].copy()

    # Resumen SOLO TIENDAS CON DOTACIÓN HOY
    resumen = universo.merge(fte_real, on="CECO", how="inner")
    resumen = resumen[resumen["DOTACION_REAL"] > 0].copy()

    resumen["BRECHA (REAL-TEORICO)"] = resumen["FTE REAL"] - resumen["FTE TEORICO"]

    pivot = resumen.pivot_table(
        index=["CECO", "GRUPO"],
        values=["FTE TEORICO", "FTE REAL", "BRECHA (REAL-TEORICO)"],
        aggfunc="sum"
    ).reset_index()

    print("Tiendas fte_aut (universo tiendas):", fte_aut["CECO"].nunique())
    print("Tiendas con dotación hoy (fte_real):", fte_real["CECO"].nunique())
    print("Tiendas en resumen:", resumen["CECO"].nunique())
    print("Filas base sin CECO (mismatch de nombre):", int(base["CECO"].isna().sum()))


    out_path = OUTPUT / "FTE_resultado.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        base.to_excel(writer, index=False, sheet_name="BASE")
        resumen.to_excel(writer, index=False, sheet_name="RESUMEN_TIENDAS")
        pivot.to_excel(writer, index=False, sheet_name="PIVOT")

    print(f"OK -> {out_path}")

    

if __name__ == "__main__":
    main()


