from __future__ import annotations

import re
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
    "bajas": pick_one("Copia_de_Bajas*.xls*"),
    "activos": pick_one("Activos_inactivos*.xls*"),
    "talana": pick_one("Lista*Empleados_Talana*.xls*"),  # <- clave
    "permisos": pick_one("PermisosAsignados*.xls*"),
    "fte_aut": pick_one("00 - FTE AUTORIZADO*.xls*"),
    "agrupador": pick_one("Agrupador 5*.xls*"),
}

print("=== Archivos seleccionados ===")
for k, v in FILES.items():
    print(f"{k}: {v.name}")
print("==============================\n")


# -----------------------------

def debug_only_gestion(path: Path):
    import pandas as pd
    xls = pd.ExcelFile(path)
    print("=== DEBUG GestionAsistencia ===")
    print("Archivo:", path)
    print("Hojas:", xls.sheet_names)

    for sh in xls.sheet_names:
        raw = pd.read_excel(path, sheet_name=sh, header=None)
        print(f"\n--- Hoja: {sh} | shape={raw.shape} ---")
        print("Primeras 8 filas (primeras 12 columnas):")
        print(raw.iloc[:8, :12].to_string(index=True, header=False))

        # buscar palabras clave en filas
        for pattern in ["rut", "run", "documento", "grupo", "tienda", "local", "sucursal", "maestra"]:
            mask = raw.apply(lambda r: r.astype(str).str.lower().str.contains(pattern, na=False).any(), axis=1)
            if mask.any():
                idxs = list(raw[mask].index[:10])
                print(f"Filas con '{pattern}': {idxs}")

# Helpers de limpieza
# -----------------------------
def normalize_rut(x) -> str:
    """
    Normaliza RUT a formato sin puntos ni guión (ej: 12.345.678-9 -> 123456789),
    manteniendo K.
    """
    if pd.isna(x):
        return ""
    s = str(x).strip().upper()
    s = re.sub(r"[^0-9K]", "", s)  # deja solo dígitos y K
    return s

def normalize_text(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).upper().replace("\xa0", " ").strip()
    s = re.sub(r"\s+", " ", s)  # colapsa tabs/múltiples espacios
    return s


# -----------------------------
# Lectura inteligente de FTE Autorizado
# (encuentra la fila header donde está CECO / NOMBRE MAESTRA)
# -----------------------------
def read_fte_aut_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None)

    # buscar fila que contenga "CECO" y "NOMBRE MAESTRA" (robusto con NaN)
    header_row = None
    for i in range(min(120, len(raw))):
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

    # Normalizar nombres
    df.columns = [str(c).strip().upper() for c in df.columns]

    needed = ["CECO", "NOMBRE MAESTRA", "FTE AUT"]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError(
            f"Faltan columnas {missing} en {sheet_name}. "
            f"Columnas disponibles: {list(df.columns)}"
        )

    df = df[needed].copy()
    df["CECO"] = df["CECO"].astype(str).str.strip()
    df["NOMBRE MAESTRA"] = df["NOMBRE MAESTRA"].map(normalize_text)
    df["FTE AUT"] = pd.to_numeric(df["FTE AUT"], errors="coerce")
    return df



def pick_latest_month_sheet(sheet_names: list[str]) -> str:
    """
    Heurística simple: prioriza hojas tipo 'ENERO_26', 'NOVIEMBRE_25', etc.
    Si no calza, usa la primera que tenga "_" y números.
    """
    months = {
        "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
        "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12
    }

    best = None  # tuple (year, month, name)
    for name in sheet_names:
        up = name.upper()
        m = re.match(r"^([A-ZÁÉÍÓÚÑ]+)[ _-]?(\d{2})$", up) or re.match(r"^([A-ZÁÉÍÓÚÑ]+)[ _-]?(\d{2})[ _-]?(\d{2})$", up)
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

    # fallback
    for name in sheet_names:
        if "_" in name:
            return name
    return sheet_names[0]

def debug_find_rut_grupo_in_excel(path: Path):
    xls = pd.ExcelFile(path)
    print("Hojas encontradas:", xls.sheet_names)

    for sh in xls.sheet_names:
        raw = pd.read_excel(path, sheet_name=sh, header=None)
        # buscamos filas donde aparezca algo tipo 'rut' en cualquier celda
        mask = raw.apply(
            lambda r: r.astype(str).str.lower().str.contains("rut|run|documento|dni", regex=True, na=False).any(),
            axis=1
        )
        hits = raw[mask]
        if len(hits) > 0:
            print(f"\n>>> En hoja '{sh}' encontré posibles filas con RUT/RUN/DOCUMENTO:")
            for idx in hits.index[:10]:
                row = raw.iloc[idx].astype(str).tolist()
                print(f"Fila {idx} -> {row[:15]}")  # muestra primeras 15 celdas
        else:
            print(f"\nHoja '{sh}': no encontré 'rut/run/documento' en las filas.")

# -----------------------------
# Lectura de tus fuentes principales (ajusta sheet_name/columnas si cambia)
# -----------------------------
def read_gestion(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Sheet1", header=1)
    df.columns = [str(c).strip().upper() for c in df.columns]

    if "IDENTIFICADOR" not in df.columns or "GRUPO" not in df.columns:
        raise ValueError(f"No encuentro IDENTIFICADOR o GRUPO. Columnas: {list(df.columns)}")

    df = df.rename(columns={"IDENTIFICADOR": "RUT"})
    df["RUT"] = df["RUT"].map(normalize_rut)
    df["GRUPO"] = df["GRUPO"].map(normalize_text)
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
    # Ajusta: mínimo RUT + GRUPO o CECO
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
    Versión simple: 1 persona = 1 FTE.
    Si tienes jornada/horas, aquí lo cambias para ponderar (ej: horas/45).
    """
    # asumimos que cada fila en base representa 1 trabajador vigente en la tienda
    grp = base.groupby(["CECO", "GRUPO"], dropna=False).agg(
        DOTACION_REAL=("RUT", "nunique")
    ).reset_index()
    grp["FTE REAL"] = grp["DOTACION_REAL"].astype(float)
    return grp

# -----------------------------
# Pipeline
# -----------------------------
def main():
    # 1) Leer fuentes
    gestion = read_gestion(FILES["gestion"])
    bajas = read_bajas(FILES["bajas"])
    activos = read_activos(FILES["activos"])
    talana = read_talana(FILES["talana"])

    # 2) FTE Autorizado: elegir hoja más reciente automáticamente
    xls = pd.ExcelFile(FILES["fte_aut"])
    month_sheet = pick_latest_month_sheet(xls.sheet_names)
    fte_aut = read_fte_aut_sheet(FILES["fte_aut"], sheet_name=month_sheet)

   # 3) Base del día (solo los que aparecen en Gestion)
base = gestion.merge(
    fte_aut,
    left_on="GRUPO",
    right_on="NOMBRE MAESTRA",
    how="left"
).drop(columns=["NOMBRE MAESTRA"])

# 4) FTE REAL desde el día
fte_real = compute_fte_real(base)  # devuelve CECO, GRUPO, DOTACION_REAL, FTE REAL

# 5) Universo completo de tiendas desde FTE Autorizado
universo = fte_aut.rename(columns={
    "NOMBRE MAESTRA": "GRUPO",
    "FTE AUT": "FTE TEORICO",
}).copy()

# 6) Resumen final: todas las tiendas, aunque no tengan gente hoy
resumen = universo.merge(fte_real, on=["CECO", "GRUPO"], how="left")
resumen["DOTACION_REAL"] = resumen["DOTACION_REAL"].fillna(0).astype(int)
resumen["FTE REAL"] = resumen["FTE REAL"].fillna(0.0)
resumen["BRECHA (REAL-TEORICO)"] = resumen["FTE REAL"] - resumen["FTE TEORICO"]


    # 7) Exportar
    out_path = OUTPUT / "FTE_resultado.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        base.to_excel(writer, index=False, sheet_name="BASE")
        resumen.to_excel(writer, index=False, sheet_name="RESUMEN_TIENDAS")
        pivot.to_excel(writer, index=False, sheet_name="PIVOT")

    print(f"OK -> {out_path}")


if __name__ == "__main__":
    main()

