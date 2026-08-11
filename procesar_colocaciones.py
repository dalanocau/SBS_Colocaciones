"""
procesar_colocaciones.py
Base 1/17 — Colocaciones SBS multientidad.

Reportes: B-2334 (Bancos), B-3220 (Financieras), C-1228 (CMACs),
          C-2228 (CRACs), C-4223 (Edpymes)

Lo específico de esta base vive acá: detección de estructura horizontal vs.
transpuesta, canon de nombres de producto y el layout de columnas de salida.
Todo lo compartido (descarga, maestro, fuzzy matching, corte de fin de mes,
clasificación SF/SMF/SMFE) viene de utils_sbs.
"""

import re

import pandas as pd

from utils_sbs import (
    BASE_DIR,
    MESES,
    ABR_MES,
    cargar_maestro,
    clasificar_50cb,
    clasificar_sf_smf,
    descargar_reporte_bytes,
    fin_de_mes,
    fuzzy_match_entidad,
)

CODIGOS_CORTE = [
    {"tipo": "Bancos", "codigo": "B-2334"},
    {"tipo": "Financieras", "codigo": "B-3220"},
    {"tipo": "CMACs", "codigo": "C-1228"},
    {"tipo": "CRACs", "codigo": "C-2228"},
    {"tipo": "Edpymes", "codigo": "C-4223"},
]

CANON_PRODUCTO = {
    "corporativo": "Corporativo", "corporativos": "Corporativo",
    "grandes empresas": "Grandes Empresas", "medianas empresas": "Medianas Empresas",
    "pequeñas empresas": "Pequeñas Empresas", "microempresas": "Microempresas",
    "micro empresas": "Microempresas", "consumo": "Consumo",
    "hipotecarios para vivienda": "Hipotecario", "hipotecario": "Hipotecario",
    "hipotecario para vivienda": "Hipotecario",
}

COLUMNAS_FINALES = [
    "mes", "Tipo", "Clasificación", "Empresa", "Empresa_Benchmark",
    "Vigentes", "Reest. y Refin.", "Atrasados", "Total",
    "Producto", "Prod.Consumo", ">50% CB",
]


# ============== utilidades locales ==============

def _norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).strip())


def _canon_producto(nombre: str) -> str:
    return CANON_PRODUCTO.get(_norm(nombre).lower(), _norm(nombre))


def _fin_de_tabla(valor_col0: str) -> bool:
    v = _norm(valor_col0)
    if not v:
        return False
    return v.upper().startswith("TOTAL") or v.lower().startswith("fuente") or v.startswith("*") or len(v) > 80


# ============== formato horizontal (entidades en filas) ==============

def _find_categoria_row_horizontal(raw: pd.DataFrame):
    for r in range(min(12, len(raw))):
        for c in range(raw.shape[1]):
            if _norm(raw.iat[r, c]).lower().startswith("corporativo"):
                count = sum(1 for cc in range(raw.shape[1]) if _norm(raw.iat[r, cc]))
                if count >= 3:
                    return r
    return None


def _find_situacion_row(raw: pd.DataFrame, cat_row: int):
    for r in range(cat_row + 1, min(cat_row + 4, len(raw))):
        count = sum(1 for c in range(raw.shape[1]) if _norm(raw.iat[r, c]).lower() == "vigentes")
        if count >= 2:
            return r
    return None


def _find_subcat_row(raw: pd.DataFrame, cat_row: int, sit_row: int):
    if sit_row == cat_row + 1:
        return None
    for r in range(cat_row + 1, sit_row):
        for c in range(raw.shape[1]):
            if "revolventes" in _norm(raw.iat[r, c]).lower():
                return r
    return None


def _extraer_bloques_horizontal(raw: pd.DataFrame, cat_row: int, subcat_row, sit_row: int):
    categorias = {c: _norm(raw.iat[cat_row, c]) for c in range(raw.shape[1])
                  if _norm(raw.iat[cat_row, c]) and "total" not in _norm(raw.iat[cat_row, c]).lower()}
    cat_cols_sorted = sorted(categorias.keys())

    subcats = {}
    if subcat_row is not None:
        subcats = {c: _norm(raw.iat[subcat_row, c]) for c in range(raw.shape[1]) if _norm(raw.iat[subcat_row, c])}
    subcat_cols_sorted = sorted(subcats.keys())

    def categoria_de(col):
        actual = None
        for cc in cat_cols_sorted:
            if cc <= col:
                actual = categorias[cc]
            else:
                break
        return actual

    def subcat_de(col):
        if not subcat_cols_sorted:
            return None
        actual = None
        for cc in subcat_cols_sorted:
            if cc <= col:
                actual = subcats[cc]
            else:
                break
        return actual

    bloques = []
    for c in range(raw.shape[1]):
        if _norm(raw.iat[sit_row, c]).lower() == "vigentes":
            producto = categoria_de(c)
            if producto is None:
                continue
            prod_consumo = subcat_de(c) if _canon_producto(producto) == "Consumo" else None
            bloques.append((_canon_producto(producto), _norm(prod_consumo) if prod_consumo else None, c, c + 1, c + 2))
    return bloques


def _parse_horizontal(raw: pd.DataFrame, tipo: str):
    cat_row = _find_categoria_row_horizontal(raw)
    if cat_row is None:
        return None
    sit_row = _find_situacion_row(raw, cat_row)
    if sit_row is None:
        return None
    subcat_row = _find_subcat_row(raw, cat_row, sit_row)
    bloques = _extraer_bloques_horizontal(raw, cat_row, subcat_row, sit_row)
    data_start = sit_row + 1

    registros = []
    for r in range(data_start, len(raw)):
        entidad_sbs = _norm(raw.iat[r, 0])
        if not entidad_sbs:
            continue
        if _fin_de_tabla(entidad_sbs):
            break
        for producto, prod_consumo, c_vig, c_reest, c_atr in bloques:
            vig = pd.to_numeric(raw.iat[r, c_vig], errors="coerce") or 0.0
            reest = pd.to_numeric(raw.iat[r, c_reest], errors="coerce") or 0.0
            atr = pd.to_numeric(raw.iat[r, c_atr], errors="coerce") or 0.0
            registros.append({
                "Tipo": tipo, "Empresa": entidad_sbs, "Producto": producto,
                "Prod.Consumo": prod_consumo, "Vigentes": vig,
                "Reest. y Refin.": reest, "Atrasados": atr,
            })
    return pd.DataFrame(registros)


# ============== formato transpuesto (entidades en columnas) ==============

def _parse_transpuesto(raw: pd.DataFrame, tipo: str):
    header_row = None
    for r in range(min(15, len(raw))):
        if _norm(raw.iat[r, 0]).lower() == "tipo de crédito" and _norm(raw.iat[r, 1]).lower() == "situación":
            header_row = r
            break
    if header_row is None:
        return None

    entidades_cols = []
    for c in range(2, raw.shape[1]):
        v = _norm(raw.iat[header_row, c])
        if v and not v.upper().startswith("TOTAL"):
            entidades_cols.append((c, v))

    registros = []
    r = header_row + 1
    categoria_actual = None
    while r < len(raw):
        col0 = _norm(raw.iat[r, 0])
        col1 = _norm(raw.iat[r, 1]).lower()
        if col0:
            categoria_actual = _canon_producto(col0)
        if col1 == "vigentes" and categoria_actual and "total" not in categoria_actual.lower():
            vig_row, reest_row, atr_row = r, r + 1, r + 2
            for c, entidad in entidades_cols:
                vig = pd.to_numeric(raw.iat[vig_row, c], errors="coerce") or 0.0
                reest = pd.to_numeric(raw.iat[reest_row, c], errors="coerce") or 0.0
                atr = pd.to_numeric(raw.iat[atr_row, c], errors="coerce") or 0.0
                registros.append({
                    "Tipo": tipo, "Empresa": entidad, "Producto": categoria_actual,
                    "Prod.Consumo": None, "Vigentes": vig,
                    "Reest. y Refin.": reest, "Atrasados": atr,
                })
            r = atr_row + 1
        else:
            r += 1
    return pd.DataFrame(registros)


def _leer_archivo(fuente, tipo: str) -> pd.DataFrame:
    raw = pd.read_excel(fuente, header=None)

    df = _parse_horizontal(raw, tipo)
    if df is not None and len(df) > 0:
        return df

    df = _parse_transpuesto(raw, tipo)
    if df is not None and len(df) > 0:
        return df

    raise ValueError("No se pudo detectar la estructura del archivo (ni horizontal ni transpuesta).")


# ============== normalización contra el maestro (fuzzy) ==============

def _normalizar_entidades(df: pd.DataFrame, maestro: pd.DataFrame):
    """
    Devuelve (df_normalizado, lista_sin_mapeo). Usa fuzzy_match_entidad de
    utils_sbs (rapidfuzz, processor=str.lower) en vez de un dict exacto, para
    tolerar variantes de nombre. Las entidades sin mapeo se excluyen del
    resultado pero no detienen el proceso.
    """
    empresas_unicas = df["Empresa"].unique()
    mapa = {}
    sin_mapeo = []
    for empresa in empresas_unicas:
        fila = fuzzy_match_entidad(empresa, maestro)
        if fila is None:
            sin_mapeo.append(empresa)
        else:
            mapa[empresa] = fila["nombre_bd"]

    df = df.copy()
    if sin_mapeo:
        df = df[~df["Empresa"].isin(sin_mapeo)].copy()

    df["Empresa_Benchmark"] = df["Empresa"].map(mapa)
    df["Clasificación"] = df["Empresa_Benchmark"].apply(clasificar_sf_smf)
    df[">50% CB"] = df["Empresa_Benchmark"].apply(clasificar_50cb)
    return df, sin_mapeo


# ============== pipeline ==============

def run(anio: int, mes_num: int) -> pd.DataFrame:
    """Descarga, procesa y exporta Colocaciones para el corte (anio, mes_num)."""
    maestro = cargar_maestro()
    mes_nombre = MESES[mes_num]
    mes_abr = ABR_MES[mes_num]

    partes = []
    for cfg in CODIGOS_CORTE:
        print(f"[Colocaciones] Descargando {cfg['codigo']} ({cfg['tipo']}) ...")
        contenido = descargar_reporte_bytes(cfg["codigo"], anio, mes_num)
        df_archivo = _leer_archivo(contenido, cfg["tipo"])
        print(f"  -> {len(df_archivo)} filas extraídas")
        partes.append(df_archivo)

    df = pd.concat(partes, ignore_index=True)
    df, sin_mapeo = _normalizar_entidades(df, maestro)

    if sin_mapeo:
        print("\n[Colocaciones][AVISO] Entidades sin mapeo (excluidas):")
        for e in sin_mapeo:
            print(f"  - {e}")

    df["Total"] = df["Vigentes"] + df["Reest. y Refin."] + df["Atrasados"]
    df["mes"] = fin_de_mes(anio, mes_num)

    resultado = df[COLUMNAS_FINALES]

    output_path = BASE_DIR / f"Colocaciones_SBS_Consolidado_{mes_abr}{anio}.xlsx"
    resultado.to_excel(output_path, index=False)
    print(f"[Colocaciones] Listo. {len(resultado)} filas exportadas a: {output_path}")

    return resultado


if __name__ == "__main__":
    # Permite seguir corriendo este archivo suelto para pruebas puntuales.
    from datetime import datetime
    ahora = datetime.now()
    run(ahora.year, ahora.month)
