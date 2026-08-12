"""
procesar_categoria_riesgo.py
Base 7/17 — Categoría de Riesgo del Cliente SBS multientidad.

Reportes: B-2309 (Bancos), B-3205 (Financieras), C-120201 (CMAC),
          C-220201 (CRAC), C-4201 (EDPYME)

Lo específico de esta base: columnas fijas por posición (no por nombre de
encabezado como en otras bases): 0=entidad, 1=Normal, 2=ConProblemasPotenciales,
3=Deficiente, 4=Dudoso, 5=Perdida, 6=TotalCreditosDirectosIndirectos. A
diferencia de la mayoría de bases, aquí se incluyen TODAS las filas TOTAL
(no solo la última) porque CMAC tiene dos totales legítimos y distintos
(con y sin CMCP Lima), no un duplicado. La columna final "Empresa" queda
como el nombre_bd resuelto (no se conserva el nombre crudo SBS aparte).
Todo lo compartido viene de utils_sbs.
"""

import re

import pandas as pd

from utils_sbs import (
    BASE_DIR,
    ABR_MES,
    cargar_maestro,
    clasificar_50cb,
    descargar_reporte_bytes,
    fin_de_mes,
    fuzzy_match_entidad,
)

CODIGOS_CORTE = [
    {"tipo": "BANCOS", "codigo": "B-2309"},
    {"tipo": "FINANCIERAS", "codigo": "B-3205"},
    {"tipo": "CMAC", "codigo": "C-120201"},
    {"tipo": "CRAC", "codigo": "C-220201"},
    {"tipo": "EDPYME", "codigo": "C-4201"},
]

UMBRAL_FUZZY = 90

COLUMNAS_FINALES = [
    "PERIODO", "Tipo", "Empresa", "Normal", "ConProblemasPotenciales",
    "Deficiente", "Dudoso", "Perdida", "TotalCreditosDirectosIndirectos", "Clasificación",
]


# ============== utilidades locales ==============

def _norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).strip())


def _es_total(v) -> bool:
    return _norm(v).upper().startswith("TOTAL")


def _es_fin(v) -> bool:
    v = _norm(v)
    if not v:
        return False
    return v.lower().startswith("nota") or v.lower().startswith("fuente") or v.startswith("*") or len(v) > 80


# ============== lectura del archivo ==============

def _find_header_row(raw: pd.DataFrame):
    for r in range(min(12, len(raw))):
        for c in range(raw.shape[1]):
            if _norm(raw.iat[r, c]).lower().startswith("normal"):
                return r
    return None


def _leer_archivo(fuente, tipo: str) -> pd.DataFrame:
    raw = pd.read_excel(fuente, header=None)
    header_row = _find_header_row(raw)
    if header_row is None:
        raise ValueError("No se encontró la fila de encabezado (Normal).")

    r = header_row + 1
    while r < len(raw) and not _norm(raw.iat[r, 0]):
        r += 1
    data_start = r

    registros = []
    for r in range(data_start, len(raw)):
        entidad_sbs = _norm(raw.iat[r, 0])
        if not entidad_sbs:
            continue
        if _es_fin(entidad_sbs):
            break

        vals = {
            "Normal": pd.to_numeric(raw.iat[r, 1], errors="coerce"),
            "ConProblemasPotenciales": pd.to_numeric(raw.iat[r, 2], errors="coerce"),
            "Deficiente": pd.to_numeric(raw.iat[r, 3], errors="coerce"),
            "Dudoso": pd.to_numeric(raw.iat[r, 4], errors="coerce"),
            "Perdida": pd.to_numeric(raw.iat[r, 5], errors="coerce"),
            "TotalCreditosDirectosIndirectos": pd.to_numeric(raw.iat[r, 6], errors="coerce"),
        }
        vals = {k: (0.0 if pd.isna(v) else v) for k, v in vals.items()}

        # a diferencia de otras bases, aquí SÍ se incluyen todas las filas TOTAL (no solo
        # la última): para CMAC hay dos totales legítimos y distintos (con y sin CMCP Lima)
        registros.append({"Tipo": tipo, "Empresa": entidad_sbs, **vals, "es_total": _es_total(entidad_sbs)})

    return pd.DataFrame(registros)


# ============== normalización contra el maestro (fuzzy) ==============

def _normalizar_entidades(df: pd.DataFrame, maestro: pd.DataFrame):
    """Devuelve (df_normalizado, lista_sin_mapeo). La columna Empresa se
    reemplaza por el nombre_bd resuelto (no se conserva el nombre crudo)."""
    empresas_unicas = df["Empresa"].unique()
    mapa = {}
    sin_mapeo = []
    for empresa in empresas_unicas:
        fila = fuzzy_match_entidad(empresa, maestro, umbral=UMBRAL_FUZZY)
        if fila is None:
            sin_mapeo.append(empresa)
        else:
            mapa[empresa] = fila["nombre_bd"]

    df = df.copy()
    if sin_mapeo:
        df = df[~df["Empresa"].isin(sin_mapeo)].copy()

    df["Empresa"] = df["Empresa"].map(mapa)
    df["Clasificación"] = df.apply(
        lambda r: "-" if r["es_total"] else clasificar_50cb(r["Empresa"]), axis=1
    )
    return df, sin_mapeo


# ============== pipeline ==============

def run(anio: int, mes_num: int) -> pd.DataFrame:
    """Descarga, procesa y exporta Categoría de Riesgo para el corte (anio, mes_num)."""
    maestro = cargar_maestro()
    mes_abr = ABR_MES[mes_num]

    partes = []
    for cfg in CODIGOS_CORTE:
        print(f"[Categoría de Riesgo] Descargando {cfg['codigo']} ({cfg['tipo']}) ...")
        contenido = descargar_reporte_bytes(cfg["codigo"], anio, mes_num)
        df_archivo = _leer_archivo(contenido, cfg["tipo"])
        print(f"  -> {len(df_archivo)} filas extraídas")
        partes.append(df_archivo)

    df = pd.concat(partes, ignore_index=True)
    df, sin_mapeo = _normalizar_entidades(df, maestro)

    if sin_mapeo:
        print("\n[Categoría de Riesgo][AVISO] Entidades sin mapeo (excluidas):")
        for e in sin_mapeo:
            print(f"  - {e}")
        print("Agrégalas a maestro_entidades.csv en GitHub y vuelve a correr si quieres incluirlas.\n")

    df["PERIODO"] = fin_de_mes(anio, mes_num)

    resultado = df[COLUMNAS_FINALES]

    output_path = BASE_DIR / f"CategoriaRiesgo_SBS_Consolidado_{mes_abr}{anio}.xlsx"
    resultado.to_excel(output_path, index=False)
    print(f"[Categoría de Riesgo] Listo. {len(resultado)} filas exportadas a: {output_path}")

    return resultado


if __name__ == "__main__":
    from datetime import datetime
    ahora = datetime.now()
    run(ahora.year, ahora.month)
