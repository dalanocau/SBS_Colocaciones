"""
procesar_ratio_liquidez.py
Base 12/17 — Ratio de Liquidez SBS multientidad.

Reportes: B-2340 (Bancos), B-3250 (Financieras), C-1244 (CMAC), C-2249 (CRAC)
Nota: sin EDPYME, no estaba en el script de referencia original para este reporte.

Lo específico de esta base: columnas fijas por posición (Activos Líquidos
MN, Pasivos CP MN, Ratio MN%, [espacio], Activos Líquidos ME, Pasivos CP
ME, Ratio ME%). Totales por sector SÍ incluidos (con ambos de CMAC, con/sin
CMCP Lima). Bug ya resuelto: el título largo del reporte contiene la misma
frase que el encabezado real ("Liquidez en Moneda Nacional"), así que la
detección exige que la celda EMPIECE con esa frase, no que solo la
contenga, para no matchear el título.
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
    {"tipo": "BANCOS", "codigo": "B-2340"},
    {"tipo": "FINANCIERAS", "codigo": "B-3250"},
    {"tipo": "CMAC", "codigo": "C-1244"},
    {"tipo": "CRAC", "codigo": "C-2249"},
]

UMBRAL_FUZZY = 90

ETIQUETAS_TOTAL_EXACTAS = {"CAJAS MUNICIPALES", "CAJAS RURALES DE AHORRO Y CRÉDITO",
                            "EMPRESAS DE CRÉDITOS", "EMPRESAS FINANCIERAS", "BANCA MÚLTIPLE"}

COLUMNAS_FINALES = ["PERIODO", "Tipo", "Empresa", "Activo", "Pasivo", "RL",
                     "ActivoE", "PasivoE", "RLE", "CLASIFICACION"]


# ============== utilidades locales ==============

def _norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).strip())


def _es_fin(v) -> bool:
    v = _norm(v)
    if not v:
        return False
    return v.lower().startswith("nota") or v.lower().startswith("fuente") or v.startswith("*") or len(v) > 80


def _es_total(v) -> bool:
    vu = _norm(v).upper()
    return vu.startswith("TOTAL") or vu in ETIQUETAS_TOTAL_EXACTAS


# ============== lectura del archivo ==============

def _find_top_row(raw: pd.DataFrame):
    """Busca la celda que EMPIEZA con 'Liquidez en Moneda Nacional' (no solo
    la contiene), porque el título completo del reporte también incluye esa
    frase dentro de una oración más larga y produciría un falso positivo."""
    for r in range(min(12, len(raw))):
        for c in range(raw.shape[1]):
            if _norm(raw.iat[r, c]).lower().startswith("liquidez en moneda nacional"):
                return r
    return None


def _leer_archivo(fuente, tipo: str) -> pd.DataFrame:
    raw = pd.read_excel(fuente, header=None)
    top_row = _find_top_row(raw)
    if top_row is None:
        raise ValueError("No se encontró la fila de encabezado (Liquidez en Moneda Nacional).")
    sub_row = top_row + 1

    r = sub_row + 1
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

        def num(c):
            v = pd.to_numeric(raw.iat[r, c], errors="coerce")
            return 0.0 if pd.isna(v) else v

        activo, pasivo, rl = num(1), num(2), num(3)     # Moneda Nacional: (a), (b), (a)/(b)
        activoE, pasivoE, rlE = num(5), num(6), num(7)  # Moneda Extranjera: (c), (d), (c)/(d)

        registros.append({
            "Tipo": tipo, "Empresa": entidad_sbs,
            "Activo": activo, "Pasivo": pasivo, "RL": rl,
            "ActivoE": activoE, "PasivoE": pasivoE, "RLE": rlE,
            "es_total": _es_total(entidad_sbs),
        })

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
    df["CLASIFICACION"] = df.apply(
        lambda r: "-" if r["es_total"] else clasificar_50cb(r["Empresa"]), axis=1
    )
    return df, sin_mapeo


# ============== pipeline ==============

def run(anio: int, mes_num: int) -> pd.DataFrame:
    """Descarga, procesa y exporta Ratio de Liquidez para el corte (anio, mes_num)."""
    maestro = cargar_maestro()
    mes_abr = ABR_MES[mes_num]

    partes = []
    for cfg in CODIGOS_CORTE:
        print(f"[Ratio de Liquidez] Descargando {cfg['codigo']} ({cfg['tipo']}) ...")
        contenido = descargar_reporte_bytes(cfg["codigo"], anio, mes_num)
        df_archivo = _leer_archivo(contenido, cfg["tipo"])
        print(f"  -> {len(df_archivo)} filas extraídas")
        partes.append(df_archivo)

    df = pd.concat(partes, ignore_index=True)
    df, sin_mapeo = _normalizar_entidades(df, maestro)

    if sin_mapeo:
        print("\n[Ratio de Liquidez][AVISO] Entidades sin mapeo (excluidas):")
        for e in sin_mapeo:
            print(f"  - {e}")
        print("Agrégalas a maestro_entidades.csv en GitHub y vuelve a correr si quieres incluirlas.\n")

    df["PERIODO"] = fin_de_mes(anio, mes_num)

    resultado = df[COLUMNAS_FINALES]

    output_path = BASE_DIR / f"RatioLiquidez_SBS_Consolidado_{mes_abr}{anio}.xlsx"
    resultado.to_excel(output_path, index=False)
    print(f"[Ratio de Liquidez] Listo. {len(resultado)} filas exportadas a: {output_path}")

    return resultado


if __name__ == "__main__":
    from datetime import datetime
    ahora = datetime.now()
    run(ahora.year, ahora.month)
