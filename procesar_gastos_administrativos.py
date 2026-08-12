"""
procesar_gastos_administrativos.py
Base 16/17 — Gastos Administrativos SBS multientidad.

Reportes: B-2348 (Bancos), B-3225 (Financieras), C-1221 (CMACs),
          C-2221 (CRACs), C-4216 (EDPYMEs)

Lo específico de esta base: 6 categorías iguales para todas las familias
(Remuneraciones a Trabajadores, Otros Gastos de Personal, Gastos del
Directorio, Honorarios Profesionales, Otros Servicios Recibidos de
Terceros, Tributos) en % + Total en soles. Totales por sector SÍ incluidos
(con la distinción Total CM / Total CM sin CMCP Lima). Novedad propia:
agrega columnas "saldo_X" por cada categoría = (%/100)*Total, el monto en
soles de cada categoría.
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
    {"tipo": "Bancos", "codigo": "B-2348"},
    {"tipo": "Financieras", "codigo": "B-3225"},
    {"tipo": "CMACs", "codigo": "C-1221"},
    {"tipo": "CRACs", "codigo": "C-2221"},
    {"tipo": "Edpymes", "codigo": "C-4216"},
]

UMBRAL_FUZZY = 90

ETIQUETAS_TOTAL_EXACTAS = {"CAJAS MUNICIPALES", "CAJAS RURALES DE AHORRO Y CRÉDITO",
                            "EMPRESAS DE CRÉDITOS", "EMPRESAS FINANCIERAS", "BANCA MÚLTIPLE"}

CATS_ORDEN = [
    "Remuneraciones_a_Trabajadores", "Otros_Gastos_de_Personal", "Gastos_del_Directorio",
    "Honorarios_Profesionales", "Otros_Servicios_Recibidos_de_Terceros", "Tributos",
]

COLUMNAS_FINALES = ["Fecha", "Tipo", "Empresa", "empresa_bench", *CATS_ORDEN, "Total",
                     *[f"saldo_{c}" for c in CATS_ORDEN], "Clasificación"]


# ============== utilidades locales ==============

def _norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).strip())


def _es_fin(v) -> bool:
    v = _norm(v)
    if not v:
        return False
    return (v.lower().startswith("nota") or v.lower().startswith("fuente")
            or v.startswith("1/") or v.startswith("2/") or v.startswith("3/") or len(v) > 80)


def _es_total(v) -> bool:
    vu = _norm(v).upper()
    return vu.startswith("TOTAL") or vu in ETIQUETAS_TOTAL_EXACTAS


def _canon_categoria(nombre: str):
    n = _norm(nombre).lower()
    if "remuneraci" in n:
        return "Remuneraciones_a_Trabajadores"
    if "otros" in n and "personal" in n:
        return "Otros_Gastos_de_Personal"
    if "directorio" in n:
        return "Gastos_del_Directorio"
    if "honorarios" in n:
        return "Honorarios_Profesionales"
    if "terceros" in n:
        return "Otros_Servicios_Recibidos_de_Terceros"
    if "tributos" in n:
        return "Tributos"
    if "total" in n:
        return "TOTAL_COL"
    return None


# ============== lectura del archivo ==============

def _find_header_row(raw: pd.DataFrame):
    for r in range(min(10, len(raw))):
        if _norm(raw.iat[r, 0]).lower() == "empresas":
            return r
    return None


def _leer_archivo(fuente, tipo: str) -> pd.DataFrame:
    raw = pd.read_excel(fuente, header=None)
    header_row = _find_header_row(raw)
    if header_row is None:
        raise ValueError("No se encontró la fila de encabezado (Empresas).")

    col_map = {}
    total_col = None
    for c in range(1, raw.shape[1]):
        canon = _canon_categoria(raw.iat[header_row, c])
        if canon == "TOTAL_COL":
            total_col = c
        elif canon:
            col_map[c] = canon

    r = header_row + 1
    while r < len(raw) and _norm(raw.iat[r, 0]) == "":
        r += 1
    data_start = r

    registros = []
    for r in range(data_start, len(raw)):
        entidad_sbs = _norm(raw.iat[r, 0])
        if not entidad_sbs:
            continue
        if _es_fin(entidad_sbs):
            break

        fila = {cat: 0.0 for cat in CATS_ORDEN}
        for c, cat in col_map.items():
            val = pd.to_numeric(raw.iat[r, c], errors="coerce")
            fila[cat] = 0.0 if pd.isna(val) else val

        total = pd.to_numeric(raw.iat[r, total_col], errors="coerce") if total_col is not None else None
        total = 0.0 if pd.isna(total) else total

        registros.append({"Tipo": tipo, "Empresa": entidad_sbs, **fila, "Total": total,
                           "es_total": _es_total(entidad_sbs)})

    return pd.DataFrame(registros)


# ============== normalización contra el maestro (fuzzy) ==============

def _normalizar_entidades(df: pd.DataFrame, maestro: pd.DataFrame):
    """Devuelve (df_normalizado, lista_sin_mapeo)."""
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

    df["empresa_bench"] = df["Empresa"].map(mapa)

    for cat in CATS_ORDEN:
        df[f"saldo_{cat}"] = df[cat] / 100 * df["Total"]

    df["Clasificación"] = df.apply(
        lambda r: "-" if r["es_total"] else clasificar_50cb(r["empresa_bench"]), axis=1
    )
    return df, sin_mapeo


# ============== pipeline ==============

def run(anio: int, mes_num: int) -> pd.DataFrame:
    """Descarga, procesa y exporta Gastos Administrativos para el corte (anio, mes_num)."""
    maestro = cargar_maestro()
    mes_abr = ABR_MES[mes_num]

    partes = []
    for cfg in CODIGOS_CORTE:
        print(f"[Gastos Administrativos] Descargando {cfg['codigo']} ({cfg['tipo']}) ...")
        contenido = descargar_reporte_bytes(cfg["codigo"], anio, mes_num)
        df_archivo = _leer_archivo(contenido, cfg["tipo"])
        print(f"  -> {len(df_archivo)} filas extraídas")
        partes.append(df_archivo)

    df = pd.concat(partes, ignore_index=True)
    df, sin_mapeo = _normalizar_entidades(df, maestro)

    if sin_mapeo:
        print("\n[Gastos Administrativos][AVISO] Entidades sin mapeo (excluidas):")
        for e in sin_mapeo:
            print(f"  - {e}")
        print("Agrégalas a maestro_entidades.csv en GitHub y vuelve a correr si quieres incluirlas.\n")

    df["Fecha"] = fin_de_mes(anio, mes_num)

    resultado = df[COLUMNAS_FINALES]

    output_path = BASE_DIR / f"GastosAdministrativos_SBS_Consolidado_{mes_abr}{anio}.xlsx"
    resultado.to_excel(output_path, index=False)
    print(f"[Gastos Administrativos] Listo. {len(resultado)} filas exportadas a: {output_path}")

    return resultado


if __name__ == "__main__":
    from datetime import datetime
    ahora = datetime.now()
    run(ahora.year, ahora.month)
