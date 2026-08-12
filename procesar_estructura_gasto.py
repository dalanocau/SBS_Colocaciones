"""
procesar_estructura_gasto.py
Base 10/17 — Estructura de Gasto (Gastos Financieros) SBS multientidad.

Reportes: B-2390 (Bancos), B-3253 (Financieras), C-1239 (CMAC),
          C-2244 (CRAC), C-4233 (EDPYME)

Lo específico de esta base: esquema "unión" de 12 categorías posibles en
porcentaje + Total en soles; para CMAC/CRAC/EDPYME faltan algunas categorías
(se completan con 0). Sin filas TOTAL por sector (a diferencia de Categoría
de Riesgo/Patrimonio/RCG) -- se excluyen todas. Bug ya resuelto: "EC TOTAL
Servicios Financieros" es el nombre real de una EDPYME, no una fila de
total -- _es_total exige startswith("TOTAL") o un match exacto contra un
set de etiquetas sin ese prefijo que la SBS a veces usa.
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
    {"tipo": "Bancos", "codigo": "B-2390"},
    {"tipo": "Financieras", "codigo": "B-3253"},
    {"tipo": "CMAC", "codigo": "C-1239"},
    {"tipo": "CRAC", "codigo": "C-2244"},
    {"tipo": "EDPYME", "codigo": "C-4233"},
]

UMBRAL_FUZZY = 90

CATS_ORDEN = [
    "Obligaciones_con_el_Público", "Dep_del_Sistema_Financiero_y_Org_Internacionales",
    "Fondos_Interbancarios", "Adeudos_y_Obligaciones_Financieras",
    "Obligaciones_en_Circulación_no_Subordinadas", "Obligaciones_en_Circulación_Subordinadas",
    "Por_Valorización_de_Inversiones", "Por_Inversiones_en_Subsidiarias_Asociadas_y_Negocios_Conjuntos",
    "Primas_al_Fondo_de_Seguro_de_Depósitos", "Diferencia_de_Cambio",
    "Productos_Financieros_Derivados", "Otros",
]

ETIQUETAS_TOTAL_EXACTAS = {"CAJAS MUNICIPALES", "CAJAS RURALES DE AHORRO Y CRÉDITO",
                            "EMPRESAS DE CRÉDITOS", "EMPRESAS FINANCIERAS", "BANCA MÚLTIPLE"}

COLUMNAS_FINALES = ["FECHA", "Tipo", "Empresa", "Empresa_Benchmark", *CATS_ORDEN,
                     "Total", ">50%_Cart_Mype"]


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
            or v.lower().startswith("1/") or v.startswith("*") or len(v) > 80)


def _es_total(v) -> bool:
    """Solo excluye si EMPIEZA con TOTAL, o coincide exacto con una etiqueta
    de total sin ese prefijo (la SBS a veces omite la palabra 'TOTAL'). Evita
    falsos positivos como 'EC TOTAL Servicios Financieros', que es el nombre
    real de una entidad, no un total."""
    vu = _norm(v).upper()
    return vu.startswith("TOTAL") or vu in ETIQUETAS_TOTAL_EXACTAS


def _canon_categoria(nombre: str):
    n = _norm(nombre).lower()
    if "no subordinadas" in n:
        return "Obligaciones_en_Circulación_no_Subordinadas"
    if "circulación" in n and "subordinadas" in n:
        return "Obligaciones_en_Circulación_Subordinadas"
    if "obligaciones con el público" in n:
        return "Obligaciones_con_el_Público"
    if "depósitos del sistema" in n:
        return "Dep_del_Sistema_Financiero_y_Org_Internacionales"
    if "fondos interbancarios" in n:
        return "Fondos_Interbancarios"
    if "adeudos y obligaciones" in n:
        return "Adeudos_y_Obligaciones_Financieras"
    if "valorización" in n:
        return "Por_Valorización_de_Inversiones"
    if "subsidiarias" in n:
        return "Por_Inversiones_en_Subsidiarias_Asociadas_y_Negocios_Conjuntos"
    if "primas al fondo" in n:
        return "Primas_al_Fondo_de_Seguro_de_Depósitos"
    if "diferencia de" in n:
        return "Diferencia_de_Cambio"
    if "derivados" in n:
        return "Productos_Financieros_Derivados"
    if "otros" in n:
        return "Otros"
    if "total" in n:
        return "TOTAL_COL"
    return None


# ============== lectura del archivo ==============

def _find_top_row(raw: pd.DataFrame):
    for r in range(min(12, len(raw))):
        for c in range(raw.shape[1]):
            if "intereses y comisiones" in _norm(raw.iat[r, c]).lower():
                return r
    return None


def _leer_archivo(fuente, tipo: str) -> pd.DataFrame:
    raw = pd.read_excel(fuente, header=None)
    top_row = _find_top_row(raw)
    if top_row is None:
        raise ValueError("No se encontró la fila de encabezado (Intereses y Comisiones).")
    sub_row = top_row + 1

    col_map = {}
    total_col = None
    for c in range(1, raw.shape[1]):
        sub_val = _norm(raw.iat[sub_row, c])
        top_val = _norm(raw.iat[top_row, c])
        canon = _canon_categoria(sub_val) if sub_val else (_canon_categoria(top_val) if top_val else None)
        if canon == "TOTAL_COL":
            total_col = c
        elif canon:
            col_map[c] = canon

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
        if _es_total(entidad_sbs):
            continue  # esta base no incluye totales por sector

        fila = {cat: 0.0 for cat in CATS_ORDEN}
        for c, cat in col_map.items():
            val = pd.to_numeric(raw.iat[r, c], errors="coerce")
            fila[cat] = 0.0 if pd.isna(val) else val

        total = pd.to_numeric(raw.iat[r, total_col], errors="coerce") if total_col is not None else None
        total = 0.0 if pd.isna(total) else total

        registros.append({"Tipo": tipo, "Empresa": entidad_sbs, **fila, "Total": total})

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

    df["Empresa_Benchmark"] = df["Empresa"].map(mapa)
    # a diferencia de otras bases: SMFE si está en la lista de 19, si no queda
    # vacía (None), no "SF" -- usar clasificar_50cb pero pisar "SF" por None
    df[">50%_Cart_Mype"] = df["Empresa_Benchmark"].apply(
        lambda nb: "SMFE" if clasificar_50cb(nb) == "SMFE" else None
    )
    return df, sin_mapeo


# ============== pipeline ==============

def run(anio: int, mes_num: int) -> pd.DataFrame:
    """Descarga, procesa y exporta Estructura de Gasto para el corte (anio, mes_num)."""
    maestro = cargar_maestro()
    mes_abr = ABR_MES[mes_num]

    partes = []
    for cfg in CODIGOS_CORTE:
        print(f"[Estructura de Gasto] Descargando {cfg['codigo']} ({cfg['tipo']}) ...")
        contenido = descargar_reporte_bytes(cfg["codigo"], anio, mes_num)
        df_archivo = _leer_archivo(contenido, cfg["tipo"])
        print(f"  -> {len(df_archivo)} filas extraídas")
        partes.append(df_archivo)

    df = pd.concat(partes, ignore_index=True)
    df, sin_mapeo = _normalizar_entidades(df, maestro)

    if sin_mapeo:
        print("\n[Estructura de Gasto][AVISO] Entidades sin mapeo (excluidas):")
        for e in sin_mapeo:
            print(f"  - {e}")
        print("Agrégalas a maestro_entidades.csv en GitHub y vuelve a correr si quieres incluirlas.\n")

    df["FECHA"] = fin_de_mes(anio, mes_num)

    resultado = df[COLUMNAS_FINALES]

    output_path = BASE_DIR / f"EstructuraGasto_SBS_Consolidado_{mes_abr}{anio}.xlsx"
    resultado.to_excel(output_path, index=False)
    print(f"[Estructura de Gasto] Listo. {len(resultado)} filas exportadas a: {output_path}")

    return resultado


if __name__ == "__main__":
    from datetime import datetime
    ahora = datetime.now()
    run(ahora.year, ahora.month)
