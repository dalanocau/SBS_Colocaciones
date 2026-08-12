"""
procesar_ingresos_financieros.py
Base 11/17 — Ingresos Financieros SBS multientidad.

Reportes: B-2347 (Bancos), B-3224 (Financieras), C-1220 (CMAC),
          C-2220 (CRAC), C-4215 (EDPYME)

Lo específico de esta base: estructura similar a Estructura de Gasto, pero
las mismas categorías para todas las familias (sin faltantes) y sin filas
TOTAL por sector. Bug propio ya resuelto: el archivo de Financieras
(B-3224) trae una columna en blanco extra al inicio -- el nombre de
entidad está en la columna 1, no en la 0 como los demás archivos -- se
detecta dinámicamente en qué columna aparece la palabra "Empresas".
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
    {"tipo": "Bancos", "codigo": "B-2347"},
    {"tipo": "Financieras", "codigo": "B-3224"},
    {"tipo": "CMAC", "codigo": "C-1220"},
    {"tipo": "CRAC", "codigo": "C-2220"},
    {"tipo": "EDPYME", "codigo": "C-4215"},
]

UMBRAL_FUZZY = 90

CATS_ORDEN = [
    "Disponible", "Fondos_Interbancarios", "Inversiones", "Créditos",
    "Por_Valorización_de_Inversiones", "Por_Inversiones_en_Subsidiarias_Asociadas_y_Negocios_Conjuntos",
    "Diferencia_de_Cambio", "Productos_Financieros_Derivados", "Otros",
]

ETIQUETAS_TOTAL_EXACTAS = {"CAJAS MUNICIPALES", "CAJAS RURALES DE AHORRO Y CRÉDITO",
                            "EMPRESAS DE CRÉDITOS", "EMPRESAS FINANCIERAS", "BANCA MÚLTIPLE"}

COLUMNAS_FINALES = ["FECHA", "Tipo", "Empresa", "Empresa_Benchmark", *CATS_ORDEN,
                     "Total", ">50% Cart. Mype"]


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
    de total sin ese prefijo. Evita falsos positivos como 'EDPYME Servicios
    Financieros TOTAL', que es el nombre real de una entidad."""
    vu = _norm(v).upper()
    return vu.startswith("TOTAL") or vu in ETIQUETAS_TOTAL_EXACTAS


def _canon_categoria(nombre: str):
    n = _norm(nombre).lower()
    if n == "disponible":
        return "Disponible"
    if "fondos interbancarios" in n:
        return "Fondos_Interbancarios"
    if n == "inversiones":
        return "Inversiones"
    if n in ("créditos", "creditos"):
        return "Créditos"
    if "valorización" in n:
        return "Por_Valorización_de_Inversiones"
    if "subsidiarias" in n:
        return "Por_Inversiones_en_Subsidiarias_Asociadas_y_Negocios_Conjuntos"
    if "diferencia de" in n:
        return "Diferencia_de_Cambio"
    if "derivados" in n:
        return "Productos_Financieros_Derivados"
    if n == "otros" or n.startswith("otros"):
        return "Otros"
    if "total" in n:
        return "TOTAL_COL"
    return None


# ============== lectura del archivo ==============

def _find_top_row_y_entidad_col(raw: pd.DataFrame):
    """Busca la fila con 'Intereses y Comisiones' y, en esa misma fila, la
    columna donde aparece 'Empresas' -- algunos archivos (ej. Financieras)
    traen una columna en blanco extra al inicio, así que el nombre de
    entidad no siempre está en la columna 0."""
    for r in range(min(12, len(raw))):
        for c in range(raw.shape[1]):
            if "intereses y comisiones" in _norm(raw.iat[r, c]).lower():
                entidad_col = 0
                for cc in range(raw.shape[1]):
                    if "empresas" in _norm(raw.iat[r, cc]).lower():
                        entidad_col = cc
                        break
                return r, entidad_col
    return None, 0


def _leer_archivo(fuente, tipo: str) -> pd.DataFrame:
    raw = pd.read_excel(fuente, header=None)
    top_row, entidad_col = _find_top_row_y_entidad_col(raw)
    if top_row is None:
        raise ValueError("No se encontró la fila de encabezado (Intereses y Comisiones).")
    sub_row = top_row + 1

    col_map = {}
    total_col = None
    for c in range(entidad_col + 1, raw.shape[1]):
        sub_val = _norm(raw.iat[sub_row, c])
        top_val = _norm(raw.iat[top_row, c])
        canon = _canon_categoria(sub_val) if sub_val else (_canon_categoria(top_val) if top_val else None)
        if canon == "TOTAL_COL":
            total_col = c
        elif canon:
            col_map[c] = canon

    r = sub_row + 1
    while r < len(raw) and not _norm(raw.iat[r, entidad_col]):
        r += 1
    data_start = r

    registros = []
    for r in range(data_start, len(raw)):
        entidad_sbs = _norm(raw.iat[r, entidad_col])
        if not entidad_sbs:
            continue
        if _es_fin(entidad_sbs):
            break
        if _es_total(entidad_sbs):
            continue

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
    # igual que Estructura de Gasto: SMFE si está en la lista de 19, si no vacía (None)
    df[">50% Cart. Mype"] = df["Empresa_Benchmark"].apply(
        lambda nb: "SMFE" if clasificar_50cb(nb) == "SMFE" else None
    )
    return df, sin_mapeo


# ============== pipeline ==============

def run(anio: int, mes_num: int) -> pd.DataFrame:
    """Descarga, procesa y exporta Ingresos Financieros para el corte (anio, mes_num)."""
    maestro = cargar_maestro()
    mes_abr = ABR_MES[mes_num]

    partes = []
    for cfg in CODIGOS_CORTE:
        print(f"[Ingresos Financieros] Descargando {cfg['codigo']} ({cfg['tipo']}) ...")
        contenido = descargar_reporte_bytes(cfg["codigo"], anio, mes_num)
        df_archivo = _leer_archivo(contenido, cfg["tipo"])
        print(f"  -> {len(df_archivo)} filas extraídas")
        partes.append(df_archivo)

    df = pd.concat(partes, ignore_index=True)
    df, sin_mapeo = _normalizar_entidades(df, maestro)

    if sin_mapeo:
        print("\n[Ingresos Financieros][AVISO] Entidades sin mapeo (excluidas):")
        for e in sin_mapeo:
            print(f"  - {e}")
        print("Agrégalas a maestro_entidades.csv en GitHub y vuelve a correr si quieres incluirlas.\n")

    df["FECHA"] = fin_de_mes(anio, mes_num)

    resultado = df[COLUMNAS_FINALES]

    output_path = BASE_DIR / f"IngresosFinancieros_SBS_Consolidado_{mes_abr}{anio}.xlsx"
    resultado.to_excel(output_path, index=False)
    print(f"[Ingresos Financieros] Listo. {len(resultado)} filas exportadas a: {output_path}")

    return resultado


if __name__ == "__main__":
    from datetime import datetime
    ahora = datetime.now()
    run(ahora.year, ahora.month)
