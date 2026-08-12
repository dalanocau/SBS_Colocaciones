"""
procesar_indicadores.py
Base 15/17 — Indicadores Financieros SBS multientidad.

Reportes: B-2401 (Bancos), B-3301 (Financieras), C-1301 (CMAC),
          C-2301 (CRAC), C-4301 (EDPYME)

Lo específico de esta base: formato TRANSPUESTO puro (entidades en
columnas, indicadores en filas agrupados por sección). ~45 indicadores
únicos reales tras canonización (limpieza de asteriscos, fecha variable
"(al DD/MM/AAAA)" de RCG, símbolos/espaciado, y un diccionario
CANON_INDICADOR para ~14 casos de wording ligeramente distinto entre
familias). Se incluyen las filas TOTAL como entidad más (igual que
Categoría de Riesgo/Patrimonio/RCG), con Clasificación="-".
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
    {"tipo": "Bancos", "codigo": "B-2401"},
    {"tipo": "Financieras", "codigo": "B-3301"},
    {"tipo": "CMAC", "codigo": "C-1301"},
    {"tipo": "CRAC", "codigo": "C-2301"},
    {"tipo": "EDPYME", "codigo": "C-4301"},
]

UMBRAL_FUZZY = 90

CANON_INDICADOR = {
    "Créditos Directos / Personal ( S/ Miles )": "Créditos Directos / Personal (S/ Miles)",
    "Depósitos / Número de Oficinas ( S/ Miles )": "Depósitos / Número de Oficinas (S/ Miles)",
    "Gastos de Administración Anualizado / Activo Productivo Promedio": "Gastos de Administración Anualizados / Activo Productivo Promedio",
    "Gastos de Operación Anualizados / Margen Financiero Total Anualizado(%)": "Gastos de Operación Anualizados / Margen Financiero Total Anualizado (%)",
    "Ingresos Financieros Anualizados / Activo Productivo Promedio": "Ingresos Financieros Anualizados / Activo Productivo Promedio (%)",
    "Pasivo Total / Capital Social y Reservas ( N° de veces )": "Pasivo Total / Capital Social y Reservas (N° de veces)",
    "Pasivo Total / Capital Social y Reservas ( Nº de veces )": "Pasivo Total / Capital Social y Reservas (N° de veces)",
    "Provisiones / Créditos Atrasados": "Provisiones / Créditos Atrasados (%)",
    "Ratio de Liquidez en M.E. (%) (promedio del mes)": "Ratio de Liquidez ME (Promedio de saldos del mes)",
    "Ratio de Liquidez en M.N. (%) (promedio del mes)": "Ratio de Liquidez MN (Promedio de saldos del mes)",
    "Utilidad Anualizada / Activo Promedio": "Utilidad Neta Anualizada / Activo Promedio",
    "Utilidad Neta Anualizada sobre Activo Promedio (%)": "Utilidad Neta Anualizada / Activo Promedio",
    "Utilidad Anualizada / Patrimonio Promedio": "Utilidad Neta Anualizada / Patrimonio Promedio",
    "Utilidad Neta Anualizada sobre Patrimonio Promedio (%)": "Utilidad Neta Anualizada / Patrimonio Promedio",
}

COLUMNAS_FINALES = ["Fecha", "Tipo", "Empresa", "Empresa_Benchmark",
                     "Sección", "Indicador", "Valor", "Clasificación"]


# ============== utilidades locales ==============

def _norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).strip())


def _limpiar_indicador(nombre) -> str:
    """Quita asteriscos de nota al pie, la fecha variable '(al DD/MM/AAAA)' de
    RCG, normaliza símbolos y espaciado, y aplica un diccionario de
    equivalencias para indicadores que la SBS nombra de forma ligeramente
    distinta según la familia."""
    n = str(nombre).replace("*", "")
    n = re.sub(r"\(al\s+\d{1,2}/\d{1,2}/\d{4}\)", "", n)
    n = n.replace("Nº", "N°")
    n = re.sub(r"\(\s+", "(", n)
    n = re.sub(r"\s+\)", ")", n)
    n = re.sub(r"\s+", " ", n).strip()
    return CANON_INDICADOR.get(n, n)


def _es_fin(v) -> bool:
    v = _norm(v)
    if not v:
        return False
    return v.lower().startswith("nota") or v.lower().startswith("fuente") or v.startswith("*") or len(v) > 90


# ============== lectura del archivo (formato transpuesto) ==============

def _find_entity_row(raw: pd.DataFrame):
    """Detecta la fila de nombres de entidad como la que tiene MÁS celdas de
    TEXTO (no numéricas) en columnas 1+ -- distingue la fila de nombres de una
    fila de indicadores con valores numéricos, que puede tener un conteo de
    celdas similar."""
    mejor_r, mejor_n = None, 0
    for r in range(min(15, len(raw))):
        texto_count = 0
        for c in range(1, raw.shape[1]):
            v = raw.iat[r, c]
            if _norm(v) == "":
                continue
            num = pd.to_numeric(v, errors="coerce")
            if pd.isna(num):
                texto_count += 1
        if texto_count > mejor_n:
            mejor_n, mejor_r = texto_count, r
    return mejor_r if mejor_n >= 3 else None


def _leer_archivo(fuente, tipo: str) -> pd.DataFrame:
    raw = pd.read_excel(fuente, header=None)
    entity_row = _find_entity_row(raw)
    if entity_row is None:
        raise ValueError("No se encontró la fila de nombres de entidad.")

    entidad_cols = [(c, _norm(raw.iat[entity_row, c])) for c in range(1, raw.shape[1])
                     if _norm(raw.iat[entity_row, c]) != ""]

    registros = []
    seccion_actual = None
    for r in range(entity_row + 1, len(raw)):
        col0 = _norm(raw.iat[r, 0])
        if not col0:
            continue
        if _es_fin(col0):
            break

        valores_fila = [pd.to_numeric(raw.iat[r, c], errors="coerce") for c, _ in entidad_cols]
        tiene_datos = any(pd.notna(v) for v in valores_fila)

        if not tiene_datos:
            # fila de sección (SOLVENCIA, CALIDAD DE ACTIVOS, etc.) -- sin datos, solo encabezado
            seccion_actual = _limpiar_indicador(col0)
            continue

        indicador = _limpiar_indicador(col0)
        for (c, empresa), val in zip(entidad_cols, valores_fila):
            registros.append({
                "Tipo": tipo, "Empresa": empresa, "Sección": seccion_actual,
                "Indicador": indicador, "Valor": None if pd.isna(val) else val,
            })

    return pd.DataFrame(registros)


# ============== normalización contra el maestro (fuzzy) ==============

def _normalizar_entidades(df: pd.DataFrame, maestro: pd.DataFrame):
    """Devuelve (df_normalizado, lista_sin_mapeo). Filas de entidad TOTAL
    (nombre_bd empieza con "Total") quedan con Clasificación="-"."""
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
    df["Clasificación"] = df["Empresa_Benchmark"].apply(clasificar_50cb)
    df.loc[df["Empresa_Benchmark"].str.startswith("Total", na=False), "Clasificación"] = "-"
    return df, sin_mapeo


# ============== pipeline ==============

def run(anio: int, mes_num: int) -> pd.DataFrame:
    """Descarga, procesa y exporta Indicadores para el corte (anio, mes_num)."""
    maestro = cargar_maestro()
    mes_abr = ABR_MES[mes_num]

    partes = []
    for cfg in CODIGOS_CORTE:
        print(f"[Indicadores] Descargando {cfg['codigo']} ({cfg['tipo']}) ...")
        contenido = descargar_reporte_bytes(cfg["codigo"], anio, mes_num)
        df_archivo = _leer_archivo(contenido, cfg["tipo"])
        print(f"  -> {len(df_archivo)} filas extraídas")
        partes.append(df_archivo)

    df = pd.concat(partes, ignore_index=True)
    df, sin_mapeo = _normalizar_entidades(df, maestro)

    if sin_mapeo:
        print("\n[Indicadores][AVISO] Entidades sin mapeo (excluidas):")
        for e in sin_mapeo:
            print(f"  - {e}")
        print("Agrégalas a maestro_entidades.csv en GitHub y vuelve a correr si quieres incluirlas.\n")

    df["Fecha"] = fin_de_mes(anio, mes_num)

    resultado = df[COLUMNAS_FINALES]

    output_path = BASE_DIR / f"Indicadores_SBS_Consolidado_{mes_abr}{anio}.xlsx"
    resultado.to_excel(output_path, index=False)
    print(f"[Indicadores] Listo. {len(resultado)} filas exportadas a: {output_path}")

    return resultado


if __name__ == "__main__":
    from datetime import datetime
    ahora = datetime.now()
    run(ahora.year, ahora.month)
