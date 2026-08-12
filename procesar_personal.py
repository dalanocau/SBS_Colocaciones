"""
procesar_personal.py
Base 3/17 — Personal SBS multientidad.

Reportes: B-2305 (Bancos), B-3202 (Financieras), C-1202 (CMACs),
          C-2202 (CRACs), C-4206 (EDPYMEs)

Lo específico de esta base: una fila por entidad (sin desglose de
producto/situación) con columnas Gerentes/Funcionarios/Empleados/Otros, y
que las filas de subtotal (ej. "TOTAL CAJAS MUNICIPALES...") se saltan con
continue -- no con break -- porque a veces hay más entidades después (ej.
CMCP Lima aparece después del subtotal de Cajas Municipales).
Todo lo compartido viene de utils_sbs.
"""

import re

import pandas as pd

from utils_sbs import (
    BASE_DIR,
    ABR_MES,
    cargar_maestro,
    clasificar_sf_smf,
    descargar_reporte_bytes,
    fin_de_mes,
    fuzzy_match_entidad,
)

CODIGOS_CORTE = [
    {"tipo": "Bancos", "codigo": "B-2305"},
    {"tipo": "Financieras", "codigo": "B-3202"},
    {"tipo": "CMACs", "codigo": "C-1202"},
    {"tipo": "CRACs", "codigo": "C-2202"},
    {"tipo": "Edpymes", "codigo": "C-4206"},
]

# Igual que Depósitos: umbral más exigente que el default de utils_sbs (80).
UMBRAL_FUZZY = 90

COLUMNAS_FINALES = [
    "Fecha", "Tipo de Entidad", "Sistema Microfinanciero", "Nacional",
    "Empresas SBS", "Empresa BD", "Gerentes", "Funcionarios", "Empleados",
    "Otros", "Total", "Clasificación >=50% MYPE",
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


# ============== lectura del archivo de personal ==============

def _find_header_row(raw: pd.DataFrame):
    for r in range(min(12, len(raw))):
        for c in range(raw.shape[1]):
            if _norm(raw.iat[r, c]).lower() == "gerentes":
                return r
    return None


def _leer_archivo(fuente, tipo: str) -> pd.DataFrame:
    raw = pd.read_excel(fuente, header=None)
    header_row = _find_header_row(raw)
    if header_row is None:
        raise ValueError("No se encontró la fila de encabezado (Gerentes).")

    cols = {}
    for c in range(raw.shape[1]):
        v = _norm(raw.iat[header_row, c]).lower()
        if v in ("gerentes", "funcionarios", "empleados", "otros"):
            cols[v] = c

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
        if _es_total(entidad_sbs):
            # fila de subtotal (ej. "TOTAL CAJAS MUNICIPALES DE AHORRO Y CRÉDITO");
            # se salta pero se sigue buscando, porque a veces hay más entidades
            # después (ej. CMCP Lima aparece después del subtotal de Cajas Municipales)
            continue
        registros.append({
            "Tipo": tipo, "Empresa": entidad_sbs,
            "Gerentes": pd.to_numeric(raw.iat[r, cols["gerentes"]], errors="coerce") or 0,
            "Funcionarios": pd.to_numeric(raw.iat[r, cols["funcionarios"]], errors="coerce") or 0,
            "Empleados": pd.to_numeric(raw.iat[r, cols["empleados"]], errors="coerce") or 0,
            "Otros": pd.to_numeric(raw.iat[r, cols["otros"]], errors="coerce") or 0,
        })
    return pd.DataFrame(registros)


# ============== normalización contra el maestro (fuzzy) ==============

def _normalizar_entidades(df: pd.DataFrame, maestro: pd.DataFrame):
    """
    Devuelve (df_normalizado, lista_sin_mapeo). A diferencia de otras bases,
    Personal también necesita las columnas microfinanciera y nacional del
    maestro (no solo nombre_bd), así que se guarda la fila completa por
    empresa en vez de solo el nombre_bd.
    """
    empresas_unicas = df["Empresa"].unique()
    filas_maestro = {}
    sin_mapeo = []
    for empresa in empresas_unicas:
        fila = fuzzy_match_entidad(empresa, maestro, umbral=UMBRAL_FUZZY)
        if fila is None:
            sin_mapeo.append(empresa)
        else:
            filas_maestro[empresa] = fila

    df = df.copy()
    if sin_mapeo:
        df = df[~df["Empresa"].isin(sin_mapeo)].copy()

    df["Empresa BD"] = df["Empresa"].map(lambda e: filas_maestro[e]["nombre_bd"])
    df["Sistema Microfinanciero"] = df["Empresa"].map(lambda e: filas_maestro[e]["microfinanciera"])
    df["Nacional"] = df["Empresa"].map(lambda e: filas_maestro[e]["nacional"])
    df["Clasificación >=50% MYPE"] = df["Empresa BD"].apply(clasificar_sf_smf)
    return df, sin_mapeo


# ============== pipeline ==============

def run(anio: int, mes_num: int) -> pd.DataFrame:
    """Descarga, procesa y exporta Personal para el corte (anio, mes_num)."""
    maestro = cargar_maestro()
    mes_abr = ABR_MES[mes_num]

    partes = []
    for cfg in CODIGOS_CORTE:
        print(f"[Personal] Descargando {cfg['codigo']} ({cfg['tipo']}) ...")
        contenido = descargar_reporte_bytes(cfg["codigo"], anio, mes_num)
        df_archivo = _leer_archivo(contenido, cfg["tipo"])
        print(f"  -> {len(df_archivo)} filas extraídas")
        partes.append(df_archivo)

    df = pd.concat(partes, ignore_index=True)
    df, sin_mapeo = _normalizar_entidades(df, maestro)

    if sin_mapeo:
        print("\n[Personal][AVISO] Entidades sin mapeo (excluidas):")
        for e in sin_mapeo:
            print(f"  - {e}")
        print("Agrégalas a maestro_entidades.csv en GitHub y vuelve a correr si quieres incluirlas.\n")

    df["Total"] = df["Gerentes"] + df["Funcionarios"] + df["Empleados"] + df["Otros"]
    df["Fecha"] = fin_de_mes(anio, mes_num)
    df = df.rename(columns={"Tipo": "Tipo de Entidad", "Empresa": "Empresas SBS"})

    resultado = df[COLUMNAS_FINALES]

    output_path = BASE_DIR / f"Personal_SBS_Consolidado_{mes_abr}{anio}.xlsx"
    resultado.to_excel(output_path, index=False)
    print(f"[Personal] Listo. {len(resultado)} filas exportadas a: {output_path}")

    return resultado


if __name__ == "__main__":
    from datetime import datetime
    ahora = datetime.now()
    run(ahora.year, ahora.month)
