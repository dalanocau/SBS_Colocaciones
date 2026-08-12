"""
procesar_castigos.py
Base 4/17 — Castigos SBS multientidad.

Reportes: B-2369 (Bancos), B-3234 (Financieras), C-1253 (CMACs),
          C-2258 (CRACs), C-4242 (EDPYMEs)

Lo específico de esta base: formato largo (una fila por Producto), 7
categorías de Tipo de crédito, valor único de Castigos por bloque. Nota:
esta base usa DOS clasificaciones distintas — "Clasificación" viene de la
columna microfinanciera del maestro (SI/NO -> SMF/SF), mientras que
"Clasificación >=50% MYPE" usa la lista de 19 SMF de siempre
(clasificar_sf_smf). No confundir una con la otra.
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
    {"tipo": "Bancos", "codigo": "B-2369"},
    {"tipo": "Financieras", "codigo": "B-3234"},
    {"tipo": "CMACs", "codigo": "C-1253"},
    {"tipo": "CRACs", "codigo": "C-2258"},
    {"tipo": "Edpymes", "codigo": "C-4242"},
]

UMBRAL_FUZZY = 90

CANON_PRODUCTO = {
    "corporativo": "Corporativo", "corporativos": "Corporativo",
    "grandes empresas": "Grandes Empresas", "medianas empresas": "Medianas Empresas",
    "pequeñas empresas": "Pequeñas Empresas", "pequeña empresas": "Pequeñas Empresas",
    "microempresas": "Microempresas", "micro empresas": "Microempresas", "microempresa": "Microempresas",
    "consumo": "Consumo",
    "hipotecarios": "Hipotecario", "hipotecario": "Hipotecario",
}

COLUMNAS_FINALES = [
    "Fecha", "Tipo", "Clasificación", "Entidad", "Entidad_Final",
    "Castigos", "Producto", "Clasificación >=50% MYPE",
]


# ============== utilidades locales ==============

def _norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).strip())


def _canon_producto(nombre: str) -> str:
    return CANON_PRODUCTO.get(_norm(nombre).lower(), _norm(nombre))


def _es_total(v) -> bool:
    return _norm(v).upper().startswith("TOTAL")


def _es_fin(v) -> bool:
    v = _norm(v)
    if not v:
        return False
    return v.lower().startswith("nota") or v.lower().startswith("fuente") or v.startswith("*") or len(v) > 80


# ============== lectura del archivo de castigos ==============

def _find_categoria_row(raw: pd.DataFrame):
    for r in range(min(12, len(raw))):
        for c in range(raw.shape[1]):
            if _norm(raw.iat[r, c]).lower().startswith("corporativo"):
                return r
    return None


def _leer_archivo(fuente, tipo: str) -> pd.DataFrame:
    raw = pd.read_excel(fuente, header=None)
    cat_row = _find_categoria_row(raw)
    if cat_row is None:
        raise ValueError("No se encontró la fila de categorías (Corporativo/Corporativos).")

    bloques = []
    for c in range(raw.shape[1]):
        v = _norm(raw.iat[cat_row, c]).lower()
        if v and not v.startswith("total") and v != "empresas":
            bloques.append((_canon_producto(v), c))

    r = cat_row + 1
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
            continue
        for producto, col in bloques:
            val = pd.to_numeric(raw.iat[r, col], errors="coerce")
            val = 0.0 if pd.isna(val) else val
            registros.append({"Tipo": tipo, "Entidad": entidad_sbs, "Producto": producto, "Castigos": val})
    return pd.DataFrame(registros)


# ============== normalización contra el maestro (fuzzy) ==============

def _normalizar_entidades(df: pd.DataFrame, maestro: pd.DataFrame):
    """
    Devuelve (df_normalizado, lista_sin_mapeo). Guarda la fila completa del
    maestro por entidad porque se necesitan tanto nombre_bd como
    microfinanciera (para la columna "Clasificación", distinta de la lista
    de 19 SMF que usa "Clasificación >=50% MYPE").
    """
    entidades_unicas = df["Entidad"].unique()
    filas_maestro = {}
    sin_mapeo = []
    for entidad in entidades_unicas:
        fila = fuzzy_match_entidad(entidad, maestro, umbral=UMBRAL_FUZZY)
        if fila is None:
            sin_mapeo.append(entidad)
        else:
            filas_maestro[entidad] = fila

    df = df.copy()
    if sin_mapeo:
        df = df[~df["Entidad"].isin(sin_mapeo)].copy()

    df["Entidad_Final"] = df["Entidad"].map(lambda e: filas_maestro[e]["nombre_bd"])
    df["Clasificación"] = df["Entidad"].map(
        lambda e: "SMF" if filas_maestro[e]["microfinanciera"] == "SI" else "SF"
    )
    df["Clasificación >=50% MYPE"] = df["Entidad_Final"].apply(clasificar_sf_smf)
    return df, sin_mapeo


# ============== pipeline ==============

def run(anio: int, mes_num: int) -> pd.DataFrame:
    """Descarga, procesa y exporta Castigos para el corte (anio, mes_num)."""
    maestro = cargar_maestro()
    mes_abr = ABR_MES[mes_num]

    partes = []
    for cfg in CODIGOS_CORTE:
        print(f"[Castigos] Descargando {cfg['codigo']} ({cfg['tipo']}) ...")
        contenido = descargar_reporte_bytes(cfg["codigo"], anio, mes_num)
        df_archivo = _leer_archivo(contenido, cfg["tipo"])
        print(f"  -> {len(df_archivo)} filas extraídas")
        partes.append(df_archivo)

    df = pd.concat(partes, ignore_index=True)
    df, sin_mapeo = _normalizar_entidades(df, maestro)

    if sin_mapeo:
        print("\n[Castigos][AVISO] Entidades sin mapeo (excluidas):")
        for e in sin_mapeo:
            print(f"  - {e}")
        print("Agrégalas a maestro_entidades.csv en GitHub y vuelve a correr si quieres incluirlas.\n")

    df["Fecha"] = fin_de_mes(anio, mes_num)

    resultado = df[COLUMNAS_FINALES]

    output_path = BASE_DIR / f"Castigos_SBS_Consolidado_{mes_abr}{anio}.xlsx"
    resultado.to_excel(output_path, index=False)
    print(f"[Castigos] Listo. {len(resultado)} filas exportadas a: {output_path}")

    return resultado


if __name__ == "__main__":
    from datetime import datetime
    ahora = datetime.now()
    run(ahora.year, ahora.month)
