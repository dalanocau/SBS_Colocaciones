"""
procesar_depositos.py
Base 2/17 — Depósitos SBS multientidad.

Reportes: B-2372 (Bancos), B-3231 (Financieras), C-1245 (CMACs), C-2250 (CRACs)
Nota: EC (Edpymes) no está incluida — no están autorizadas por la SBS a
captar depósitos del público.

Lo específico de esta base: detección de la fila de categorías por la
palabra "ahorros", el canon de producto (A la vista / Ahorro / A plazo /
CTS — "Depósitos Totales" se excluye a propósito) y el layout de columnas
de salida (Pers Nat / Pers Jur sin fines de lucro / Otras Pers Jur).
Todo lo compartido viene de utils_sbs.
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
    {"tipo": "Bancos", "codigo": "B-2372"},
    {"tipo": "Financieras", "codigo": "B-3231"},
    {"tipo": "CMACs", "codigo": "C-1245"},
    {"tipo": "CRACs", "codigo": "C-2250"},
]

# Umbral de similitud más alto que el default de utils_sbs (80) porque en
# Depósitos varias entidades tienen nombres parecidos entre sí (bancos vs.
# sus "Sucursales") y un umbral más laxo generaba falsos positivos.
UMBRAL_FUZZY = 90

CANON_PRODUCTO_DEP = {
    "depósitos a la vista": "A la vista",
    "depósitos de ahorros": "Ahorro",
    "depósitos a plazo": "A plazo",
    "depósitos cts": "CTS",
}

COLUMNAS_FINALES = [
    "mes", "Tipo", "Clasificación", "Empresa", "Empresa_Benchmark",
    "Pers Nat", "Pers Jur sin fines de lucro", "Otras Pers Jur", "Total",
    "Producto", ">50% CB",
]


# ============== utilidades locales ==============

def _norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).strip())


def _fin_de_tabla(valor_col0) -> bool:
    v = _norm(valor_col0)
    if not v:
        return False
    if v.replace(".", "").isdigit():  # filas basura tipo "0" que a veces trae la SBS
        return True
    return v.upper().startswith("TOTAL") or v.lower().startswith("fuente") or v.startswith("*") or len(v) > 80


# ============== lectura del archivo de depósitos ==============

def _find_categoria_row(raw: pd.DataFrame):
    for r in range(min(12, len(raw))):
        for c in range(raw.shape[1]):
            if "ahorros" in _norm(raw.iat[r, c]).lower():
                return r
    return None


def _leer_archivo(fuente, tipo: str) -> pd.DataFrame:
    raw = pd.read_excel(fuente, header=None)
    cat_row = _find_categoria_row(raw)
    if cat_row is None:
        raise ValueError("No se encontró la fila de categorías (Depósitos de Ahorros).")
    subcat_row = cat_row + 1

    bloques = []
    for c in range(raw.shape[1]):
        v = _norm(raw.iat[cat_row, c]).lower()
        if v in CANON_PRODUCTO_DEP:  # "Depósitos Totales" queda fuera a propósito
            bloques.append((CANON_PRODUCTO_DEP[v], c, c + 1, c + 2))

    r = subcat_row + 1
    while r < len(raw) and not _norm(raw.iat[r, 0]):
        r += 1
    data_start = r

    registros = []
    for r in range(data_start, len(raw)):
        entidad_sbs = _norm(raw.iat[r, 0])
        if not entidad_sbs:
            continue
        if _fin_de_tabla(entidad_sbs):
            break
        for producto, c_nat, c_jur, c_otras in bloques:
            nat = pd.to_numeric(raw.iat[r, c_nat], errors="coerce") or 0.0
            jur = pd.to_numeric(raw.iat[r, c_jur], errors="coerce") or 0.0
            otras = pd.to_numeric(raw.iat[r, c_otras], errors="coerce") or 0.0
            registros.append({
                "Tipo": tipo, "Empresa": entidad_sbs, "Producto": producto,
                "Pers Nat": nat, "Pers Jur sin fines de lucro": jur, "Otras Pers Jur": otras,
            })
    return pd.DataFrame(registros)


# ============== normalización contra el maestro (fuzzy) ==============

def _normalizar_entidades(df: pd.DataFrame, maestro: pd.DataFrame):
    """
    Devuelve (df_normalizado, lista_sin_mapeo). Usa fuzzy_match_entidad de
    utils_sbs (rapidfuzz, processor=str.lower, ya con limpieza de asteriscos)
    con el umbral más exigente de esta base (90). Las entidades sin mapeo se
    excluyen del resultado pero no detienen el proceso.
    """
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
    df["Clasificación"] = df["Empresa_Benchmark"].apply(clasificar_sf_smf)
    df[">50% CB"] = df["Empresa_Benchmark"].apply(clasificar_50cb)
    return df, sin_mapeo


# ============== pipeline ==============

def run(anio: int, mes_num: int) -> pd.DataFrame:
    """Descarga, procesa y exporta Depósitos para el corte (anio, mes_num)."""
    maestro = cargar_maestro()
    mes_abr = ABR_MES[mes_num]

    partes = []
    for cfg in CODIGOS_CORTE:
        print(f"[Depósitos] Descargando {cfg['codigo']} ({cfg['tipo']}) ...")
        contenido = descargar_reporte_bytes(cfg["codigo"], anio, mes_num)
        df_archivo = _leer_archivo(contenido, cfg["tipo"])
        print(f"  -> {len(df_archivo)} filas extraídas")
        partes.append(df_archivo)

    df = pd.concat(partes, ignore_index=True)
    df, sin_mapeo = _normalizar_entidades(df, maestro)

    if sin_mapeo:
        print("\n[Depósitos][AVISO] Entidades sin mapeo (excluidas):")
        for e in sin_mapeo:
            print(f"  - {e}")
        print("Agrégalas a maestro_entidades.csv en GitHub y vuelve a correr si quieres incluirlas.\n")

    df["Total"] = df["Pers Nat"] + df["Pers Jur sin fines de lucro"] + df["Otras Pers Jur"]
    df["mes"] = fin_de_mes(anio, mes_num)

    resultado = df[COLUMNAS_FINALES]

    output_path = BASE_DIR / f"Depositos_SBS_Consolidado_{mes_abr}{anio}.xlsx"
    resultado.to_excel(output_path, index=False)
    print(f"[Depósitos] Listo. {len(resultado)} filas exportadas a: {output_path}")

    return resultado


if __name__ == "__main__":
    from datetime import datetime
    ahora = datetime.now()
    run(ahora.year, ahora.month)
