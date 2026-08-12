"""
procesar_patrimonio_efectivo.py
Base 8/17 — Patrimonio Efectivo SBS multientidad.

Reportes: B-2370 (Bancos), B-3252 (Financieras), C-1257 (CMAC),
          C-2262 (CRAC), C-4246 (EDPYMEs)

Lo específico de esta base: para Bancos/Financieras las columnas a/b/c ya
vienen en soles (Total=a+b+c); para CMAC/CRAC/EDPYME a/b/c vienen en
PORCENTAJE (suman 100%) y solo Total está en soles -- hay que convertir.
Igual que Categoría de Riesgo: se incluyen TODAS las filas TOTAL (CMAC
tiene 2 legítimos y distintos, con y sin CMCP Lima).
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

# es_porcentaje: True = las columnas a/b/c vienen en % (CMAC/CRAC/EDPYME);
# False = ya en soles (Bancos/Financieras)
CODIGOS_CORTE = [
    {"tipo": "BANCOS", "codigo": "B-2370", "es_porcentaje": False},
    {"tipo": "FINANCIERAS", "codigo": "B-3252", "es_porcentaje": False},
    {"tipo": "CMAC", "codigo": "C-1257", "es_porcentaje": True},
    {"tipo": "CRAC", "codigo": "C-2262", "es_porcentaje": True},
    {"tipo": "EDPYMES", "codigo": "C-4246", "es_porcentaje": True},
]

UMBRAL_FUZZY = 90

COLUMNAS_FINALES = [
    "PERIODO", "Empresa", "Empresa Benchmark", "Tipo", "Cap. Ord Nivel 1",
    "Cap. Adi Nivel 1", "Nivel_2", "Total", "Nivel_1_Soles", "Nivel 2 Soles", "Clasificación",
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
    for r in range(min(15, len(raw))):
        for c in range(raw.shape[1]):
            if _norm(raw.iat[r, c]).upper() == "ENTIDAD":
                return r
    return None


def _leer_archivo(fuente, tipo: str, es_porcentaje: bool) -> pd.DataFrame:
    raw = pd.read_excel(fuente, header=None)
    header_row = _find_header_row(raw)
    if header_row is None:
        raise ValueError("No se encontró la fila de encabezado (ENTIDAD).")

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

        a = pd.to_numeric(raw.iat[r, 1], errors="coerce"); a = 0.0 if pd.isna(a) else a
        b = pd.to_numeric(raw.iat[r, 2], errors="coerce"); b = 0.0 if pd.isna(b) else b
        c = pd.to_numeric(raw.iat[r, 3], errors="coerce"); c = 0.0 if pd.isna(c) else c
        total = pd.to_numeric(raw.iat[r, 4], errors="coerce"); total = 0.0 if pd.isna(total) else total

        if es_porcentaje:
            nivel1_soles = (a + b) / 100 * total
            nivel2_soles = c / 100 * total
        else:
            nivel1_soles = a + b
            nivel2_soles = c

        # a diferencia de otras bases, aquí SÍ se incluyen todas las filas TOTAL (igual que en
        # Categoría de Riesgo): para CMAC hay dos totales legítimos y distintos (con y sin CMCP Lima)
        registros.append({
            "Tipo": tipo, "Empresa": entidad_sbs,
            "Cap. Ord Nivel 1": a, "Cap. Adi Nivel 1": b, "Nivel_2": c, "Total": total,
            "Nivel_1_Soles": nivel1_soles, "Nivel 2 Soles": nivel2_soles,
            "es_total": _es_total(entidad_sbs),
        })

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

    df["Empresa Benchmark"] = df["Empresa"].map(mapa)
    df["Clasificación"] = df.apply(
        lambda r: "-" if r["es_total"] else clasificar_50cb(r["Empresa Benchmark"]), axis=1
    )
    return df, sin_mapeo


# ============== pipeline ==============

def run(anio: int, mes_num: int) -> pd.DataFrame:
    """Descarga, procesa y exporta Patrimonio Efectivo para el corte (anio, mes_num)."""
    maestro = cargar_maestro()
    mes_abr = ABR_MES[mes_num]

    partes = []
    for cfg in CODIGOS_CORTE:
        print(f"[Patrimonio Efectivo] Descargando {cfg['codigo']} ({cfg['tipo']}) ...")
        contenido = descargar_reporte_bytes(cfg["codigo"], anio, mes_num)
        df_archivo = _leer_archivo(contenido, cfg["tipo"], cfg["es_porcentaje"])
        print(f"  -> {len(df_archivo)} filas extraídas")
        partes.append(df_archivo)

    df = pd.concat(partes, ignore_index=True)
    df, sin_mapeo = _normalizar_entidades(df, maestro)

    if sin_mapeo:
        print("\n[Patrimonio Efectivo][AVISO] Entidades sin mapeo (excluidas):")
        for e in sin_mapeo:
            print(f"  - {e}")
        print("Agrégalas a maestro_entidades.csv en GitHub y vuelve a correr si quieres incluirlas.\n")

    df["PERIODO"] = fin_de_mes(anio, mes_num)

    resultado = df[COLUMNAS_FINALES]

    output_path = BASE_DIR / f"PatrimonioEfectivo_SBS_Consolidado_{mes_abr}{anio}.xlsx"
    resultado.to_excel(output_path, index=False)
    print(f"[Patrimonio Efectivo] Listo. {len(resultado)} filas exportadas a: {output_path}")

    return resultado


if __name__ == "__main__":
    from datetime import datetime
    ahora = datetime.now()
    run(ahora.year, ahora.month)
