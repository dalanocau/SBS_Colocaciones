"""
procesar_clientes_credito.py
Base 5/17 — Clientes de Crédito SBS multientidad.

Reportes: B-230803 (Bancos), B-3218 (Financieras), C-1231 (CMACs),
          C-2231 (CRACs), C-4226 (EDPYMEs)

Lo específico de esta base: a diferencia de Colocaciones/Castigos, el Total
NO es la suma de las 7 categorías (un cliente puede tener créditos de
varios tipos y se contaría doble) — se toma el "(X) Total" tal cual lo
reporta la SBS, incluida la fila TOTAL por sector como una entidad más
(con Clasificación="-"). También hay un bug de detección de encabezado ya
resuelto: en Bancos/Financieras hay una fila previa con solo el nombre
corto de categoría (sin columna Total) antes de la fila real de
encabezado — hay que quedarse con la ÚLTIMA coincidencia de "corporativo",
no la primera, o se pierde la columna Total.
Todo lo compartido viene de utils_sbs.
"""

import re

import pandas as pd

from utils_sbs import (
    BASE_DIR,
    ABR_MES,
    cargar_maestro,
    clasificar_50cb,
    clasificar_sf_smf,
    descargar_reporte_bytes,
    fin_de_mes,
    fuzzy_match_entidad,
)

CODIGOS_CORTE = [
    {"tipo": "Bancos", "codigo": "B-230803"},
    {"tipo": "Financieras", "codigo": "B-3218"},
    {"tipo": "CMACs", "codigo": "C-1231"},
    {"tipo": "CRACs", "codigo": "C-2231"},
    {"tipo": "Edpymes", "codigo": "C-4226"},
]

UMBRAL_FUZZY = 90

# Lista (no dict) porque se matchea por substring, en orden, contra el
# encabezado crudo de columna (ej. "Deudores Corporativos" contiene "corporativo").
CANON_PRODUCTO = [
    ("corporativo", "Corporativo"),
    ("grandes empresas", "Grandes Empresas"),
    ("medianas empresas", "Medianas Empresas"),
    ("pequeñas empresas", "Pequeñas Empresas"),
    ("microempresas", "Microempresas"),
    ("consumo", "Consumo"),
    ("hipotecario", "Hipotecario"),
    ("total", "(X) Total"),
]

COLUMNAS_FINALES = [
    "Fecha", "Tipo", "Clasificación", "Empresa", "Empresa Benchmark",
    "Nº de clientes", "Producto", ">50% CB",
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
    return (v.lower().startswith("nota") or v.lower().startswith("fuente")
            or v.startswith("*") or v.startswith("http") or len(v) > 80)


def _clasificar_columna(header_text: str):
    h = _norm(header_text).lower()
    for clave, canon in CANON_PRODUCTO:
        if clave in h:
            return canon
    return None


# ============== lectura del archivo de clientes de crédito ==============

def _find_header_row(raw: pd.DataFrame):
    """
    Toma la ÚLTIMA fila que menciona 'corporativo' en la ventana de
    búsqueda: en Bancos/Financieras hay una fila previa con solo el nombre
    corto de categoría (sin la columna Total) y luego la fila real de
    encabezado ('Deudores Corporativos'...Total de deudores) -- quedarnos
    con la primera perdería la columna Total.
    """
    ultimo = None
    for r in range(min(12, len(raw))):
        for c in range(raw.shape[1]):
            if "corporativo" in _norm(raw.iat[r, c]).lower():
                ultimo = r
                break
    return ultimo


def _leer_archivo(fuente, tipo: str) -> pd.DataFrame:
    raw = pd.read_excel(fuente, header=None)
    header_row = _find_header_row(raw)
    if header_row is None:
        raise ValueError("No se encontró la fila de encabezado (Corporativo).")

    bloques = []
    for c in range(raw.shape[1]):
        canon = _clasificar_columna(raw.iat[header_row, c])
        if canon:
            bloques.append((canon, c))

    r = header_row + 1
    while r < len(raw) and not _norm(raw.iat[r, 0]):
        r += 1
    data_start = r

    registros = []
    total_pendiente = None  # se sobreescribe con cada fila TOTAL; solo queda la última (el gran total real)
    for r in range(data_start, len(raw)):
        entidad_sbs = _norm(raw.iat[r, 0])
        if not entidad_sbs:
            continue
        if _es_fin(entidad_sbs):
            break

        valores = {}
        for producto, col in bloques:
            val = pd.to_numeric(raw.iat[r, col], errors="coerce")
            valores[producto] = 0.0 if pd.isna(val) else val

        if _es_total(entidad_sbs):
            total_pendiente = (entidad_sbs, valores)
            continue

        for producto, val in valores.items():
            registros.append({"Tipo": tipo, "Empresa": entidad_sbs, "Producto": producto,
                               "Nº de clientes": val, "es_total": False})

    if total_pendiente:
        entidad_sbs, valores = total_pendiente
        for producto, val in valores.items():
            registros.append({"Tipo": tipo, "Empresa": entidad_sbs, "Producto": producto,
                               "Nº de clientes": val, "es_total": True})

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
        lambda r: "-" if r["es_total"] else clasificar_sf_smf(r["Empresa Benchmark"]), axis=1
    )
    df[">50% CB"] = df.apply(
        lambda r: None if r["es_total"] else clasificar_50cb(r["Empresa Benchmark"]), axis=1
    )
    return df, sin_mapeo


# ============== pipeline ==============

def run(anio: int, mes_num: int) -> pd.DataFrame:
    """Descarga, procesa y exporta Clientes de Crédito para el corte (anio, mes_num)."""
    maestro = cargar_maestro()
    mes_abr = ABR_MES[mes_num]

    partes = []
    for cfg in CODIGOS_CORTE:
        print(f"[Clientes de Crédito] Descargando {cfg['codigo']} ({cfg['tipo']}) ...")
        contenido = descargar_reporte_bytes(cfg["codigo"], anio, mes_num)
        df_archivo = _leer_archivo(contenido, cfg["tipo"])
        print(f"  -> {len(df_archivo)} filas extraídas")
        partes.append(df_archivo)

    df = pd.concat(partes, ignore_index=True)
    df, sin_mapeo = _normalizar_entidades(df, maestro)

    if sin_mapeo:
        print("\n[Clientes de Crédito][AVISO] Entidades sin mapeo (excluidas):")
        for e in sin_mapeo:
            print(f"  - {e}")
        print("Agrégalas a maestro_entidades.csv en GitHub y vuelve a correr si quieres incluirlas.\n")

    df["Fecha"] = fin_de_mes(anio, mes_num)

    resultado = df[COLUMNAS_FINALES]

    output_path = BASE_DIR / f"ClientesCredito_SBS_Consolidado_{mes_abr}{anio}.xlsx"
    resultado.to_excel(output_path, index=False)
    print(f"[Clientes de Crédito] Listo. {len(resultado)} filas exportadas a: {output_path}")

    return resultado


if __name__ == "__main__":
    from datetime import datetime
    ahora = datetime.now()
    run(ahora.year, ahora.month)
