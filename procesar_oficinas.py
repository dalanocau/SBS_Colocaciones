"""
procesar_oficinas.py
Base 13/17 — Oficinas por Zona Geográfica SBS multientidad.

Reportes: B-2303 (Bancos), B-3201 (Financieras), C-1201 (CMAC),
          C-2201 (CRAC), C-4205 (EDPYME)

Lo específico de esta base: formato largo (una fila por Empresa+
Departamento), sin filas TOTAL por sector, sin columna Total agregada.
Nombres de departamento normalizados (con/sin tilde según familia).
"Sucursales en el Exterior" tratado como un departamento más (solo aplica
a Bancos). Nota: el fuzzy matching de utils_sbs ya incluye
processor=str.lower por defecto (el fix descubierto justo en esta base
originalmente), así que no requiere nada especial aquí.
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
    {"tipo": "Bancos", "codigo": "B-2303"},
    {"tipo": "Financieras", "codigo": "B-3201"},
    {"tipo": "CMAC", "codigo": "C-1201"},
    {"tipo": "CRAC", "codigo": "C-2201"},
    {"tipo": "Edpyme", "codigo": "C-4205"},
]

UMBRAL_FUZZY = 90

ETIQUETAS_TOTAL_EXACTAS = {"CAJAS MUNICIPALES", "CAJAS RURALES DE AHORRO Y CRÉDITO",
                            "EMPRESAS DE CRÉDITOS", "EMPRESAS FINANCIERAS", "BANCA MÚLTIPLE"}

CANON_DEPTO = {
    "amazonas": "Amazonas", "ancash": "Ancash", "apurimac": "Apurímac", "apurímac": "Apurímac",
    "arequipa": "Arequipa", "ayacucho": "Ayacucho", "cajamarca": "Cajamarca", "callao": "Callao",
    "cusco": "Cusco", "huancavelica": "Huancavelica", "huanuco": "Huánuco", "huánuco": "Huánuco",
    "ica": "Ica", "junin": "Junín", "junín": "Junín", "la libertad": "La Libertad", "lambayeque": "Lambayeque",
    "lima": "Lima", "loreto": "Loreto", "madre de dios": "Madre de Dios", "moquegua": "Moquegua",
    "pasco": "Pasco", "piura": "Piura", "puno": "Puno", "san martin": "San Martín", "san martín": "San Martín",
    "tacna": "Tacna", "tumbes": "Tumbes", "ucayali": "Ucayali",
    "sucursales en el exterior": "Sucursales en el Exterior",
}

COLUMNAS_FINALES = ["Fecha", "Empresa", "Empresa_Benchmark", "Nº de Agencias",
                     "Departamento", "Tipo", "Clasificación"]


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


def _canon_depto(nombre: str):
    return CANON_DEPTO.get(_norm(nombre).lower())


# ============== lectura del archivo ==============

def _find_header_row(raw: pd.DataFrame):
    for r in range(min(10, len(raw))):
        for c in range(raw.shape[1]):
            if _norm(raw.iat[r, c]).lower() == "empresas":
                return r
    return None


def _leer_archivo(fuente, tipo: str) -> pd.DataFrame:
    raw = pd.read_excel(fuente, header=None)
    header_row = _find_header_row(raw)
    if header_row is None:
        raise ValueError("No se encontró la fila de encabezado (Empresas).")

    col_map = {}
    for c in range(1, raw.shape[1]):
        depto = _canon_depto(raw.iat[header_row, c])
        if depto:
            col_map[c] = depto

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
            continue
        for c, depto in col_map.items():
            val = pd.to_numeric(raw.iat[r, c], errors="coerce")
            val = 0.0 if pd.isna(val) else val
            registros.append({"Tipo": tipo, "Empresa": entidad_sbs, "Departamento": depto, "Nº de Agencias": val})

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
    df["Clasificación"] = df["Empresa_Benchmark"].apply(clasificar_50cb)
    return df, sin_mapeo


# ============== pipeline ==============

def run(anio: int, mes_num: int) -> pd.DataFrame:
    """Descarga, procesa y exporta Oficinas por Zona Geográfica para el corte (anio, mes_num)."""
    maestro = cargar_maestro()
    mes_abr = ABR_MES[mes_num]

    partes = []
    for cfg in CODIGOS_CORTE:
        print(f"[Oficinas] Descargando {cfg['codigo']} ({cfg['tipo']}) ...")
        contenido = descargar_reporte_bytes(cfg["codigo"], anio, mes_num)
        df_archivo = _leer_archivo(contenido, cfg["tipo"])
        print(f"  -> {len(df_archivo)} filas extraídas")
        partes.append(df_archivo)

    df = pd.concat(partes, ignore_index=True)
    df, sin_mapeo = _normalizar_entidades(df, maestro)

    if sin_mapeo:
        print("\n[Oficinas][AVISO] Entidades sin mapeo (excluidas):")
        for e in sin_mapeo:
            print(f"  - {e}")
        print("Agrégalas a maestro_entidades.csv en GitHub y vuelve a correr si quieres incluirlas.\n")

    df["Fecha"] = fin_de_mes(anio, mes_num)

    resultado = df[COLUMNAS_FINALES]

    output_path = BASE_DIR / f"Oficinas_SBS_Consolidado_{mes_abr}{anio}.xlsx"
    resultado.to_excel(output_path, index=False)
    print(f"[Oficinas] Listo. {len(resultado)} filas exportadas a: {output_path}")

    return resultado


if __name__ == "__main__":
    from datetime import datetime
    ahora = datetime.now()
    run(ahora.year, ahora.month)
