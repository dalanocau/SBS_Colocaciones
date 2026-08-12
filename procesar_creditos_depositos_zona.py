"""
procesar_creditos_depositos_zona.py
Base 14/17 — Créditos y Depósitos por Zona Geográfica SBS multientidad.

Reportes: B-2358 (Bancos), B-3241 (Financieras), C-1234 (CMACs),
          C-2234 (CRACs), C-4228 (EDPYMEs)

Lo específico de esta base: reporte a nivel de OFICINA (no de entidad),
requiere forward-fill de Empresa/Departamento/Provincia/Distrito (solo
aparecen en la primera fila de cada bloque). Bloques de producto varían
por familia: Bancos/Financieras tienen A la vista+Ahorro+A plazo+Créditos;
CMAC/CRAC solo Ahorro+A plazo+Créditos (sin A la vista); EDPYME solo
Créditos. La columna "Región CAQP" queda vacía a propósito (no se calcula).
Fila "Total general" al final de cada archivo se excluye.
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
    {"tipo": "Bancos", "codigo": "B-2358", "familia": "con_vista"},
    {"tipo": "Financieras", "codigo": "B-3241", "familia": "con_vista"},
    {"tipo": "CMACs", "codigo": "C-1234", "familia": "sin_vista"},
    {"tipo": "CRACs", "codigo": "C-2234", "familia": "sin_vista"},
    {"tipo": "Edpymes", "codigo": "C-4228", "familia": "solo_creditos"},
]

UMBRAL_FUZZY = 90

BLOQUES_POR_FAMILIA = {
    "con_vista": [("A la vista", 5, 6, 7), ("Ahorro", 8, 9, 10), ("A plazo", 11, 12, 13), ("Créditos", 15, 16, 17)],
    "sin_vista": [("Ahorro", 5, 6, 7), ("A plazo", 8, 9, 10), ("Créditos", 12, 13, 14)],
    "solo_creditos": [("Créditos", 5, 6, 7)],
}

COLUMNAS_FINALES = ["Fecha", "Empresa", "Empresa_Benchmark", "Tipo", "Clasificación",
                     "Departamento", "Provincia", "Distrito", "Código de oficina",
                     "MN", "ME", "Total", "Producto", "Región CAQP", ">50% CB"]


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


def _es_total_general(v) -> bool:
    return _norm(v).lower() == "total general"


# ============== lectura del archivo (con forward-fill) ==============

def _find_header_row(raw: pd.DataFrame):
    for r in range(min(10, len(raw))):
        if _norm(raw.iat[r, 0]).lower() == "empresa":
            return r
    return None


def _leer_archivo(fuente, tipo: str, familia: str) -> pd.DataFrame:
    raw = pd.read_excel(fuente, header=None)
    header_row = _find_header_row(raw)
    if header_row is None:
        raise ValueError("No se encontró la fila de encabezado (Empresa).")
    sub_row = header_row + 1

    r = sub_row + 1
    while r < len(raw) and _norm(raw.iat[r, 4]) == "" and _norm(raw.iat[r, 0]) == "":
        r += 1
    data_start = r

    fin = len(raw)
    for r in range(data_start, len(raw)):
        col0 = _norm(raw.iat[r, 0])
        if _es_total_general(col0):
            fin = r
            break
        if col0 == "" and _norm(raw.iat[r, 4]) == "":
            resto_vacio = all(_norm(raw.iat[r, c]) == "" for c in range(1, min(raw.shape[1], 8)))
            if resto_vacio:
                fin = r
                break
        if _es_fin(col0):
            fin = r
            break

    # Forward-fill de Empresa/Departamento/Provincia/Distrito -- equivalente a
    # Ir a Especial > Blancos + "=celda superior" + Ctrl+Enter en Excel.
    bloque = raw.iloc[data_start:fin, [0, 1, 2, 3]].copy()
    bloque.columns = ["Empresa", "Departamento", "Provincia", "Distrito"]
    bloque = bloque.replace("", pd.NA)
    bloque = bloque.ffill()

    bloques = BLOQUES_POR_FAMILIA[familia]
    registros = []
    for i, r in enumerate(range(data_start, fin)):
        empresa = _norm(bloque.iloc[i]["Empresa"])
        depto = _norm(bloque.iloc[i]["Departamento"])
        prov = _norm(bloque.iloc[i]["Provincia"])
        dist = _norm(bloque.iloc[i]["Distrito"])
        codigo = _norm(raw.iat[r, 4])
        if not empresa:
            continue
        for producto, c_mn, c_me, c_tot in bloques:
            mn = pd.to_numeric(raw.iat[r, c_mn], errors="coerce"); mn = 0.0 if pd.isna(mn) else mn
            me = pd.to_numeric(raw.iat[r, c_me], errors="coerce"); me = 0.0 if pd.isna(me) else me
            tot = pd.to_numeric(raw.iat[r, c_tot], errors="coerce"); tot = 0.0 if pd.isna(tot) else tot
            registros.append({
                "Tipo": tipo, "Empresa": empresa, "Departamento": depto, "Provincia": prov,
                "Distrito": dist, "Código de oficina": codigo, "Producto": producto,
                "MN": mn, "ME": me, "Total": tot,
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

    df["Empresa_Benchmark"] = df["Empresa"].map(mapa)
    df["Clasificación"] = df["Empresa_Benchmark"].apply(clasificar_sf_smf)
    df[">50% CB"] = df["Empresa_Benchmark"].apply(clasificar_50cb)
    return df, sin_mapeo


# ============== pipeline ==============

def run(anio: int, mes_num: int) -> pd.DataFrame:
    """Descarga, procesa y exporta Créditos y Depósitos por Zona Geográfica para el corte (anio, mes_num)."""
    maestro = cargar_maestro()
    mes_abr = ABR_MES[mes_num]

    partes = []
    for cfg in CODIGOS_CORTE:
        print(f"[Créditos/Depósitos Zona] Descargando {cfg['codigo']} ({cfg['tipo']}) ...")
        contenido = descargar_reporte_bytes(cfg["codigo"], anio, mes_num)
        df_archivo = _leer_archivo(contenido, cfg["tipo"], cfg["familia"])
        print(f"  -> {len(df_archivo)} filas extraídas")
        partes.append(df_archivo)

    df = pd.concat(partes, ignore_index=True)
    df, sin_mapeo = _normalizar_entidades(df, maestro)

    if sin_mapeo:
        print("\n[Créditos/Depósitos Zona][AVISO] Entidades sin mapeo (excluidas):")
        for e in sin_mapeo:
            print(f"  - {e}")
        print("Agrégalas a maestro_entidades.csv en GitHub y vuelve a correr si quieres incluirlas.\n")

    df["Fecha"] = fin_de_mes(anio, mes_num)
    df["Región CAQP"] = None  # queda vacía, no se calcula

    resultado = df[COLUMNAS_FINALES]

    output_path = BASE_DIR / f"CredDepZonaGeo_SBS_Consolidado_{mes_abr}{anio}.xlsx"
    resultado.to_excel(output_path, index=False)
    print(f"[Créditos/Depósitos Zona] Listo. {len(resultado)} filas exportadas a: {output_path}")

    return resultado


if __name__ == "__main__":
    from datetime import datetime
    ahora = datetime.now()
    run(ahora.year, ahora.month)
