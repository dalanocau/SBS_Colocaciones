"""
procesar_clientes_ahorro.py
Base 6/17 — Clientes de Ahorro SBS multientidad.

Reportes: B-2373 (Bancos), B-3232 (Financieras), C-1250 (CMACs), C-2255 (CRACs)
Nota: EC (Edpymes) no incluida — no autorizada a captar ahorros del público
(mismo motivo que en Depósitos).

Lo específico de esta base: misma estructura de 3 columnas por producto que
Depósitos (Personas Naturales / Pers. Jur. sin fines de lucro / Otras Pers.
Jur.), combinada con el manejo de "(X) Total" y fila TOTAL por sector como
entidad más (igual que Clientes de Crédito). Fix propio: filas basura tipo
"0" deben saltarse con continue, no cortar la lectura con break, porque a
veces aparecen ANTES de la fila TOTAL real que sí se quiere incluir.
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
    {"tipo": "Bancos", "codigo": "B-2373"},
    {"tipo": "Financieras", "codigo": "B-3232"},
    {"tipo": "CMACs", "codigo": "C-1250"},
    {"tipo": "CRACs", "codigo": "C-2255"},
]

UMBRAL_FUZZY = 90

CANON_PRODUCTO_DEP = {
    "depósitos a la vista": "A la vista",
    "depósitos de ahorros": "Ahorro",
    "depósitos de ahorro": "Ahorro",
    "depósitos a plazo": "A plazo",
    "depósitos cts": "CTS",
}

COLUMNAS_FINALES = [
    "Fecha", "Tipo", "Clasificación", "Empresa", "Empresa Benchmark",
    "Personas Naturales", "Personas Jurídicas sin fines de lucro",
    "Otras Personas Jurídicas", "Total", "Producto", ">50% mype",
]


# ============== utilidades locales ==============

def _norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).strip())


def _es_basura(v) -> bool:
    v = _norm(v)
    return v != "" and v.replace(".", "").isdigit()


def _es_total(v) -> bool:
    return _norm(v).upper().startswith("TOTAL")


def _es_fin(v) -> bool:
    v = _norm(v)
    if not v:
        return False
    return (v.lower().startswith("nota") or v.lower().startswith("fuente")
            or v.startswith("*") or v.startswith("http") or len(v) > 80)


# ============== lectura del archivo ==============

def _find_categoria_row(raw: pd.DataFrame):
    for r in range(min(12, len(raw))):
        for c in range(raw.shape[1]):
            if "ahorro" in _norm(raw.iat[r, c]).lower():
                return r
    return None


def _leer_archivo(fuente, tipo: str) -> pd.DataFrame:
    raw = pd.read_excel(fuente, header=None)
    cat_row = _find_categoria_row(raw)
    if cat_row is None:
        raise ValueError("No se encontró la fila de categorías (Depósitos de Ahorro/s).")

    bloques_normales, bloque_total = [], None
    for c in range(raw.shape[1]):
        v = _norm(raw.iat[cat_row, c]).lower()
        if v in CANON_PRODUCTO_DEP:
            bloques_normales.append((CANON_PRODUCTO_DEP[v], c, c + 1, c + 2))
        elif "depósitos totales" in v:
            bloque_total = ("(X) Total", c, c + 1, c + 2)
    todos_bloques = bloques_normales + ([bloque_total] if bloque_total else [])

    r = cat_row + 1
    while r < len(raw) and (not _norm(raw.iat[r, 0]) or "persona" in _norm(raw.iat[r, 1]).lower()):
        r += 1
    data_start = r

    registros = []
    total_pendiente = None
    for r in range(data_start, len(raw)):
        entidad_sbs = _norm(raw.iat[r, 0])
        if not entidad_sbs or _es_basura(entidad_sbs):
            continue
        if _es_fin(entidad_sbs):
            break

        valores = {}
        for producto, c1, c2, c3 in todos_bloques:
            nat = pd.to_numeric(raw.iat[r, c1], errors="coerce"); nat = 0.0 if pd.isna(nat) else nat
            jur = pd.to_numeric(raw.iat[r, c2], errors="coerce"); jur = 0.0 if pd.isna(jur) else jur
            otras = pd.to_numeric(raw.iat[r, c3], errors="coerce"); otras = 0.0 if pd.isna(otras) else otras
            valores[producto] = (nat, jur, otras)

        if _es_total(entidad_sbs):
            total_pendiente = (entidad_sbs, valores)  # se sobreescribe; solo queda el último (el total real)
            continue

        for producto, (nat, jur, otras) in valores.items():
            registros.append({"Tipo": tipo, "Empresa": entidad_sbs, "Producto": producto,
                               "Personas Naturales": nat, "Personas Jurídicas sin fines de lucro": jur,
                               "Otras Personas Jurídicas": otras, "es_total": False})

    if total_pendiente:
        entidad_sbs, valores = total_pendiente
        for producto, (nat, jur, otras) in valores.items():
            registros.append({"Tipo": tipo, "Empresa": entidad_sbs, "Producto": producto,
                               "Personas Naturales": nat, "Personas Jurídicas sin fines de lucro": jur,
                               "Otras Personas Jurídicas": otras, "es_total": True})

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
    df["Total"] = df["Personas Naturales"] + df["Personas Jurídicas sin fines de lucro"] + df["Otras Personas Jurídicas"]
    df["Clasificación"] = df.apply(
        lambda r: "-" if r["es_total"] else clasificar_sf_smf(r["Empresa Benchmark"]), axis=1
    )
    df[">50% mype"] = df.apply(
        lambda r: None if r["es_total"] else clasificar_50cb(r["Empresa Benchmark"]), axis=1
    )
    return df, sin_mapeo


# ============== pipeline ==============

def run(anio: int, mes_num: int) -> pd.DataFrame:
    """Descarga, procesa y exporta Clientes de Ahorro para el corte (anio, mes_num)."""
    maestro = cargar_maestro()
    mes_abr = ABR_MES[mes_num]

    partes = []
    for cfg in CODIGOS_CORTE:
        print(f"[Clientes de Ahorro] Descargando {cfg['codigo']} ({cfg['tipo']}) ...")
        contenido = descargar_reporte_bytes(cfg["codigo"], anio, mes_num)
        df_archivo = _leer_archivo(contenido, cfg["tipo"])
        print(f"  -> {len(df_archivo)} filas extraídas")
        partes.append(df_archivo)

    df = pd.concat(partes, ignore_index=True)
    df, sin_mapeo = _normalizar_entidades(df, maestro)

    if sin_mapeo:
        print("\n[Clientes de Ahorro][AVISO] Entidades sin mapeo (excluidas):")
        for e in sin_mapeo:
            print(f"  - {e}")
        print("Agrégalas a maestro_entidades.csv en GitHub y vuelve a correr si quieres incluirlas.\n")

    df["Fecha"] = fin_de_mes(anio, mes_num)

    resultado = df[COLUMNAS_FINALES]

    output_path = BASE_DIR / f"ClientesAhorro_SBS_Consolidado_{mes_abr}{anio}.xlsx"
    resultado.to_excel(output_path, index=False)
    print(f"[Clientes de Ahorro] Listo. {len(resultado)} filas exportadas a: {output_path}")

    return resultado


if __name__ == "__main__":
    from datetime import datetime
    ahora = datetime.now()
    run(ahora.year, ahora.month)
