"""
procesar_rcg.py
Base 9/17 — Ratio de Capital Global (RCG) SBS multientidad.

Reportes RCG: B-2402 (Bancos), B-3302 (Financieras), C-1252 (CMAC),
              C-2257 (CRAC), C-4241 (EDPYME)

Lo específico de esta base: la columna "Patrimonio" se trae de la base
Patrimonio Efectivo (mismo corte), cruzada por Empresa Benchmark -- este
pipeline descarga y procesa también los 5 archivos de Patrimonio Efectivo
internamente para hacer ese join. El cruce es TOLERANTE A FALLOS: si algún
archivo de Patrimonio Efectivo no está disponible (no salió aún para ese
corte), se salta ese archivo y esas entidades quedan con Patrimonio vacío
(NaN), sin afectar el resto del RCG -- a diferencia de la descarga RCG
misma, que si falla sí detiene todo el proceso.
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

CODIGOS_RCG = [
    {"tipo": "BANCOS", "codigo": "B-2402"},
    {"tipo": "FINANCIERAS", "codigo": "B-3302"},
    {"tipo": "CMAC", "codigo": "C-1252"},
    {"tipo": "CRAC", "codigo": "C-2257"},
    {"tipo": "EDPYME", "codigo": "C-4241"},
]

# Mismos 5 archivos de Patrimonio Efectivo, necesarios para cruzar la columna
# "Patrimonio". Este cruce es tolerante a fallos: si alguno no está disponible
# para el corte, se salta y las entidades correspondientes quedan con
# Patrimonio vacío, sin afectar el resto del RCG.
CODIGOS_PE = [
    {"tipo": "BANCOS", "codigo": "B-2370"},
    {"tipo": "FINANCIERAS", "codigo": "B-3252"},
    {"tipo": "CMAC", "codigo": "C-1257"},
    {"tipo": "CRAC", "codigo": "C-2262"},
    {"tipo": "EDPYMES", "codigo": "C-4246"},
]

UMBRAL_FUZZY = 90

COLUMNAS_FINALES = [
    "PERIODO", "Tipo", "Empresa", "Empresa Benchmark",
    "Creditos", "Mercado", "Operacional", "Total",
    "Creditos_A", "Mercado_A", "Operacional_A", "Total_A",
    "Patrimonio", "CORE_CAPITAL", "RCO", "RCG", "Clasificación",
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


def _find_header_row(raw: pd.DataFrame, etiquetas: set):
    for r in range(min(15, len(raw))):
        for c in range(raw.shape[1]):
            if _norm(raw.iat[r, c]).upper() in etiquetas:
                return r
    return None


# ============== lectura de RCG ==============

def _leer_archivo_rcg(fuente, tipo: str) -> pd.DataFrame:
    raw = pd.read_excel(fuente, header=None)
    header_row = _find_header_row(raw, {"EMPRESAS", "ENTIDAD"})
    if header_row is None:
        raise ValueError("No se encontró la fila de encabezado (EMPRESAS/ENTIDAD).")

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

        def num(c):
            v = pd.to_numeric(raw.iat[r, c], errors="coerce")
            return 0.0 if pd.isna(v) else v

        a, b, c_ = num(1), num(2), num(3)              # Requerimiento PE: crédito, mercado, operacional
        d, e, f = num(5), num(6), num(7)                # APR: crédito, mercado, operacional
        core, n1apr, rcg = num(10), num(11), num(12)    # Ratios %

        registros.append({
            "Tipo": tipo, "Empresa": entidad_sbs,
            "Creditos": a, "Mercado": b, "Operacional": c_,
            "Creditos_A": d, "Mercado_A": e, "Operacional_A": f,
            "CORE_CAPITAL": core, "RCO": n1apr, "RCG": rcg,
            "es_total": _es_total(entidad_sbs),
        })

    return pd.DataFrame(registros)


# ============== lectura de Patrimonio Efectivo (solo para el cruce) ==============

def _leer_archivo_patrimonio(fuente, tipo: str) -> pd.DataFrame:
    raw = pd.read_excel(fuente, header=None)
    header_row = _find_header_row(raw, {"ENTIDAD"})
    if header_row is None:
        raise ValueError("No se encontró la fila de encabezado (ENTIDAD) en Patrimonio Efectivo.")

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
        total = pd.to_numeric(raw.iat[r, 4], errors="coerce")
        total = 0.0 if pd.isna(total) else total
        registros.append({"Tipo": tipo, "Empresa": entidad_sbs, "PatrimonioTotal": total,
                           "es_total": _es_total(entidad_sbs)})

    return pd.DataFrame(registros)


# ============== normalización contra el maestro (fuzzy, con cache compartido) ==============

def _resolver_entidades(df: pd.DataFrame, maestro: pd.DataFrame, cache: dict, sin_mapeo: set):
    """Resuelve Empresa Benchmark reutilizando un cache compartido entre RCG
    y Patrimonio Efectivo, para no repetir el fuzzy match del mismo nombre."""
    def resolver(nombre_crudo):
        if nombre_crudo in cache:
            return cache[nombre_crudo]
        fila = fuzzy_match_entidad(nombre_crudo, maestro, umbral=UMBRAL_FUZZY)
        resultado = fila["nombre_bd"] if fila is not None else None
        cache[nombre_crudo] = resultado
        if resultado is None:
            sin_mapeo.add(nombre_crudo)
        return resultado

    df = df.copy()
    df["Empresa Benchmark"] = df["Empresa"].apply(resolver)
    return df


# ============== pipeline ==============

def run(anio: int, mes_num: int) -> pd.DataFrame:
    """Descarga, procesa y exporta RCG para el corte (anio, mes_num), cruzando Patrimonio Efectivo."""
    maestro = cargar_maestro()
    mes_abr = ABR_MES[mes_num]
    cache = {}
    sin_mapeo = set()

    # --- RCG: base principal, si falla algo aquí sí se detiene todo el proceso ---
    partes_rcg = []
    for cfg in CODIGOS_RCG:
        print(f"[RCG] Descargando {cfg['codigo']} ({cfg['tipo']}) ...")
        contenido = descargar_reporte_bytes(cfg["codigo"], anio, mes_num)
        df_archivo = _leer_archivo_rcg(contenido, cfg["tipo"])
        print(f"  -> {len(df_archivo)} filas extraídas")
        partes_rcg.append(df_archivo)
    df_rcg = pd.concat(partes_rcg, ignore_index=True)

    # --- Patrimonio Efectivo: solo para el cruce de la columna "Patrimonio".
    # Si un archivo falla (no salió aún para este corte, o solo salió alguno
    # de los 5), se salta ese archivo y se continúa -- no se detiene el RCG. ---
    partes_pe = []
    for cfg in CODIGOS_PE:
        print(f"[RCG] Descargando Patrimonio Efectivo {cfg['codigo']} ({cfg['tipo']}) para el cruce...")
        try:
            contenido = descargar_reporte_bytes(cfg["codigo"], anio, mes_num)
            df_archivo = _leer_archivo_patrimonio(contenido, cfg["tipo"])
            print(f"  -> {len(df_archivo)} filas extraídas")
            partes_pe.append(df_archivo)
        except Exception as e:
            print(f"  [RCG][AVISO] No se pudo obtener Patrimonio Efectivo de {cfg['tipo']} ({cfg['codigo']}): {e}")
            print(f"  -> Las entidades de {cfg['tipo']} quedarán con Patrimonio vacío en el resultado final.\n")

    df_pe = pd.concat(partes_pe, ignore_index=True) if partes_pe else pd.DataFrame(columns=["Empresa", "PatrimonioTotal"])

    df_rcg = _resolver_entidades(df_rcg, maestro, cache, sin_mapeo)
    if len(df_pe):
        df_pe = _resolver_entidades(df_pe, maestro, cache, sin_mapeo)

    if sin_mapeo:
        print("\n[RCG][AVISO] Entidades sin mapeo (excluidas):")
        for e in sorted(sin_mapeo):
            print(f"  - {e}")
        print("Agrégalas a maestro_entidades.csv en GitHub y vuelve a correr si quieres incluirlas.\n")
        df_rcg = df_rcg[~df_rcg["Empresa"].isin(sin_mapeo)].copy()
        if len(df_pe):
            df_pe = df_pe[~df_pe["Empresa"].isin(sin_mapeo)].copy()

    pe_map = df_pe.set_index("Empresa Benchmark")["PatrimonioTotal"].to_dict() if len(df_pe) else {}
    df_rcg["Patrimonio"] = df_rcg["Empresa Benchmark"].map(pe_map)

    sin_patrimonio = df_rcg[df_rcg["Patrimonio"].isna()]
    if len(sin_patrimonio):
        print(f"[RCG][AVISO] {len(sin_patrimonio)} filas quedaron sin Patrimonio Efectivo cruzado (columna vacía):")
        for tipo, grupo in sin_patrimonio.groupby("Tipo"):
            print(f"  - {tipo}: {grupo['Empresa Benchmark'].nunique()} entidades")

    df_rcg["Total"] = df_rcg["Creditos"] + df_rcg["Mercado"] + df_rcg["Operacional"]
    df_rcg["Total_A"] = df_rcg["Creditos_A"] + df_rcg["Mercado_A"] + df_rcg["Operacional_A"]
    df_rcg["Clasificación"] = df_rcg.apply(
        lambda r: "-" if r["es_total"] else clasificar_50cb(r["Empresa Benchmark"]), axis=1
    )
    df_rcg["PERIODO"] = fin_de_mes(anio, mes_num)

    resultado = df_rcg[COLUMNAS_FINALES]

    output_path = BASE_DIR / f"RCG_SBS_Consolidado_{mes_abr}{anio}.xlsx"
    resultado.to_excel(output_path, index=False)
    print(f"[RCG] Listo. {len(resultado)} filas exportadas a: {output_path}")

    return resultado


if __name__ == "__main__":
    from datetime import datetime
    ahora = datetime.now()
    run(ahora.year, ahora.month)
