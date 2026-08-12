"""
procesar_eeff.py
Base 17/17 (última) — EEFF (Estados Financieros: BG + ER) SBS multientidad.

Reportes: B-2201 (Bancos), B-3101 (Financieras), C-1101 (CMAC),
          C-2101 (CRAC), C-4103 (EDPYME)

La base más compleja del proyecto: cada archivo trae 2 hojas (BG=Balance
General, ER=Estado de Ganancias y Pérdidas), con nombre de hoja distinto
por familia (Bancos/Financieras usan "1"/"2"; CMAC/CRAC/EDPYME usan
nombres propios). Formato "doble transpuesto": los bloques de columnas
repetidos son solo artefacto de impresión, cada entidad ya tiene sus
propias 3 columnas MN/ME/TOTAL en la fila completa. El fuzzy matching de
cuentas contables (contra las listas canónicas CANON_ER/CANON_BG) es
independiente del fuzzy matching de entidades (contra el maestro) -- usa
un umbral propio (88) y matchea contra las listas canónicas, no contra el
maestro, así que se implementa localmente con rapidfuzz en vez de
utils_sbs.fuzzy_match_entidad. Se excluyen entidades TOTAL (no están en
la estructura EEFF, a diferencia de otras bases). El archivo BG se corta
al llegar a la sección de CUENTAS CONTINGENTES (fuera de balance, no está
en la estructura canónica).
Todo lo compartido para descarga/maestro/entidades viene de utils_sbs.
"""

import re
from io import BytesIO

import pandas as pd
from rapidfuzz import fuzz, process

from utils_sbs import (
    BASE_DIR,
    ABR_MES,
    cargar_maestro,
    clasificar_50cb,
    descargar_reporte_bytes,
    fin_de_mes,
    fuzzy_match_entidad,
)

UMBRAL_FUZZY = 90              # para nombres de entidad
UMBRAL_FUZZY_CUENTA = 88       # para nombres de cuenta contable (BG/ER)

# Hoja "1"/"2" para Bancos y Financieras; nombres de hoja propios para CMAC/CRAC/EDPYME
CODIGOS_CORTE = [
    {"tipo": "Bancos", "codigo": "B-2201", "hoja_bg": "1", "hoja_er": "2"},
    {"tipo": "Financieras", "codigo": "B-3101", "hoja_bg": "1", "hoja_er": "2"},
    {"tipo": "CMAC", "codigo": "C-1101", "hoja_bg": "bg_cm", "hoja_er": "gyp_cm"},
    {"tipo": "CRAC", "codigo": "C-2101", "hoja_bg": "bg_cr", "hoja_er": "gyp_cr"},
    {"tipo": "EDPYME", "codigo": "C-4103", "hoja_bg": "bg_edp", "hoja_er": "gyp_edp"},
]

MANUAL_ER = {
    "PROVISIONES PARA CRÉDITOS DIRECTOS": "(4) PROVISIONES PARA INCOBRABILIDAD DE CRÉDITOS",
    "UTILIDAD (PÉRDIDA) POR VENTA DE CARTERA": "(8) GANANCIA (PÉRDIDA) POR VENTA DE CARTERA",
    "UTILIDAD (PÉRDIDA) POR VENTA DE CARTERA CREDITICIA": "(8) GANANCIA (PÉRDIDA) POR VENTA DE CARTERA",
    "RESULTADO ANTES DE IMPUESTO A LA RENTA": "(14) UTILIDAD (PÉRDIDA) ANTES DE PARTICIPACIONES E  IMPUESTO A LA RENTA",
    "RESULTADO ANTES DEL IMPUESTO A LA RENTA": "(14) UTILIDAD (PÉRDIDA) ANTES DE PARTICIPACIONES E  IMPUESTO A LA RENTA",
    "RESULTADO NETO DEL EJERCICIO": "(17) UTILIDAD (PÉRDIDA) NETA",
}
MANUAL_BG = {
    "INSTITUCIONES DEL PAÍS": "(B4.1) Instituciones Financieras del País",
    "INSTITUCIONES DEL EXTERIOR Y ORGANISMOS INTERNACIONALES": "(B4.2) Empresas del Exterior y Organismos Internacionales",
}
DETIENE_BG_EN = "CONTINGENTES"  # sección fuera de balance, no está en la estructura canónica

CANON_ER = [
    '(1) INGRESOS FINANCIEROS',
    '     (1.1) Disponibles',
    '     (1.2) Fondos Interbancarios',
    '     (1.3) Inversiones',
    '     (1.4) Créditos Directos',
    '     (1.5) Ganancias por Valorización de Inversiones ',
    '     (1.6) Ganancias por Inversiones en Subsidiarias, Asociadas y Negocios Conjuntos',
    '     (1.7) Diferencia de Cambio',
    '     (1.8) Ganancias en Productos Financieros Derivados',
    '     (1.9) Reajuste por Indexación',
    '     (1.10) Otros',
    '(2) GASTOS FINANCIEROS',
    '     (2.1) Obligaciones con el Público',
    '     (2.2) Depósitos del Sistema Financiero y Organismos Financieros Internacionales',
    '     (2.3) Fondos Interbancarios',
    '     (2.4) Adeudos y Obligaciones Financieras',
    '     (2.5) Obligaciones en Circulación no Subordinadas',
    '     (2.6) Obligaciones en Circulación Subordinadas',
    '     (2.7) Pérdida por Valorización de Inversiones',
    '     (2.8) Pérdidas por Inversiones en Subsidiarias, Asociadas y Negocios Conjuntos',
    '     (2.9) Primas al Fondo de Seguro de Depósitos',
    '     (2.10) Diferencia de Cambio',
    '     (2.11) Pérdidas en Productos Financieros Derivados',
    '     (2.12) Reajuste por Indexación',
    '     (2.13) Otros',
    '(3) MARGEN FINANCIERO BRUTO',
    '(4) PROVISIONES PARA INCOBRABILIDAD DE CRÉDITOS',
    '     (4.1) Provisiones para Desvalorización de Inversiones',
    '     (4.2) Provisiones para Incobrabilidad de Créditos',
    '(5) MARGEN FINANCIERO NETO',
    '(6) INGRESOS POR SERVICIOS FINANCIEROS',
    '     (6.1) Cuentas por Cobrar',
    '     (6.2) Créditos Indirectos',
    '     (6.3) Fideicomisos y Comisiones de Confianza',
    '     (6.4) Ingresos Diversos',
    '(7) GASTOS POR SERVICIOS FINANCIEROS',
    '     (7.1) Cuentas por Pagar',
    '     (7.2) Créditos Indirectos',
    '     (7.3) Fideicomisos y Comisiones de Confianza',
    '     (7.4) Gastos Diversos',
    '(8) GANANCIA (PÉRDIDA) POR VENTA DE CARTERA',
    '(9) MARGEN OPERACIONAL',
    '(10) GASTOS ADMINISTRATIVOS',
    '     (10.1) Personal',
    '     (10.2) Directorio',
    '     (10.3) Servicios Recibidos de Terceros',
    '     (10.4) Impuestos y Contribuciones',
    '(11) MARGEN OPERACIONAL NETO',
    '(12) PROVISIONES, DEPRECIACIÓN Y AMORTIZACIÓN',
    '     (12.1) Provisiones para Contingencias y Otras',
    '     (12.2) Provisiones para Créditos Indirectos',
    '     (12.3) Provisiones por Pérdida por Deterioro de Inversiones',
    '     (12.4) Provisiones para Incobrabilidad de Cuentas por Cobrar',
    '     (12.5) Provisiones para Bienes Realizables, Recibidos en Pago y Adjudicados',
    '     (12.6) Otras Provisiones',
    '     (12.7) Depreciación',
    '     (12.8) Amortización',
    '(13) OTROS INGRESOS Y GASTOS',
    '     (13.1) Ingresos (Gastos) por Recuperación de Créditos',
    '     (13.2) Ingresos (Gastos) Extraordinarios',
    '     (13.3) Ingresos (Gastos) de Ejercicios Anteriores',
    '(14) UTILIDAD (PÉRDIDA) ANTES DE PARTICIPACIONES E  IMPUESTO A LA RENTA',
    '(15) PARTICIPACIÓN DE TRABAJADORES',
    '(16) IMPUESTO A LA RENTA',
    '(17) UTILIDAD (PÉRDIDA) NETA',
]

CANON_BG = [
    '      (A1) DISPONIBLE',
    '             (A1.1) Caja',
    '             (A1.2) Bancos y Corresponsales',
    '             (A1.3) Canje',
    '             (A1.4) Otros',
    '      (A2) FONDOS INTERBANCARIOS',
    '      (A3) INVERSIONES NETAS DE PROVISIONES E INGRESOS NO DEVENGADOS',
    '             (A3.1) Negociables para Intermediación Financiera',
    '             (A3.2) Inversiones a Valor Razonable con Cambios en Resultados',
    '             (A3.3) Inversiones Disponibles para la Venta',
    '             (A3.4) Inversiones a Vencimiento',
    '             (A3.5) Permanentes',
    '             (A3.6) Inversiones en subsidiarias y asociadas',
    '             (A3.7) Inversiones en Commodities',
    '             (A3.8) Provisiones',
    '             (A3.9) Ingresos por Compraventa de Valores no Devengados',
    '      (A4) CRÉDITOS NETOS DE PROVISIONES E INGRESOS NO DEVENGADOS',
    '             (A4.1) Vigentes',
    '                      (A4.1.1) Cuentas Corrientes',
    '                      (A4.1.2) Tarjetas de Crédito',
    '                      (A4.1.3) Descuentos',
    '                      (A4.1.4) Factoring',
    '                      (A4.1.5) Préstamos',
    '                      (A4.1.6) Arrendamiento',
    '                      (A4.1.7) Hipotecarios',
    '                      (A4.1.8) Comercio Exterior',
    '                      (A4.1.9) Créditos por Liquidar',
    '                      (A4.1.10) Otros',
    '             (A4.2) Refinanciados y Reestructurados',
    '             (A4.3) Atrasados',
    '                      (A4.3.1) Vencidos',
    '                      (A4.3.2) En Cobranza Judicial',
    '             (A4.4) Provisiones',
    '             (A4.5) Intereses y Comisiones  no Devengados',
    '      (A5) CUENTAS POR COBRAR NETAS DE PROVISIONES',
    '      (A6) RENDIMIENTOS DEVENGADOS POR COBRAR',
    '             (A6.1) Disponible',
    '             (A6.2) Fondos Interbancarios',
    '             (A6.3) Inversiones',
    '             (A6.4) Créditos',
    '             (A6.5) Cuentas por Cobrar',
    '      (A7) BIENES REALIZABLES, RECIBIDOS EN PAGO, ADJUDICADOS Y FUERA DE USO NETOS',
    '      (A8) INMUEBLE, MOBILIARIO Y EQUIPO NETO',
    '      (A9) OTROS ACTIVOS',
    '(A) TOTAL ACTIVO',
    '      (B1) OBLIGACIONES CON EL PÚBLICO',
    '             (B1.1) Depósitos A La Vista',
    '             (B1.2) Depósitos de Ahorros',
    '             (B1.3) Depósitos a Plazo',
    '                      (B1.3.1) Certificados Bancarios y de Depósitos',
    '                      (B1.3.2) Cuentas a Plazo',
    '                      (B1.3.3) C.T.S.',
    '                      (B1.3.4) Otros',
    '             (B1.4) Depósitos Restringidos',
    '             (B1.5) Otras Obligaciones',
    '                      (B1.5.1) A la Vista',
    '                      (B1.5.2) Relacionadas con Inversiones',
    '      (B2) DEPÓSITOS DEL SISTEMA FINANCIERO Y ORGANISMOS INTERNACIONALES',
    '             (B2.1) Depósitos a la Vista',
    '             (B2.2) Depósitos de Ahorros',
    '             (B2.3) Depósitos a Plazo',
    '      (B3) FONDOS INTERBANCARIOS',
    '      (B4) ADEUDOS Y OBLIGACIONES FINANCIERAS',
    '             (B4.1) Instituciones Financieras del País',
    '             (B4.2) Empresas del Exterior y Organismos Internacionales',
    '      (B5) OBLIGACIONES EN CIRCULACIÓN NO SUBORDINADAS',
    '             (B5.1) Bonos de Arrendamiento Financiero',
    '             (B5.2) Instrumentos Hipotecarios',
    '             (B5.3) Otros Instrumentos de Deuda',
    '      (B6) CUENTAS POR PAGAR NETAS',
    '      (B7) INTERESES Y OTROS GASTOS DEVENGADOS POR PAGAR',
    '             (B7.1) Obligaciones con el Público',
    '             (B7.2) Depósitos del Sistema Financiero y Organismos Internacionales',
    '             (B7.3) Fondos Interbancarios',
    '             (B7.4) Adeudos y Obligaciones Financieras',
    '             (B7.5) Obligaciones en Circulación no Subordinadas',
    '             (B7.6) Cuentas por Pagar',
    '      (B8) OTROS PASIVOS',
    '      (B9) PROVISIONES',
    '             (B9.1) Créditos Indirectos',
    '             (B9.2) Otras Provisiones',
    '      (B10) OBLIGACIONES EN CIRCULACIÓN SUBORDINADAS 1/',
    '(B) TOTAL PASIVO',
    '(C) PATRIMONIO',
    '      (C1) Capital Social',
    '      (C2) Capital Adicional y Ajustes al Patrimonio',
    '      (C3) Capital Adicional',
    '      (C4) Reservas',
    '      (C5) Ajustes al Patrimonio',
    '      (C6) Resultados Acumulados',
    '      (C7) Resultados no realizados',
    '      (C8) Resultados Netos del Ejercicio',
    '(D) TOTAL PASIVO Y PATRIMONIO',
]

COLS_META = ["code", "MES", "Tipo", "MICROFINAN.", "NACIONAL", "Empresa", "NOMB_CORREG", "Moneda"]


# ============== utilidades locales ==============

def _norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).strip())


def _es_encabezado_o_fin(v) -> bool:
    v = _norm(v)
    vl = v.lower()
    if not v:
        return True
    if vl.startswith("balance general") or vl.startswith("estado de ganancias"):
        return True
    if re.match(r"^\d{4}-\d{2}-\d{2}", v):
        return True
    if vl.startswith("(en miles") or vl.startswith("( en miles"):
        return True
    if vl.startswith("tipo de cambio") or v.startswith("*") or vl.startswith("mediante") or v.startswith("1/"):
        return True
    if len(v) > 100:
        return True
    return False


def _es_total_entidad(v) -> bool:
    return _norm(v).upper().startswith("TOTAL")


def _find_situacion_row(raw: pd.DataFrame):
    for r in range(min(10, len(raw))):
        count = sum(1 for c in range(raw.shape[1]) if _norm(raw.iat[r, c]).upper() == "MN")
        if count >= 2:
            return r
    return None


def _build_canon_lookup(canon_list):
    canon_strip = [(c, re.sub(r"^\s*\(\S+\)\s*", "", c).strip()) for c in canon_list]
    choices = [s for _, s in canon_strip]
    mapa = {s: c for c, s in canon_strip}
    return choices, mapa


# ============== lectura de una hoja (BG o ER) -- misma lógica para ambas ==============

def _leer_hoja(contenido_bytes: bytes, sheet_name: str, canon_list: list, manual_dict: dict, detiene_en=None):
    raw = pd.read_excel(BytesIO(contenido_bytes), sheet_name=sheet_name, header=None)
    sit_row = _find_situacion_row(raw)
    if sit_row is None:
        raise ValueError(f"No se encontró la fila de situación (MN/ME/TOTAL) en la hoja '{sheet_name}'.")
    entity_row = sit_row - 1

    entity_starts = [(c, _norm(raw.iat[entity_row, c])) for c in range(raw.shape[1])
                      if _norm(raw.iat[sit_row, c]).upper() == "MN"]
    # excluye entidades TOTAL (Total Banca Múltiple, etc.) -- esta base no las incluye
    entity_starts = [(c, e) for c, e in entity_starts if not _es_total_entidad(e)]

    choices, mapa_cuenta = _build_canon_lookup(canon_list)
    cache_cuenta = {}

    def resolver_cuenta(label):
        if label in cache_cuenta:
            return cache_cuenta[label]
        if label.upper() in manual_dict:
            cache_cuenta[label] = manual_dict[label.upper()]
            return cache_cuenta[label]
        match = process.extractOne(label, choices, scorer=fuzz.WRatio, processor=str.lower)
        resultado = mapa_cuenta[match[0]] if match and match[1] >= UMBRAL_FUZZY_CUENTA else None
        cache_cuenta[label] = resultado
        return resultado

    datos = {empresa: {"MN": {}, "ME": {}, "TOTAL": {}} for _, empresa in entity_starts}

    for r in range(sit_row + 1, len(raw)):
        label = _norm(raw.iat[r, 0])
        if not label:
            continue
        if _es_encabezado_o_fin(label):
            break
        if detiene_en and label.upper() == detiene_en:
            break
        canon = resolver_cuenta(label)
        if canon is None:
            continue  # cuenta no reconocida (footer residual, etc.) -- se ignora
        for c, empresa in entity_starts:
            mn = pd.to_numeric(raw.iat[r, c], errors="coerce")
            me = pd.to_numeric(raw.iat[r, c + 1], errors="coerce")
            tot = pd.to_numeric(raw.iat[r, c + 2], errors="coerce")
            datos[empresa]["MN"][canon] = 0.0 if pd.isna(mn) else mn
            datos[empresa]["ME"][canon] = 0.0 if pd.isna(me) else me
            datos[empresa]["TOTAL"][canon] = 0.0 if pd.isna(tot) else tot

    return datos


def _datos_a_dataframe(datos: dict, canon_list: list, tipo: str) -> pd.DataFrame:
    filas = []
    for empresa, monedas in datos.items():
        for moneda, valores in monedas.items():
            fila = {"Tipo": tipo, "Empresa": empresa, "Moneda": moneda}
            for col in canon_list:
                fila[col] = valores.get(col, 0.0)
            filas.append(fila)
    return pd.DataFrame(filas)


# ============== normalización contra el maestro (fuzzy) ==============

def _normalizar_entidades(df: pd.DataFrame, maestro: pd.DataFrame):
    """
    Devuelve (df_normalizado, lista_sin_mapeo). Usa fuzzy_match_entidad de
    utils_sbs, que ya limpia asteriscos internamente antes de matchear (el
    mismo fix que se aplicó a la limpieza de nombres de cuenta). Necesita
    también microfinanciera y nacional del maestro (igual que Personal).
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

    df["NOMB_CORREG"] = df["Empresa"].map(lambda e: filas_maestro[e]["nombre_bd"])
    df["MICROFINAN."] = df["Empresa"].map(lambda e: filas_maestro[e]["microfinanciera"])
    df["NACIONAL"] = df["Empresa"].map(lambda e: filas_maestro[e]["nacional"])
    # SMFE si está en la lista de 19, si no vacía (None) -- igual que Estructura de
    # Gasto/Ingresos Financieros, no "SF"
    df[">50% CB"] = df["NOMB_CORREG"].apply(lambda nb: "SMFE" if clasificar_50cb(nb) == "SMFE" else None)
    return df, sin_mapeo


# ============== pipeline ==============

def run(anio: int, mes_num: int):
    """
    Descarga, procesa y exporta EEFF (BG+ER) para el corte (anio, mes_num).
    Devuelve (resultado_er, resultado_bg) -- a diferencia de las otras 16
    bases, esta exporta 2 hojas en el mismo Excel.
    """
    maestro = cargar_maestro()
    mes_abr = ABR_MES[mes_num]

    todos_er, todos_bg = [], []
    sin_mapeo_totales = set()

    for cfg in CODIGOS_CORTE:
        print(f"[EEFF] Descargando {cfg['codigo']} ({cfg['tipo']}) ...")
        contenido_bytesio = descargar_reporte_bytes(cfg["codigo"], anio, mes_num)
        contenido_bytes = contenido_bytesio.getvalue()

        datos_er = _leer_hoja(contenido_bytes, cfg["hoja_er"], CANON_ER, MANUAL_ER)
        datos_bg = _leer_hoja(contenido_bytes, cfg["hoja_bg"], CANON_BG, MANUAL_BG, detiene_en=DETIENE_BG_EN)
        print(f"  -> {len(datos_er)} entidades en ER, {len(datos_bg)} entidades en BG")

        df_er = _datos_a_dataframe(datos_er, CANON_ER, cfg["tipo"])
        df_bg = _datos_a_dataframe(datos_bg, CANON_BG, cfg["tipo"])

        df_er, sm_er = _normalizar_entidades(df_er, maestro)
        sin_mapeo_totales.update(sm_er)
        df_bg, sm_bg = _normalizar_entidades(df_bg, maestro)
        sin_mapeo_totales.update(sm_bg)

        todos_er.append(df_er)
        todos_bg.append(df_bg)

    if sin_mapeo_totales:
        print("\n[EEFF][AVISO] Entidades sin mapeo (excluidas):")
        for e in sorted(sin_mapeo_totales):
            print(f"  - {e}")
        print("Agrégalas a maestro_entidades.csv en GitHub y vuelve a correr si quieres incluirlas.\n")

    resultado_er = pd.concat(todos_er, ignore_index=True)
    resultado_bg = pd.concat(todos_bg, ignore_index=True)

    mes_periodo = fin_de_mes(anio, mes_num)
    for df in (resultado_er, resultado_bg):
        df["MES"] = mes_periodo
        df["code"] = (df["NOMB_CORREG"].astype(str) + df["Moneda"].astype(str))

    resultado_er = resultado_er[COLS_META + CANON_ER + [">50% CB"]]
    resultado_bg = resultado_bg[COLS_META + CANON_BG + [">50% CB"]]

    output_path = BASE_DIR / f"EEFF_SBS_Consolidado_{mes_abr}{anio}.xlsx"
    with pd.ExcelWriter(output_path) as writer:
        resultado_er.to_excel(writer, sheet_name="er", index=False)
        resultado_bg.to_excel(writer, sheet_name="bg", index=False)

    print(f"[EEFF] Listo. ER: {len(resultado_er)} filas, BG: {len(resultado_bg)} filas. Exportado a: {output_path}")

    return resultado_er, resultado_bg


if __name__ == "__main__":
    from datetime import datetime
    ahora = datetime.now()
    run(ahora.year, ahora.month)
