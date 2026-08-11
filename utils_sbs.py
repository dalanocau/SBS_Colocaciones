"""
utils_sbs.py
Módulo compartido para el proyecto SBS Multientidad (17 bases).

Centraliza la lógica que hoy está duplicada en cada procesar_X.py:
- Descarga de reportes SBS con reintentos
- Carga (con cache) del maestro de entidades desde GitHub
- Fuzzy matching contra el maestro (con el fix processor=str.lower)
- Limpieza de asteriscos en nombres de entidad/cuenta
- Detección de filas TOTAL (startswith + set de etiquetas sin prefijo)
- Cálculo del corte de fin de mes (último día real del mes)
- Clasificación SF / SMF / SMFE

Todos los procesar_X.py deben hacer:
    from utils_sbs import *
en vez de reimplementar cada una de estas piezas.
"""

import calendar
import io
import time
from pathlib import Path

import pandas as pd
import requests
from rapidfuzz import process, fuzz

# ---------------------------------------------------------------------------
# Configuración general
# ---------------------------------------------------------------------------

BASE_DIR = Path.home() / "Downloads"

MAESTRO_URL = (
    "https://raw.githubusercontent.com/dalanocau/SBS_Colocaciones/"
    "refs/heads/main/maestro_entidades.csv"
)

MESES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Setiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

ABR_MES = {
    1: "en", 2: "fe", 3: "mr", 4: "ab", 5: "my", 6: "jn",
    7: "jl", 8: "ag", 9: "se", 10: "oc", 11: "no", 12: "di",
}

URL_TEMPLATE = (
    "https://intranet2.sbs.gob.pe/estadistica/financiera/{anio}/"
    "{mes_nombre}/{codigo}-{abr_mes}{anio_corto}.xls"
)

# La lista de entidades SMF (Sistema Microfinanciero) para las columnas
# "Clasificación" / ">50% CB" / ">=50% MYPE" NO se hardcodea acá: vive en la
# columna "clasificacion" del maestro (nombre_bd -> SF/SMF/-). Ver
# _lista_smf() más abajo, que la deriva del maestro y la cachea.
_LISTA_SMF_CACHE = None

# Etiquetas de fila TOTAL que la SBS a veces usa SIN la palabra "TOTAL" al
# inicio (además del caso normal que sí empieza con "TOTAL").
ETIQUETAS_TOTAL_SIN_PREFIJO = {
    "CAJAS MUNICIPALES",
    "CAJAS RURALES DE AHORRO Y CRÉDITO",
    "EMPRESAS DE CRÉDITOS",
    "EMPRESAS FINANCIERAS",
    "BANCA MÚLTIPLE",
}

_MAESTRO_CACHE = None


# ---------------------------------------------------------------------------
# Descarga
# ---------------------------------------------------------------------------

def construir_url(codigo: str, anio: int, mes_num: int) -> str:
    """Arma la URL de descarga de un reporte SBS para un código y corte dados."""
    return URL_TEMPLATE.format(
        anio=anio,
        mes_nombre=MESES[mes_num],
        codigo=codigo,
        abr_mes=ABR_MES[mes_num],
        anio_corto=str(anio)[-2:],
    )


def descargar_reporte(
    codigo: str,
    anio: int,
    mes_num: int,
    reintentos: int = 3,
    espera_seg: float = 2.0,
    sheet_name=0,
    verify_ssl: bool = False,
):
    """
    Descarga un reporte SBS y lo devuelve como DataFrame (o dict de DataFrames
    si sheet_name=None y el archivo tiene varias hojas, como en EEFF).

    Reintenta ante fallas de red/timeout. Lanza la última excepción si se
    agotan los reintentos, para que el llamador decida si abortar solo esa
    base (ver patrón de robustez usado en RCG) o toda la corrida.
    """
    url = construir_url(codigo, anio, mes_num)
    ultimo_error = None

    for intento in range(1, reintentos + 1):
        try:
            resp = requests.get(url, verify=verify_ssl, timeout=30)
            resp.raise_for_status()
            return pd.read_excel(io.BytesIO(resp.content), sheet_name=sheet_name)
        except Exception as e:
            ultimo_error = e
            if intento < reintentos:
                time.sleep(espera_seg * intento)

    raise RuntimeError(
        f"No se pudo descargar {codigo} ({anio}-{mes_num:02d}) tras "
        f"{reintentos} intentos: {ultimo_error}"
    )


# ---------------------------------------------------------------------------
# Maestro de entidades
# ---------------------------------------------------------------------------

def cargar_maestro(forzar_recarga: bool = False) -> pd.DataFrame:
    """
    Descarga el maestro_entidades.csv desde GitHub y lo cachea en memoria
    para que las 17 bases de una misma corrida no vuelvan a pegarle a la red.
    """
    global _MAESTRO_CACHE
    if _MAESTRO_CACHE is None or forzar_recarga:
        _MAESTRO_CACHE = pd.read_csv(MAESTRO_URL)
    return _MAESTRO_CACHE


# ---------------------------------------------------------------------------
# Limpieza de nombres
# ---------------------------------------------------------------------------

def limpiar_asteriscos(texto: str) -> str:
    """Quita TODOS los asteriscos de nota al pie (no solo al final)."""
    if not isinstance(texto, str):
        return texto
    return texto.replace("*", "").strip()


def normalizar_nombre(texto: str) -> str:
    """Limpieza estándar antes de fuzzy match: asteriscos + espacios extra."""
    if not isinstance(texto, str):
        return texto
    return " ".join(limpiar_asteriscos(texto).split())


# ---------------------------------------------------------------------------
# Fuzzy matching
# ---------------------------------------------------------------------------

def fuzzy_match_entidad(nombre_sbs: str, maestro: pd.DataFrame = None, umbral: int = 80):
    """
    Matchea un nombre crudo de la SBS contra la columna nombre_sbs del
    maestro, usando processor=str.lower (fix del bug de mayúsculas
    encontrado en Oficinas por Zona Geográfica).

    Devuelve la fila del maestro (Series) si hay match >= umbral, o None.
    """
    if maestro is None:
        maestro = cargar_maestro()

    nombre_limpio = normalizar_nombre(nombre_sbs)
    choices = maestro["nombre_sbs"].tolist()

    resultado = process.extractOne(
        nombre_limpio,
        choices,
        scorer=fuzz.WRatio,
        processor=str.lower,
    )

    if resultado is None or resultado[1] < umbral:
        return None

    match_nombre = resultado[0]
    return maestro.loc[maestro["nombre_sbs"] == match_nombre].iloc[0]


# ---------------------------------------------------------------------------
# Detección de filas TOTAL
# ---------------------------------------------------------------------------

def es_fila_total(texto: str) -> bool:
    """
    True si el texto corresponde a una fila de TOTAL por sector.
    Exige que EMPIECE con "TOTAL" (no que solo lo contenga, para no
    confundir entidades reales como "EC TOTAL Servicios Financieros"),
    más el set de etiquetas que la SBS a veces usa sin ese prefijo.
    """
    if not isinstance(texto, str):
        return False
    limpio = normalizar_nombre(texto).upper()
    if limpio.startswith("TOTAL"):
        return True
    return limpio in ETIQUETAS_TOTAL_SIN_PREFIJO


# ---------------------------------------------------------------------------
# Corte de fin de mes
# ---------------------------------------------------------------------------

def fin_de_mes(anio: int, mes_num: int):
    """Devuelve el último día real del mes (calendar.monthrange)."""
    ultimo_dia = calendar.monthrange(anio, mes_num)[1]
    return f"{anio}-{mes_num:02d}-{ultimo_dia:02d}"


# ---------------------------------------------------------------------------
# Clasificación SF / SMF / SMFE
# ---------------------------------------------------------------------------

def _lista_smf(maestro: pd.DataFrame = None) -> set:
    """
    Deriva la lista de nombre_bd marcados SMF en el maestro (columna
    'clasificacion'), cacheada en memoria. Hoy son 19 entidades, pero se lee
    del CSV para no tener que tocar código si la lista cambia.
    """
    global _LISTA_SMF_CACHE
    if _LISTA_SMF_CACHE is None:
        if maestro is None:
            maestro = cargar_maestro()
        _LISTA_SMF_CACHE = set(
            maestro.loc[maestro["clasificacion"] == "SMF", "nombre_bd"]
        )
    return _LISTA_SMF_CACHE


def clasificar_sf_smf(nombre_bd: str) -> str:
    """SMF si nombre_bd está marcado SMF en el maestro, SF si no."""
    return "SMF" if nombre_bd in _lista_smf() else "SF"


def clasificar_50cb(nombre_bd: str) -> str:
    """SMFE si nombre_bd está marcado SMF en el maestro, SF si no."""
    return "SMFE" if nombre_bd in _lista_smf() else "SF"
