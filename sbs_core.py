# -*- coding: utf-8 -*-
"""
sbs_core.py — Módulo compartido para todos los pipelines SBS multientidad.

Centraliza: descarga robusta (tolerante a archivos faltantes), carga del
maestro de entidades, selector de corte (Tkinter), y resolución de nombres
de entidad (fuzzy matching insensible a mayúsculas y asteriscos).

Cualquier fix aquí se propaga automáticamente a todos los pipelines que
importen este archivo — no hace falta editar cada uno por separado.
"""
import re
import time
import tkinter as tk
from tkinter import ttk
from io import BytesIO
from pathlib import Path
from datetime import date, datetime
import calendar

import pandas as pd
import requests
import urllib3
from rapidfuzz import fuzz, process

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ============== CONFIGURACIÓN GLOBAL (un solo lugar para cambiarla) ==============

# Carpeta de salida de TODOS los pipelines. Cambia esta línea si algún día quieres
# que los Excel finales vayan a otro lugar que no sea Descargas.
BASE_DIR = Path.home() / "Downloads"
BASE_DIR.mkdir(parents=True, exist_ok=True)

MAESTRO_URL = "https://raw.githubusercontent.com/dalanocau/SBS_Colocaciones/refs/heads/main/maestro_entidades.csv"
MAESTRO_PATH_LOCAL = BASE_DIR / "maestro_entidades.csv"

MAX_REINTENTOS = 3
ESPERA_ENTRE_REINTENTOS = 3
UMBRAL_FUZZY_ENTIDAD = 90

MESES = {
    1: ("Enero", "en"), 2: ("Febrero", "fe"), 3: ("Marzo", "ma"), 4: ("Abril", "ab"),
    5: ("Mayo", "my"), 6: ("Junio", "jn"), 7: ("Julio", "jl"), 8: ("Agosto", "ag"),
    9: ("Setiembre", "se"), 10: ("Octubre", "oc"), 11: ("Noviembre", "no"), 12: ("Diciembre", "di"),
}

LISTA_SMF_NOMBRE_BD = {
    "Mibanco", "Compartamos Banco", "Financiera Confianza", "Financiera Proempresa",
    "Financiera Qapaq", "Financiera Surgir", "CMAC Arequipa", "CMAC Huancayo",
    "CMAC Piura", "CMAC Cusco", "CMAC Trujillo", "CMAC Ica", "CMAC Tacna",
    "CMAC Maynas", "CMAC Paita", "CMAC Del Santa", "CRAC Los Andes",
    "CRAC Prymera", "Edpyme Alternativa",
}

FIRMAS_EXCEL_VALIDAS = (b"\xd0\xcf\x11\xe0", b"PK")


# ============== UTILIDADES DE TEXTO ==============

def norm(s) -> str:
    """Normaliza espacios en blanco de cualquier valor (celda de Excel, texto, etc.)."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).strip())


def es_fin_de_tabla(v, prefijos_extra=()) -> bool:
    """Detecta filas de pie de página / notas al final de un archivo SBS."""
    v = norm(v)
    if not v:
        return False
    vl = v.lower()
    prefijos = ("nota", "fuente", "mediante", "1/", "2/", "3/") + prefijos_extra
    if v.startswith("*"):
        return True
    if any(vl.startswith(p.lower()) or v.startswith(p) for p in prefijos):
        return True
    if len(v) > 100:
        return True
    return False


def es_total_entidad(v, etiquetas_exactas=frozenset()) -> bool:
    """True si el texto es una fila/columna de TOTAL por sector (no una entidad real).
    Exige que EMPIECE con TOTAL (no solo que lo contenga) para no confundir nombres
    reales como 'EC TOTAL Servicios Financieros' con una fila de total."""
    vu = norm(v).upper()
    return vu.startswith("TOTAL") or vu in etiquetas_exactas


# ============== SELECTOR DE CORTE (TKINTER) ==============

def seleccionar_corte(titulo_ventana: str = "Seleccionar corte SBS"):
    """Abre una ventana para elegir año y mes. Devuelve (anio:str, nmes:int)."""
    seleccion = {}

    def confirmar():
        seleccion["anio"] = combo_anio.get()
        seleccion["nmes"] = combo_mes.current() + 1
        ventana.destroy()

    ventana = tk.Tk()
    ventana.title(titulo_ventana)
    ventana.geometry("320x180")
    ventana.resizable(False, False)

    tk.Label(ventana, text="Año:", font=("Segoe UI", 10)).pack(pady=(15, 0))
    anio_actual = datetime.now().year
    anios = [str(a) for a in range(anio_actual - 3, anio_actual + 1)]
    combo_anio = ttk.Combobox(ventana, values=anios, state="readonly", width=15)
    combo_anio.set(str(anio_actual))
    combo_anio.pack()

    tk.Label(ventana, text="Mes:", font=("Segoe UI", 10)).pack(pady=(10, 0))
    nombres_meses = [v[0] for v in MESES.values()]
    combo_mes = ttk.Combobox(ventana, values=nombres_meses, state="readonly", width=15)
    combo_mes.current(datetime.now().month - 1)
    combo_mes.pack()

    tk.Button(ventana, text="Procesar", command=confirmar, width=15).pack(pady=20)
    ventana.mainloop()

    if not seleccion:
        raise RuntimeError("No se seleccionó ningún corte (se cerró la ventana).")
    return seleccion["anio"], seleccion["nmes"]


def corte_fin_de_mes(anio: str, nmes: int) -> date:
    ultimo_dia = calendar.monthrange(int(anio), nmes)[1]
    return date(int(anio), nmes, ultimo_dia)


def url_base_sbs(anio: str, nmes: int) -> str:
    mes_nombre, _ = MESES[nmes]
    return f"https://intranet2.sbs.gob.pe/estadistica/financiera/{anio}/{mes_nombre}/"


# ============== DESCARGA ROBUSTA (tolerante a archivos faltantes) ==============

def _es_excel_valido(contenido: bytes) -> bool:
    return contenido[:4].startswith(FIRMAS_EXCEL_VALIDAS[0]) or contenido[:2] == FIRMAS_EXCEL_VALIDAS[1]


def descargar_archivo_sbs(codigo: str, anio: str, mes_abr: str, url_base: str) -> bytes:
    """Descarga UN archivo. Reintenta MAX_REINTENTOS veces. Si falla, levanta
    RuntimeError — usa descargar_multiples() si quieres que un archivo faltante
    no tumbe el resto del proceso."""
    nombre_archivo = f"{codigo}-{mes_abr}{anio}.xls"
    url = url_base + nombre_archivo

    ultimo_error = None
    for intento in range(1, MAX_REINTENTOS + 1):
        try:
            resp = requests.get(url, verify=False, timeout=30)
            resp.raise_for_status()
            if not _es_excel_valido(resp.content):
                muestra = resp.content[:200]
                if b"<html" in muestra.lower() or b"<!doctype" in muestra.lower():
                    raise ValueError(
                        f"La SBS devolvió una página web en vez del Excel para {codigo} "
                        f"(posible error de red/autenticación, o el archivo no existe todavía para este corte)."
                    )
                raise ValueError(f"Contenido no parece un Excel válido para {codigo} ({len(resp.content)} bytes).")
            return resp.content
        except (requests.RequestException, ValueError) as e:
            ultimo_error = e
            if intento < MAX_REINTENTOS:
                print(f"  [aviso] intento {intento}/{MAX_REINTENTOS} falló para {codigo}: {e}. Reintentando...")
                time.sleep(ESPERA_ENTRE_REINTENTOS)

    raise RuntimeError(f"No se pudo descargar {codigo} tras {MAX_REINTENTOS} intentos. Último error: {ultimo_error}")


def descargar_multiples(codigos_config: list, anio: str, nmes: int, campo_codigo: str = "codigo"):
    """Descarga varios archivos SBS, SIN dejar que un archivo faltante tumbe el resto.

    codigos_config: lista de dicts, cada uno con al menos {campo_codigo: "B-1234", "tipo": "Bancos", ...}
    Devuelve (contenidos, fallidos):
      - contenidos: dict {tipo: bytes} SOLO de los que sí se descargaron
      - fallidos:   lista de dicts {tipo, codigo, error} de los que no se pudieron obtener

    Uso típico:
        contenidos, fallidos = descargar_multiples(CODIGOS_CORTE, ANIO, NMES)
        for cfg in CODIGOS_CORTE:
            if cfg["tipo"] not in contenidos:
                continue  # este archivo no salió este corte, se omite sin romper nada
            ... procesar contenidos[cfg["tipo"]] ...
    """
    url_base = url_base_sbs(anio, nmes)
    _, mes_abr = MESES[nmes]
    contenidos, fallidos = {}, []

    for cfg in codigos_config:
        codigo = cfg[campo_codigo]
        tipo = cfg.get("tipo", codigo)
        print(f"Descargando {codigo} ({tipo}) ...")
        try:
            contenidos[tipo] = descargar_archivo_sbs(codigo, anio, mes_abr, url_base)
        except Exception as e:
            print(f"  [OMITIDO] {tipo} ({codigo}) no se pudo obtener para este corte: {e}")
            fallidos.append({"tipo": tipo, "codigo": codigo, "error": str(e)})

    if fallidos:
        print(f"\n[AVISO] {len(fallidos)} archivo(s) no disponibles para este corte, se omiten del resultado:")
        for f in fallidos:
            print(f"  - {f['tipo']} ({f['codigo']})")
        print("Esto es normal si la SBS aún no publicó todos los archivos del mes — el resto del")
        print("proceso continúa igual; vuelve a correr más tarde si quieres completar los faltantes.\n")

    return contenidos, fallidos


# ============== CARGA DEL MAESTRO (GitHub con fallback local) ==============

def cargar_maestro() -> pd.DataFrame:
    for intento in range(1, MAX_REINTENTOS + 1):
        try:
            resp = requests.get(MAESTRO_URL, verify=False, timeout=20)
            resp.raise_for_status()
            maestro = pd.read_csv(BytesIO(resp.content), encoding="utf-8-sig")
            print("Maestro de entidades cargado desde GitHub.")
            # guarda copia local de respaldo para la próxima vez que GitHub falle
            try:
                maestro.to_csv(MAESTRO_PATH_LOCAL, index=False, encoding="utf-8-sig")
            except Exception:
                pass
            break
        except Exception as e:
            print(f"  [aviso] intento {intento}/{MAX_REINTENTOS} falló al leer el maestro desde GitHub: {e}")
            if intento < MAX_REINTENTOS:
                time.sleep(ESPERA_ENTRE_REINTENTOS)
    else:
        if not MAESTRO_PATH_LOCAL.exists():
            raise RuntimeError(f"No se pudo cargar el maestro ni desde GitHub ni localmente ({MAESTRO_PATH_LOCAL}).")
        print(f"No se pudo cargar desde GitHub, usando copia local: {MAESTRO_PATH_LOCAL}")
        maestro = pd.read_csv(MAESTRO_PATH_LOCAL, encoding="utf-8-sig")

    maestro["nombre_sbs"] = maestro["nombre_sbs"].apply(norm)
    return maestro


# ============== RESOLUCIÓN DE NOMBRES DE ENTIDAD (fuzzy, robusto) ==============

class ResolverEntidades:
    """Envuelve el maestro + caché + fuzzy matching para resolver nombres de
    entidad SBS a su nombre_bd normalizado. Insensible a mayúsculas/minúsculas
    y a asteriscos de nota al pie (ej. 'CMAC Arequipa***' -> 'CMAC Arequipa')."""

    def __init__(self, maestro: pd.DataFrame, umbral: int = UMBRAL_FUZZY_ENTIDAD):
        self.mapa_bd = maestro.set_index("nombre_sbs")["nombre_bd"].to_dict()
        self.mapa_micro = maestro.set_index("nombre_sbs")["microfinanciera"].to_dict() \
            if "microfinanciera" in maestro.columns else {}
        self.mapa_nac = maestro.set_index("nombre_sbs")["nacional"].to_dict() \
            if "nacional" in maestro.columns else {}
        self.lista_nombres = maestro["nombre_sbs"].tolist()
        self.umbral = umbral
        self._cache = {}
        self.correcciones = []
        self.sin_match = []

    def resolver(self, nombre_crudo: str):
        """Devuelve el nombre_bd correspondiente, o None si no se pudo mapear."""
        if nombre_crudo in self._cache:
            return self._cache[nombre_crudo]
        if nombre_crudo in self.mapa_bd:
            self._cache[nombre_crudo] = nombre_crudo
            return nombre_crudo

        sin_asteriscos = re.sub(r"\*+", "", nombre_crudo).strip()
        if sin_asteriscos in self.mapa_bd:
            self._cache[nombre_crudo] = sin_asteriscos
            return sin_asteriscos

        match = process.extractOne(sin_asteriscos, self.lista_nombres, scorer=fuzz.WRatio, processor=str.lower)
        if match and match[1] >= self.umbral:
            nombre_matcheado, score, _ = match
            self.correcciones.append((nombre_crudo, nombre_matcheado, score))
            self._cache[nombre_crudo] = nombre_matcheado
            return nombre_matcheado

        self.sin_match.append(nombre_crudo)
        self._cache[nombre_crudo] = None
        return None

    def nombre_bd(self, nombre_sbs_resuelto):
        return self.mapa_bd.get(nombre_sbs_resuelto)

    def microfinanciera(self, nombre_sbs_resuelto):
        return self.mapa_micro.get(nombre_sbs_resuelto)

    def nacional(self, nombre_sbs_resuelto):
        return self.mapa_nac.get(nombre_sbs_resuelto)

    def es_smf(self, nombre_bd_valor) -> bool:
        return nombre_bd_valor in LISTA_SMF_NOMBRE_BD

    def imprimir_resumen(self):
        if self.correcciones:
            print("\n[INFO] Nombres corregidos automáticamente por matching difuso:")
            vistos = set()
            for orig, matched, score in self.correcciones:
                if orig not in vistos:
                    print(f"  '{orig}' -> '{matched}' (similitud {score:.0f}%)")
                    vistos.add(orig)
        sin_match_unicos = sorted(set(self.sin_match))
        if sin_match_unicos:
            print("\n[AVISO] Entidades sin mapeo, ni exacto ni difuso (excluidas del resultado):")
            for e in sin_match_unicos:
                print(f"  - {e}")
            print("Agrégalas al maestro_entidades.csv en GitHub y vuelve a correr si quieres incluirlas.\n")
        return sin_match_unicos
