"""
orquestador_ui.py
Punto de entrada único para correr las 17 bases del proyecto SBS Multientidad.

Flujo:
  Pantalla 1 (Modo)      -> "Procesar todo" / "Procesar seleccionados"
  Pantalla 2 (Checklist) -> solo si eligió "seleccionados"; las 17 bases
                             aparecen marcadas, y si ya se corrió ese mismo
                             corte antes, las ya procesadas salen desmarcadas
  Pantalla 3 (Año/Mes)   -> común a ambos modos
  -> ejecuta cada base seleccionada, sin que una falla tumbe a las demás
  -> guarda qué bases se corrieron para ese corte (para la próxima vez)
  -> muestra un resumen final (OK / ERROR por base)

IMPORTANTE: este archivo asume que cada procesar_X.py fue refactorizado de
"script standalone" a una función run(anio, mes_num) -> None (o que guarda
su propio Excel y no retorna nada). Esa refactorización es el siguiente paso
para cada una de las 17 bases; aquí abajo se importan como placeholders.
"""

import json
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk

from utils_sbs import BASE_DIR, MESES
from procesar_colocaciones import run as _run_colocaciones

# ---------------------------------------------------------------------------
# Registro de bases: nombre visible -> función run(anio, mes_num)
#
# Reemplazar cada import/placeholder por el real cuando el script
# correspondiente esté refactorizado a función. Mientras tanto, cada entrada
# puede apuntar a una función dummy para probar la UI de punta a punta.
# ---------------------------------------------------------------------------

def _pendiente(nombre):
    """Placeholder: usar hasta que el procesar_X.py real esté refactorizado."""
    def _run(anio, mes_num):
        raise NotImplementedError(f"'{nombre}' aún no está conectado al orquestador")
    return _run


BASES = {
    "Colocaciones": _run_colocaciones,
    "Depósitos": _pendiente("Depósitos"),
    "Personal": _pendiente("Personal"),
    "Castigos": _pendiente("Castigos"),
    "Clientes de Crédito": _pendiente("Clientes de Crédito"),
    "Clientes de Ahorro": _pendiente("Clientes de Ahorro"),
    "Categoría de Riesgo del Cliente": _pendiente("Categoría de Riesgo del Cliente"),
    "Patrimonio Efectivo": _pendiente("Patrimonio Efectivo"),
    "RCG": _pendiente("RCG"),
    "Estructura de Gasto": _pendiente("Estructura de Gasto"),
    "Ingresos Financieros": _pendiente("Ingresos Financieros"),
    "Ratio de Liquidez": _pendiente("Ratio de Liquidez"),
    "Oficinas por Zona Geográfica": _pendiente("Oficinas por Zona Geográfica"),
    "Créditos y Depósitos por Zona Geográfica": _pendiente("Créditos y Depósitos por Zona Geográfica"),
    "Indicadores": _pendiente("Indicadores"),
    "Gastos Administrativos": _pendiente("Gastos Administrativos"),
    "EEFF": _pendiente("EEFF"),
}

ESTADO_PATH = BASE_DIR / "orquestador_estado.json"


# ---------------------------------------------------------------------------
# Persistencia de la última selección por corte
# ---------------------------------------------------------------------------

def _cargar_estado() -> dict:
    if ESTADO_PATH.exists():
        try:
            return json.loads(ESTADO_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _guardar_estado(corte: str, bases_ok: list):
    estado = _cargar_estado()
    ya_procesadas = set(estado.get(corte, []))
    ya_procesadas.update(bases_ok)
    estado[corte] = sorted(ya_procesadas)
    ESTADO_PATH.parent.mkdir(parents=True, exist_ok=True)
    ESTADO_PATH.write_text(json.dumps(estado, ensure_ascii=False, indent=2), encoding="utf-8")


def _bases_ya_procesadas(corte: str) -> set:
    return set(_cargar_estado().get(corte, []))


# ---------------------------------------------------------------------------
# Ejecución
# ---------------------------------------------------------------------------

def ejecutar(bases_seleccionadas: list, anio: int, mes_num: int):
    corte = f"{anio}-{mes_num:02d}"
    resultados = {}

    for nombre in bases_seleccionadas:
        func = BASES[nombre]
        try:
            func(anio, mes_num)
            resultados[nombre] = "OK"
        except Exception as e:
            resultados[nombre] = f"ERROR: {e}"

    ok = [n for n, r in resultados.items() if r == "OK"]
    _guardar_estado(corte, ok)

    resumen = "\n".join(f"{'✅' if r == 'OK' else '❌'} {n} — {r}" for n, r in resultados.items())
    messagebox.showinfo("Resumen de procesamiento", f"Corte {corte}\n\n{resumen}")


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

class OrquestadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Orquestador SBS Multientidad")
        self.modo = None
        self.checkbox_vars = {}
        self._pantalla_modo()

    def _limpiar(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    # -- Pantalla 1: Modo -----------------------------------------------

    def _pantalla_modo(self):
        self._limpiar()
        frame = ttk.Frame(self.root, padding=20)
        frame.pack()

        ttk.Label(frame, text="¿Qué deseas procesar?", font=("", 12, "bold")).pack(pady=(0, 15))
        ttk.Button(
            frame, text="Procesar todo", width=30,
            command=lambda: self._elegir_modo("todo"),
        ).pack(pady=5)
        ttk.Button(
            frame, text="Procesar seleccionados", width=30,
            command=lambda: self._elegir_modo("seleccionados"),
        ).pack(pady=5)

    def _elegir_modo(self, modo):
        self.modo = modo
        if modo == "todo":
            self.bases_elegidas = list(BASES.keys())
            self._pantalla_corte()
        else:
            self._pantalla_checklist()

    # -- Pantalla 2: Checklist (solo modo "seleccionados") ---------------

    def _pantalla_checklist(self):
        self._limpiar()
        frame = ttk.Frame(self.root, padding=20)
        frame.pack()

        ttk.Label(frame, text="Selecciona las bases a procesar", font=("", 12, "bold")).pack(pady=(0, 10))

        # Pre-marcar según lo ya corrido para el corte más reciente conocido,
        # si existe; si no hay historial, todas quedan marcadas.
        ultimo_corte = self._ultimo_corte_registrado()
        ya_procesadas = _bases_ya_procesadas(ultimo_corte) if ultimo_corte else set()

        self.checkbox_vars = {}
        for nombre in BASES:
            marcado = nombre not in ya_procesadas
            var = tk.BooleanVar(value=marcado)
            self.checkbox_vars[nombre] = var
            ttk.Checkbutton(frame, text=nombre, variable=var).pack(anchor="w")

        ttk.Button(frame, text="Continuar", command=self._confirmar_checklist).pack(pady=(15, 0))

    def _ultimo_corte_registrado(self):
        estado = _cargar_estado()
        if not estado:
            return None
        return sorted(estado.keys())[-1]

    def _confirmar_checklist(self):
        self.bases_elegidas = [n for n, v in self.checkbox_vars.items() if v.get()]
        if not self.bases_elegidas:
            messagebox.showwarning("Atención", "Selecciona al menos una base.")
            return
        self._pantalla_corte()

    # -- Pantalla 3: Año / Mes (común a ambos modos) ---------------------

    def _pantalla_corte(self):
        self._limpiar()
        frame = ttk.Frame(self.root, padding=20)
        frame.pack()

        ttk.Label(frame, text="Selecciona el corte", font=("", 12, "bold")).pack(pady=(0, 10))

        anio_actual = datetime.now().year
        anios = [str(a) for a in range(anio_actual - 3, anio_actual + 1)]

        ttk.Label(frame, text="Año").pack(anchor="w")
        self.combo_anio = ttk.Combobox(frame, values=anios, state="readonly")
        self.combo_anio.set(str(anio_actual))
        self.combo_anio.pack(fill="x", pady=(0, 10))

        ttk.Label(frame, text="Mes").pack(anchor="w")
        self.combo_mes = ttk.Combobox(frame, values=list(MESES.values()), state="readonly")
        self.combo_mes.set(MESES[datetime.now().month])
        self.combo_mes.pack(fill="x", pady=(0, 15))

        ttk.Button(frame, text="Procesar", command=self._procesar).pack()

    def _procesar(self):
        anio = int(self.combo_anio.get())
        mes_nombre = self.combo_mes.get()
        mes_num = [k for k, v in MESES.items() if v == mes_nombre][0]

        self.root.withdraw()
        ejecutar(self.bases_elegidas, anio, mes_num)
        self.root.destroy()


def main():
    root = tk.Tk()
    OrquestadorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
