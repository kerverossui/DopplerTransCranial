"""
Ultrasonografía Neurocrítica — v1.0
=====================================
Archivo único integrado.
Módulos:
  · Vaina de Nervio Óptico (VNO bilateral)
  · Doppler Transcraneal TCD (DopplerACMApp v3.1)

── COMPILACIÓN A .EXE (PyInstaller) ──────────────────────────────────────────
El informe clínico en Word usa la librería 'python-docx'. Un .exe compilado
NO puede instalarla en tiempo de ejecución (no hay pip ni intérprete
accesibles dentro del .exe) — debe quedar incluida DESDE LA COMPILACIÓN:

    1)  pip install matplotlib numpy python-docx
    2)  pyinstaller --collect-data docx  [...tus demás opciones...]  archivo.py

El flag --collect-data docx es OBLIGATORIO: python-docx trae una plantilla
interna (docx/templates/default.docx) que PyInstaller no detecta solo; sin
ese flag el .exe arranca bien pero revienta al crear el primer documento
(Document()). Si usas un .spec en vez de la línea de comandos, agrega el
equivalente con collect_data_files('docx') en datas=[...].
"""

_BUILD_ID = "2026-10-17-co2-y-auditoria-datos"
# ^ Si este texto NO aparece en el título de las ventanas al abrir la app,
#   estás ejecutando un archivo .py distinto al que se te entregó (una copia
#   vieja, otra ruta, o un .exe compilado antes de este cambio).

# Autor de LA APLICACIÓN (no de informes, casos ni reportes clínicos).
# El blob está ofuscado (XOR + Base64) para que no se edite de un vistazo;
# no es cifrado fuerte. Los informes DOCX / PNG / .ncu clínicos siguen
# atribuyéndose al médico que los firma, nunca a este crédito.
_APP_AUTHOR_BLOB = "CjF7EQQqOBEeIiBdbgcgQy8tdXIrMSNQIDcwQg=="
_APP_AUTHOR_KEY  = b"NCU1"


def _autor_app():
    import base64 as _b64
    raw = _b64.b64decode(_APP_AUTHOR_BLOB)
    return bytes(b ^ _APP_AUTHOR_KEY[i % len(_APP_AUTHOR_KEY)]
                 for i, b in enumerate(raw)).decode("utf-8")


def _meta_app():
    return {
        "autor_app": _autor_app(),
        "rol": "autor_software",
        "build": _BUILD_ID,
        "nota": ("Autor de la aplicación. Los informes, casos y reportes "
                 "clínicos son responsabilidad del médico que los elabora."),
    }


# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 0 — CHEQUEO E INSTALACIÓN DE DEPENDENCIAS AL ARRANCAR
# ══════════════════════════════════════════════════════════════════════════════
#   Se ejecuta ANTES de importar matplotlib/numpy/docx a nivel de módulo, para
#   poder detectar que faltan sin que la app misma reviente al arrancar (un
#   ImportError en un import de nivel de módulo mata el proceso antes de que
#   exista ninguna ventana para avisarlo). Por eso este bloque usa solo
#   librería estándar (sys, os, subprocess, importlib, tkinter).
#
#   Solo cubre lo que puede romper el arranque o una función central:
#     - matplotlib, numpy  → CRÍTICOS: toda la app depende de ellos para dibujar.
#     - python-docx        → IMPORTANTE: solo el informe en Word lo necesita;
#                             si falta, el resto de la app sigue funcionando.
#   anthropic / edge-tts / pygame son funciones opcionales (comentario IA,
#   narración por voz) que YA manejan su propia ausencia con avisos puntuales
#   en el momento de usarlas — no se fuerzan aquí para no pedir instalar (ni
#   red) cosas que la mayoría de sesiones no usa.
import sys, os, subprocess, importlib.util
import tkinter as _tk
from tkinter import messagebox as _messagebox

_REQUISITOS = [
    # (paquete para pip, módulo para import, es_critico)
    ("matplotlib",  "matplotlib", True),
    ("numpy",       "numpy",      True),
    ("python-docx", "docx",       False),
]


def _dependencias_faltantes():
    return [(pip_name, mod_name, critico)
            for pip_name, mod_name, critico in _REQUISITOS
            if importlib.util.find_spec(mod_name) is None]


def _verificar_dependencias_arranque():
    """Chequea (y en modo desarrollo instala) las dependencias de terceros
    ANTES de que el resto del archivo intente importarlas. No crea ninguna
    ventana si todo está presente (camino rápido, arranque sin demora)."""
    faltan = _dependencias_faltantes()
    if not faltan:
        return  # camino rápido — nada que hacer en el 99% de los arranques

    congelado = getattr(sys, "frozen", False)   # True si es un .exe de PyInstaller
    criticos_faltan  = [f for f in faltan if f[2]]

    # Ventana Tk oculta, solo para poder mostrar diálogos en este punto tan
    # temprano (todavía no existe la ventana principal de la app).
    root = _tk.Tk()
    root.withdraw()

    try:
        if congelado:
            # Dentro de un .exe compilado NUNCA se intenta pip: no hay
            # intérprete ni pip accesibles ahí dentro (sys.executable es el
            # propio .exe). Solo se puede avisar y, si falta algo crítico,
            # impedir que la app siga (arrancaría rota de todas formas).
            nombres = ", ".join(p for p, _, _ in faltan)
            detalle_criticos = (
                "\n\nFaltan componentes CRÍTICOS (matplotlib/numpy) — la "
                "aplicación no puede continuar en este estado."
                if criticos_faltan else
                "\n\nSolo falta un componente no crítico; la aplicación "
                "puede continuar, pero el informe en Word no funcionará "
                "hasta reconstruir el .exe.")
            _messagebox.showerror(
                "Faltan componentes en este .exe",
                "Esta versión compilada no incluye: " + nombres + ".\n\n"
                "No se puede instalar dentro de un .exe ya compilado: hay "
                "que reconstruirlo incluyéndolo. En la máquina donde se "
                "compila:\n\n"
                "    pip install " + " ".join(p for p, _, _ in faltan) + "\n"
                "    pyinstaller --collect-data docx  [...opciones...]\n\n"
                "(--collect-data docx solo hace falta si python-docx está "
                "en la lista de arriba.)" + detalle_criticos,
                parent=root)
            if criticos_faltan:
                root.destroy()
                sys.exit(1)
            return  # solo faltaba algo no crítico (docx) → seguir igual

        # Modo desarrollo (python archivo.py): se puede instalar con pip.
        nombres = ", ".join(p for p, _, _ in faltan)
        if not _messagebox.askyesno(
                "Faltan componentes",
                "Esta aplicación necesita las siguientes librerías, que no "
                "están instaladas:\n\n    " + nombres + "\n\n"
                "¿Instalarlas ahora automáticamente?\n"
                "(requiere conexión a internet; puede tardar un minuto)",
                parent=root):
            if criticos_faltan:
                _messagebox.showerror(
                    "No se puede continuar",
                    "matplotlib y/o numpy son indispensables para esta "
                    "aplicación. Instálalos manualmente y vuelve a abrir:\n\n"
                    "    pip install " + " ".join(p for p, _, c in faltan if c),
                    parent=root)
                root.destroy()
                sys.exit(1)
            return  # rechazó instalar, pero solo faltaba algo no crítico

        try:
            proc = subprocess.run(
                [sys.executable, "-m", "pip", "install", "--quiet",
                 "--disable-pip-version-check"] + [p for p, _, _ in faltan],
                capture_output=True, text=True, timeout=300)
        except Exception as e:
            proc = None
            error_txt = str(e)

        ok = proc is not None and proc.returncode == 0
        if not ok:
            detalle = (error_txt if proc is None else
                       (proc.stderr or proc.stdout or "Error desconocido")[-800:])
            _messagebox.showerror(
                "Error de instalación",
                "No se pudieron instalar automáticamente:\n\n" + detalle +
                "\n\nInstálalas manualmente con:\n"
                "    pip install " + nombres, parent=root)

        # Re-chequear qué sigue faltando después del intento de instalación.
        sigue_faltando = _dependencias_faltantes()
        criticos_siguen = [f for f in sigue_faltando if f[2]]
        if criticos_siguen:
            _messagebox.showerror(
                "No se puede continuar",
                "matplotlib y/o numpy siguen sin poder instalarse. La "
                "aplicación no puede abrir sin ellos.\n\n"
                "Instálalos manualmente con:\n"
                "    pip install " + " ".join(p for p, _, c in criticos_siguen),
                parent=root)
            root.destroy()
            sys.exit(1)
        elif not sigue_faltando and ok:
            _messagebox.showinfo("Listo",
                                 "Todas las librerías se instalaron correctamente.",
                                 parent=root)
    finally:
        try:
            root.destroy()
        except Exception:
            pass


_verificar_dependencias_arranque()
print("Ultrasonografía Neurocrítica — build:", _BUILD_ID)


# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 1 — IMPORTS Y PALETA COMPARTIDA
# ══════════════════════════════════════════════════════════════════════════════
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.gridspec as gridspec
from matplotlib.figure import Figure
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# -- Figuras SIN pyplot (corrección del proceso que no se cerraba) --
#   plt.figure()/plt.subplots() con el backend TkAgg crean, por cada figura,
#   una ventana raíz tk.Tk() OCULTA (la del "gestor de figuras" de pyplot).
#   Mientras quede una sola abierta, root.mainloop() no termina al cerrar la
#   ventana principal: el .exe desaparece de la pantalla pero sigue vivo en
#   el Administrador de tareas. Una Figure creada directamente no tiene
#   gestor ni ventana propia: se dibuja en Tk con FigureCanvasTkAgg y se
#   exporta con savefig() exactamente igual.
def _nueva_figura(**kw):
    fig = Figure(**kw)
    FigureCanvasAgg(fig)        # canvas Agg: savefig, draw() y get_renderer()
    return fig


def _nuevos_subplots(nrows=1, ncols=1, **kw):
    fig = _nueva_figura(**kw)
    return fig, fig.subplots(nrows, ncols)


def _cerrar_figura(fig):
    """Libera la memoria de una figura. Seguro con None o ya cerrada."""
    if fig is None:
        return
    try:
        fig.clear()
    except Exception:
        pass

# -- Tipografia global matplotlib (reportes PNG) --
#   El sistema usa sans-serif tecnica de palo seco. matplotlib por defecto usa
#   DejaVu Sans (generica, no matchea la UI). Forzamos una pila sans con
#   prioridad a fuentes tecnicas/condensadas disponibles en Windows, con
#   fallback seguro. Esto alinea los PNG con la UI Tk (Segoe UI).
matplotlib.rcParams["font.family"]      = "sans-serif"
matplotlib.rcParams["font.sans-serif"]  = ["Segoe UI", "Arial Narrow",
                                           "DejaVu Sans Condensed", "Arial",
                                           "DejaVu Sans"]
matplotlib.rcParams["font.style"]       = "normal"   # nunca italic (regla del sistema)
matplotlib.rcParams["mathtext.default"] = "regular"
import json, pickle, traceback, tempfile, queue, threading, logging, weakref
from datetime import datetime


# ── Procesos hijo (audio) y apagado ───────────────────────────────────────────
#   Un subprocess no muere solo cuando muere la app: un PowerShell de audio
#   podía quedar huérfano en el Administrador de tareas. Se registran aquí
#   (referencias débiles: no impiden que se liberen) y se matan al salir.
_HIJOS = weakref.WeakSet()


def _registrar_hijo(proc):
    try:
        _HIJOS.add(proc)
    except TypeError:
        pass


def _apagado_final():
    """Idempotente. Mata procesos hijo vivos y cierra pygame si se cargó."""
    for p in list(_HIJOS):
        try:
            if p.poll() is None:
                p.terminate()
                try:
                    p.wait(timeout=2)
                except subprocess.TimeoutExpired:
                    p.kill()
        except Exception:
            pass
    pg = sys.modules.get("pygame")
    if pg is not None:
        try:
            if pg.mixer.get_init():
                pg.mixer.quit()
            pg.quit()
        except Exception:
            pass


# ── Registro de errores ───────────────────────────────────────────────────────
#   En el .exe "sin consola" sys.stderr es None: cualquier excepción en un
#   botón desaparecía sin dejar rastro. Ahora va a errores.log en la carpeta
#   de configuración (ver _ruta_config).
_log = logging.getLogger("ultrasonografia")


def _configurar_log():
    try:
        from logging.handlers import RotatingFileHandler
        h = RotatingFileHandler(_ruta_config("errores.log"), maxBytes=512_000,
                                backupCount=2, encoding="utf-8")
        h.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        _log.addHandler(h)
        _log.setLevel(logging.INFO)
    except Exception:
        _log.addHandler(logging.NullHandler())

    def _hook(tipo, valor, tb):
        _log.error("Excepción no capturada:\n%s",
                   "".join(traceback.format_exception(tipo, valor, tb)))
    sys.excepthook = _hook
    threading.excepthook = lambda a: _hook(a.exc_type, a.exc_value, a.exc_traceback)

# -- python-docx (informe clínico exportable) --
#   A esta altura _verificar_dependencias_arranque() ya intentó instalarlo si
#   faltaba (modo desarrollo) o ya avisó que el .exe necesita reconstruirse
#   (modo congelado). Este try/except es el respaldo final: cubre el caso de
#   que el usuario haya rechazado la instalación al arrancar, o que solo
#   faltara python-docx (no crítico) en un .exe — la app sigue abriendo igual
#   y _asegurar_docx_disponible() reintentará cuando se use el botón.
_DOCX_OK = False
Document = Inches = Pt = RGBColor = WD_ALIGN_PARAGRAPH = WD_TABLE_ALIGNMENT = None
try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    _DOCX_OK = True
except ImportError:
    pass


def _asegurar_docx_disponible(parent=None):
    """Verifica si python-docx está disponible.

    - En modo DESARROLLO (ejecutando con `python archivo.py`): si falta,
      pregunta una vez y, si el usuario acepta, lo instala automáticamente
      con pip (mismo intérprete que corre la app) sin salir del programa.
    - En modo EXE COMPILADO (PyInstaller, sys.frozen=True): NUNCA intenta
      pip — dentro de un .exe congelado no hay intérprete de Python ni pip
      accesibles (sys.executable apunta al propio .exe), así que un intento
      de instalar ahí solo produce un error confuso. En su lugar avisa que
      el .exe necesita reconstruirse con la librería ya incluida.

    Devuelve True si al terminar el módulo queda disponible para usarse.
    Se hace lazy (no al arrancar la app) para no demorar el inicio ni pedir
    permiso de red sin que el usuario haya pedido explícitamente un informe.
    """
    global _DOCX_OK, Document, Inches, Pt, RGBColor, WD_ALIGN_PARAGRAPH, WD_TABLE_ALIGNMENT
    if _DOCX_OK:
        return True

    if getattr(sys, "frozen", False):
        messagebox.showerror(
            "Falta un componente en este .exe",
            "Esta versión compilada no incluye la librería 'python-docx', "
            "necesaria para generar el informe en Word.\n\n"
            "No se puede instalar dentro de un .exe ya compilado: hay que "
            "reconstruirlo incluyéndola. En la máquina donde se compila:\n\n"
            "    pip install python-docx\n"
            "    pyinstaller --collect-data docx  [...resto de tus opciones...]\n\n"
            "El flag --collect-data docx es obligatorio: python-docx trae "
            "una plantilla interna (default.docx) que PyInstaller no "
            "detecta solo; sin ese flag el .exe arranca pero falla al "
            "crear el primer documento.",
            parent=parent)
        return False

    if not messagebox.askyesno(
            "Falta un componente",
            "Para generar el informe en Word se necesita la librería "
            "'python-docx', que no está instalada.\n\n"
            "¿Instalarla ahora automáticamente?\n"
            "(requiere conexión a internet; toma unos segundos)",
            parent=parent):
        return False

    try:
        proc = subprocess.run(
            [sys.executable, "-m", "pip", "install", "--quiet", "--disable-pip-version-check",
             "python-docx"],
            capture_output=True, text=True, timeout=120)
    except Exception as e:
        messagebox.showerror(
            "Error de instalación",
            "No se pudo ejecutar pip automáticamente:\n" + str(e) +
            "\n\nInstálalo manualmente abriendo una terminal y ejecutando:\n"
            "    pip install python-docx", parent=parent)
        return False

    if proc.returncode != 0:
        detalle = (proc.stderr or proc.stdout or "Error desconocido")[-800:]
        messagebox.showerror(
            "Error de instalación",
            "La instalación automática de python-docx falló:\n\n" + detalle +
            "\n\nInstálalo manualmente con:\n    pip install python-docx",
            parent=parent)
        return False

    try:
        from docx import Document as _Document
        from docx.shared import Inches as _Inches, Pt as _Pt, RGBColor as _RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH as _WDA
        from docx.enum.table import WD_TABLE_ALIGNMENT as _WDT
    except ImportError as e:
        messagebox.showerror(
            "Error",
            "pip terminó sin errores pero el módulo python-docx sigue sin "
            "poder importarse:\n" + str(e), parent=parent)
        return False

    Document, Inches, Pt, RGBColor = _Document, _Inches, _Pt, _RGBColor
    WD_ALIGN_PARAGRAPH, WD_TABLE_ALIGNMENT = _WDA, _WDT
    _DOCX_OK = True
    messagebox.showinfo("Listo", "python-docx se instaló correctamente.", parent=parent)
    return True


# Citas a la guía que versiones anteriores incrustaban en Estenosis_texto
# (y por eso reaparecían en "4. IMPRESIÓN DIAGNÓSTICA" al reabrir un .ncu
# o al exportar un análisis generado con un .exe viejo).
_CITA_GUIA_ESTENOSIS_PATTERNS = (
    r"\s*según los criterios de velocidad Doppler establecidos en la "
    r"Guía Técnica de Procedimiento del INCN\s*\(CPMS\s*93320\.01\)\.?",
    r"\s*según los criterios(?: de velocidad Doppler)? establecidos en "
    r"la Guía Técnica(?: de Procedimiento)?(?: del INCN)?(?:\s*\([^)]*\))?\.?",
    r"\s*según la Guía Técnica de Procedimiento(?: del INCN)?(?:\s*\([^)]*\))?\.?",
)


def _limpiar_cita_guia_estenosis(texto):
    """Quita la cita a la guía INCN del párrafo de impresión diagnóstica."""
    if not texto:
        return texto
    import re as _re
    limpio = str(texto)
    for pat in _CITA_GUIA_ESTENOSIS_PATTERNS:
        limpio = _re.sub(pat, "", limpio, flags=_re.IGNORECASE)
    limpio = _re.sub(r"[ \t]+", " ", limpio).strip()
    if limpio and limpio[-1] not in ".!?":
        limpio += "."
    return limpio


# ══════════════════════════════════════════════════════════════════════════════
# INFORME DOPPLER TRANSCRANEAL BILATERAL (ACM + ACA + ACP, ambos lados)
#   Requerido para el procedimiento facturado CPMS 93320.01 (auditoría SIS
#   INCN, sep-2026). La Guía Técnica (RD N° 253-2025-DG-INCN, 6.3.3-B) pide
#   insonar ACM, ACA, ACP (y bifurcación de ACI) de ambos lados. Si una
#   arteria no se visualiza, se informa como "no se logra visualizar" con
#   su motivo. El VNO NO forma parte de este informe.
# ══════════════════════════════════════════════════════════════════════════════
ARTERIAS_INFORME = [("ACM", "arteria cerebral media"),
                    ("ACA", "arteria cerebral anterior"),
                    ("ACP", "arteria cerebral posterior")]
LADOS_INFORME = [("D", "Derecho", "derecha", "Der"),
                 ("I", "Izquierdo", "izquierda", "Izq")]
# Profundidades de insonación: Guía Técnica INCN CPMS 93320.01, 6.3.3-B
PROF_GUIA = {"ACM": "40–65", "ACA": "65–80", "ACP": "55–70"}
# Umbrales de estenosis de la guía:
#   ACM (Doppler): VM > 100 moderada, > 128 severa.
#   ACA / ACP: la guía indica que con Doppler NO hay parámetros fijos; solo
#   con Dúplex da VPS > 155 (ACA) y > 145 (ACP) — se usan como criterio
#   orientador y se rotula así en el informe.
UMBRAL_VPS_DUPLEX = {"ACM": 220.0, "ACA": 155.0, "ACP": 145.0}
VENTANA_OPC = ["Adecuada / Buena", "Inadecuada", "Ausente"]
# Coherencia clínica: una arteria NO visualizada no puede tener ventana
# "Adecuada / Buena" (si la ventana es buena, la arteria se ve). Una arteria
# medida no puede tener ventana "Ausente". Por eso las opciones dependen de
# si hay medición.
VENTANA_OPC_MEDIDA    = ["Adecuada / Buena", "Inadecuada"]
VENTANA_OPC_NO_VISTA  = ["Inadecuada", "Ausente"]


def _ventana_coherente(ventana, medida):
    """Corrige combinaciones imposibles (p. ej. .ncu antiguo)."""
    if medida:
        return ventana if ventana in VENTANA_OPC_MEDIDA else "Adecuada / Buena"
    return ventana if ventana in VENTANA_OPC_NO_VISTA else "Inadecuada"


def _motivo_no_vis(ventana):
    return ("ventana acústica temporal ausente" if ventana == "Ausente"
            else "ventana acústica temporal inadecuada")


def _vm_de(data):
    """Velocidad media (TAMV = campo TAmax del equipo)."""
    if not data:
        return None
    v = data.get("Estenosis_VM")
    if v is None:
        v = data.get("TAmax")
    try:
        return float(v) if v not in (None, "") else None
    except (TypeError, ValueError):
        return None


def _fnum(v, dec=1):
    try:
        return f"{float(v):.{dec}f}"
    except (TypeError, ValueError):
        return "—"


CO2_NORMAL = (35, 45)   # mmHg — fuera de rango: avisar (confunde VM / PI / PIC)


def _co2_txt(data):
    """'38 mmHg' / '30 mmHg (hipocapnia)' / '—'. Un solo formato para todas
    las salidas (monitor, captura, figura, panel, texto IA)."""
    v = (data or {}).get("PaCO2")
    if v is None:
        return "—"
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "—"
    s = f"{v:g} mmHg"
    if v < CO2_NORMAL[0]:
        s += " (hipocapnia)"
    elif v > CO2_NORMAL[1]:
        s += " (hipercapnia)"
    return s


def _co2_fuera(data):
    v = (data or {}).get("PaCO2")
    try:
        return v is not None and not (CO2_NORMAL[0] <= float(v) <= CO2_NORMAL[1])
    except (TypeError, ValueError):
        return False


def _pam_txt(data):
    v = (data or {}).get("PAM")
    return "—" if v is None else f"{_fnum(v, 0)} mmHg"


def _renderizar_ambas_acm_imprimible(acm_d, acm_i):
    """Figura imprimible con ACM derecha e izquierda lado a lado."""
    pal = _paleta(True)
    fig, axs = _nuevos_subplots(1, 2, figsize=(14, 4.2), facecolor=pal["BG"])
    for ax, lado, data in ((axs[0], "Derecha", acm_d), (axs[1], "Izquierda", acm_i)):
        OtrasArteriasApp._dibujar_ax(
            None, ax, "ACM", lado, data,
            pal["BG2"], pal["BORDER"], pal["TEXT"], pal["TEXT2"],
            printable=True, texto_sin_dato="No se logra visualizar")
    fig.tight_layout()
    return fig


def _conclusion_tcd(mediciones, ventanas):
    """Conclusión redactada en prosa. mediciones[(art, 'D'|'I')] = data|None."""
    frases = []
    # ACM — criterio Doppler de la guía (VM)
    for k, _, lado_txt, _ in LADOS_INFORME:
        d = mediciones.get(("ACM", k))
        vm = _vm_de(d)
        if not d:
            continue
        if vm is None:
            frases.append(f"En la arteria cerebral media {lado_txt} no se registró "
                          "velocidad media (TAMV), por lo que no se aplica el "
                          "criterio de estenosis.")
        elif vm > 128:
            frases.append(f"La arteria cerebral media {lado_txt} presenta velocidad "
                          f"media de {vm:.1f} cm/s (> 128 cm/s), hallazgo compatible "
                          "con estenosis severa.")
        elif vm > 100:
            frases.append(f"La arteria cerebral media {lado_txt} presenta velocidad "
                          f"media de {vm:.1f} cm/s (entre 100 y 128 cm/s), hallazgo "
                          "compatible con estenosis moderada.")
        else:
            frases.append(f"La arteria cerebral media {lado_txt} presenta velocidad "
                          f"media de {vm:.1f} cm/s, por debajo de 100 cm/s, sin "
                          "evidencia de estenosis significativa.")
    aprox = [lado_txt for k, _, lado_txt, _ in LADOS_INFORME
             if (mediciones.get(("ACM", k)) or {}).get("VM_aproximada")]
    if aprox:
        frases.append("La velocidad media de la arteria cerebral media " +
                      " y ".join(aprox) + " no se registró como TAMV (campo TAmax "
                      "del equipo); se usó un valor sustituto (TAmean o (VPS + 2·VFD)/3), "
                      "por lo que debe interpretarse con cautela.")
    # ACA / ACP — VPS vs umbral dúplex (orientador)
    elevadas, normales = [], []
    for art, nom in ARTERIAS_INFORME[1:]:
        for k, _, lado_txt, _ in LADOS_INFORME:
            d = mediciones.get((art, k))
            if not d:
                continue
            ps = d.get("PS")
            try:
                ps = float(ps)
            except (TypeError, ValueError):
                continue
            if ps > UMBRAL_VPS_DUPLEX[art]:
                elevadas.append(f"{nom} {lado_txt} (VPS {ps:.0f} cm/s > "
                                f"{UMBRAL_VPS_DUPLEX[art]:.0f} cm/s)")
            else:
                normales.append(f"{art} {lado_txt}")
    if normales:
        _lst = (", ".join(normales[:-1]) + " y " + normales[-1]) if len(normales) > 1 else normales[0]
        frases.append("Las velocidades de " + _lst +
                      " no alcanzan los umbrales de estenosis.")
    if elevadas:
        frases.append("Se registra velocidad de pico sistólico por encima del umbral "
                      "orientador de estenosis moderada en: " + "; ".join(elevadas) +
                      ". Se sugiere correlación con estudio angiográfico.")
    # Asimetría > 30 % (guía)
    asim = []
    for art, nom in ARTERIAS_INFORME:
        vd = _vm_de(mediciones.get((art, "D")))
        vi = _vm_de(mediciones.get((art, "I")))
        if vd and vi:
            dif = abs(vd - vi) / max(vd, vi) * 100.0
            if dif > 30:
                asim.append(f"{nom} ({dif:.0f} %, mayor a "
                            f"{'derecha' if vd > vi else 'izquierda'})")
    if asim:
        frases.append("Existe asimetría de velocidades medias entre ambos lados "
                      "superior al 30 % en: " + "; ".join(asim) + ".")
    # IP fuera de rango
    ip_alto = []
    for art, nom in ARTERIAS_INFORME:
        for k, _, lado_txt, _ in LADOS_INFORME:
            d = mediciones.get((art, k))
            if d and d.get("PI") is not None:
                if float(d["PI"]) > DopplerACMApp.RANGOS_ARTERIA[art][7]:
                    ip_alto.append(f"{art} {lado_txt} ({float(d['PI']):.2f})")
    if ip_alto:
        frases.append("Índice de pulsatilidad elevado en " + ", ".join(ip_alto) +
                      ", patrón de aumento de resistencia distal que requiere "
                      "correlación clínica.")
    # Arterias no visualizadas
    no_vis = []
    for art, nom in ARTERIAS_INFORME:
        for k, _, lado_txt, _ in LADOS_INFORME:
            if not mediciones.get((art, k)):
                no_vis.append(f"{nom} {lado_txt} ({_motivo_no_vis(ventanas.get((art, k)))})")
    if no_vis:
        frases.append("No se logra visualizar: " + "; ".join(no_vis) + ".")
    else:
        frases.append("Se lograron insonar las seis arterias evaluadas "
                      "(ACM, ACA y ACP de ambos lados).")
    if not any(mediciones.values()):
        frases = ["No se logró obtener señal Doppler en ninguna de las arterias "
                  "evaluadas."] + frases[-1:]
    return " ".join(frases)


def _construir_informe_tcd_bilateral_docx(pac, hc, mediciones, ventanas,
                                          imagen_path):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)
    NEGRO = RGBColor(0, 0, 0)

    def _par(texto="", bold=False, size=10.5, align=None, space_after=4):
        p = doc.add_paragraph()
        if align is not None:
            p.alignment = align
        p.paragraph_format.space_after = Pt(space_after)
        run = p.add_run(texto)
        run.bold = bold
        run.font.size = Pt(size)
        run.font.name = "Calibri"
        run.font.color.rgb = NEGRO
        return p

    def _heading(texto, space_before=8):
        p = _par(texto, bold=True, size=11.5, space_after=3)
        p.paragraph_format.space_before = Pt(space_before)
        return p

    def _cell(c, texto, bold=False, size=8.5, center=True):
        c.text = ""
        p = c.paragraphs[0]
        if center:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(texto)
        r.bold = bold
        r.font.size = Pt(size)
        r.font.name = "Calibri"

    _par("MINISTERIO DE SALUD  ·  INSTITUTO NACIONAL DE CIENCIAS NEUROLÓGICAS",
         bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=0)
    _par("Departamento de Emergencia, Unidad de Cuidados Intensivos",
         size=9, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    _par("INFORME DE ECOGRAFÍA DOPPLER TRANSCRANEAL",
         bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)

    if imagen_path and os.path.isfile(imagen_path):
        doc.add_picture(imagen_path, width=Inches(7.0))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 1. Datos
    _heading("1.  DATOS DEL PACIENTE", space_before=2)
    t = doc.add_table(rows=1, cols=4)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    for c, (txt, b) in zip(t.rows[0].cells, [("Paciente:", True), (pac or "—", False),
                                             ("HC / ID:", True), (hc or "—", False)]):
        _cell(c, txt, bold=b, size=10, center=False)
    r2 = t.add_row().cells
    for c, (txt, b) in zip(r2, [("Fecha y hora:", True),
                                (datetime.now().strftime("%d/%m/%Y  %H:%M"), False),
                                ("Estudio:", True), ("Doppler transcraneal bilateral", False)]):
        _cell(c, txt, bold=b, size=10, center=False)

    # 2. Técnica
    _heading("2.  TÉCNICA")
    _par("Se realizó estudio de ultrasonografía Doppler transcraneal con transductor "
         "de 2 MHz a través de la ventana acústica temporal bilateral, evaluando la "
         "arteria cerebral media (profundidad 40–65 mm), la arteria cerebral anterior "
         "(65–80 mm) y la arteria cerebral posterior (55–70 mm) de ambos lados.")

    def _anchos_tabla(tabla, anchos):
        tabla.autofit = False
        for _col, _w in zip(tabla.columns, anchos):
            _col.width = Inches(_w)
        for _row in tabla.rows:
            for _c, _w in zip(_row.cells, anchos):
                _c.width = Inches(_w)

    def _cell_ventana(c, ventana):
        c.text = ""
        for i, opc in enumerate(VENTANA_OPC):
            p = c.paragraphs[0] if i == 0 else c.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
            marca = "[X]" if opc == ventana else "[ ]"
            r = p.add_run(f"{marca} {opc}")
            r.bold = (opc == ventana)
            r.font.size = Pt(8)
            r.font.name = "Calibri"

    # 3. Hallazgos — tabla por arteria, con ventana acústica de cada una
    _heading("3.  HALLAZGOS")
    cab = ["Arteria", "Lado", "Prof. (mm)", "Ventana acústica",
           "VPS", "VFD", "VM", "IP", "IR"]
    tb = doc.add_table(rows=1, cols=len(cab))
    tb.style = "Table Grid"
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c, h in zip(tb.rows[0].cells, cab):
        _cell(c, h, bold=True)
    for art, _ in ARTERIAS_INFORME:
        for k, lado_full, _, _ in LADOS_INFORME:
            d = mediciones.get((art, k))
            ven = _ventana_coherente(ventanas.get((art, k)), bool(d))
            row = tb.add_row().cells
            _cell(row[0], art, bold=True)
            _cell(row[1], lado_full)
            _cell(row[2], PROF_GUIA[art])
            _cell_ventana(row[3], ven)
            if d:
                _cell(row[4], _fnum(d.get("PS")))
                _cell(row[5], _fnum(d.get("ED")))
                _cell(row[6], ("~" if d.get("VM_aproximada") else "") + _fnum(_vm_de(d)))
                _cell(row[7], _fnum(d.get("PI"), 2))
                _cell(row[8], _fnum(d.get("RI"), 2))
            else:
                m = row[4].merge(row[8])
                _cell(m, f"No se logra visualizar — {_motivo_no_vis(ven)}", size=8)
    _anchos_tabla(tb, [0.65, 0.85, 0.75, 1.45, 0.6, 0.6, 0.6, 0.55, 0.55])
    _par("Velocidades en cm/s. VPS: velocidad de pico sistólico; VFD: velocidad de "
         "final diastólico; VM: velocidad media (TAMV); ~VM: valor sustituto, "
         "TAMV no registrada; IP: índice de pulsatilidad "
         "de Gosling; IR: índice de resistencia de Pourcelot.", size=7.5, space_after=2)

    # 4. Valores de referencia — tabla aparte
    _heading("4.  VALORES DE REFERENCIA")
    cab2 = ["Arteria", "Prof. (mm)", "VPS", "VFD", "VM", "IP",
            "Umbral de estenosis"]
    umbral = {"ACM": "VM > 100 moderada · > 128 severa",
              "ACA": "No fijo con Doppler · dúplex VPS > 155",
              "ACP": "No fijo con Doppler · dúplex VPS > 145"}
    tr = doc.add_table(rows=1, cols=len(cab2))
    tr.style = "Table Grid"
    tr.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c, h in zip(tr.rows[0].cells, cab2):
        _cell(c, h, bold=True)
    for art, _ in ARTERIAS_INFORME:
        r = DopplerACMApp.RANGOS_ARTERIA[art]
        row = tr.add_row().cells
        for c, t in zip(row, [art, PROF_GUIA[art], f"{r[0]}–{r[1]}", f"{r[2]}–{r[3]}",
                              f"{r[4]}–{r[5]}", f"{r[6]}–{r[7]}", umbral[art]]):
            _cell(c, t, bold=(t == art))
    _anchos_tabla(tr, [0.65, 0.8, 0.8, 0.8, 0.8, 0.8, 2.55])
    _par("Rangos normales de VPS, VFD, VM e IP: referencia bibliográfica (la Guía "
         "Técnica no establece valores normales). Profundidades, "
         "umbrales de estenosis y asimetría de VM entre lados > 30 %: Guía Técnica "
         "INCN. En ACA/ACP el criterio dúplex se usa como orientador.",
         size=7.5, space_after=2)

    # 5. Conclusión
    _heading("5.  CONCLUSIÓN")
    _par(_conclusion_tcd(mediciones, ventanas), size=10.5)

    doc.add_paragraph().paragraph_format.space_before = Pt(16)
    _par("_" * 40, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=0)
    _par("Médico responsable", align=WD_ALIGN_PARAGRAPH.CENTER, size=10, space_after=0)
    _par("Informe elaborado conforme a los lineamientos de la Guía Técnica de "
         "Procedimiento: \u201cEcocardiografía Doppler, Transcraneal\u201d "
         "(RD N° 253-2025-DG-INCN, INCN).",
         size=8, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=0)
    return doc


# ── Paleta clínica ────────────────────────────────────────────────────────────
#
#   Sistema de color inspirado en monitores de paciente clínicos.
#   Principio: fondo negro absoluto, colores de dato saturados al máximo,
#   codificación semántica estricta — cada color tiene un significado fijo.
#
#   BG       → negro puro de pantalla OLED/TFT de monitor
#   BG2      → superficie elevada (paneles, headers) — gris carbón muy oscuro
#   BG3      → campos de entrada, celdas de dato — gris oscuro medio
#   BORDER   → separadores de canal y bordes de campo — casi invisibles
#   ACCENT   → cian eléctrico: SpO₂, links, foco activo (≡ canal SpO₂ del monitor)
#   GREEN    → verde néon: ECG / FC / estado OK (≡ canal ECG del monitor)
#   YELLOW   → amarillo puro: alarma media / RR / advertencia (≡ canal RR del monitor)
#   RED      → rojo brillante: alarma crítica / presión arterial invasiva
#   PINK     → rojo-naranja suave: uso auxiliar (PIC elevada, zona de riesgo)
#   TEXT     → blanco puro: valores numéricos principales
#   TEXT2    → gris claro: etiquetas secundarias, unidades, subtexto
#   PURPLE   → NO se usa en el sistema — reasignado a magenta-ámbar para PA pulmonar
#   PURPLE2  → variante más clara del anterior
#   GRAY     → gris medio-oscuro: elementos inactivos, placeholders, guías
#
BG      = "#000000"   # negro absoluto — fondo de pantalla monitor
BG2     = "#111111"   # gris carbón — header, paneles elevados
BG3     = "#1C1C1C"   # gris oscuro — campos de entrada, celdas
BORDER  = "#2A2A2A"   # gris muy oscuro — separadores de canal
ACCENT  = "#00FFFF"   # cian eléctrico — SpO₂, foco activo
GREEN   = "#00FF00"   # verde néon — ECG / FC / estado normal
YELLOW  = "#FFFF00"   # amarillo puro — RR / alarma media
RED     = "#FF4040"   # rojo alarma — presión invasiva / alarma crítica (AA sobre BG3)
PINK    = "#FF6040"   # rojo-naranja — PIC elevada / zona límite
TEXT    = "#FFFFFF"   # blanco puro — valores numéricos principales
TEXT2   = "#C8C8C8"   # gris claro legible — etiquetas de sección, valores 2°
TEXT3   = "#909090"   # gris medio — unidades, subtexto fino, notas
PURPLE  = "#FFB300"   # ámbar-dorado — PA pulmonar / canal auxiliar
PURPLE2 = "#FFCC44"   # ámbar claro — variante auxiliar
GOLD    = "#D4AF37"   # dorado — botón de informe Doppler TCD
GRAY    = "#333333"   # gris medio — inactivos, placeholders, líneas guía

# ── Paleta "imprimible" (fondo blanco, bajo consumo de tinta) ─────────────────
#   Se usa SOLO al exportar/guardar informes cuando el usuario elige la
#   versión imprimible. La UI en pantalla y la versión "pantalla" de los PNG
#   siguen usando la paleta oscura de arriba sin cambios.
_PALETTE_SCREEN = dict(BG=BG, BG2=BG2, BG3=BG3, BORDER=BORDER,
                        ACCENT=ACCENT, GREEN=GREEN, YELLOW=YELLOW, RED=RED, PINK=PINK,
                        TEXT=TEXT, TEXT2=TEXT2, TEXT3=TEXT3,
                        PURPLE=PURPLE, PURPLE2=PURPLE2, GRAY=GRAY)

_PALETTE_PRINT = dict(
    BG="#FFFFFF", BG2="#FFFFFF", BG3="#F0F0F0", BORDER="#999999",
    ACCENT="#0072B2", GREEN="#1A7A1A", YELLOW="#B8860B", RED="#B00000", PINK="#C1440E",
    TEXT="#000000", TEXT2="#3A3A3A", TEXT3="#606060",
    PURPLE="#9A6300", PURPLE2="#8A5A00", GRAY="#777777",
)


def _paleta(printable=False):
    """Devuelve la paleta a usar: oscura (pantalla) o clara (imprimible/ahorro tinta)."""
    return _PALETTE_PRINT if printable else _PALETTE_SCREEN


# Mapa de traducción de colores: cualquier color de la paleta de PANTALLA que
# venga precalculado (p.ej. pic_color, lind_col) se convierte a su equivalente
# oscuro/legible de la paleta IMPRIMIBLE. Así ningún verde/amarillo/cian
# brillante se cuela al informe de fondo blanco (ilegible en blanco y negro).
_SCREEN_TO_PRINT = {
    _PALETTE_SCREEN["BG"]:     _PALETTE_PRINT["BG"],
    _PALETTE_SCREEN["BG2"]:    _PALETTE_PRINT["BG2"],
    _PALETTE_SCREEN["BG3"]:    _PALETTE_PRINT["BG3"],
    _PALETTE_SCREEN["BORDER"]: _PALETTE_PRINT["BORDER"],
    _PALETTE_SCREEN["ACCENT"]: _PALETTE_PRINT["ACCENT"],
    _PALETTE_SCREEN["GREEN"]:  _PALETTE_PRINT["GREEN"],
    _PALETTE_SCREEN["YELLOW"]: _PALETTE_PRINT["YELLOW"],
    _PALETTE_SCREEN["RED"]:    _PALETTE_PRINT["RED"],
    _PALETTE_SCREEN["PINK"]:   _PALETTE_PRINT["PINK"],
    _PALETTE_SCREEN["TEXT"]:   _PALETTE_PRINT["TEXT"],
    _PALETTE_SCREEN["TEXT2"]:  _PALETTE_PRINT["TEXT2"],
    _PALETTE_SCREEN["TEXT3"]:  _PALETTE_PRINT["TEXT3"],
    _PALETTE_SCREEN["PURPLE"]: _PALETTE_PRINT["PURPLE"],
    _PALETTE_SCREEN["PURPLE2"]:_PALETTE_PRINT["PURPLE2"],
    _PALETTE_SCREEN["GRAY"]:   _PALETTE_PRINT["GRAY"],
    "#FF2020":                 _PALETTE_PRINT["RED"],   # rojo alterno usado en algunos módulos
}


def _map_color(color, printable):
    """Traduce un color de pantalla a su equivalente imprimible si aplica."""
    if not printable or color is None:
        return color
    try:
        return _SCREEN_TO_PRINT.get(str(color).upper(),
                                    _SCREEN_TO_PRINT.get(str(color), color))
    except Exception:
        return color

# ── Tipografía clínica ────────────────────────────────────────────────────────
#
#   SISTEMA TIPOGRÁFICO CLÍNICO — reglas estrictas:
#
#   1. UNA SOLA FAMILIA de display: sans-serif técnica de palo seco.
#      El estándar es una propietaria condensada; el equivalente más fiel
#      disponible en Windows es "Segoe UI" (humanista, limpia, densa).
#      "Consolas" SOLO para datos numéricos tabulares donde la alineación
#      monospace importa (tablas, campos de entrada de cifras).
#
#   2. BOLD se reserva para JERARQUÍA: valores numéricos principales,
#      etiquetas de parámetro (ECG/SpO2), títulos de sección y acciones
#      primarias. El texto corrido de interfaz va en REGULAR.
#
#   3. ITALIC ESTÁ PROHIBIDO. El sistema no usa cursiva.
#      Las notas/aclaraciones van en regular pequeño, nunca en italic.
#
#   4. JERARQUÍA DE TAMAÑO BRUTAL: el salto entre un valor grande y su
#      etiqueta es abrupto, no progresivo. El número manda.
#
#   ── Familias ──
_FAM_UI   = "Segoe UI"   # interfaz, etiquetas, títulos (display principal)
_FAM_DATA = "Consolas"   # datos numéricos tabulares, campos de cifra
#
#   ── Roles ──
F_TITLE  = (_FAM_UI,   14, "bold")    # título principal de ventana — header compacto, no dominante
F_HEAD   = (_FAM_UI,   11, "bold")    # encabezados de sección
F_BODY   = (_FAM_UI,   10)            # texto de interfaz general (regular)
F_MONO   = (_FAM_DATA, 10)            # datos numéricos, campos de entrada
F_SMALL  = (_FAM_UI,    8)            # subtexto, notas al pie (regular)
F_LABEL  = (_FAM_UI,   12, "bold")    # etiquetas de parámetro (ECG/SpO2 style)
F_STATUS = (_FAM_UI,    9)            # barra de estado (regular)
#
#   ── Roles de DATO numérico (se usa monospace bold para cifras) ──
F_VALUE_XL = (_FAM_DATA, 28, "bold")  # valor numérico dominante (IL, etc.)
F_VALUE_M  = (_FAM_DATA, 12, "bold")  # valor numérico medio (escalas, dato chico)
F_DATA     = (_FAM_DATA, 10)          # dato tabular regular
#
#   ── Etiquetas de CANAL (estilo "ECG"/"SpO2" del monitor) ──
#   Las etiquetas de canal van en la familia de DISPLAY (sans-serif)
#   bold, NO en la monospace de datos. Tamano grande pero subordinado al valor.
F_CHANNEL   = (_FAM_UI, 15, "bold")   # etiqueta de canal grande (ACM DERECHA)
F_CHANNEL_S = (_FAM_UI, 11, "bold")   # etiqueta de canal mediana / estado


def _mostrar_creditos(parent=None):
    """Ventana de créditos del SOFTWARE, no del informe clínico."""
    autor = _autor_app()
    win = tk.Toplevel(parent) if parent is not None else tk.Toplevel()
    win.title("Créditos de la aplicación")
    win.configure(bg=BG)
    win.resizable(False, False)
    if parent is not None:
        try:
            win.transient(parent)
        except Exception:
            pass
    tk.Label(win, text="Ultrasonografía Neurocrítica",
             bg=BG, fg=ACCENT, font=F_TITLE).pack(padx=28, pady=(22, 4))
    tk.Label(win, text="Autor de la aplicación",
             bg=BG, fg=TEXT2, font=F_SMALL).pack()
    tk.Label(win, text=autor,
             bg=BG, fg=TEXT, font=F_HEAD).pack(pady=(6, 12))
    tk.Label(win,
             text=("Este crédito identifica al autor del software.\n"
                   "Los informes, casos .ncu y reportes clínicos\n"
                   "pertenecen al médico que los elabora y firma."),
             bg=BG, fg=TEXT2, font=F_BODY, justify="center").pack(padx=28)
    tk.Label(win, text="Build  " + _BUILD_ID,
             bg=BG, fg=TEXT3, font=F_SMALL).pack(pady=(12, 4))
    tk.Button(win, text="Cerrar", command=win.destroy,
              bg=BG3, fg=TEXT, font=F_BODY, relief="flat",
              padx=16, pady=6, cursor="hand2").pack(pady=(8, 18))
    win.update_idletasks()
    try:
        if parent is not None:
            x = parent.winfo_rootx() + (parent.winfo_width() - win.winfo_width()) // 2
            y = parent.winfo_rooty() + (parent.winfo_height() - win.winfo_height()) // 2
            win.geometry("+%d+%d" % (max(x, 0), max(y, 0)))
    except Exception:
        pass

# ── Umbrales VNO (diámetro de vaina del nervio óptico, mm) ───────────────────
#   < 5.0       → normal
#   5.0 – 5.9   → zona gris / limítrofe (hallazgo no concluyente, correlacionar)
#   ≥ 6.0       → dilatada (sugestiva de PIC elevada)
VNO_NORMAL_MAX  = 5.0   # límite superior de normalidad
VNO_ELEVADA_MIN = 6.0   # umbral de dilatación franca


def _vno_cat(val):
    """Clasifica un diámetro de VNO (mm) en zona normal / gris / elevada.

    Devuelve (clave, etiqueta, color_ui, color_png). Centraliza el criterio
    para no duplicarlo entre la UI Tk y las figuras matplotlib.
    """
    if val < VNO_NORMAL_MAX:
        return ("normal",  "Normal",    GREEN,  "#00FF00")
    if val < VNO_ELEVADA_MIN:
        return ("gris",    "Limítrofe", YELLOW, "#FFFF00")
    return     ("elevada", "Dilatada",  RED,    "#FF2020")


def _vno_global(od, oi):
    """Categoría conjunta de ambos ojos: prevalece la zona más severa."""
    orden = {"normal": 0, "gris": 1, "elevada": 2}
    c_od, c_oi = _vno_cat(od), _vno_cat(oi)
    return c_od if orden[c_od[0]] >= orden[c_oi[0]] else c_oi


# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 1b — PERSISTENCIA DE CASOS  (formato binario .ncu)
# ══════════════════════════════════════════════════════════════════════════════
#   Formato propio para guardar/abrir un caso (paciente + VNO + TCD bilateral).
#   pickle binario → ilegible en bloc de notas, sin dependencias externas.
#   No es cifrado: solo evita la edición casual en texto plano.
#   El archivo vive JUNTO al ejecutable/script (no dentro del exe), por lo que
#   funciona igual en desarrollo y compilado con PyInstaller.

CASO_MAGIC   = b"NCU\x01"   # firma de 4 bytes — identifica el formato propio
CASO_VERSION = 1


def _base_dir():
    """Directorio del ejecutable (compilado) o del script (desarrollo)."""
    if getattr(sys, "frozen", False):          # PyInstaller: exe real
        return os.path.dirname(os.path.abspath(sys.executable))
    argv0 = sys.argv[0] if sys.argv and sys.argv[0] else __file__
    d = os.path.dirname(os.path.abspath(argv0))
    return d or os.path.expanduser("~")


def _carpeta_salida(sub="informes_tcd"):
    """Subcarpeta GARANTIZADA escribible para auto-guardar informes.

    NUNCA usa os.getcwd(): cuando la app se abre por asociación de archivo,
    'Ejecutar como administrador' o desde un acceso directo, el CWD puede ser
    C:\\Windows\\System32 (no escribible) → [Errno 13] Permission denied.
    Prueba candidatos en orden y devuelve el primero donde SÍ puede escribir
    (con verificación real de escritura, no solo creación de carpeta).
    """
    home = os.path.expanduser("~")
    bases = []
    if home and home != "~":
        bases.append(os.path.join(home, "Documents", "UltrasonografiaNeurocritica"))
        bases.append(os.path.join(home, "UltrasonografiaNeurocritica"))
    try:
        bases.append(_base_dir())
    except Exception:
        pass
    bases.append(os.path.join(tempfile.gettempdir(), "UltrasonografiaNeurocritica"))
    for base in bases:
        try:
            destino = os.path.join(base, sub)
            os.makedirs(destino, exist_ok=True)
            probe = os.path.join(destino, ".probe")
            with open(probe, "w") as _f:
                _f.write("")
            os.remove(probe)
            return destino
        except OSError:
            continue
    return tempfile.gettempdir()


def _ruta_config(nombre="ia_settings.json"):
    """Ruta escribible y PERSISTENTE para configuración (p. ej. ajustes de IA).

    Usa %APPDATA%/UltrasonografiaNeurocritica en Windows (o ~/.config en otros).
    Evita el directorio de __file__: bajo PyInstaller onefile ese directorio es
    temporal (_MEIxxxx) y se borra al cerrar, por lo que la config no persiste.
    """
    appdata = os.environ.get("APPDATA")           # Windows
    if appdata:
        base = os.path.join(appdata, "UltrasonografiaNeurocritica")
    else:                                         # macOS / Linux
        base = os.path.join(os.path.expanduser("~"), ".config",
                            "UltrasonografiaNeurocritica")
    try:
        os.makedirs(base, exist_ok=True)
    except OSError:
        base = _carpeta_salida("config")          # fallback garantizado escribible
    return os.path.join(base, nombre)


def guardar_caso(data, path):
    """Serializa el caso a un .ncu con cabecera de firma + versión.

    `app` es metadato del SOFTWARE (autor de la app). No sustituye
    al médico responsable del caso clínico contenido en `data`.
    """
    payload = {"version": CASO_VERSION, "data": data, "app": _meta_app()}
    # Escritura atómica: se escribe a un temporal en la MISMA carpeta y luego
    # se reemplaza. Si la app se cae o se corta la luz a mitad de guardar, el
    # .ncu anterior queda intacto en vez de quedar truncado e ilegible.
    tmp = path + ".tmp"
    with open(tmp, "wb") as f:
        f.write(CASO_MAGIC)
        pickle.dump(payload, f, protocol=pickle.HIGHEST_PROTOCOL)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, path)


class _UnpicklerCaso(pickle.Unpickler):
    """pickle.load() normal ejecuta código arbitrario si el archivo fue
    manipulado: la firma 'NCU\x01' no lo impide, cualquiera puede copiarla.
    Este unpickler solo acepta los tipos de datos que un caso realmente
    contiene (contenedores básicos, fechas y escalares/arreglos numpy) y
    rechaza cualquier otra clase o función."""
    _PERMITIDOS = {
        ("builtins", "dict"), ("builtins", "list"), ("builtins", "tuple"),
        ("builtins", "set"), ("builtins", "frozenset"), ("builtins", "str"),
        ("builtins", "int"), ("builtins", "float"), ("builtins", "bool"),
        ("builtins", "complex"), ("builtins", "bytes"), ("builtins", "bytearray"),
        ("collections", "OrderedDict"),
        ("datetime", "datetime"), ("datetime", "date"),
        ("datetime", "time"), ("datetime", "timedelta"),
        ("numpy", "dtype"), ("numpy", "ndarray"),
        ("numpy.core.multiarray", "scalar"), ("numpy._core.multiarray", "scalar"),
        ("numpy.core.multiarray", "_reconstruct"),
        ("numpy.core.numeric", "_frombuffer"),
        ("numpy._core.numeric", "_frombuffer"),
        ("numpy._core.multiarray", "_reconstruct"),
    }

    def find_class(self, module, name):
        if (module, name) in self._PERMITIDOS:
            return super().find_class(module, name)
        raise pickle.UnpicklingError(
            "El caso contiene un tipo no permitido (" + module + "." + name +
            "). Por seguridad no se abre.")


def cargar_caso(path):
    """Lee un .ncu y devuelve el dict de datos.

    Rechaza con ValueError cualquier archivo cuya firma no coincida, para no
    intentar deserializar contenido ajeno (pickle solo debe abrir lo propio).
    """
    with open(path, "rb") as f:
        if f.read(len(CASO_MAGIC)) != CASO_MAGIC:
            raise ValueError("Archivo no reconocido (no es un caso .ncu válido)")
        payload = _UnpicklerCaso(f).load()
    if not isinstance(payload, dict) or "data" not in payload:
        raise ValueError("Estructura de caso inválida")
    return payload["data"]


# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 2 — MÓDULO DOPPLER TCD (v3.1 — DopplerACMApp)
# ══════════════════════════════════════════════════════════════════════════════

def _apply_ttk_dark_theme(root):
    """
    Aplica el tema ttk oscuro clínico (Combobox, Notebook, etc.) al
    intérprete Tcl de `root`. Es la MISMA paleta que usa DopplerACMApp,
    factorizada aquí para que cualquier ventana (panel principal, panel
    "Otras arterias", etc.) tenga combobox/desplegables oscuros y no el
    estilo nativo claro del sistema operativo.

    Idempotente y barata: puede llamarse varias veces (una por ventana)
    sin efectos secundarios — ttk.Style() es compartido por interprete.
    """
    s = ttk.Style()
    try:
        s.theme_use("clam")
    except Exception:
        pass
    s.configure("TFrame",        background=BG)
    s.configure("TLabel",        background=BG,  foreground=TEXT,  font=F_BODY)
    s.configure("TNotebook",     background=BG,  borderwidth=0)
    s.configure("TNotebook.Tab", background=BG3, foreground=TEXT2, font=F_BODY,
                padding=[14, 6])
    s.configure("TCheckbutton",  background=BG,  foreground=TEXT2, font=F_SMALL)
    s.configure("TCombobox",
                fieldbackground=BG3,
                foreground=TEXT,
                background=BG3,
                selectforeground=TEXT,
                selectbackground=ACCENT,
                bordercolor=BORDER,
                lightcolor=BG3,
                darkcolor=BG3,
                arrowcolor=TEXT,
                font=F_MONO)
    s.map("TCombobox",
          fieldbackground=[("readonly", BG3), ("disabled", BG2)],
          foreground=[("readonly", TEXT), ("disabled", TEXT2)],
          background=[("readonly", BG3), ("active", BG3)],
          bordercolor=[("focus", ACCENT), ("!focus", BORDER)],
          arrowcolor=[("readonly", TEXT), ("disabled", TEXT2)])
    # Lista desplegable (Listbox interno del Combobox) — sin esto queda
    # con el fondo blanco/nativo del sistema operativo.
    root.option_add("*TCombobox*Listbox.background",       BG3)
    root.option_add("*TCombobox*Listbox.foreground",       TEXT)
    root.option_add("*TCombobox*Listbox.selectBackground", ACCENT)
    root.option_add("*TCombobox*Listbox.selectForeground", BG)
    root.option_add("*TCombobox*Listbox.font",             F_MONO)
    root.option_add("*TCombobox*Listbox.relief",           "flat")
    s.map("TNotebook.Tab",
          background=[("selected", BG2)],
          foreground=[("selected", ACCENT)])


class DopplerACMApp:

    # Rangos normales TAmean (cm/s) y PI por arteria
    # (PS_min, PS_max, ED_min, ED_max, TAmean_min, TAmean_max, PI_min, PI_max)
    RANGOS_ARTERIA = {
        "ACM":             (60, 140, 25,  65, 40,  85, 0.55, 1.25),
        "ACA":             (40, 100, 20,  55, 30,  70, 0.55, 1.25),
        "ACP":             (30,  85, 15,  45, 20,  55, 0.55, 1.25),
        "ACI":             (40, 100, 20,  55, 30,  70, 0.55, 1.25),
        "ACE":             (40, 100, 15,  45, 25,  60, 0.55, 1.25),
        "AB":              (30,  85, 15,  45, 25,  60, 0.55, 1.25),
        "AV":              (30,  80, 12,  40, 20,  50, 0.55, 1.25),
        "PICA":            (20,  65, 10,  35, 15,  45, 0.55, 1.25),
        "AICA":            (20,  65, 10,  35, 15,  45, 0.55, 1.25),
        "Sifón Carotídeo": (40, 100, 20,  55, 30,  70, 0.55, 1.25),
        "Oftálmica":       (15,  50,  5,  20, 10,  30, 0.55, 1.35),
    }

    CAMPOS = [
        ("PS  – Pico Sistólico",        "PS",     "cm/s",  60,   140),
        ("ED  – Final Diastólico",       "ED",     "cm/s",  25,    65),
        ("RI  – Índice Resistividad",    "RI",     "",      0.45, 0.75),
        ("PI  – Índice Pulsatilidad",    "PI",     "",      0.55, 1.25),
        ("S/D – Sistólico/Diastólico",   "SD",     "",      1.4,   3.8),
        ("TAmax  [≡ TAMV equipo = VM]",      "TAmax",  "cm/s",  40,    85),
        ("TAmean [no usar Czos./IL]",   "TAmean", "cm/s",  40,    85),
        ("HR  – Frecuencia Cardíaca",    "HR",     "lpm",   45,   120),
    ]

    # ── Init ──────────────────────────────────────────────────────────────────
    def __init__(self, root, on_result=None):
        self.root = root
        self.root.title("Doppler Transcraneal ACM  [" + _BUILD_ID + "]")
        self.root.geometry("1220x980")
        self.root.configure(bg=BG)
        self.root.resizable(True, True)
        self.on_result  = on_result   # callback(dict) → NeurocriticalApp
        self.last_result = None

        self.canvas_fig = None
        self.last_fig   = None
        self._lind_il_result  = None
        self._lind_tamax_acm  = None
        self._pi_autocalc     = False
        self._informe_texto   = ""
        self._estenosis_sev   = None
        self._estenosis_texto = ""
        self._estenosis_vm    = None
        # Estado de voz — declarado ANTES de construir el notebook: _tab_opciones
        # lo redefine, pero si esa pestaña fallara al construirse, _edge_tts_run
        # seguiría encontrando los atributos y no reventaría con AttributeError.
        self._edge_proc       = None
        self._edge_pygame     = None
        self._edge_stop       = False

        # Los hilos (IA, voz) NO deben tocar Tk directamente: Tkinter no es
        # seguro entre hilos. Encolan funciones y el hilo principal las
        # ejecuta en _drenar_cola_ui().
        self._cola_ui      = queue.Queue()
        self._after_tick   = None
        self._after_cola   = None
        self._cerrada      = False
        self.root.bind("<Destroy>", self._on_destroy, add="+")

        self._apply_styles()
        self._build_header()
        self._build_notebook()
        self._build_statusbar()

    # ── Estilos ttk ───────────────────────────────────────────────────────────
    def _apply_styles(self):
        _apply_ttk_dark_theme(self.root)

    # ── Header ────────────────────────────────────────────────────────────────
    def _build_header(self):
        h = tk.Frame(self.root, bg=BG2, height=64)
        h.pack(fill=tk.X)
        h.pack_propagate(False)
        tk.Label(h, text="⬡  DOPPLER TRANSCRANEAL",
                 bg=BG2, fg=ACCENT, font=F_TITLE).pack(side=tk.LEFT, padx=22, pady=14)
        tk.Label(h, text="Arteria Cerebral Media  –  Análisis Hemodinámico",
                 bg=BG2, fg=TEXT2, font=F_BODY).pack(side=tk.LEFT)
        self.clock_lbl = tk.Label(h, text="", bg=BG2, fg=TEXT2, font=F_SMALL)
        self.clock_lbl.pack(side=tk.RIGHT, padx=20)
        tk.Button(h, text="ⓘ  Créditos",
                  bg=BG2, fg=TEXT2, font=F_SMALL,
                  relief="flat", bd=0, cursor="hand2",
                  activebackground=BG3, activeforeground=ACCENT,
                  command=lambda: _mostrar_creditos(self.root)
                  ).pack(side=tk.RIGHT, padx=8)
        self._tick()
        self._drenar_cola_ui()
        tk.Frame(self.root, bg=ACCENT, height=2).pack(fill=tk.X)

    def _tick(self):
        if self._cerrada:
            return
        try:
            self.clock_lbl.config(text=datetime.now().strftime("%d/%m/%Y   %H:%M:%S"))
            self._after_tick = self.root.after(1000, self._tick)
        except tk.TclError:          # la ventana ya se destruyó
            self._after_tick = None

    # ── Puente hilo → UI y cierre limpio ──────────────────────────────────────
    def _ui_after(self, fn):
        """Llamable desde cualquier hilo: encola fn para el hilo de Tk."""
        if not self._cerrada:
            self._cola_ui.put(fn)

    def _drenar_cola_ui(self):
        if self._cerrada:
            return
        try:
            while True:
                fn = self._cola_ui.get_nowait()
                try:
                    fn()
                except tk.TclError:
                    pass
        except queue.Empty:
            pass
        try:
            self._after_cola = self.root.after(100, self._drenar_cola_ui)
        except tk.TclError:
            self._after_cola = None

    def _on_destroy(self, event):
        # <Destroy> llega también por cada widget hijo; solo interesa la ventana.
        if event.widget is self.root:
            self.cerrar()

    def cerrar(self):
        """Libera todo lo que podría mantener vivo el proceso o la memoria:
        temporizadores de Tk, reproducción de voz (pygame / PowerShell /
        afplay / mpg123) y la figura de matplotlib. Idempotente."""
        if self._cerrada:
            return
        self._cerrada = True
        for aid in (self._after_tick, self._after_cola):
            if aid:
                try:
                    self.root.after_cancel(aid)
                except tk.TclError:
                    pass
        self._detener_audio()
        try:
            if self.canvas_fig is not None:
                self.canvas_fig.get_tk_widget().destroy()
        except tk.TclError:
            pass
        self.canvas_fig = None
        _cerrar_figura(self.last_fig)
        self.last_fig = None

    # ── Notebook ──────────────────────────────────────────────────────────────
    def _build_notebook(self):
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=14, pady=8)
        self._tab_medicion()
        self._tab_lindegaard()
        self._tab_ia()
        self._tab_opciones()
        self._tab_referencia()

    # =========================================================================
    # PESTAÑA 1 — Medición
    # =========================================================================
    def _tab_medicion(self):
        tab = ttk.Frame(self.nb)
        self.nb.add(tab, text="  📊  Medición  ")

        # ── Panel izquierdo ───────────────────────────────────────────────────
        left = tk.Frame(tab, bg=BG, width=360)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 8))
        left.pack_propagate(False)

        # Datos paciente
        self._seccion(left, "DATOS DEL PACIENTE")
        pf = tk.Frame(left, bg=BG)
        pf.pack(fill=tk.X, pady=2)

        self._lbl(pf, "Nombre:").grid(row=0, column=0, sticky="w", pady=3)
        self.ent_nombre = self._entry(pf, 26)
        self.ent_nombre.grid(row=0, column=1, padx=6, pady=3)

        self._lbl(pf, "ID / HC:").grid(row=1, column=0, sticky="w", pady=3)
        self.ent_hc = self._entry(pf, 26)
        self.ent_hc.grid(row=1, column=1, padx=6, pady=3)

        self._lbl(pf, "Arteria:").grid(row=2, column=0, sticky="w", pady=3)
        self.arteria_var = tk.StringVar(value="ACM")
        self.arteria_cb = ttk.Combobox(pf, textvariable=self.arteria_var,
                     values=["ACM", "ACA", "ACP", "ACI", "ACE",
                             "AB", "PICA", "AICA", "AV",
                             "Sifón Carotídeo", "Oftálmica"],
                     width=18, state="readonly")
        self.arteria_cb.grid(row=2, column=1, sticky="w", padx=6)
        self.arteria_cb.bind("<<ComboboxSelected>>", self._on_arteria_change)

        self._lbl(pf, "Lado:").grid(row=3, column=0, sticky="w", pady=3)
        self.lado_var = tk.StringVar(value="Derecho")
        ttk.Combobox(pf, textvariable=self.lado_var,
                     values=["Derecho", "Izquierdo", "Bilateral"],
                     width=12, state="readonly").grid(row=3, column=1, sticky="w", padx=6)

        # Parámetros Doppler
        self._seccion(left, "PARÁMETROS DOPPLER")
        self.entries       = {}
        self.status_labels = {}
        self.vars          = {}

        ef = tk.Frame(left, bg=BG)
        ef.pack(fill=tk.X, pady=2)

        for texto, clave, unidad, vmin, vmax in self.CAMPOS:
            row = tk.Frame(ef, bg=BG)
            row.pack(fill=tk.X, pady=2)

            tk.Label(row, text=texto, bg=BG, fg=TEXT2,
                     font=F_SMALL, width=26, anchor="w").pack(side=tk.LEFT)

            var = tk.StringVar()
            var.trace_add("write", lambda *a, k=clave: self._on_change(k))
            self.vars[clave] = var

            ent = tk.Entry(row, textvariable=var, width=9,
                           bg=BG3, fg=TEXT, font=F_MONO,
                           relief="flat", insertbackground=TEXT,
                           highlightthickness=1,
                           highlightbackground=BORDER,
                           highlightcolor=ACCENT)
            ent.pack(side=tk.LEFT, padx=4)

            if unidad:
                tk.Label(row, text=unidad, bg=BG, fg=TEXT2,
                         font=F_SMALL, width=5).pack(side=tk.LEFT)

            sl = tk.Label(row, text="", bg=BG, font=F_SMALL, width=11, anchor="w")
            sl.pack(side=tk.LEFT)

            self.entries[clave]       = {"entry": ent, "min": vmin, "max": vmax}
            self.status_labels[clave] = sl

        # Auto-calc
        self.auto_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(left, text="Autocalcular RI / PI / S/D si están vacíos",
                        variable=self.auto_var).pack(anchor="w", pady=4)

        # PAM
        pam_f = tk.Frame(left, bg=BG)
        pam_f.pack(fill=tk.X, pady=4)
        tk.Label(pam_f, text="PAM (mmHg) – opcional:", bg=BG, fg=YELLOW,
                 font=F_SMALL).pack(side=tk.LEFT)
        self.pam_ent = tk.Entry(pam_f, width=8, bg=BG3, fg=YELLOW,
                                font=F_MONO, relief="flat",
                                insertbackground=YELLOW,
                                highlightthickness=1,
                                highlightbackground=BORDER,
                                highlightcolor=YELLOW)
        self.pam_ent.pack(side=tk.LEFT, padx=6)

        # PaCO₂ / EtCO₂ (debajo de la PAM). Referencia 35–45 mmHg: fuera de
        # ese rango, PI / Bellner / Czosnyka / VM pueden estar confundidos
        # por el CO₂ y conviene interpretarlos con cautela.
        co2_f = tk.Frame(left, bg=BG)
        co2_f.pack(fill=tk.X, pady=4)
        tk.Label(co2_f, text="PaCO₂ / EtCO₂ (mmHg) – opcional:", bg=BG,
                 fg=YELLOW, font=F_SMALL).pack(side=tk.LEFT)
        self.co2_ent = tk.Entry(co2_f, width=8, bg=BG3, fg=YELLOW,
                                font=F_MONO, relief="flat",
                                insertbackground=YELLOW,
                                highlightthickness=1,
                                highlightbackground=BORDER,
                                highlightcolor=YELLOW)
        self.co2_ent.pack(side=tk.LEFT, padx=6)

        # Botones
        self._seccion(left, "ACCIONES")
        bf = tk.Frame(left, bg=BG)
        bf.pack(fill=tk.X, pady=4)
        self._btn(bf, "▶  Generar Análisis", self._safe_generar, ACCENT).pack(fill=tk.X, pady=2)
        self._btn(bf, "💾  Guardar PNG (pantalla)",   self._guardar_png,             BG3).pack(fill=tk.X, pady=2)
        self._btn(bf, "🖨  Guardar PNG (imprimible)", self._guardar_png_imprimible,  BG3).pack(fill=tk.X, pady=2)
        self._btn(bf, "🗑  Limpiar",          self._limpiar,      BG3).pack(fill=tk.X, pady=2)

        # ── Panel derecho ─────────────────────────────────────────────────────
        right = tk.Frame(tab, bg=BG)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.graph_frame = tk.Frame(right, bg=BG)
        self.graph_frame.pack(fill=tk.BOTH, expand=True)

        # Placeholder
        self.placeholder = tk.Label(
            self.graph_frame,
            text="La gráfica Doppler aparecerá aquí\n\nIngresa PS y ED y pulsa  ▶  Generar Análisis",
            bg=BG, fg=BORDER, font=(_FAM_UI, 13))
        self.placeholder.place(relx=0.5, rely=0.5, anchor="center")

        # ── Informe clínico exportable (reemplaza la caja de copiar/pegar) ─────
        #   El texto del informe se sigue construyendo internamente
        #   (self._informe_texto) para el IL y otros usos, pero ya no se
        #   muestra en pantalla: en su lugar hay un botón que exporta un
        #   informe DOCX imprimible (imagen + texto en formato de oración),
        #   siguiendo los lineamientos de la Guía Técnica INCN CPMS 93320.01.
        self._informe_texto = ""
        informe_f = tk.Frame(right, bg=BG2)
        informe_f.pack(fill=tk.X, padx=4, pady=(0, 4))

        self.estenosis_lbl = tk.Label(informe_f, text="Estenosis ACM: —",
                                      bg=BG2, fg=TEXT2, font=F_BODY, anchor="w")
        self.estenosis_lbl.pack(side=tk.LEFT, padx=12, pady=10)

        # El informe (bilateral ACM/ACA/ACP) se genera SOLO desde el botón
        # único del dashboard principal — ver NeurocriticalApp.

    # =========================================================================
    # PESTAÑA 2 — Índice de Lindegaard
    # =========================================================================
    def _tab_lindegaard(self):
        self.tab_lind = ttk.Frame(self.nb)
        self.nb.add(self.tab_lind, text="  🧠  Vasoespasmo / IL  ")
        # Deshabilitada hasta que haya duplex válido
        self.nb.tab(1, state="disabled")

        tab = self.tab_lind

        # ── Encabezado informativo ────────────────────────────────────────────
        info = tk.Frame(tab, bg=BG2)
        info.pack(fill=tk.X, padx=16, pady=(14, 4))
        tk.Label(info,
                 text="ÍNDICE DE LINDEGAARD  (IL)  —  Detección de vasoespasmo ACM",
                 bg=BG2, fg=ACCENT, font=F_HEAD).pack(side=tk.LEFT, padx=12, pady=10)

        tk.Label(tab,
                 text="Esta pestaña se activa automáticamente al completar PS y ED en la pestaña Medición.",
                 bg=BG, fg=TEXT2, font=F_SMALL).pack(anchor="w", padx=20, pady=(0, 6))

        # ── Datos del duplex (solo lectura, se llenan desde _generar) ─────────
        self._seccion(tab, "DATOS DEL DUPLEX TRANSCRANEAL (ACM intracraneal)")

        duplex_f = tk.Frame(tab, bg=BG)
        duplex_f.pack(fill=tk.X, padx=20, pady=4)

        self._lbl(duplex_f, "Paciente:").grid(row=0, column=0, sticky="w", pady=4)
        self.lind_info_pac = tk.Label(duplex_f, text="—", bg=BG, fg=TEXT, font=F_MONO)
        self.lind_info_pac.grid(row=0, column=1, sticky="w", padx=10)

        self._lbl(duplex_f, "Arteria evaluada:").grid(row=1, column=0, sticky="w", pady=4)
        self.lind_info_art = tk.Label(duplex_f, text="—", bg=BG, fg=TEXT, font=F_MONO)
        self.lind_info_art.grid(row=1, column=1, sticky="w", padx=10)

        self._lbl(duplex_f, "TAMV ACM (campo TAmax) — SE USA:").grid(row=2, column=0, sticky="w", pady=4)
        self.lind_info_tamax = tk.Label(duplex_f, text="—  cm/s", bg=BG, fg=GREEN, font=F_MONO)
        self.lind_info_tamax.grid(row=2, column=1, sticky="w", padx=10)

        self._lbl(duplex_f, "TAmean equipo (no usado IL):").grid(row=3, column=0, sticky="w", pady=4)
        self.lind_info_tamean = tk.Label(duplex_f, text="—  cm/s", bg=BG, fg=TEXT2, font=F_MONO)
        self.lind_info_tamean.grid(row=3, column=1, sticky="w", padx=10)

        # ── Input ACI extracraneal ────────────────────────────────────────────
        self._seccion(tab, "ACI EXTRACRANEAL  (ventana submandibular — ingresar ahora)")

        aci_f = tk.Frame(tab, bg=BG)
        aci_f.pack(fill=tk.X, padx=20, pady=4)

        self._lbl(aci_f, "VM ACI extracraneal (cm/s):").grid(row=0, column=0, sticky="w", pady=6)
        self.lind_aci_var = tk.StringVar()
        self.lind_aci_ent = tk.Entry(aci_f, textvariable=self.lind_aci_var, width=10,
                                     bg=BG3, fg=TEXT, font=F_MONO, relief="flat",
                                     insertbackground=TEXT,
                                     highlightthickness=1,
                                     highlightbackground=BORDER,
                                     highlightcolor=ACCENT)
        self.lind_aci_ent.grid(row=0, column=1, padx=8, pady=6)
        tk.Label(aci_f, text="cm/s", bg=BG, fg=TEXT2, font=F_SMALL).grid(row=0, column=2, sticky="w")

        tk.Label(aci_f,
                 text="⚑  ACI extracraneal ipsilateral  ·  vent. submandibular  ·  45–50 mm  ·  ingresar TAMV (≡ campo TAmax del equipo)",
                 bg=BG, fg=TEXT2, font=(_FAM_UI, 7)
                 ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 4))

        # El IL usa SIEMPRE la TAMV (campo "TAmax" del equipo) en ambos vasos.
        tk.Label(aci_f,
                 text="ℹ  IL = TAMV ACM / TAMV ACI-extrac.  —  ambos = campo \"TAmax\" del equipo (mean flow velocity)",
                 bg=BG, fg=TEXT2, font=(_FAM_UI, 7)
                 ).grid(row=2, column=0, columnspan=3, sticky="w", pady=4)

        # ── Resultado IL ──────────────────────────────────────────────────────
        self._seccion(tab, "RESULTADO — ÍNDICE DE LINDEGAARD")

        res_f = tk.Frame(tab, bg=BG2, bd=0)
        res_f.pack(fill=tk.X, padx=20, pady=8)

        # IL grande
        self.lind_il_lbl = tk.Label(res_f, text="IL = —",
                                    bg=BG2, fg=TEXT2,
                                    font=F_VALUE_XL)
        self.lind_il_lbl.pack(pady=(16, 4))

        self.lind_interp_lbl = tk.Label(res_f, text="Ingresa VM ACI extracraneal para calcular",
                                        bg=BG2, fg=TEXT2,
                                        font=F_BODY)
        self.lind_interp_lbl.pack(pady=(0, 4))

        self.lind_detalle_lbl = tk.Label(res_f, text="",
                                         bg=BG2, fg=TEXT2,
                                         font=F_SMALL)
        self.lind_detalle_lbl.pack(pady=(0, 16))

        # ── Escala referencia ─────────────────────────────────────────────────
        self._seccion(tab, "ESCALA DE REFERENCIA")

        esc_f = tk.Frame(tab, bg=BG)
        esc_f.pack(fill=tk.X, padx=20, pady=6)

        escalas = [
            ("IL < 3",   "Sin vasoespasmo ACM",         GREEN),
            ("IL 3–6",   "Vasoespasmo ACM probable",     YELLOW),
            ("IL > 6",   "Vasoespasmo ACM SEVERO",       RED),
        ]
        for i, (rng, desc, col) in enumerate(escalas):
            tk.Label(esc_f, text=rng,  bg=BG, fg=col,   font=F_VALUE_M,
                     width=10, anchor="w").grid(row=i, column=0, sticky="w", pady=3, padx=(0,10))
            tk.Label(esc_f, text=desc, bg=BG, fg=TEXT,  font=F_BODY,
                     anchor="w").grid(row=i, column=1, sticky="w", pady=3)

        tk.Label(tab,
                 text="Fuente: Lindegaard KF et al. J Neurosurg 69:677-684 (1988)",
                 bg=BG, fg=TEXT2, font=(_FAM_UI, 7)).pack(anchor="w", padx=20, pady=(4, 2))
        tk.Label(tab,
                 text="Nota: IL discrimina vasoespasmo verdadero de hiperemia global. "
                      "Correlacionar siempre con clínica y angio-TC/RM.",
                 bg=BG, fg=TEXT2, font=(_FAM_UI, 7),
                 wraplength=600, justify="left").pack(anchor="w", padx=20, pady=(0, 8))

        # Trace para cálculo en vivo
        self.lind_aci_var.trace_add(
            "write", lambda *a: (self._calc_lindegaard_live(),
                                 self._sincronizar_il_monitor()))

    # =========================================================================
    # PESTAÑA 3 — IA  (contexto clínico + comentario)
    # =========================================================================
    def _tab_ia(self):
        tab = ttk.Frame(self.nb)
        self.nb.add(tab, text="  🤖  IA  ")

        self._seccion(tab, "CONTEXTO CLÍNICO")
        tk.Label(tab,
                 text="La IA integra estos datos al comentar el Doppler. "
                      "Configura proveedor y API key en  ⚙ Opciones.",
                 bg=BG, fg=TEXT2, font=F_SMALL).pack(anchor="w", padx=24, pady=(0, 8))

        ctx_f = tk.Frame(tab, bg=BG)
        ctx_f.pack(fill=tk.X, padx=24)

        self._lbl(ctx_f, "Diagnóstico / Motivo:").grid(
            row=0, column=0, sticky="nw", pady=4)
        self.ia_ctx_dx = tk.Entry(ctx_f, width=62,
                                  bg=BG3, fg=TEXT, font=F_MONO,
                                  relief="flat", insertbackground=TEXT,
                                  highlightthickness=1,
                                  highlightbackground=BORDER,
                                  highlightcolor=ACCENT)
        self.ia_ctx_dx.grid(row=0, column=1, padx=10, pady=4, sticky="w")
        tk.Label(ctx_f,
                 text="p.ej. HSA Fisher 3, D4 post-clipaje, nimodipino activo",
                 bg=BG, fg=TEXT2, font=F_SMALL).grid(row=1, column=1, sticky="w", padx=10)

        self._lbl(ctx_f, "Hallazgos clínicos:").grid(
            row=2, column=0, sticky="nw", pady=(10, 4))
        self.ia_ctx_hx = tk.Text(ctx_f, width=62, height=4,
                                 bg=BG3, fg=TEXT, font=F_MONO,
                                 relief="flat", insertbackground=TEXT,
                                 wrap="word", padx=6, pady=4,
                                 highlightthickness=1,
                                 highlightbackground=BORDER)
        self.ia_ctx_hx.grid(row=2, column=1, padx=10, pady=4, sticky="w")
        tk.Label(ctx_f,
                 text="p.ej. Glasgow 12, déficit motor izquierdo, TC sin nuevo infarto",
                 bg=BG, fg=TEXT2, font=F_SMALL).grid(row=3, column=1, sticky="w", padx=10)

        bf = tk.Frame(tab, bg=BG)
        bf.pack(anchor="w", padx=24, pady=(10, 4))
        self._btn(bf, "🤖  Generar comentario IA",
                  self._generar_comentario_ia, ACCENT).pack(side=tk.LEFT, padx=4)
        self._btn(bf, "🗑  Limpiar",
                  self._limpiar_contexto_ia, BG3).pack(side=tk.LEFT, padx=4)

        self.ia_comentario_lbl = tk.Label(
            tab, text="", bg=BG, fg=TEXT2,
            font=F_SMALL, wraplength=720, justify="left")
        self.ia_comentario_lbl.pack(anchor="w", padx=24)

        self._seccion(tab, "COMENTARIO")
        self.ia_result_box = tk.Text(tab, height=7, bg=BG2, fg=TEXT,
                                     font=F_MONO, relief="flat",
                                     state="disabled", wrap="word",
                                     padx=12, pady=8)
        self.ia_result_box.pack(fill=tk.X, padx=24, pady=(0, 8))

    # ── Contexto IA helpers ───────────────────────────────────────────────────
    def _limpiar_contexto_ia(self):
        self.ia_ctx_dx.delete(0, tk.END)
        self.ia_ctx_hx.delete("1.0", tk.END)
        self.ia_result_box.config(state="normal")
        self.ia_result_box.delete("1.0", tk.END)
        self.ia_result_box.config(state="disabled")
        self.ia_comentario_lbl.config(text="")

    def _generar_comentario_ia(self):
        prov = self.ia_prov.get()
        key  = self.ia_key.get().strip()
        if prov == "Desactivado":
            self.ia_comentario_lbl.config(
                text="⚠ Activa un proveedor en  ⚙ Opciones.", fg=YELLOW)
            return
        if not key:
            self.ia_comentario_lbl.config(
                text="⚠ Introduce una API key en  ⚙ Opciones.", fg=YELLOW)
            return
        informe = (self._informe_texto or "").strip()
        if not informe:
            self.ia_comentario_lbl.config(
                text="⚠ Genera primero un análisis Doppler.", fg=YELLOW)
            return
        dx  = self.ia_ctx_dx.get().strip()
        hx  = self.ia_ctx_hx.get("1.0", tk.END).strip()
        ctx = ""
        if dx: ctx += f"Diagnóstico / motivo: {dx}\n"
        if hx: ctx += f"Hallazgos clínicos: {hx}\n"
        if not ctx:
            ctx = "(Sin contexto clínico — interpretación paramétrica únicamente)"
        prompt = (
            "Eres un médico intensivista experto en neuromonitorización y Doppler transcraneal. "
            "Se te proporciona el resultado numérico de un Doppler transcraneal y el contexto clínico. "
            "Genera un comentario clínico conciso (máximo 120 palabras) en español, "
            "integrado con el contexto, sin repetir los números del informe. "
            "Si los datos son insuficientes para una conclusión, indícalo. "
            "No inventes hallazgos.\n\n"
            f"=== RESULTADO DOPPLER ===\n{informe}\n\n"
            f"=== CONTEXTO CLÍNICO ===\n{ctx}"
        )
        self.ia_comentario_lbl.config(text="⏳ Consultando IA…", fg=TEXT2)
        self.root.update_idletasks()
        import threading
        threading.Thread(
            target=self._llamar_ia_comentario,
            args=(prov, key, self.ia_model.get(), prompt),
            daemon=True).start()

    def _llamar_ia_raw(self, prov, key, model, prompt, max_tokens=350):
        """Llama al proveedor y devuelve texto. Soporta OpenAI-compat y Anthropic."""
        _PROV_META = {
            # proveedor : (base_url o None, usa_sdk_openai)
            "OpenAI":      (None,                          True),
            "Anthropic":   (None,                          False),
            "Grok (xAI)":  ("https://api.x.ai/v1",        True),
            "Google":      ("https://generativelanguage.googleapis.com/v1beta/openai/", True),
            "DeepSeek":    ("https://api.deepseek.com/v1", True),
            "Mistral":     ("https://api.mistral.ai/v1",   True),
        }
        meta = _PROV_META.get(prov)
        if meta is None:
            raise ValueError(f"Proveedor '{prov}' no reconocido.")
        base_url, use_openai = meta
        if use_openai:
            from openai import OpenAI
            kwargs = {"api_key": key}
            if base_url:
                kwargs["base_url"] = base_url
            r = OpenAI(**kwargs).chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=max_tokens)
            return r.choices[0].message.content.strip()
        else:  # Anthropic nativo
            import anthropic
            r = anthropic.Anthropic(api_key=key).messages.create(
                model=model, max_tokens=max_tokens,
                messages=[{"role": "user", "content": prompt}])
            return r.content[0].text.strip()

    def _llamar_ia_comentario(self, prov, key, model, prompt):
        try:
            resultado = self._llamar_ia_raw(prov, key, model, prompt)
            def _mostrar():
                self.ia_comentario_lbl.config(text="✓ Comentario generado:", fg=GREEN)
                self.ia_result_box.config(state="normal")
                self.ia_result_box.delete("1.0", tk.END)
                self.ia_result_box.insert("1.0", resultado)
                self.ia_result_box.config(state="disabled")
            self._ui_after(_mostrar)
        except Exception as e:
            # Capturar el mensaje AQUÍ: en Python 3 'e' se desvincula al cerrar
            # el except, y este lambda se ejecuta diferido (root.after) cuando 'e'
            # ya no existe → NameError. Guardamos el texto antes.
            msg = "✗ " + str(e)[:120]
            self._ui_after(lambda: self.ia_comentario_lbl.config(
                text=msg, fg=RED))

    # =========================================================================
    # PESTAÑA 4 — Opciones  (API keys + voz)
    # =========================================================================
    def _tab_opciones(self):
        tab = ttk.Frame(self.nb)
        self.nb.add(tab, text="  ⚙  Opciones  ")

        # ── IA ────────────────────────────────────────────────────────────────
        self._seccion(tab, "INTELIGENCIA ARTIFICIAL")
        tk.Label(tab,
                 text="Las claves se guardan en ia_settings.json (solo en tu equipo).",
                 bg=BG, fg=TEXT2, font=F_SMALL).pack(anchor="w", padx=24, pady=(0, 8))

        f = tk.Frame(tab, bg=BG)
        f.pack(fill=tk.X, padx=24)

        # Proveedor
        self._lbl(f, "Proveedor:").grid(row=0, column=0, sticky="w", pady=6)
        self.ia_prov = tk.StringVar(value="Desactivado")
        ttk.Combobox(f, textvariable=self.ia_prov,
                     values=["OpenAI", "Anthropic", "Google",
                             "Grok (xAI)", "DeepSeek", "Mistral", "Desactivado"],
                     width=18, state="readonly").grid(row=0, column=1, sticky="w", padx=10)

        # Hint de modelos (actualizado al cambiar proveedor)
        self.ia_model_hint = tk.Label(f, text="", bg=BG, fg=TEXT2, font=F_SMALL)
        self.ia_model_hint.grid(row=0, column=2, sticky="w", padx=6)

        # Bind para actualizar hint
        self.ia_prov.trace_add("write", self._on_prov_change)

        # API Key
        self._lbl(f, "API Key:").grid(row=1, column=0, sticky="w", pady=6)
        self.ia_key = tk.StringVar()
        self.ia_key_ent = tk.Entry(f, textvariable=self.ia_key, width=52, show="*",
                                   bg=BG3, fg=TEXT, font=F_MONO,
                                   relief="flat", insertbackground=TEXT,
                                   highlightthickness=1,
                                   highlightbackground=BORDER,
                                   highlightcolor=ACCENT)
        self.ia_key_ent.grid(row=1, column=1, padx=10)
        self._btn(f, "👁", self._toggle_key, BG3).grid(row=1, column=2, padx=4)

        # Modelo — Entry libre, sin dropdown
        self._lbl(f, "Modelo:").grid(row=2, column=0, sticky="w", pady=6)
        self.ia_model = tk.StringVar(value="")
        self.ia_model_ent = tk.Entry(f, textvariable=self.ia_model, width=36,
                                     bg=BG3, fg=TEXT, font=F_MONO,
                                     relief="flat", insertbackground=TEXT,
                                     highlightthickness=1,
                                     highlightbackground=BORDER,
                                     highlightcolor=ACCENT)
        self.ia_model_ent.grid(row=2, column=1, sticky="w", padx=10)

        bf_ia = tk.Frame(tab, bg=BG)
        bf_ia.pack(anchor="w", padx=24, pady=(10, 4))
        self._btn(bf_ia, "💾  Guardar",         self._save_ia,  ACCENT).pack(side=tk.LEFT, padx=4)
        self._btn(bf_ia, "📂  Cargar",          self._load_ia,  BG3).pack(side=tk.LEFT, padx=4)
        self._btn(bf_ia, "🧪  Probar conexión", self._test_ia,  BG3).pack(side=tk.LEFT, padx=4)

        self.ia_status = tk.Label(tab, text="", bg=BG, fg=TEXT2, font=F_SMALL)
        self.ia_status.pack(anchor="w", padx=24)

        # ── Voz ───────────────────────────────────────────────────────────────
        self._seccion(tab, "SÍNTESIS DE VOZ  —  Edge TTS")
        tk.Label(tab,
                 text="Voces neurales de Microsoft Edge · requiere internet · sin API key"
                      "  ·  instala:  pip install edge-tts pygame",
                 bg=BG, fg=TEXT2, font=F_SMALL).pack(anchor="w", padx=24, pady=(0, 8))

        vf = tk.Frame(tab, bg=BG)
        vf.pack(fill=tk.X, padx=24)

        self._lbl(vf, "Narración auto:").grid(row=0, column=0, sticky="w", pady=6)
        self.ia_voice = tk.BooleanVar(value=False)
        ttk.Checkbutton(vf, text="Activar narración al generar análisis",
                        variable=self.ia_voice).grid(row=0, column=1, sticky="w", padx=10)

        self._lbl(vf, "Voz:").grid(row=1, column=0, sticky="w", pady=6)
        self.edge_voice = tk.StringVar(value="es-PE-AlexNeural")
        ttk.Combobox(vf, textvariable=self.edge_voice,
                     values=["es-PE-AlexNeural", "es-PE-CamilaNeural",
                             "es-ES-AlvaroNeural", "es-ES-ElviraNeural",
                             "es-MX-JorgeNeural", "es-MX-DaliaNeural",
                             "es-AR-TomasNeural", "es-AR-ElenaNeural"],
                     width=28, state="readonly").grid(row=1, column=1, sticky="w", padx=10)

        self._lbl(vf, "Velocidad:").grid(row=2, column=0, sticky="w", pady=6)
        self.edge_rate = tk.StringVar(value="+0%")
        ttk.Combobox(vf, textvariable=self.edge_rate,
                     values=["-20%", "-10%", "+0%", "+10%", "+20%", "+30%"],
                     width=10, state="readonly").grid(row=2, column=1, sticky="w", padx=10)

        bf_voz = tk.Frame(tab, bg=BG)
        bf_voz.pack(anchor="w", padx=24, pady=(10, 4))
        self._btn(bf_voz, "🔊  Probar voz", self._test_edge_tts, BG3).pack(side=tk.LEFT, padx=4)
        self._btn(bf_voz, "⏹  Detener",    self._stop_edge_tts, BG3).pack(side=tk.LEFT, padx=4)

        self.voice_status = tk.Label(tab, text="", bg=BG, fg=TEXT2, font=F_SMALL)
        self.voice_status.pack(anchor="w", padx=24)

        # Estado interno de reproducción de voz
        #   _edge_proc   → subprocess del reproductor (PowerShell / afplay / mpg123)
        #   _edge_pygame → módulo pygame vivo mientras suena por esa vía
        #   _edge_stop   → bandera cooperativa: el hilo de reproducción la
        #                  consulta para cortar. Sin ella, "⏹ Detener" no tenía
        #                  ningún efecto en la ruta pygame (bucle ininterrumpible).
        self._edge_proc   = None
        self._edge_pygame = None
        self._edge_stop   = False

        self._load_ia(silent=True)

    # Hints de modelos por proveedor (solo referencia, el campo es libre)
    _MODELOS_HINT = {
        "OpenAI":      "GPT-5, GPT-5.4, gpt-4o, gpt-4o-mini",
        "Anthropic":   "claude-opus-4-6, claude-sonnet-4-6",
        "Google":      "gemini-2.5-pro, gemini-2.5-flash",
        "Grok (xAI)":  "grok-4, grok-4-0709, grok-3",
        "DeepSeek":    "deepseek-chat, deepseek-reasoner",
        "Mistral":     "mistral-large-latest, mistral-small-latest",
        "Desactivado": "",
    }

    def _on_prov_change(self, *_):
        prov = self.ia_prov.get()
        hint = self._MODELOS_HINT.get(prov, "")
        self.ia_model_hint.config(text=hint if hint else "")
        # Sugerir modelo por defecto si campo vacío
        defaults = {
            "OpenAI":      "gpt-4o",
            "Anthropic":   "claude-sonnet-4-6",
            "Google":      "gemini-2.5-flash",
            "Grok (xAI)":  "grok-4",
            "DeepSeek":    "deepseek-chat",
            "Mistral":     "mistral-large-latest",
            "Desactivado": "",
        }
        if not self.ia_model.get().strip():
            self.ia_model.set(defaults.get(prov, ""))

    def _toggle_key(self):
        self.ia_key_ent.config(
            show="" if self.ia_key_ent.cget("show") == "*" else "*")

    def _save_ia(self):
        cfg = {
            "provider":   self.ia_prov.get(),
            "api_key":    self.ia_key.get(),
            "model":      self.ia_model.get(),
            "voice":      self.ia_voice.get(),
            "edge_voice": self.edge_voice.get(),
            "edge_rate":  self.edge_rate.get(),
            "ctx_dx":     self.ia_ctx_dx.get().strip(),
            "ctx_hx":     self.ia_ctx_hx.get("1.0", tk.END).strip(),
        }
        path = _ruta_config("ia_settings.json")
        try:
            tmp = path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as fp:
                json.dump(cfg, fp, indent=2, ensure_ascii=False)
            os.replace(tmp, path)
            self.ia_status.config(text="✓ Guardado en " + path, fg=GREEN)
        except OSError as e:
            self.ia_status.config(
                text="✗ No se pudo guardar: " + str(e)[:80], fg=RED)

    def _load_ia(self, silent=False):
        path = _ruta_config("ia_settings.json")
        if not os.path.exists(path):
            return
        try:
            with open(path, encoding="utf-8") as fp:
                cfg = json.load(fp)
            self.ia_prov.set(cfg.get("provider", "Desactivado"))
            self.ia_key.set(cfg.get("api_key", ""))
            self.ia_model.set(cfg.get("model", ""))
            self.ia_voice.set(cfg.get("voice", False))
            self.edge_voice.set(cfg.get("edge_voice", "es-PE-AlexNeural"))
            self.edge_rate.set(cfg.get("edge_rate", "+0%"))
            ctx_dx = cfg.get("ctx_dx", "")
            ctx_hx = cfg.get("ctx_hx", "")
            if ctx_dx:
                self.ia_ctx_dx.delete(0, tk.END)
                self.ia_ctx_dx.insert(0, ctx_dx)
            if ctx_hx:
                self.ia_ctx_hx.delete("1.0", tk.END)
                self.ia_ctx_hx.insert("1.0", ctx_hx)
            if not silent:
                self.ia_status.config(text="✓ Configuración cargada", fg=GREEN)
        except Exception as e:
            if not silent:
                self.ia_status.config(text="Error al cargar: " + str(e), fg=RED)

    def _test_ia(self):
        key  = self.ia_key.get().strip()
        prov = self.ia_prov.get()
        if prov == "Desactivado":
            self.ia_status.config(text="IA desactivada.", fg=TEXT2); return
        if not key:
            self.ia_status.config(text="⚠ Introduce una API key primero.", fg=YELLOW); return
        if not self.ia_model.get().strip():
            self.ia_status.config(text="⚠ Escribe un nombre de modelo.", fg=YELLOW); return
        self.ia_status.config(text="Probando…", fg=TEXT2)
        self.root.update_idletasks()
        try:
            resp = self._llamar_ia_raw(prov, key, self.ia_model.get().strip(),
                                       "Responde solo: OK", max_tokens=5)
            self.ia_status.config(text=f"✓ {prov} OK: {resp}", fg=GREEN)
        except Exception as e:
            self.ia_status.config(text="✗ " + str(e)[:100], fg=RED)

    # ── Edge TTS ──────────────────────────────────────────────────────────────
    def _edge_tts_speak(self, texto):
        """Sintetiza texto con edge-tts en un hilo separado (no bloquea UI)."""
        import threading
        # Reiniciar la bandera: si el usuario pulsó "Detener" en la narración
        # anterior, sin esto la nueva se cortaría de inmediato.
        self._edge_stop = False
        threading.Thread(target=self._edge_tts_run, args=(texto,), daemon=True).start()

    def _edge_tts_run(self, texto):
        """Hilo: genera MP3 con edge-tts y lo reproduce sin abrir apps externas."""
        import asyncio
        try:
            import edge_tts
        except ImportError:
            self._ui_after(lambda: self.voice_status.config(
                text="⚠ Instala edge-tts:  pip install edge-tts", fg=YELLOW))
            return

        voice = self.edge_voice.get()
        rate  = self.edge_rate.get()

        async def _sintetizar(out_path):
            comm = edge_tts.Communicate(texto, voice=voice, rate=rate)
            await comm.save(out_path)

        tmp_path = None
        try:
            self._ui_after(lambda: self.voice_status.config(
                text="🔊 Sintetizando…", fg=TEXT2))
            with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as tmp:
                tmp_path = tmp.name

            asyncio.run(_sintetizar(tmp_path))

            self._ui_after(lambda: self.voice_status.config(
                text="🔊 Reproduciendo…", fg=GREEN))

            if sys.platform.startswith("win"):
                self._edge_proc = None
                reproducido = False
                # Opción 1: pygame (sin subprocess, sin ventanas)
                try:
                    import pygame
                    if not pygame.mixer.get_init():
                        pygame.mixer.init()
                    if self._edge_stop:
                        raise RuntimeError("detenido antes de reproducir")
                    pygame.mixer.music.load(tmp_path)
                    pygame.mixer.music.play()
                    self._edge_pygame = pygame     # referencia para "⏹ Detener"
                    import time
                    # El flag _edge_stop permite cortar la reproducción; antes
                    # este bucle era ininterrumpible y "Detener" no hacía nada.
                    while pygame.mixer.music.get_busy() and not self._edge_stop:
                        time.sleep(0.05)
                    pygame.mixer.music.stop()
                    try:
                        pygame.mixer.music.unload()   # suelta el MP3 para poder borrarlo
                    except Exception:
                        pass
                    self._edge_pygame = None
                    reproducido = True
                except Exception:
                    self._edge_pygame = None
                # Opción 2: PowerShell MediaPlayer completamente oculto
                if not reproducido and not self._edge_stop:
                    uri = tmp_path.replace("\\", "/")
                    ps = (
                        "Add-Type -AssemblyName PresentationCore; "
                        "$mp = New-Object System.Windows.Media.MediaPlayer; "
                        "$mp.Open([System.Uri]'" + uri + "'); "
                        "$mp.Play(); "
                        "Start-Sleep -Milliseconds 300; "
                        "$t = 0; "
                        "while (-not $mp.NaturalDuration.HasTimeSpan -and $t -lt 100)"
                        "{ Start-Sleep -Milliseconds 100; $t++ }; "
                        "if (-not $mp.NaturalDuration.HasTimeSpan) { $mp.Close(); exit 1 }; "
                        "$d=[int]$mp.NaturalDuration.TimeSpan.TotalMilliseconds; "
                        "Start-Sleep -Milliseconds $d; "
                        "$mp.Stop(); $mp.Close()"
                    )
                    proc = subprocess.Popen(
                        ["powershell", "-NoProfile", "-NonInteractive",
                         "-WindowStyle", "Hidden", "-c", ps],
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                        creationflags=0x08000000)  # CREATE_NO_WINDOW
                    self._edge_proc = proc         # ahora sí es terminable
                    _registrar_hijo(proc)
                    proc.wait()
                    self._edge_proc = None

            elif sys.platform == "darwin":
                self._edge_proc = subprocess.Popen(
                    ["afplay", tmp_path],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                _registrar_hijo(self._edge_proc)
                self._edge_proc.wait()
                self._edge_proc = None

            else:  # Linux
                self._edge_proc = None
                for player in ["mpg123", "mpg321", "ffplay"]:
                    try:
                        args = (["ffplay", "-nodisp", "-autoexit",
                                 "-loglevel", "quiet", tmp_path]
                                if player == "ffplay"
                                else [player, "-q", tmp_path])
                        p = subprocess.Popen(
                            args,
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                        self._edge_proc = p     # terminable desde "⏹ Detener"
                        _registrar_hijo(p)
                        p.wait()
                        self._edge_proc = None
                        break
                    except FileNotFoundError:
                        continue

            _fin_txt, _fin_col = (("⏹ Detenido", TEXT2) if self._edge_stop
                                  else ("✓ Listo", GREEN))
            self._ui_after(lambda t=_fin_txt, c=_fin_col:
                            self.voice_status.config(text=t, fg=c))

        except Exception as e:
            # Mismo fix que en _llamar_ia_comentario: capturar el texto antes de
            # que 'e' se desvincule al cerrar el except (el lambda es diferido).
            msg = "✗ " + str(e)[:120]
            self._ui_after(lambda: self.voice_status.config(
                text=msg, fg=RED))
        finally:
            if tmp_path:
                try:
                    os.unlink(tmp_path)
                except Exception:
                    pass

    def _test_edge_tts(self):
        """Prueba la voz con una frase de ejemplo."""
        self._edge_tts_speak(
            "Prueba de síntesis de voz Edge. "
            "Arteria cerebral media derecha. "
            "Pico sistólico dentro de rangos normales.")

    def _stop_edge_tts(self):
        """Detiene la reproducción activa, sea cual sea la vía usada.

        Hay tres rutas de reproducción y antes solo se cubría una:
          · pygame (Windows, vía preferente) → bucle en el hilo: se corta con
            la bandera _edge_stop y un stop() directo sobre el mixer.
          · PowerShell / afplay / mpg123 → subprocess: terminate().
          · Síntesis aún en curso (antes de sonar) → la bandera evita que
            arranque la reproducción al terminar de sintetizar.
        """
        detenido = self._detener_audio()
        try:
            self.voice_status.config(
                text="⏹ Detenido" if detenido else "⏹ Detención solicitada",
                fg=TEXT2)
        except tk.TclError:
            pass

    def _detener_audio(self):
        """Corta cualquier reproducción sin tocar la UI. Devuelve True si
        había algo sonando. Lo usan "⏹ Detener" y el cierre de la ventana."""
        self._edge_stop = True
        detenido = False

        pg = self._edge_pygame
        if pg is not None:
            try:
                if pg.mixer.get_init():
                    pg.mixer.music.stop()
                    detenido = True
            except Exception:
                pass
            self._edge_pygame = None

        proc = self._edge_proc
        if proc is not None:
            try:
                if proc.poll() is None:
                    proc.terminate()
                    detenido = True
            except Exception:
                pass
            self._edge_proc = None
        return detenido

    def _narrar_resultado(self, ps, ed, pi, pic_label, arteria, lado, estados,
                          pic_bellner_str="", lind_data=None,
                          alertas=(), co2=None):
        """Narra el resultado del análisis si la voz está activada."""
        if not self.ia_voice.get():
            return
        est_ps  = estados.get("PS",  ("", ""))[0]
        est_ed  = estados.get("ED",  ("", ""))[0]
        est_pi  = estados.get("PI",  ("", ""))[0]

        def _clean(s):
            return s.replace("✓ ", "").replace("↑ ", "").replace("↓ ", "")

        texto = (
            f"Análisis Doppler Transcraneal. "
            f"Arteria {arteria}, lado {lado}. "
            f"Pico sistólico {round(ps, 0)} centímetros por segundo, {_clean(est_ps)}. "
            f"Velocidad diastólica {round(ed, 0)} centímetros por segundo, {_clean(est_ed)}. "
        )
        if pi is not None:
            texto += f"Índice de pulsatilidad {round(pi, 2)}, {_clean(est_pi)}. "
        # Czosnyka (pic_label limpio de símbolos para TTS)
        pic_voz = pic_label.replace("↑", "").replace("↓", "").strip()
        texto += f"Presión intracraneal por Czosnyka: {pic_voz}. "
        # Bellner
        if pic_bellner_str:
            bell_voz = pic_bellner_str.replace("↑", "").replace("↓", "").strip()
            texto += f"Presión intracraneal por Bellner: {bell_voz}. "
        # Lindegaard
        if lind_data:
            il, _, _, lind_txt, _ = lind_data
            texto += f"Índice de Lindegaard {round(il, 1)}. {lind_txt}."
        if _co2_fuera({"PaCO2": co2}):
            texto += (f"Atención: dióxido de carbono {co2:g} milímetros de mercurio, "
                      "fuera de rango; interpretar con cautela. ")
        for _, t in (alertas or ()):
            texto += "Alerta: " + t.replace("VM", "velocidad media").replace(
                "cm/s", "centímetros por segundo").replace("IL", "índice de Lindegaard") + ". "
        self._edge_tts_speak(texto)

    # =========================================================================
    # PESTAÑA 3 — Referencia
    # =========================================================================
    def _tab_referencia(self):
        tab = ttk.Frame(self.nb)
        self.nb.add(tab, text="  ℹ  Referencia  ")

        txt = tk.Text(tab, bg=BG2, fg=TEXT, font=F_MONO,
                      relief="flat", wrap="word", padx=18, pady=14)
        txt.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        txt.insert("1.0", """\
DOPPLER TRANSCRANEAL — VALORES DE REFERENCIA  (todas las arterias)
════════════════════════════════════════════════════════════════════════════════

  PS = Pico Sistólico    ED = Final Diastólico

  ══ NOMENCLATURA DE ESTE EQUIPO ═══════════════════════════════════
  Campo "TAmax"  en este equipo = TAMV = Mean Flow Velocity = velocidad media
                 ⚠ NO es el TAP (Time-Averaged Peak / envolvente pico)
                 Se usa para: Czosnyka (CPP/PIC), Lindegaard (IL), PI, Soustiel
  Campo "TAmean" en este equipo = campo redundante / no usar para fórmulas
  ══════════════════════════════════════════════════════════════════════

  Nomenclatura estándar (literatura):
    TAP / TAmax  = Time-Averaged Peak velocity  (envolvente espectral pico)
    TAMV / MFV   = Mean Flow Velocity = TAP / 2 ≈     ← lo que da este equipo
    TAmean       = mismo que TAMV en la mayoría de equipos

  RI = (PS−ED)/PS        PI = (PS−ED)/TAMV     S/D = PS/ED
  Si no se mide TAMV:   TAMV ≈ (PS + 2·ED) / 3
  RI normal < 0.75  ·  PI normal 0.55–1.25 (por arteria, ver tabla)  ·  S/D 1.4–3.8

  ⚠ Czosnyka, Lindegaard y Soustiel: usar siempre campo "TAmax" de este equipo
    (= TAMV). El campo "TAmean" de este equipo NO se usa en estas fórmulas.

────────────────────────────────────────────────────────────────────────────────
TABLA RESUMEN
────────────────────────────────────────────────────────────────────────────────
Arteria            Ventana       Prof.       PS cm/s   ED cm/s   TAmax cm/s    PI
────────────────────────────────────────────────────────────────────────────────
ACM                Temporal      30–65 mm    60–140    25–65     40–85         0.55–1.25
ACA                Temporal      60–75 mm    40–100    20–55     30–70         0.55–1.25
ACP (P1/P2)        Temporal      60–70 mm    30–85     15–45     20–55         0.55–1.25
ACI / Sifón Car.   Temporal/Orb  65–80 mm    40–100    20–55     30–70         0.55–1.25
ACE                Submandib.    30–55 mm    40–100    15–45     25–60         0.55–1.25
Arteria Basilar    Suboccipital  80–120 mm   30–85     15–45     25–60         0.55–1.25
Arteria Vertebral  Suboccipital  60–75 mm    30–80     12–40     20–50         0.55–1.25
PICA               Suboccipital  50–70 mm    20–65     10–35     15–45         0.55–1.25
AICA               Suboccipital  50–70 mm    20–65     10–35     15–45         0.55–1.25
Oftálmica          Orbitaria     45–55 mm    15–50      5–20     10–30         0.55–1.35
Sifón Carotídeo    Orbitaria     65–80 mm    40–100    20–55     30–70         0.55–1.25
────────────────────────────────────────────────────────────────────────────────
  Fuente de profundidad, dirección de flujo y VM media ± DE: Naqvi 2013, Tabla 1
  (adaptada de Nicoletto y Burkman 2009) — refs. 7 y 10. Sin dato en esa tabla
  (se mantienen los valores previos): ACE, PICA, AICA y los rangos de PS, ED,
  RI, PI y TAmax. La Tabla 1 da media ± DE, no rango; los rangos de TAmax de
  esta app NO se reemplazaron.
────────────────────────────────────────────────────────────────────────────────

════════════════════════════════════════════════════════════════════════════════
DETALLE POR ARTERIA
════════════════════════════════════════════════════════════════════════════════

─── ACM — Arteria Cerebral Media ───────────────────────────────────────────────
  Ventana temporal  ·  30–65 mm  ·  flujo HACIA el transductor (+)
  PS: 60–140 cm/s     ED: 25–65 cm/s     TAmean: 40–85 cm/s
  RI: 0.45–0.75       PI: 0.55–1.25      S/D: 1.4–3.8
  ▸ Arteria más estudiada en TCD. Referencia para índice de Lindegaard.
  ▸ VM normal del adulto (Naqvi 2013): 55 ± 12 cm/s.
  ▸ Vasoespasmo: leve 120–150 · moderado 150–200 · severo > 200 cm/s (TAmax)
  ▸ Índice Lindegaard = VM ACM intracraneal / VM ACI extracraneal  →  > 3 probable · > 6 severo

─── ACA — Arteria Cerebral Anterior ────────────────────────────────────────────
  Ventana temporal  ·  60–75 mm  ·  flujo LEJOS del transductor (−)
  PS: 40–100 cm/s     ED: 20–55 cm/s     TAmean: 30–70 cm/s
  RI: 0.45–0.75       PI: 0.55–1.25
  ▸ VM normal del adulto (Naqvi 2013): 50 ± 11 cm/s.
  ▸ Velocidades ~70–80 % de las de ACM ipsilateral.
  ▸ Asimetría > 20 % entre lados sugiere patología.
  ▸ Vasoespasmo: leve 90–120 · moderado 120–150 · severo > 150 cm/s (TAmax)

─── ACP — Arteria Cerebral Posterior ───────────────────────────────────────────
  Ventana temporal  ·  P1: 60–70 mm (hacia) / P2: 60–70 mm (lejos)
  PS: 30–85 cm/s      ED: 15–45 cm/s     TAmean: 20–55 cm/s
  RI: 0.45–0.75       PI: 0.55–1.25
  ▸ VM normal del adulto (Naqvi 2013): P1 39 ± 10 · P2 40 ± 10 cm/s.
  ▸ Velocidades menores que ACM; útil para valorar circulación posterior.
  ▸ Vasoespasmo: leve 70–90 · moderado 90–110 · severo > 110 cm/s (TAmax)

─── ACI — Arteria Carótida Interna (segmento intracraneal / sifón) ─────────────
  Ventana temporal u orbitaria  ·  65–80 mm (transorbital)
  PS: 40–100 cm/s     ED: 20–55 cm/s     TAmean: 30–70 cm/s
  RI: 0.45–0.75       PI: 0.55–1.25
  ▸ Denominador del índice de Lindegaard: VM ACI extracraneal ipsilateral
    (ventana submandibular). No usar ACI intracraneal.
  ▸ VM normal del adulto (Naqvi 2013): supraclinoidea 41 ± 11 (lejos) ·
    paraselar 47 ± 14 cm/s (hacia).
  ▸ Ventana orbitaria: ⚠ reducir MI < 0.23 (protección ocular).

─── ACE — Arteria Carótida Externa ─────────────────────────────────────────────
  Ventana submandibular  ·  30–55 mm  ·  flujo hacia transductor (+)
  PS: 40–100 cm/s     ED: 15–45 cm/s     TAmean: 25–60 cm/s
  RI: 0.45–0.75       PI: 0.55–1.25
  ▸ Alta resistencia comparada con ACI; RI elevado puede ser normal.
  ▸ Útil para evaluar colaterales y flujo carotídeo externo.

─── AB — Arteria Basilar ───────────────────────────────────────────────────────
  Ventana suboccipital (foramen magno)  ·  80–120 mm  ·  flujo LEJOS (−)
  PS: 30–85 cm/s      ED: 15–45 cm/s     TAmean: 25–60 cm/s
  RI: 0.45–0.75       PI: 0.55–1.25
  ▸ VM normal del adulto (Naqvi 2013): 41 ± 10 cm/s.
  ▸ Vasoespasmo: leve 80–100 · moderado 100–140 · severo > 140 cm/s (TAmax)
  ▸ Índice Soustiel = TAmax AB / TAmax AV  →  > 2 probable · > 3 severo
  ▸ Asimetría vertebral previa puede indicar oclusión parcial.

─── AV — Arteria Vertebral ─────────────────────────────────────────────────────
  Ventana suboccipital  ·  60–75 mm  ·  flujo LEJOS (−)
  PS: 30–80 cm/s      ED: 12–40 cm/s     TAmean: 20–50 cm/s
  RI: 0.45–0.75       PI: 0.55–1.25
  ▸ VM normal del adulto (Naqvi 2013): 38 ± 10 cm/s.
  ▸ Denominador del índice de Soustiel.
  ▸ Asimetría > 30 % entre AV derecha e izquierda: considerar hipoplasia.
  ▸ Vasoespasmo: leve 60–80 · moderado 80–110 · severo > 110 cm/s (TAmax)

─── PICA — Arteria Posteroinferior del Cerebelo ────────────────────────────────
  Ventana suboccipital  ·  50–70 mm
  PS: 20–65 cm/s      ED: 10–35 cm/s     TAmean: 15–45 cm/s
  RI: 0.45–0.75       PI: 0.55–1.25
  ▸ Velocidades bajas; puede confundirse con AV distal.
  ▸ Oclusión → síndrome de Wallenberg (infarto bulbar lateral).

─── AICA — Arteria Anteroinferior del Cerebelo ─────────────────────────────────
  Ventana suboccipital  ·  50–70 mm
  PS: 20–65 cm/s      ED: 10–35 cm/s     TAmean: 15–45 cm/s
  RI: 0.45–0.75       PI: 0.55–1.25
  ▸ Difícil insonación; valores similares a PICA.
  ▸ Irriga puente inferior, cerebelo anteroinferior y oído interno.

─── Oftálmica ──────────────────────────────────────────────────────────────────
  Ventana orbitaria  ·  45–55 mm  ·  ⚠ MI < 0.23 OBLIGATORIO
  PS: 15–50 cm/s      ED: 5–20 cm/s      TAmean: 10–30 cm/s
  RI: 0.55–0.80       PI: 0.55–1.35   (mayor resistencia que intracraneales)
  ▸ VM normal del adulto (Naqvi 2013): 21 ± 5 cm/s · resistencia alta.
  ▸ Flujo normal hacia el ojo (+). Inversión → estenosis ACI severa ipsilateral.
  ▸ Colateral clásica: ACI → oftálmica → ACE en oclusiones carotídeas.

─── Sifón Carotídeo ────────────────────────────────────────────────────────────
  Ventana orbitaria  ·  65–80 mm  ·  ⚠ MI < 0.23
  PS: 40–100 cm/s     ED: 20–55 cm/s     TAmean: 30–70 cm/s
  RI: 0.45–0.75       PI: 0.55–1.25
  ▸ Segmento C2–C4 de la ACI. Dirección de flujo variable según porción:
    supraclinoidea lejos (−) · paraselar hacia (+)  (Naqvi 2013).
  ▸ Estenosis aquí no queda reflejada en el índice de Lindegaard clásico.

════════════════════════════════════════════════════════════════════════════════
ESTIMACIÓN PIC no invasiva  —  dos métodos  (solo ACM)
════════════════════════════════════════════════════════════════════════════════

MÉTODO 1 — Czosnyka et al. J Neurosurg 1998;88(5):802-8
  CPPe = PAM × (ED / TAMV) + 14     [TAMV = campo "TAmax" de este equipo]
  PICe = PAM − CPPe
  Requiere: PAM y TAMV (campo TAmax). Solo ACM.
  Validación: TBI. Mejor utilidad: DESCARTAR HTE (VPN alto, IMPRESSIT-2).
  ⚠ Fórmula original usa TAP (envolvente pico); este equipo solo da TAMV.

MÉTODO 2 — Bellner et al. Surg Neurol 2004;62(1):45-51
  PICe = 10.93 × PI − 1.28
  Requiere: solo PI de Gosling. No requiere PAM. Solo ACM.
  Validación: 81 pac. (HSA 57% + TBI 26%), r=0.938, SD=±2.5 mmHg (rango 5–40).
  ⚠ LIMITACIÓN: resultado no replicado con igual precisión en cohortes externas.
    Usar como orientación, no como valor absoluto.

  Interpretación PICe (ambos métodos):
      < 15 mmHg   →  Normal          (límite fisiológico adulto)
      15–20 mmHg  →  Elevada    ↑    (zona de alerta)
      > 20 mmHg   →  ALTA       ↑↑   (umbral de acción clínica BTF/NCS 2023)

  CPP objetivo: > 60 mmHg en adultos (> 50 en niños).
  ⚠ Estimación orientativa. Correlacionar siempre con clínica e imagen.
  ⚠ Aplicar solo en ACM.

════════════════════════════════════════════════════════════════════════════════
CRITERIOS DE VASOESPASMO
════════════════════════════════════════════════════════════════════════════════
  [TAMV = campo "TAmax" de este equipo = Mean Flow Velocity]

  ACM (TAMV)    Leve: 120–150   Moderado: 150–200   Severo: > 200 cm/s
  ACA (TAMV)    Leve:  90–120   Moderado: 120–150   Severo: > 150 cm/s
  ACP (TAMV)    Leve:  70–90    Moderado:  90–110   Severo: > 110 cm/s
  AB  (TAMV)    Leve:  80–100   Moderado: 100–140   Severo: > 140 cm/s
  AV  (TAMV)    Leve:  60–80    Moderado:  80–110   Severo: > 110 cm/s

  Índice de Lindegaard (IL) = TAMV ACM intracraneal / TAMV ACI extracraneal
  ACI extracraneal = carótida interna ipsilateral, ventana submandibular
  (retromandibular), 45–50 mm; VM normal 30 ± 9 cm/s (Naqvi 2013)
  Ingresar TAMV de ACI extracraneal = campo "TAmax" del equipo medido en submandibular
        > 3  →  vasoespasmo ACM probable
        > 6  →  vasoespasmo ACM severo

  Índice de Soustiel = TAMV AB / TAMV AV extracraneal  (campo "TAmax" del equipo)
        > 2  →  vasoespasmo basilar probable
        > 3  →  vasoespasmo basilar severo

════════════════════════════════════════════════════════════════════════════════
PROFUNDIDADES DE INSONACIÓN
════════════════════════════════════════════════════════════════════════════════
  Ventana temporal
    ACM              30–65 mm   (+) hacia transductor
    ACA              60–75 mm   (−) lejos del transductor
    ACP P1           60–70 mm   (+) hacia transductor
    ACP P2           60–70 mm   (−) lejos del transductor
    ACI sifón        60–70 mm   variable

  Ventana suboccipital (foramen magno)
    Arteria Basilar    80–120 mm  (−) lejos
    Arteria Vertebral   60–75 mm  (−) lejos
    PICA                50–70 mm

  Ventana orbitaria  ⚠ MI < 0.23
    Oftálmica           45–55 mm   (+) hacia transductor
    Sifón carotídeo     65–80 mm   supraclinoideo (−) · paraselar (+)

  Ventana submandibular
    ACI extracraneal    45–50 mm
    ACE                 30–55 mm

════════════════════════════════════════════════════════════════════════════════
BIBLIOGRAFÍA  (formato Vancouver)
════════════════════════════════════════════════════════════════════════════════
  1. Aaslid R, Markwalder TM, Nornes H. Noninvasive transcranial Doppler
     ultrasound recording of flow velocity in basal cerebral arteries.
     J Neurosurg. 1982;57(6):769-74.

  2. Lindegaard KF, Nornes H, Bakke SJ, Sorteberg W, Nakstad P. Cerebral
     vasospasm diagnosis by means of angiography and blood velocity
     measurements. Acta Neurochir (Wien). 1989;100(1-2):12-24.

  3. Czosnyka M, Matta BF, Smielewski P, Kirkpatrick PJ, Pickard JD.
     Cerebral perfusion pressure in head-injured patients: a noninvasive
     assessment using transcranial Doppler ultrasonography.
     J Neurosurg. 1998;88(5):802-8.

  4. Soustiel JF, Shik V, Shreiber R, Tavor Y, Goldsher D. Basilar
     vasospasm diagnosis: investigation of a modified "Lindegaard Index"
     based on imaging studies and blood velocity measurements of the
     basilar artery. Stroke. 2002;33(1):72-7.

  5. Sloan MA, Alexandrov AV, Tegeler CH, Spencer MP, Caplan LR,
     Feldmann E, et al. Assessment: transcranial Doppler ultrasonography:
     report of the Therapeutics and Technology Assessment Subcommittee of
     the American Academy of Neurology. Neurology. 2004;62(9):1468-81.

  6. Alexandrov AV, Sloan MA, Wong LK, Douville C, Razumovsky AY,
     Koroshetz WJ, et al. Practice standards for transcranial Doppler
     ultrasound: part I—test performance. J Neuroimaging. 2007;17(1):11-8.

  7. Naqvi J, Yap KH, Ahmad G, Ghosh J. Transcranial Doppler ultrasound:
     a review of the physical principles and major applications in
     critical care. Int J Vasc Med. 2013;2013:629378.
     (Tabla 1: insonación y VM normal por arteria · Tabla 3: grados de
     vasoespasmo · Tabla 2: factores que modifican la VM, incl. PCO2.)

  8. Robba C, Pozzebon S, Moro B, Vincent JL, Creteur J, Taccone FS.
     Multimodal non-invasive assessment of intracranial hypertension:
     an observational study. Crit Care. 2020;24(1):379.

  9. Rasulo FA, Calza S, Robba C, Taccone FS, Biasucci DG, Badenes R,
     et al. Transcranial Doppler as a screening test to exclude
     intracranial hypertension in brain-injured patients: the
     IMPRESSIT-2 prospective multicenter international study.
     Crit Care. 2022;26(1):110.

 10. Nicoletto HA, Burkman MH. Transcranial Doppler series part II:
     performing a transcranial Doppler. Am J Electroneurodiagnostic
     Technol. 2009;49(1):14-27.  (Fuente de la Tabla 1 de la ref. 7:
     profundidad, dirección de flujo y VM normal por arteria.)

 11. Bellner J, Romner B, Reinstrup P, Kristiansson KA, Ryding E, Brandt L.
     Transcranial Doppler sonography pulsatility index (PI) reflects
     intracranial pressure (ICP). Surg Neurol. 2004;62(1):45-51.
""")
        txt.config(state="disabled")

    # ── Lindegaard helpers ────────────────────────────────────────────────────
    def _calc_lindegaard_live(self):
        """Calcula IL en tiempo real. SIEMPRE usa TAMV (campo TAmax del equipo).

        El IL de Lindegaard se define con mean flow velocity (TAMV) en ambos
        vasos. En este equipo, el campo "TAmax" contiene la TAMV. Por eso el IL
        SOLO se calcula con un TAMV ACM medido (no con la aproximación
        (PS+2.ED)/3, que introduciría error en una decisión clínica).
        """
        # VM ACM = campo TAmax (TAMV) medido. Debe existir y ser > 0.
        acm_val = self._lind_tamax_acm
        if acm_val is None or acm_val <= 0:
            self.lind_il_lbl.config(   text="IL = —", fg=TEXT2)
            self.lind_interp_lbl.config(
                text="Falta TAMV ACM medido (campo TAmax) — no se calcula con aproximación",
                fg=YELLOW)
            self.lind_detalle_lbl.config(text="", fg=TEXT2)
            self._lind_il_result = None
            return
        try:
            aci = float(self.lind_aci_var.get().strip())
            if aci <= 0:
                self.lind_interp_lbl.config(text="ACI debe ser > 0", fg=YELLOW)
                self._lind_il_result = None
                return
            il = acm_val / aci
            if il > 6:
                col   = RED
                interp = "Vasoespasmo ACM SEVERO"
            elif il > 3:
                col   = YELLOW
                interp = "Vasoespasmo ACM probable"
            else:
                col   = GREEN
                interp = "Sin vasoespasmo ACM"
            self.lind_il_lbl.config(   text=f"IL = {il:.2f}", fg=col)
            self.lind_interp_lbl.config(text=interp,           fg=col)
            self.lind_detalle_lbl.config(
                text=f"TAMV ACM {acm_val:.1f} cm/s  /  TAMV ACI-ex {aci:.1f} cm/s",
                fg=TEXT2)
            # Guardar para uso en figura
            self._lind_il_result = (il, acm_val, aci, interp, col)
            # Si ya hay un análisis generado, reflejar el IL en el informe de texto
            self._anexar_il_a_informe(il, acm_val, aci, interp)
        except (ValueError, ZeroDivisionError):
            self.lind_il_lbl.config(   text="IL = —",                         fg=TEXT2)
            self.lind_interp_lbl.config(text="Ingresa TAMV ACI extracraneal para calcular", fg=TEXT2)
            self.lind_detalle_lbl.config(text="",                               fg=TEXT2)
            self._lind_il_result = None

    def _anexar_il_a_informe(self, il, acm_val, aci, interp):
        """Añade/actualiza la línea del IL en el informe de texto interno ya generado.

        El IL se ingresa en la pestaña Lindegaard DESPUÉS de generar el análisis,
        así que el informe original no lo incluye. Aquí lo anexamos sin regenerar
        toda la figura. La figura PNG ya guardada no se modifica; para incluir el
        IL en el PNG hay que pulsar de nuevo "Generar Análisis".

        Desde que la caja de informe visible fue reemplazada por el botón de
        exportación DOCX, este método opera sobre self._informe_texto (string
        interno) en vez de un widget Tk.
        """
        try:
            contenido = (self._informe_texto or "").rstrip("\n")
            if not contenido.strip():
                return  # no hay informe que actualizar
            linea_il = ("Índice de Lindegaard = " + str(round(il, 2))
                        + "  (TAMV ACM " + str(round(acm_val, 1))
                        + " / ACI-ex " + str(round(aci, 1)) + " cm/s)"
                        + "  →  " + interp)
            # Quitar línea IL previa si existe, para no duplicar
            lineas = [ln for ln in contenido.split("\n")
                      if not ln.startswith("Índice de Lindegaard")]
            lineas.append(linea_il)
            self._informe_texto = "\n".join(lineas)
        except Exception:
            pass

    def _sincronizar_il_monitor(self):
        """El IL se ingresa DESPUÉS de generar el análisis. Si ya hay un
        resultado enviado al monitor principal, se le actualiza el IL y se
        reenvía para que la alerta de vasoespasmo lo tome en cuenta."""
        lr = self.last_result
        if not lr or lr.get("arteria") != "ACM":
            return
        il = self._lind_il_result[0] if self._lind_il_result else None
        il = round(il, 2) if il is not None else None
        if lr.get("IL") == il:
            return
        lr["IL"] = il
        if self.on_result:
            try:
                self.on_result(lr)
            except Exception:
                pass

    def _get_lindegaard_data(self):
        """Devuelve resultado IL para la figura o None."""
        return getattr(self, '_lind_il_result', None)

    def _habilitar_tab_lindegaard(self, nombre, arteria, lado, tamean, tamax):
        """Llamado desde _generar para activar la pestaña IL y poblar datos.

        El Índice de Lindegaard SOLO está validado para la ACM (ACM intracraneal
        / ACI extracraneal ipsilateral). Para cualquier otra arteria se deshabilita
        la pestaña y no se calcula IL, evitando un índice clínicamente sin sentido.
        """
        if arteria != "ACM":
            self.nb.tab(1, state="disabled")
            self._lind_tamax_acm  = None
            self._lind_il_result  = None
            self.lind_info_pac.config(   text=nombre)
            self.lind_info_art.config(   text=f"{arteria}  {lado}  —  IL no aplicable")
            self.lind_info_tamean.config(text="—  cm/s")
            self.lind_info_tamax.config( text="—  cm/s")
            self.lind_il_lbl.config(     text="IL = —", fg=TEXT2)
            self.lind_interp_lbl.config(
                text="El Índice de Lindegaard solo aplica a la ACM", fg=YELLOW)
            self.lind_detalle_lbl.config(text="", fg=TEXT2)
            return
        self.nb.tab(1, state="normal")
        self._lind_tamax_acm  = tamax
        self.lind_info_pac.config(   text=nombre)
        self.lind_info_art.config(   text=f"{arteria}  {lado}")
        self.lind_info_tamean.config(text=f"{tamean:.1f}  cm/s")
        tamax_txt = f"{tamax:.1f}  cm/s" if tamax else "—"
        self.lind_info_tamax.config( text=tamax_txt)
        # Recalcular si ya hay ACI ingresada
        self._calc_lindegaard_live()

    # =========================================================================
    # Status bar
    # =========================================================================
    def _build_statusbar(self):
        bar = tk.Frame(self.root, bg=BG2, height=26)
        bar.pack(fill=tk.X, side=tk.BOTTOM)
        bar.pack_propagate(False)
        tk.Frame(bar, bg=ACCENT, width=3).pack(side=tk.LEFT)
        self.status_lbl = tk.Label(bar,
                                   text="Listo  ·  Ingresa PS y ED para comenzar",
                                   bg=BG2, fg=TEXT2, font=F_SMALL, anchor="w")
        self.status_lbl.pack(side=tk.LEFT, padx=10)
        tk.Label(bar, text="App: " + _autor_app(),
                 bg=BG2, fg=TEXT3, font=F_SMALL, anchor="e"
                 ).pack(side=tk.RIGHT, padx=10)

    def _set_status(self, msg, color=TEXT2):
        self.status_lbl.config(text=msg, fg=color)

    # =========================================================================
    # Helpers de construcción UI
    # =========================================================================
    def _seccion(self, parent, titulo):
        f = tk.Frame(parent, bg=BG)
        f.pack(fill=tk.X, pady=(10, 2))
        tk.Label(f, text=titulo, bg=BG, fg=ACCENT,
                 font=F_HEAD).pack(side=tk.LEFT)
        tk.Frame(f, bg=BORDER, height=1).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=8, pady=4)

    def _lbl(self, parent, texto):
        return tk.Label(parent, text=texto, bg=BG, fg=TEXT2, font=F_SMALL)

    def _entry(self, parent, width=20):
        return tk.Entry(parent, width=width, bg=BG3, fg=TEXT,
                        font=F_MONO, relief="flat", insertbackground=TEXT,
                        highlightthickness=1,
                        highlightbackground=BORDER,
                        highlightcolor=ACCENT)

    def _btn(self, parent, texto, cmd, color=BG3):
        """Boton estilo monitor clínico. El color NUNCA rellena el boton.

        - color=BG3 (default) -> boton normal: fondo oscuro, texto gris claro.
        - color=ACCENT/GREEN/... -> boton PRIMARIO: fondo oscuro, texto y
          borde del color indicado. El color vive en texto+borde, no en relleno.
        """
        if color == BG3:
            return tk.Button(parent, text=texto, command=cmd,
                             bg=BG3, fg=TEXT2, font=F_BODY,
                             relief="flat", bd=0, cursor="hand2",
                             activebackground=BORDER, activeforeground=TEXT,
                             padx=10, pady=5)
        # Boton primario: el color pasa a texto + borde fino
        return tk.Button(parent, text=texto, command=cmd,
                         bg=BG3, fg=color, font=F_BODY,
                         relief="flat", bd=0, cursor="hand2",
                         highlightthickness=1, highlightbackground=color,
                         highlightcolor=color,
                         activebackground=BORDER, activeforeground=color,
                         padx=10, pady=5)

    # =========================================================================
    # Validación en tiempo real
    # =========================================================================
    def _on_change(self, clave):
        self._validar(clave, self.vars[clave].get().strip())

    def _on_arteria_change(self, event=None):
        art = self.arteria_var.get()
        r = self.RANGOS_ARTERIA.get(art, self.RANGOS_ARTERIA["ACM"])
        ps_min, ps_max, ed_min, ed_max, tam_min, tam_max, pi_min, pi_max = r
        # Actualizar rangos dinámicamente
        ri_min, ri_max = 0.45, 0.75
        sd_min, sd_max = 1.4, 3.8
        updates = {
            "PS":     (ps_min, ps_max),
            "ED":     (ed_min, ed_max),
            "RI":     (ri_min, ri_max),
            "PI":     (pi_min, pi_max),
            "SD":     (sd_min, sd_max),
            "TAmax":  (tam_min, tam_max),  # TAmax = TAMV en este equipo; mismos rangos que TAmean
            "TAmean": (tam_min, tam_max),
            "HR":     (45, 120),
        }
        for k, (mn, mx) in updates.items():
            self.entries[k]["min"] = mn
            self.entries[k]["max"] = mx
            self._validar(k, self.vars[k].get().strip())

    def _validar(self, clave, valor):
        lbl = self.status_labels[clave]
        if not valor:
            lbl.config(text="", fg=TEXT2); return
        try:
            v = float(valor)
            mn, mx = self.entries[clave]["min"], self.entries[clave]["max"]
            if mn <= v <= mx:
                lbl.config(text="✓ Normal",  fg=GREEN)
            elif mn * 0.80 <= v <= mx * 1.25:
                lbl.config(text="⚠ Límite",  fg=YELLOW)
            else:
                lbl.config(text="✗ Anormal", fg=RED)
        except ValueError:
            lbl.config(text="?", fg=TEXT2)

    # =========================================================================
    # Generar gráfica
    # =========================================================================
    def _safe_generar(self):
        try:
            self._generar()
        except Exception as e:
            msg = "Error crítico:\n" + str(e) + "\n\n" + traceback.format_exc()
            messagebox.showerror("Error", msg[:900])
            self._set_status("Error al generar análisis", RED)

    def _fval(self, clave, default=None):
        """Devuelve float del campo o default si vacío/inválido."""
        try:
            return float(self.vars[clave].get().strip())
        except ValueError:
            return default

    def _generar(self):
        ps = self._fval("PS")
        ed = self._fval("ED")

        if not ps or not ed or ps <= 0 or ed <= 0:
            messagebox.showwarning("Datos incompletos",
                                   "PS y ED son obligatorios y deben ser > 0.")
            return

        # PAM OPCIONAL: solo se requiere para Czosnyka (PIC/CPP).
        # Bellner (PI) e Índice de Lindegaard NO la necesitan, así que si falta
        # se omite Czosnyka pero el resto del análisis se genera igual.
        try:
            pam = float(self.pam_ent.get())
            if pam <= 0:
                pam = None
        except (ValueError, AttributeError):
            pam = None

        # PaCO₂ / EtCO₂ OPCIONAL: no entra en ninguna fórmula; solo se registra
        # y avisa si está fuera de 35–45 mmHg (puede confundir PI/VM/PIC).
        try:
            co2 = float(self.co2_ent.get())
            if co2 <= 0:
                co2 = None
        except (ValueError, AttributeError):
            co2 = None

        # HR sin valor por defecto: antes se asumían 70 lpm y ese número
        # aparecía en figura, informe e IA como si se hubiera medido.
        hr     = self._fval("HR")
        tamean = self._fval("TAmean")
        tamax  = self._fval("TAmax")

        # ¿La TAMV es medida o aproximada? Si no se ingresó TAmean ni TAmax,
        # se sustituye por (PS+2·ED)/3 — una APROXIMACIÓN. Lo registramos para
        # marcar después la estimación de Bellner que dependa de ella.
        tamean_aprox = (not tamean or tamean <= 0)
        if not tamean or tamean <= 0:
            tamean = (ps + 2 * ed) / 3

        ri = self._fval("RI")
        pi = self._fval("PI")
        sd = self._fval("SD")
        auto_used = False
        self._pi_autocalc = False

        if self.auto_var.get():
            if ri is None:
                ri = (ps - ed) / ps
                self.vars["RI"].set(f"{ri:.3f}")
                auto_used = True
            if sd is None:
                sd = ps / ed
                self.vars["SD"].set(f"{sd:.2f}")
                auto_used = True
            if pi is None:
                # PI de Gosling = (PS - ED) / TAMV
                # En este equipo, "TAmax" contiene la TAMV (velocidad media).
                # Se usa tamax (= TAMV del equipo) como denominador.
                # TAmean es campo redundante; se usa solo si tamax no fue ingresado.
                v_pi = tamax if (tamax is not None and tamax > 0) else tamean
                if v_pi and v_pi > 0:
                    pi = (ps - ed) / v_pi
                    self.vars["PI"].set(f"{pi:.3f}")
                    auto_used = True
                    self._pi_autocalc = True
                else:
                    self._pi_autocalc = False
            else:
                self._pi_autocalc = False

        for k in self.entries:
            self._validar(k, self.vars[k].get())

        # ── Onda Doppler (4 ciclos) ───────────────────────────────────────────
        period = (60.0 / hr) if hr and hr > 20 else 1.0
        t_norm = [0.00, 0.04, 0.10, 0.18, 0.26, 0.38, 0.52, 0.68, 0.84, 1.00]
        v_norm = [0.00, 0.22, 0.60, 1.00, 0.84, 0.62, 0.48, 0.36, 0.24, 0.00]
        all_t, all_v = [], []
        for c in range(4):
            off = c * period
            for tn, vn in zip(t_norm, v_norm):
                all_t.append(off + tn * period)
                all_v.append(ed + (ps - ed) * vn)

        # ── PIC / CPP — Czosnyka et al. 1998 (solo ACM) ────────────────────
        # CPPe = MAP × (ED / TAMV) + 14
        # PICe = MAP − CPPe
        # NOTA NOMENCLATURA DE EQUIPO:
        #   Este equipo reporta como "TAmax" la velocidad MEDIA (TAMV / mean flow velocity).
        #   La fórmula de Czosnyka original usa TAP (Time-Averaged Peak = envolvente pico).
        #   En ausencia de TAP explícito, se usa TAMV como mejor aproximación disponible,
        #   pero hay que tener en cuenta que Czosnyka fue validado con TAP.
        #   El campo "TAmean" de este equipo NO se usa aquí.
        arteria_actual = self.arteria_var.get()
        # pam ya viene de arriba (puede ser None si no se ingresó)

        # ── Estenosis de ACM (Doppler) — solo aplica a la arteria cerebral media ──
        # VM = TAMV = campo "TAmax" del equipo; respaldo en "TAmean" si TAmax
        # no fue ingresado (misma lógica de respaldo usada para IL/Czosnyka/PI).
        vm_estenosis = tamax if (tamax is not None and tamax > 0) else \
                       (tamean if (tamean is not None and tamean > 0) else None)
        if arteria_actual == "ACM":
            self._estenosis_sev, self._estenosis_texto, estenosis_color = \
                self._evaluar_estenosis_acm_doppler(vm_estenosis)
            self._estenosis_texto = _limpiar_cita_guia_estenosis(self._estenosis_texto)
            self._estenosis_vm = vm_estenosis
            etiquetas_sev = {"sin_estenosis": "Sin estenosis",
                              "moderada": "MODERADA", "severa": "SEVERA"}
            self.estenosis_lbl.config(
                text="Estenosis ACM: " + etiquetas_sev.get(self._estenosis_sev, "—"),
                fg=estenosis_color)
        else:
            self._estenosis_sev, self._estenosis_texto, self._estenosis_vm = None, "", None
            self.estenosis_lbl.config(
                text="Estenosis ACM: — (criterio solo válido para ACM)", fg=TEXT2)

        pic_label, pic_color = "No calculable", TEXT2
        cpp_str = ""
        ppc_czos_val = None

        # Para Czosnyka usamos el campo "TAmax" del equipo (= TAMV en este ecógrafo).
        # TAmean del equipo es un segundo campo redundante que NO se usa aquí.
        v_tap = tamax if (tamax is not None and tamax > 0) else None

        # ── Bellner 2004: ICP = 10.93 × PI − 1.28  (Surg Neurol 62:45-51) ────────────
        # Solo requiere PI de Gosling. No requiere PAM. Solo ACM.
        # Correlación r=0.938 en población original (HSA+TBI). SD=±2.5 mmHg (5–40 mmHg).
        # LIMITACIón: no replicado con igual precisión en otras cohortes. Orientativo.
        pic_bellner_str, pic_bellner_color = "", TEXT2
        pb = None   # valor numérico PIC-Bellner (mmHg), se usa luego para PPC-Bellner
        # Si el PI fue autocalculado con datos incoherentes (ED >= TAMV), no es fiable.
        pi_incoherente = (getattr(self, "_pi_autocalc", False)
                          and tamax is not None and tamax > 0 and ed >= tamax)
        # ¿El PI proviene de una TAMV APROXIMADA? Ocurre si fue autocalculado y no
        # había TAmax medido y la TAmean también fue aproximada (PS+2·ED)/3. En ese
        # caso Bellner se basa en datos estimados → se marca y no se afirma "Normal".
        pi_de_aprox = (getattr(self, "_pi_autocalc", False)
                       and (tamax is None or tamax <= 0) and tamean_aprox)
        if arteria_actual == "ACM" and pi is not None and pi > 0 and not pi_incoherente:
            pb = 10.93 * pi - 1.28
            if pb > 20:
                pic_bellner_color = RED
                pic_bellner_str = f"ALTA  {pb:.0f} mmHg"
            elif pb > 15:
                pic_bellner_color = YELLOW
                pic_bellner_str = f"Elevada  {pb:.0f} mmHg"
            elif pb >= 5:
                pic_bellner_color = GREEN
                pic_bellner_str = f"Normal  {pb:.0f} mmHg"
            else:
                pic_bellner_color = TEXT2
                pic_bellner_str = f"{pb:.0f} mmHg (revisar PI)"
            if pi_de_aprox:
                # Rigor: no afirmar normalidad con un PI derivado de VM estimada.
                pic_bellner_str += "  ⚠ PI desde VM aprox."
                if pic_bellner_color == GREEN:
                    pic_bellner_color = YELLOW

        if arteria_actual != "ACM":
            pic_label = "Solo aplicable en ACM"
            pic_color = TEXT2
            cpp_str = "Solo aplicable en ACM"
        elif pam is None:
            pic_label = "Ingrese PAM"
            pic_color = TEXT2
            cpp_str = "Czosnyka requiere PAM (mmHg)"
        elif v_tap is None:
            pic_label = "Ingrese TAmax (VM del equipo)"
            pic_color = TEXT2
            cpp_str = "Czosnyka requiere TAmax (VM del equipo)"
        elif ed >= v_tap:
            # Coherencia física: en flujo arterial ED < TAMV siempre.
            pic_label = "Datos incoherentes (ED ≥ TAmax)"
            pic_color = YELLOW
            cpp_str = "Revisar trazado de envolvente"
        else:
            cppe = pam * (ed / v_tap) + 14
            pice = pam - cppe

            # Umbrales: normal < 15 mmHg (fisiológico adulto),
            # ≥20 mmHg = umbral de alarma clínica (BTF / NCS 2023)
            if pice > 20:
                pic_label = f"ALTA  {pice:.0f} mmHg"
                pic_color = RED
            elif pice > 15:
                pic_label = f"Elevada  {pice:.0f} mmHg"
                pic_color = YELLOW
            elif pice >= 5:
                pic_label = f"Normal  {pice:.0f} mmHg"
                pic_color = GREEN
            else:
                # PICe muy baja o negativa → fórmula fuera de rango fisiológico
                pic_label = f"{pice:.0f} mmHg  (revisar datos)"
                pic_color = TEXT2

            cpp_str = f"CPP estimada ≈ {cppe:.0f} mmHg"
            if cppe < 50:
                cpp_str += "  ⚠ Muy baja"
            elif cppe < 60:
                cpp_str += "  ⚠ Límite"
            ppc_czos_val = cppe

        # ── PPC (Presión de Perfusión Cerebral) ────────────────────────────────
        # PPC-Czosnyka: es la misma CPPe ya calculada arriba (Czosnyka 1998).
        # PPC-Bellner : PPC = PAM − PICe(Bellner) — requiere PAM además del PI
        #               (Bellner solo da PIC; la PPC se deriva restando de PAM).
        #               Ambas son estimaciones NO invasivas, solo válidas en ACM.
        ppc_bellner_str, ppc_bellner_color = "", TEXT2
        if pam is not None and pb is not None:
            ppc_b = pam - pb
            ppc_bellner_val = round(ppc_b, 0)
            if ppc_b < 50:
                ppc_bellner_color = RED
                ppc_bellner_str = f"{ppc_b:.0f} mmHg  ⚠ Muy baja"
            elif ppc_b < 60:
                ppc_bellner_color = YELLOW
                ppc_bellner_str = f"{ppc_b:.0f} mmHg  ⚠ Límite"
            else:
                ppc_bellner_color = GREEN
                ppc_bellner_str = f"{ppc_b:.0f} mmHg"
        else:
            ppc_bellner_val = None
            if arteria_actual != "ACM":
                ppc_bellner_str = "Solo aplicable en ACM"
            elif pam is None:
                ppc_bellner_str = "Requiere PAM"
            else:
                ppc_bellner_str = "Requiere PI (Bellner)"

        # ── Datos del paciente ────────────────────────────────────────────────
        nombre  = self.ent_nombre.get().strip() or "Sin nombre"
        hc      = self.ent_hc.get().strip()     or "—"
        lado    = self.lado_var.get()
        arteria = self.arteria_var.get()
        now_str = datetime.now().strftime("%d/%m/%Y  %H:%M:%S")

        # ── Habilitar pestaña Lindegaard con datos del duplex ─────────────────
        self._habilitar_tab_lindegaard(nombre, arteria, lado, tamean, tamax)

        # ── Construir figura ──────────────────────────────────────────────────
        if self.canvas_fig:
            self.canvas_fig.get_tk_widget().destroy()
            self.canvas_fig = None
        if self.last_fig:
            _cerrar_figura(self.last_fig)
            self.last_fig = None
        if self.placeholder:
            self.placeholder.place_forget()

        # ── Lindegaard data ───────────────────────────────────────────────────
        lind_data = self._get_lindegaard_data()

        # ── Estados de cada índice (bajo/normal/alto) ─────────────────────────
        # Se calcula ANTES de dibujar la figura porque también lo usa la
        # narración por voz (_narrar_resultado), fuera de _dibujar(). Los
        # colores aquí son solo para la tabla en pantalla; la narración usa
        # únicamente el texto del estado, no el color.
        _r = self.RANGOS_ARTERIA.get(arteria, self.RANGOS_ARTERIA["ACM"])
        _ps_mn, _ps_mx, _ed_mn, _ed_mx, _tm_mn, _tm_mx, _pi_mn, _pi_mx = _r

        def _estado(val, vmin, vmax):
            """Devuelve (símbolo, color) según si el valor está bajo/normal/alto."""
            if val is None:
                return ("—", TEXT2)
            if val < vmin:
                return ("↓ bajo", ACCENT)
            if val > vmax:
                return ("↑ alto", RED)
            return ("normal", GREEN)

        estados = {
            "PS":     _estado(ps,     _ps_mn,      _ps_mx),
            "ED":     _estado(ed,     _ed_mn,      _ed_mx),
            "RI":     _estado(ri,     0.45,         0.75),
            "PI":     _pi_estado(pi, arteria),   # criterio único por arteria
            "SD":     _estado(sd,     1.4,          3.8),
            "TAmean": _estado(tamean, _tm_mn,      _tm_mx),
            "TAmax":  _estado(tamax,  _tm_mn,      _tm_mx),  # TAmax = TAMV: mismos rangos que TAmean
            "HR":     _estado(hr,     45,           120),
        }

        # Datos para el aviso de alerta en la figura: mismas claves y criterio
        # que last_result → _alertas_tcd (monitor principal).
        _datos_alerta = {
            "arteria":       arteria,
            "TAmax":         tamax,
            "ED":            ed,
            "IL":            round(lind_data[0], 2) if lind_data else None,
            "VM_aproximada": not (tamax is not None and tamax > 0),
            "Estenosis_VM":  self._estenosis_vm,
        }

        # ── Construcción de la figura (parametrizable pantalla/imprimible) ───
        # _dibujar() se guarda en self._dibujar_figura para poder regenerar
        # más tarde una versión imprimible (fondo blanco, bajo consumo de
        # tinta) sin tener que repetir toda la lógica de cálculo de arriba.
        def _dibujar(printable=False):
            pal = _paleta(printable)
            BG, BG3, BORDER       = pal["BG"], pal["BG3"], pal["BORDER"]
            ACCENT, GREEN, YELLOW = pal["ACCENT"], pal["GREEN"], pal["YELLOW"]
            RED, PINK             = pal["RED"], pal["PINK"]
            TEXT, TEXT2           = pal["TEXT"], pal["TEXT2"]

            # Colores precalculados fuera de esta función (con la paleta de
            # pantalla): se traducen a la paleta imprimible para no arrastrar
            # verdes/amarillos brillantes ilegibles en blanco y negro.
            _pic_color         = _map_color(pic_color, printable)
            _pic_bellner_color = _map_color(pic_bellner_color, printable)

            # ── Layout: 3 filas — onda / [tabla + PIC + (IL)] / pie paciente ─
            fig = _nueva_figura(figsize=(11, 8), facecolor=BG)
            if lind_data:
                gs = gridspec.GridSpec(3, 3, figure=fig,
                                       height_ratios=[2.8, 1.1, 0.18],
                                       hspace=0.38, wspace=0.30)
            else:
                gs  = gridspec.GridSpec(3, 2, figure=fig,
                                        height_ratios=[2.8, 1.1, 0.18],
                                        hspace=0.38, wspace=0.30)

            # — Subplot 1: onda principal (siempre full-width) —
            ax1 = fig.add_subplot(gs[0, :])
            ax1.set_facecolor(BG)
            # Trazo fino sin relleno bajo la curva
            ax1.plot(all_t, all_v, color=PINK, linewidth=1.5, zorder=3)
            ax1.axhline(ps,     color=GREEN,  ls="--", lw=1.5,
                        label="PS = " + str(round(ps, 1)) + " cm/s")
            ax1.axhline(ed,     color=ACCENT, ls="--", lw=1.5,
                        label="ED = " + str(round(ed, 1)) + " cm/s")
            # Línea de Vmean: priorizar TAmax (TAP, estándar TCD neuro), si no hay
            # mostrar TAmean como fallback con etiqueta clara.
            if tamax is not None and tamax > 0:
                ax1.axhline(tamax, color=YELLOW, ls=":",  lw=1.5,
                            label="TAmax = " + str(round(tamax, 1)) + " cm/s")
            elif tamean and tamean > 0:
                ax1.axhline(tamean, color=YELLOW, ls=":",  lw=1.5,
                            label=(("VM aprox. (PS+2ED)/3 = " if tamean_aprox
                                    else "TAmean = ")
                                   + str(round(tamean, 1)) + " cm/s"))

            for c in range(1, 4):
                ax1.axvline(c * period, color=BORDER, lw=1, ls=":", alpha=0.5)

            # Anotaciones en primer ciclo
            peak_t = t_norm[3] * period
            ax1.annotate("PS " + str(round(ps, 0)),
                         xy=(peak_t, ps), xytext=(peak_t + 0.04, ps + 5),
                         color=GREEN, fontsize=8,
                         arrowprops=dict(arrowstyle="->", color=GREEN, lw=1))
            ax1.annotate("ED " + str(round(ed, 0)),
                         xy=(0.85 * period, ed), xytext=(0.85 * period + 0.04, ed - 7),
                         color=ACCENT, fontsize=8,
                         arrowprops=dict(arrowstyle="->", color=ACCENT, lw=1))

            title = "Onda Doppler Transcraneal  —  " + arteria + " " + lado
            if auto_used:
                title += "   [índices auto-calculados]"
            ax1.set_title(title, color=TEXT, fontsize=11, fontweight="bold", pad=12)
            ax1.set_xlabel("Tiempo (s)",       color=TEXT2, fontsize=9)
            ax1.set_ylabel("Velocidad (cm/s)", color=TEXT2, fontsize=9)
            ax1.tick_params(colors=TEXT2, labelsize=8)
            for sp in ax1.spines.values():
                sp.set_edgecolor(BORDER)
            leg = ax1.legend(loc="upper right", fontsize=8,
                             facecolor=BG3, edgecolor=BORDER, labelcolor=TEXT)
            leg.get_frame().set_alpha(0.92)
            _alertas_en_ax(ax1, _datos_alerta, RED, YELLOW, printable,
                           x=0.01, y=0.97, fontsize=10)

            # — Subplot 2: tabla de índices —
            ax2 = fig.add_subplot(gs[1, 0])
            ax2.set_facecolor(BG)
            ax2.axis("off")

            # Convertir a strings ANTES de usarlos en la tabla (evita el bug)
            ri_s    = str(round(ri, 3))    if ri    is not None else "—"
            pi_s    = str(round(pi, 3))    if pi    is not None else "—"
            sd_s    = str(round(sd, 2))    if sd    is not None else "—"
            tamax_s = str(round(tamax, 1)) if tamax is not None else "—"

            # Rangos para la arteria seleccionada
            _r = self.RANGOS_ARTERIA.get(arteria, self.RANGOS_ARTERIA["ACM"])
            _ps_mn, _ps_mx, _ed_mn, _ed_mx, _tm_mn, _tm_mx, _pi_mn, _pi_mx = _r

            def _estado_fig(val, vmin, vmax):
                """(símbolo, color) con la paleta local de esta figura."""
                if val is None:
                    return ("—", TEXT2)
                if val < vmin:
                    return ("↓ bajo", ACCENT)
                if val > vmax:
                    return ("↑ alto", RED)
                return ("normal", GREEN)

            estados_fig = {
                "PS":     _estado_fig(ps,     _ps_mn,      _ps_mx),
                "ED":     _estado_fig(ed,     _ed_mn,      _ed_mx),
                "RI":     _estado_fig(ri,     0.45,         0.75),
                "PI":     _pi_estado(pi, arteria, pal),   # criterio único
                "SD":     _estado_fig(sd,     1.4,          3.8),
                "TAmean": _estado_fig(tamean, _tm_mn,      _tm_mx),
                "TAmax":  _estado_fig(tamax,  _tm_mn,      _tm_mx),  # TAmax = TAMV: mismos rangos que TAmean
                "HR":     _estado_fig(hr,     45,           120),
            }

            filas = [
                ("PS",     str(round(ps, 1)),     "cm/s",  estados_fig["PS"]),
                ("ED",     str(round(ed, 1)),     "cm/s",  estados_fig["ED"]),
                ("RI",     ri_s,                  "",      estados_fig["RI"]),
                ("PI",     pi_s,                  "",      estados_fig["PI"]),
                ("S/D",    sd_s,                  "",      estados_fig["SD"]),
                ("TAmean", ("~" if tamean_aprox else "") + str(round(tamean, 1)),
                                                  "cm/s",  estados_fig["TAmean"]),
                ("TAmax",  tamax_s,               "cm/s",  estados_fig["TAmax"]),
                ("HR",     _fnum(hr, 0),          "lpm",   estados_fig["HR"]),
                ("PAM",    _fnum(pam, 0),         "mmHg",  ("—", TEXT2) if pam is None
                                                           else ("", TEXT2)),
                ("PaCO₂",  _fnum(co2, 0),         "mmHg",  _estado_fig(co2, *CO2_NORMAL)),
            ]

            N = len(filas)
            ax2.set_xlim(0, 4)
            ax2.set_ylim(0, N + 0.9)

            # Encabezado de columnas
            ax2.text(0.05, N + 0.65, "ÍNDICES",
                     color=ACCENT, fontsize=8, fontweight="bold", va="center")
            ax2.text(1.55, N + 0.65, "Valor",
                     color=TEXT2, fontsize=7, va="center", ha="center")
            ax2.text(2.55, N + 0.65, "Unidad",
                     color=TEXT2, fontsize=7, va="center", ha="center")
            ax2.text(3.95, N + 0.65, "Estado",
                     color=TEXT2, fontsize=7, va="center", ha="right")
            ax2.axhline(N + 0.45, color=ACCENT, lw=0.8, alpha=0.5)

            for i, (label, val, unit, (est_txt, est_col)) in enumerate(reversed(filas)):
                y = i + 0.5
                bg_alpha = 0.06 if i % 2 == 0 else 0.0
                if bg_alpha:
                    ax2.axhspan(y - 0.5, y + 0.5, color=TEXT, alpha=bg_alpha, zorder=0)
                ax2.text(0.05, y, label, color=TEXT2, fontsize=8,  va="center")
                ax2.text(1.55, y, val,   color=TEXT,  fontsize=9,  va="center",
                         fontweight="bold", ha="center")
                ax2.text(2.55, y, unit,  color=TEXT2, fontsize=7,  va="center", ha="center")
                ax2.text(3.95, y, est_txt, color=est_col, fontsize=8, va="center",
                         ha="right", fontweight="bold")
                ax2.axhline(y + 0.5, color=BORDER, lw=0.4, alpha=0.4)

            # — Subplot 3: PIC / CPP —
            ax3 = fig.add_subplot(gs[1, 1])
            ax3.set_facecolor(BG)
            ax3.axis("off")
            ax3.text(0.5, 0.99, "ESTIMACIÓN PIC  (no invasiva)",
                     transform=ax3.transAxes, color=ACCENT,
                     fontsize=8, fontweight="bold", ha="center", va="top")

            # ─ Czosnyka 1998 ─
            ax3.text(0.02, 0.88, "Czosnyka 1998",
                     transform=ax3.transAxes, color=TEXT2, fontsize=7,
                     fontweight="bold", va="top")
            ax3.text(0.02, 0.81, "CPPe = PAM×(ED/TAMV)+14",
                     transform=ax3.transAxes, color=TEXT2, fontsize=7,
                     va="top")
            ax3.text(0.02, 0.73, pic_label,
                     transform=ax3.transAxes, color=_pic_color,
                     fontsize=11, fontweight="bold", va="top")
            if cpp_str:
                ax3.text(0.02, 0.63, cpp_str,
                         transform=ax3.transAxes, color=TEXT,
                         fontsize=8, va="top")

            # línea divisoria
            ax3.plot([0.02, 0.98], [0.54, 0.54], color=BORDER, lw=0.6, alpha=0.7,
                     transform=ax3.transAxes, clip_on=False)

            # ─ Bellner 2004 ─
            ax3.text(0.02, 0.52, "Bellner 2004",
                     transform=ax3.transAxes, color=TEXT2, fontsize=7,
                     fontweight="bold", va="top")
            ax3.text(0.02, 0.45, "ICP = 10.93×PI − 1.28",
                     transform=ax3.transAxes, color=TEXT2, fontsize=7,
                     va="top")
            if pic_bellner_str:
                ax3.text(0.02, 0.37, pic_bellner_str,
                         transform=ax3.transAxes, color=_pic_bellner_color,
                         fontsize=11, fontweight="bold", va="top")
                pi_origen = "calc." if getattr(self, "_pi_autocalc", False) else "medido"
                ax3.text(0.02, 0.27, f"PI = {pi_s} ({pi_origen})",
                         transform=ax3.transAxes, color=TEXT2,
                         fontsize=8, va="top")
            else:
                ax3.text(0.02, 0.37, "Requiere PI (ACM)",
                         transform=ax3.transAxes, color=TEXT2,
                         fontsize=8, va="top")

            # nota limitación (anclada al borde inferior, sin solaparse con el PI)
            ax3.text(0.5, 0.00,
                     "Normal < 15 · Alarma ≥ 20 mmHg\n"
                     "Czosnyka usa TAMV · Bellner orientativo",
                     transform=ax3.transAxes, color=TEXT2,
                     fontsize=6, ha="center", va="bottom")

            # — Subplot 4: Índice de Lindegaard (solo si activo) —
            if lind_data:
                il, lind_acm, lind_aci, _, lind_col = lind_data
                lind_col = _map_color(lind_col, printable)
                ax_il = fig.add_subplot(gs[1, 2])
                ax_il.set_facecolor(BG)
                ax_il.axis("off")

                # Título
                ax_il.text(0.5, 0.97, "ÍNDICE DE LINDEGAARD",
                           transform=ax_il.transAxes, color=ACCENT,
                           fontsize=8, fontweight="bold", ha="center", va="top")

                # Línea separadora
                ax_il.plot([0.05, 0.95], [0.88, 0.88], color=ACCENT, lw=0.6, alpha=0.5,
                           transform=ax_il.transAxes, clip_on=False)

                # IL valor grande
                ax_il.text(0.5, 0.72,
                           f"{il:.2f}",
                           transform=ax_il.transAxes, color=lind_col,
                           fontsize=22, fontweight="bold", ha="center", va="center")

                # Interpretación
                if il > 6:
                    interp = "Vasoespasmo ACM\nSEVERO"
                elif il > 3:
                    interp = "Vasoespasmo ACM\nprobable"
                else:
                    interp = "Sin vasoespasmo\nACM"
                ax_il.text(0.5, 0.48, interp,
                           transform=ax_il.transAxes, color=lind_col,
                           fontsize=8, fontweight="bold", ha="center", va="center")

                # Datos usados
                ax_il.text(0.5, 0.30,
                           f"TAMV ACM {lind_acm:.0f} cm/s",
                           transform=ax_il.transAxes, color=TEXT2,
                           fontsize=7, ha="center")
                ax_il.text(0.5, 0.22,
                           f"TAMV ACI-ex {lind_aci:.0f} cm/s",
                           transform=ax_il.transAxes, color=TEXT2,
                           fontsize=7, ha="center")
                ax_il.text(0.5, 0.14,
                           "(vent. submandibular · 45–50 mm)",
                           transform=ax_il.transAxes, color=TEXT2,
                           fontsize=6, ha="center")

                # Fórmula y referencia
                ax_il.text(0.5, 0.07,
                           "IL = TAMV ACM / TAMV ACI-ex",
                           transform=ax_il.transAxes, color=TEXT2,
                           fontsize=6, ha="center")
                ax_il.text(0.5, 0.01,
                           "> 6 Severo · > 3 Probable  |  Lindegaard 1989",
                           transform=ax_il.transAxes, color=TEXT2,
                           fontsize=6, ha="center")

            # — Fila inferior: datos del paciente —
            ax_pie = fig.add_subplot(gs[2, :])
            ax_pie.set_facecolor(BG)
            ax_pie.axis("off")
            # Línea separadora sutil
            ax_pie.plot([0.02, 0.98], [1.0, 1.0], color=BORDER, lw=0.6, alpha=0.7,
                        transform=ax_pie.transAxes, clip_on=False)
            # Nombre (izquierda, mayor énfasis)
            ax_pie.text(0.02, 0.42, nombre,
                        transform=ax_pie.transAxes,
                        color=TEXT, fontsize=9, fontweight="bold", va="center")
            # HC y timestamp (centro-izquierda, secundario)
            ax_pie.text(0.02, -0.12,
                        "HC: " + hc + "   ·   " + now_str + "   ·   " + arteria + " " + lado,
                        transform=ax_pie.transAxes,
                        color=TEXT2, fontsize=8, va="center")

            # tight_layout no es compatible con subplots de texto puro; el espaciado
            # ya está definido en GridSpec (hspace/wspace). Ajuste de márgenes manual.
            fig.subplots_adjust(left=0.06, right=0.97, top=0.95, bottom=0.04)
            return fig

        self._dibujar_figura = _dibujar
        fig = _dibujar(False)

        self.canvas_fig = FigureCanvasTkAgg(fig, master=self.graph_frame)
        self.canvas_fig.draw()
        self.canvas_fig.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self.last_fig = fig

        # Auto-guardar en subcarpeta informes_tcd/ (carpeta SIEMPRE escribible).
        # Un fallo de guardado NUNCA debe abortar el análisis ya visible en
        # pantalla: por eso savefig va protegido con try/except.
        safe_nombre = "".join(
            c for c in nombre if c.isalnum() or c in " _-"
        ).strip().replace(" ", "_") or "paciente"
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_dir = _carpeta_salida("informes_tcd")
        auto_path = os.path.join(out_dir, safe_nombre + "_" + ts + ".png")
        try:
            fig.savefig(auto_path, dpi=180, bbox_inches="tight", facecolor=BG)
        except OSError as e:
            auto_path = None
            self._set_status("No se pudo auto-guardar el PNG: " + str(e)[:80],
                             YELLOW)

        # Auto-guardar también la versión imprimible (fondo blanco, bajo
        # consumo de tinta) — es la que normalmente se pega/imprime.
        auto_path_imp = None
        if auto_path:
            try:
                auto_path_imp = os.path.join(
                    os.path.dirname(auto_path) or ".",
                    os.path.splitext(os.path.basename(auto_path))[0] + "_imprimible.png")
                fig_imp = self._dibujar_figura(True)
                fig_imp.savefig(auto_path_imp, dpi=180, bbox_inches="tight",
                                 facecolor=fig_imp.get_facecolor())
                _cerrar_figura(fig_imp)
            except Exception:
                auto_path_imp = None

        # Actualizar informe
        self._actualizar_informe(
            ps, ed, ri, pi, sd, tamean, tamax, hr,
            pic_label, cpp_str, pam,
            now_str, nombre, hc, lado, arteria,
            lind_data,
            pic_bellner_str=pic_bellner_str,
            co2=co2, tamean_aprox=tamean_aprox,
            alertas=_alertas_tcd(_datos_alerta))

        # auto_path puede ser None si falló el auto-guardado (ver arriba).
        # Concatenar None reventaba aquí con TypeError y _safe_generar mostraba
        # "Error crítico" sobre un análisis que en realidad SÍ se generó bien.
        if auto_path:
            self._set_status("✓ Generado  ·  Auto-guardado: " + auto_path, GREEN)
        else:
            self._set_status("✓ Generado  ·  sin auto-guardado (ver aviso)", YELLOW)

        # Aviso por CO₂ fuera de rango (35–45 mmHg)
        if co2 is not None and not (35 <= co2 <= 45):
            tipo = "hipocapnia" if co2 < 35 else "hipercapnia"
            self._set_status(
                "⚠ CO₂ " + f"{co2:g}" + " mmHg (" + tipo + "): PI / Bellner / "
                "Czosnyka / VM pueden estar confundidos — interpretar con cautela",
                YELLOW)

        # ── Exportar resumen a pantalla principal ─────────────────────────────
        self.last_result = {
            "timestamp":  now_str,
            "paciente":   nombre,
            "hc":         hc,
            "arteria":    arteria,
            "lado":       lado,
            "PS":         ps,
            "ED":         ed,
            "RI":         round(ri, 3)    if ri    is not None else None,
            "PI":         round(pi, 3)    if pi    is not None else None,
            "TAmax":      tamax,
            # Solo el valor INGRESADO. Si faltaba, tamean se calculó como
            # (PS+2·ED)/3 para uso interno y no debe mostrarse como medido.
            "TAmean":     None if tamean_aprox else tamean,
            "TAmean_aproximada": tamean_aprox,
            "HR":         hr,
            "SD":         round(sd, 2) if sd is not None else None,
            "PAM":        pam,
            "PaCO2":      co2,
            "PIC_label":  pic_label,
            "PIC_bellner":pic_bellner_str,
            "PPC_czosnyka_val": round(ppc_czos_val, 0) if ppc_czos_val is not None else None,
            "PPC_czosnyka_str": cpp_str,
            "PPC_bellner_val":  ppc_bellner_val,
            "PPC_bellner_str":  ppc_bellner_str,
            "PPC_bellner_color": ppc_bellner_color,
            "auto_path":  auto_path,
            "auto_path_imp": auto_path_imp,
            # Para las alertas del monitor principal (_alertas_tcd):
            "IL":            round(lind_data[0], 2) if lind_data else None,
            "VM_aproximada": not (tamax is not None and tamax > 0),
            "Estenosis_severidad": self._estenosis_sev,
            "Estenosis_texto":     self._estenosis_texto,
            "Estenosis_VM":        self._estenosis_vm,
        }
        if self.on_result:
            try:
                self.on_result(self.last_result)
            except Exception:
                pass

        # Narración automática (Edge TTS, opcional)
        self._narrar_resultado(ps, ed, pi, pic_label, arteria, lado, estados,
                               pic_bellner_str=pic_bellner_str, lind_data=lind_data,
                               alertas=_alertas_tcd(_datos_alerta), co2=co2)

    # ── Estenosis de ACM (Doppler transcraneal) ─────────────────────────────────
    #   Criterios según Guía Técnica de Procedimiento INCN:
    #   "Ecocardiografía Doppler, Transcraneal" CPMS 93320.01
    #   (RD N° 253-2025-DG-INCN, aprobada 27/06/2025), sección 6.3.3.
    #
    #   Usando el equipo de DOPPLER (no dúplex color) el examen es positivo
    #   para estenosis de ACM si la VELOCIDAD MEDIA de flujo sanguíneo cerebral:
    #       > 100 cm/s  →  estenosis MODERADA
    #       > 128 cm/s  →  estenosis SEVERA
    #   La guía también exige, para hablar de estenosis con más certeza, que la
    #   relación VM segmento estenótico / preestenótico sea ≥ 2 y que la
    #   asimetría entre lados no supere 30% — ninguno de los dos es medible con
    #   una sola medición unilateral, así que se documentan como advertencia,
    #   no como criterio diagnóstico adicional.
    #
    #   La "velocidad media" que exige la guía es la TAMV (Time-Averaged Mean
    #   Velocity). En este equipo, ese valor se ingresa en el campo "TAmax"
    #   (ver notas de nomenclatura en _generar); si no se ingresó, se usa
    #   "TAmean" como respaldo.
    UMBRAL_ACM_MODERADA = 100.0   # cm/s
    UMBRAL_ACM_SEVERA   = 128.0   # cm/s

    def _evaluar_estenosis_acm_doppler(self, vm):
        """Devuelve (severidad, texto_interpretativo, color) para VM de ACM.

        severidad ∈ {None, "sin_estenosis", "moderada", "severa"}
        None → no hay VM disponible para evaluar.
        """
        if vm is None or vm <= 0:
            return (None,
                    "No se dispone de una velocidad media (TAMV) válida para "
                    "valorar criterios de estenosis de la arteria cerebral media.",
                    TEXT2)

        if vm > self.UMBRAL_ACM_SEVERA:
            texto = (
                f"La velocidad media de flujo sanguíneo cerebral en la arteria "
                f"cerebral media es de {vm:.1f} cm/s, valor que supera el "
                f"umbral de {self.UMBRAL_ACM_SEVERA:.0f} cm/s, hallazgo "
                f"compatible con ESTENOSIS SEVERA de la arteria cerebral media.")
            return ("severa", texto, RED)

        if vm > self.UMBRAL_ACM_MODERADA:
            texto = (
                f"La velocidad media de flujo sanguíneo cerebral en la arteria "
                f"cerebral media es de {vm:.1f} cm/s, valor comprendido entre "
                f"{self.UMBRAL_ACM_MODERADA:.0f} y {self.UMBRAL_ACM_SEVERA:.0f} "
                f"cm/s, hallazgo compatible con ESTENOSIS MODERADA de la "
                f"arteria cerebral media.")
            return ("moderada", texto, YELLOW)

        texto = (
            f"La velocidad media de flujo sanguíneo cerebral en la arteria "
            f"cerebral media es de {vm:.1f} cm/s, valor por debajo del umbral "
            f"de {self.UMBRAL_ACM_MODERADA:.0f} cm/s, sin evidencia de "
            f"estenosis significativa de la arteria cerebral media.")
        return ("sin_estenosis", texto, GREEN)

    # ── Informe de texto ──────────────────────────────────────────────────────
    def _actualizar_informe(self, ps, ed, ri, pi, sd, tamean, tamax, hr,
                            pic_label, cpp_str, pam,
                            now_str, nombre, hc, lado, arteria="ACM",
                            lind_data=None, pic_bellner_str="",
                            co2=None, tamean_aprox=False, alertas=()):
        # Construir strings seguros ANTES del f-string
        ri_s    = str(round(ri, 3))    if ri    is not None else "—"
        pi_s    = str(round(pi, 3))    if pi    is not None else "—"
        sd_s    = str(round(sd, 2))    if sd    is not None else "—"
        tamax_s = str(round(tamax, 1)) if tamax is not None else "—"

        lineas = [
            "══ INFORME  " + now_str + " ══",
            "Paciente: " + nombre + "   HC: " + hc + "   " + arteria + " " + lado,
            ("PS=" + str(round(ps, 1))
             + "  ED=" + str(round(ed, 1))
             + "  RI=" + ri_s
             + "  PI=" + pi_s
             + "  S/D=" + sd_s
             # Nomenclatura del equipo: TAmax = TAMV = VM. TAmean es otro campo.
             + "  VM/TAMV(TAmax)=" + tamax_s
             + "  TAmean=" + (("~" + str(round(tamean, 1)) + " (aprox.)")
                              if tamean_aprox else str(round(tamean, 1)))
             + "  HR=" + _fnum(hr, 0)),
            "PIC Czosnyka: " + pic_label,
        ]
        if cpp_str:
            lineas.append(cpp_str)
        if pic_bellner_str:
            lineas.append("PIC Bellner:  " + pic_bellner_str + "  (ICP=10.93×PI−1.28)")
        if pam is not None:
            lineas.append("PAM = " + str(round(pam)) + " mmHg")
        lineas.append("PaCO2/EtCO2 = " + _co2_txt({"PaCO2": co2}))
        for _, t in (alertas or ()):
            lineas.append("ALERTA: " + t)
        if lind_data:
            il, lind_acm, lind_aci, lind_txt, _ = lind_data
            lineas.append(
                "Índice de Lindegaard = " + str(round(il, 2))
                + "  (ACM " + str(round(lind_acm, 1))
                + " / ACI " + str(round(lind_aci, 1)) + " cm/s)"
                + "  →  " + lind_txt)
        if arteria == "ACM" and getattr(self, "_estenosis_texto", ""):
            lineas.append("Estenosis ACM: " + self._estenosis_texto)

        self._informe_texto = "\n".join(lineas)

    # ── Guardar PNG (manual) ──────────────────────────────────────────────────
    def _guardar_png(self):
        if not self.last_fig:
            messagebox.showinfo("Sin gráfica", "Genera primero el análisis."); return
        nombre  = self.ent_nombre.get().strip() or "paciente"
        ts      = datetime.now().strftime("%Y%m%d_%H%M%S")
        default = nombre.replace(" ", "_") + "_" + ts + ".png"
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG", "*.png"), ("Todos", "*.*")],
            initialfile=default,
            title="Guardar gráfica Doppler")
        if path:
            try:
                self.last_fig.savefig(path, dpi=200, bbox_inches="tight",
                                      facecolor=BG)
                self._set_status("✓ Guardado: " + path, GREEN)
            except OSError as e:
                messagebox.showerror("Error al guardar",
                                     "No se pudo guardar el PNG:\n" + str(e))
                self._set_status("✗ No se pudo guardar", RED)

    # ── Guardar PNG imprimible (fondo blanco, bajo consumo de tinta) ──────────
    def _guardar_png_imprimible(self):
        if not getattr(self, "_dibujar_figura", None):
            messagebox.showinfo("Sin gráfica", "Genera primero el análisis."); return
        nombre  = self.ent_nombre.get().strip() or "paciente"
        ts      = datetime.now().strftime("%Y%m%d_%H%M%S")
        default = nombre.replace(" ", "_") + "_" + ts + "_imprimible.png"
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG", "*.png"), ("Todos", "*.*")],
            initialfile=default,
            title="Guardar gráfica Doppler (versión imprimible)")
        if path:
            fig_imp = self._dibujar_figura(True)
            try:
                fig_imp.savefig(path, dpi=200, bbox_inches="tight",
                                 facecolor=fig_imp.get_facecolor())
                self._set_status("✓ Guardado (imprimible): " + path, GREEN)
            except OSError as e:
                messagebox.showerror("Error al guardar",
                                     "No se pudo guardar el PNG:\n" + str(e))
                self._set_status("✗ No se pudo guardar", RED)
            finally:
                _cerrar_figura(fig_imp)

    # ── Limpiar ───────────────────────────────────────────────────────────────
    def _limpiar(self):
        for k in self.vars:
            self.vars[k].set("")
            self.status_labels[k].config(text="")
        self.pam_ent.delete(0, tk.END)
        self.co2_ent.delete(0, tk.END)
        self.ent_nombre.delete(0, tk.END)
        self.ent_hc.delete(0, tk.END)
        # Reset Lindegaard tab
        self.lind_aci_var.set("")
        self._lind_il_result  = None
        self._lind_tamax_acm  = None
        self.lind_il_lbl.config(   text="IL = —",                          fg=TEXT2)
        self.lind_interp_lbl.config(text="Ingresa VM ACI extracraneal para calcular", fg=TEXT2)
        self.lind_detalle_lbl.config(text="",                                fg=TEXT2)
        self.lind_info_pac.config(   text="—")
        self.lind_info_art.config(   text="—")
        self.lind_info_tamean.config(text="—  cm/s")
        self.lind_info_tamax.config( text="—  cm/s")
        self.nb.tab(1, state="disabled")
        self.nb.select(0)
        self._informe_texto = ""
        self._estenosis_sev, self._estenosis_texto, self._estenosis_vm = None, "", None
        self.estenosis_lbl.config(text="Estenosis ACM: —", fg=TEXT2)
        if self.canvas_fig:
            self.canvas_fig.get_tk_widget().destroy()
            self.canvas_fig = None
        if self.last_fig:
            _cerrar_figura(self.last_fig)
            self.last_fig = None
        # Invalidar la figura diferida y el último resultado. Sin esto,
        # "Guardar PNG (imprimible)" seguía regenerando la figura del paciente
        # ANTERIOR desde el closure — riesgo de informe mal atribuido.
        self._dibujar_figura = None
        self.last_result     = None
        self._pi_autocalc    = False
        self.placeholder.place(relx=0.5, rely=0.5, anchor="center")
        self._set_status("Campos limpiados", TEXT2)


# =============================================================================


# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 3 — ÍCONOS (canvas drawing helpers)
# ══════════════════════════════════════════════════════════════════════════════

def _draw_nervio_optico(canvas, x, y, color, scale=1.0):
    """Ícono esquemático vaina del nervio óptico: globo + tallo."""
    r  = int(28 * scale)
    sw = int(10 * scale)
    sh = int(36 * scale)
    canvas.create_rectangle(x - sw//2, y, x + sw//2, y + sh,
                             fill=color, outline="", tags="icon")
    canvas.create_oval(x - r, y - r*2, x + r, y,
                       fill=color, outline="", tags="icon")


def _tcd_cycle(cx, cy, W, H, ed_frac=0.12):
    """
    Morfología TCD paramétrica realista basada en ED/PS (ed_frac).

    ed_frac bajo (< 0.20)  → alta resistencia / PIC elevada:
        - Bajada sistólica muy rápida y pronunciada
        - Sin incisura dícrota visible (PI alto aplana S2)
        - Meseta diastólica muy baja y plana

    ed_frac alto (> 0.40)  → baja resistencia / hiperemia:
        - Bajada sistólica gradual
        - Incisura dícrota bien visible
        - Diástole alta y sostenida
    """
    ed_frac = max(0.04, min(ed_frac, 0.70))

    # Pendiente de bajada post-pico: más abrupta si ed_frac bajo
    # t_nadir: tiempo donde llega al mínimo diastólico
    t_nadir = 0.28 + ed_frac * 0.25        # 0.29..0.45

    # Incisura dícrota: solo visible si ed_frac > 0.22
    has_notch = ed_frac > 0.22
    notch_h   = ed_frac * 0.55 + 0.04 if has_notch else ed_frac + 0.02
    s2_h      = ed_frac * 0.65 + 0.03 if has_notch else ed_frac + 0.01

    ctrl = [
        (0.00,     ed_frac * 0.85),
        (0.04,     ed_frac * 0.85),
        (0.08,     1.00),               # Pico sistólico
        (0.13,     ed_frac + 0.18),     # bajada rápida post-pico
        (t_nadir,  notch_h),            # incisura / mínimo
        (t_nadir + 0.06, s2_h),         # onda dícrota (suave si PI alto)
        (t_nadir + 0.14, ed_frac + 0.04),
        (0.55,     ed_frac + 0.01),
        (0.75,     ed_frac),            # meseta diastólica
        (0.92,     ed_frac * 0.90),
        (1.00,     ed_frac * 0.85),
    ]
    pts = []
    for t, h in ctrl:
        h = max(0.0, min(h, 1.0))
        pts += [cx + int(t * W), cy - int(h * H)]
    return pts


def _draw_doppler_wave(canvas, x, y, color, scale=1.0, w=None, h=None,
                       ps=None, ed=None):
    """
    Dibuja 4 ciclos TCD.
    Si se pasan ps y ed (en cm/s), la morfología refleja la relación ED/PS real:
      - H se usa como escala máxima (ya escalada externamente)
      - ed_frac = ED/PS modula altura diastólica y S2
    """
    TW = w if w else int(90 * scale)
    H  = h if h else int(55 * scale)
    # Grosor de onda: trazo fino fijo — 1px canvas pequeño, 2px en grande.
    # NO escala con el ancho del canvas para evitar trazos gruesos en pantallas anchas.
    lw = 2 if TW > 500 else 1

    # Calcular ed_frac desde parámetros reales si están disponibles
    if ps and ps > 0 and ed is not None and ed >= 0:
        ed_frac = min(ed / ps, 0.65)
    else:
        ed_frac = 0.12   # default morfología alta resistencia

    n_cycles = 4
    gap      = int(TW * 0.012)
    cw       = (TW - gap * (n_cycles - 1)) // n_cycles

    all_pts = []
    for i in range(n_cycles):
        cx = x + i * (cw + gap)
        all_pts.extend(_tcd_cycle(cx, y, cw, H, ed_frac=ed_frac))

    # Línea de base — 1px
    canvas.create_line(x, y, x + TW, y, fill=color,
                       width=1, tags="icon")
    if len(all_pts) >= 4:
        # smooth=False: tkinter smooth engrosa visualmente el trazo.
        # Línea directa entre puntos, limpia, sin antialiasing exagerado.
        canvas.create_line(*all_pts, fill=color, width=lw,
                           smooth=False, tags="icon", joinstyle="miter", capstyle="butt")


# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 4 — MÓDULO VNO
# ══════════════════════════════════════════════════════════════════════════════

class VNOWindow(tk.Toplevel):
    """Ventana de medición de Vaina de Nervio Óptico bilateral."""

    def __init__(self, parent, callback_result):
        super().__init__(parent)
        self.parent          = parent
        self.callback_result = callback_result

        self.title("Vaina de Nervio Óptico — Medición")
        self.geometry("520x540")
        self.resizable(False, False)
        self.configure(bg=BG)
        self.grab_set()
        self._build()

    def _build(self):
        tk.Label(self, text="Vaina de Nervio Óptico",
                 font=F_TITLE, bg=BG, fg=ACCENT
                 ).pack(pady=(22, 4))
        tk.Label(self,
                 text="Medición ecográfica bilateral · corte 3 mm retrorbitario",
                 font=F_SMALL, bg=BG, fg=TEXT2).pack()

        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=20, pady=12)

        form = tk.Frame(self, bg=BG)
        form.pack(fill="x", padx=30)

        tk.Label(form, text="",          width=18, bg=BG).grid(row=0, column=0)
        tk.Label(form, text="mm",        font=F_LABEL, bg=BG, fg=TEXT2, width=10
                 ).grid(row=0, column=1)
        tk.Label(form, text="Resultado", font=F_LABEL, bg=BG, fg=TEXT2, width=16
                 ).grid(row=0, column=2)

        self.entries  = {}
        self.res_lbls = {}
        for i, (key, label) in enumerate([("OD", "Ojo Derecho"),
                                           ("OI", "Ojo Izquierdo")]):
            tk.Label(form, text=label, font=F_LABEL, bg=BG, fg=TEXT,
                     anchor="w", width=18
                     ).grid(row=i+1, column=0, pady=12, sticky="w")

            var = tk.StringVar()
            tk.Entry(form, textvariable=var, width=10,
                     font=F_VALUE_M, bg=BG3, fg=TEXT,
                     insertbackground=TEXT, relief="flat",
                     highlightthickness=1, highlightbackground=BORDER,
                     highlightcolor=ACCENT, justify="center"
                     ).grid(row=i+1, column=1, padx=8)
            var.trace_add("write", lambda *_, k=key: self._eval(k))
            self.entries[key] = var

            lbl = tk.Label(form, text="—", font=F_BODY, bg=BG, fg=TEXT2, width=16)
            lbl.grid(row=i+1, column=2)
            self.res_lbls[key] = lbl

        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=20, pady=14)

        gf = tk.Frame(self, bg=BG2, highlightthickness=1,
                      highlightbackground=BORDER)
        gf.pack(fill="x", padx=30, pady=(0, 12))
        self.global_lbl  = tk.Label(gf, text="Sin datos suficientes",
                                    font=F_HEAD,
                                    bg=BG2, fg=TEXT2, pady=10)
        self.global_lbl.pack()
        self.interp_lbl  = tk.Label(gf, text="", font=F_SMALL,
                                    bg=BG2, fg=TEXT2, pady=2)
        self.interp_lbl.pack()

        nf = tk.Frame(self, bg=BG)
        nf.pack(fill="x", padx=30, pady=(0, 8))
        tk.Label(nf, text="Nota / hallazgos adicionales:",
                 font=F_SMALL, bg=BG, fg=TEXT2, anchor="w").pack(anchor="w")
        self.nota_text = tk.Text(nf, height=3, width=50, font=F_BODY,
                                  bg=BG3, fg=TEXT, insertbackground=TEXT,
                                  relief="flat", highlightthickness=1,
                                  highlightbackground=BORDER,
                                  highlightcolor=ACCENT)
        self.nota_text.pack(fill="x", pady=4)

        bf = tk.Frame(self, bg=BG)
        bf.pack(pady=8)
        tk.Button(bf, text="✓  Registrar medición",
                  font=F_LABEL, bg=BG3, fg=ACCENT,
                  relief="flat", bd=0, padx=20, pady=9,
                  highlightthickness=1, highlightbackground=ACCENT,
                  activebackground=BORDER, activeforeground=ACCENT,
                  cursor="hand2",
                  command=self._registrar).pack(side="left", padx=8)
        tk.Button(bf, text="🖼  Guardar informe PNG",
                  font=F_BODY, bg=BG3, fg=TEXT2,
                  relief="flat", bd=0, padx=14, pady=9,
                  activebackground=BORDER, activeforeground=TEXT,
                  cursor="hand2",
                  command=self._guardar_png).pack(side="left", padx=8)
        tk.Button(bf, text="Cancelar",
                  font=F_BODY, bg=BG3, fg=TEXT2,
                  relief="flat", bd=0, padx=14, pady=9,
                  activebackground=BORDER, cursor="hand2",
                  command=self.destroy).pack(side="left", padx=8)

    def _eval(self, key):
        try:
            val = float(self.entries[key].get())
            if val <= 0: raise ValueError
            clave, etiqueta, color, _ = _vno_cat(val)
            simbolo = {"normal": "✓", "gris": "≈", "elevada": "⬆"}[clave]
            self.res_lbls[key].config(
                text=f"{simbolo} {etiqueta}  ({val:.1f} mm)", fg=color)
        except ValueError:
            self.res_lbls[key].config(text="—", fg=TEXT2)
        self._eval_global()

    def _eval_global(self):
        vals = {}
        for k, var in self.entries.items():
            try:
                v = float(var.get())
                if v > 0: vals[k] = v
            except ValueError:
                pass
        if len(vals) < 2:
            self.global_lbl.config(text="Sin datos suficientes", fg=TEXT2)
            self.interp_lbl.config(text="")
            return
        od, oi = vals.get("OD"), vals.get("OI")
        if od is None or oi is None:
            self.global_lbl.config(text="Completar ambos lados", fg=TEXT2)
            self.interp_lbl.config(text="")
            return
        prom = (od + oi) / 2
        clave, _etq, color, _ = _vno_global(od, oi)
        if clave == "elevada":
            self.global_lbl.config(
                text=f"⬆  Vaina dilatada  —  prom. {prom:.2f} mm", fg=color)
            self.interp_lbl.config(
                text="Sugestivo de PIC elevada (≥ 6 mm)  ·  correlacionar clínicamente",
                fg=YELLOW)
        elif clave == "gris":
            self.global_lbl.config(
                text=f"≈  Zona gris  —  prom. {prom:.2f} mm", fg=color)
            self.interp_lbl.config(
                text="Limítrofe (5.0–5.9 mm)  ·  no concluyente, correlacionar clínicamente",
                fg=YELLOW)
        else:
            self.global_lbl.config(
                text=f"✓  Vaina normal  —  prom. {prom:.2f} mm", fg=color)
            self.interp_lbl.config(
                text="Sin evidencia ecográfica de PIC elevada por VNO", fg=TEXT2)

    def _guardar_png(self):
        """Genera y guarda un informe PNG de la medición VNO actual."""
        vals = {}
        for k, var in self.entries.items():
            try:
                v = float(var.get())
                if v > 0: vals[k] = v
            except ValueError:
                pass
        if len(vals) < 2:
            messagebox.showwarning("Datos incompletos",
                                   "Ingresar medición de ambos ojos antes de guardar.",
                                   parent=self)
            return

        od, oi  = vals["OD"], vals["OI"]
        prom    = (od + oi) / 2
        clave, _etq, _col_ui, _ = _vno_global(od, oi)
        nota    = self.nota_text.get("1.0", "end").strip()
        ts      = datetime.now().strftime("%d/%m/%Y %H:%M")

        try:

            fig, ax = _nuevos_subplots(figsize=(7, 4.5))
            fig.patch.set_facecolor("#000000")
            ax.set_facecolor("#111111")
            for spine in ax.spines.values():
                spine.set_color("#2A2A2A")
            ax.tick_params(colors="#FFFFFF")

            # Barras OD / OI (color según zona: verde / amarillo / rojo)
            colores = [_vno_cat(od)[3], _vno_cat(oi)[3]]
            bars = ax.bar(["OD", "OI"], [od, oi], color=colores,
                          width=0.4, edgecolor="#2A2A2A")
            ax.axhline(VNO_NORMAL_MAX, color="#FFFF00", linewidth=1.2,
                       linestyle="--", label=f"Zona gris ≥ {VNO_NORMAL_MAX} mm")
            ax.axhline(VNO_ELEVADA_MIN, color="#FF2020", linewidth=1.4,
                       linestyle="--", label=f"Dilatada ≥ {VNO_ELEVADA_MIN} mm")

            for bar, val in zip(bars, [od, oi]):
                ax.text(bar.get_x() + bar.get_width()/2,
                        bar.get_height() + 0.08,
                        f"{val:.1f} mm", ha="center", va="bottom",
                        color="#FFFFFF", fontsize=15, fontweight="bold")

            ax.set_ylim(0, max(od, oi, VNO_ELEVADA_MIN) * 1.35)
            ax.set_ylabel("mm", color="#FFFFFF")
            ax.set_title("Vaina de Nervio Óptico — Medición Bilateral",
                         color="#FFFFFF", fontsize=16, pad=10)
            ax.legend(facecolor="#1C1C1C", edgecolor="#2A2A2A",
                      labelcolor="#FFFFFF", fontsize=11)

            _txt = {"normal":  "Vaina normal — sin evidencia de PIC elevada",
                    "gris":    "Zona gris (5.0–5.9 mm) — limítrofe, correlacionar",
                    "elevada": "↑ VAINA DILATADA — sugestivo PIC elevada"}
            _col = {"normal": "#00FF00", "gris": "#FFFF00", "elevada": "#FF2020"}
            estado_txt, estado_col = _txt[clave], _col[clave]
            fig.text(0.5, 0.01,
                     f"Prom {prom:.2f} mm  ·  {estado_txt}\n"
                     f"{ts}{'  ·  ' + nota if nota else ''}",
                     ha="center", va="bottom", fontsize=10,
                     color=estado_col, wrap=True)

            path = filedialog.asksaveasfilename(
                parent=self,
                defaultextension=".png",
                filetypes=[("PNG", "*.png")],
                initialfile=f"VNO_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                title="Guardar informe VNO")
            if path:
                fig.savefig(path, dpi=180, bbox_inches="tight",
                            facecolor=fig.get_facecolor())
                messagebox.showinfo("Guardado", f"Informe guardado en:\n{path}",
                                    parent=self)
            _cerrar_figura(fig)
        except Exception as e:
            messagebox.showerror("Error al guardar",
                                 f"No se pudo generar el PNG:\n{e}", parent=self)

    def _registrar(self):
        vals = {}
        for k, var in self.entries.items():
            try:
                v = float(var.get())
                if v > 0: vals[k] = v
            except ValueError:
                pass
        if len(vals) < 2:
            messagebox.showwarning("Datos incompletos",
                                   "Ingresar medición de ambos ojos antes de registrar.",
                                   parent=self)
            return
        od, oi  = vals["OD"], vals["OI"]
        clave, etiqueta, _, _ = _vno_global(od, oi)
        self.callback_result({
            "timestamp": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "OD_mm":    od,
            "OI_mm":    oi,
            "promedio": round((od + oi) / 2, 2),
            "categoria": clave,            # "normal" | "gris" | "elevada"
            "cat_label": etiqueta,
            "elevada":  clave == "elevada",
            "nota":     self.nota_text.get("1.0", "end").strip(),
        })
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 5 — PANTALLA PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════


# ── ZONA DE COLOR DEL PI — FUENTE ÚNICA DE VERDAD ────────────────────────────
#
#   ANTES: cada vista tenía sus propios umbrales.
#     · Panel principal e informe consolidado → 1.0 / 1.4 fijos, iguales para
#       toda arteria (y comentados como "según Bellner", cosa que no eran:
#       con Bellner, PI 1.4 ≡ PIC ≈ 14 mmHg, todavía normal).
#     · Panel "Otras arterias" y tabla de índices → rango por arteria.
#   Resultado: un PI de 1.30 en ACM salía amarillo en una vista y rojo en otra,
#   dentro incluso de la MISMA figura (tabla vs. línea de datos).
#
#   AHORA: criterio único POR ARTERIA, tomado de RANGOS_ARTERIA — que es
#   exactamente lo que publica la pestaña Referencia y lo único que distingue
#   ACM (PI 0.55–1.25) de Oftálmica (0.55–1.35).
#
#   Zonas:
#     bajo    < PI_min              patrón de baja resistencia (hiperemia,
#                                   shunt AV, amortiguación distal)
#     normal    PI_min – PI_max
#     limite    hasta PI_max×1.15   margen de tolerancia de medición
#     alto    > PI_max×1.15
#
#   Para endurecer o relajar el criterio hay UN solo lugar que tocar:
#   DopplerACMApp.RANGOS_ARTERIA (posiciones 6 y 7) y PI_LIMITE_FACTOR.
PI_LIMITE_FACTOR = 1.15

_PI_TEXTO = {"sin_dato": "—",        "bajo":   "↓ bajo",
             "normal":   "normal",   "limite": "⚠ límite",
             "alto":     "↑ alto"}

# Zona → nombre de tag en los widgets tk.Text (info de los paneles)
_PI_TAG = {"sin_dato": "dim",      "bajo":   "bajo",
           "normal":   "verde",    "limite": "amarillo",
           "alto":     "rojo"}

# Zona → clave de paleta (sirve tanto para la de pantalla como la imprimible)
_PI_PAL_KEY = {"sin_dato": "TEXT2", "bajo":   "ACCENT",
               "normal":   "GREEN", "limite": "YELLOW",
               "alto":     "RED"}


def _pi_rango(arteria):
    """(PI_min, PI_max) de referencia para la arteria. Fallback: ACM."""
    r = DopplerACMApp.RANGOS_ARTERIA.get(
        arteria, DopplerACMApp.RANGOS_ARTERIA["ACM"])
    return r[6], r[7]


def _pi_zona(pi, arteria="ACM"):
    """Clasifica un PI: 'sin_dato' | 'bajo' | 'normal' | 'limite' | 'alto'."""
    if pi is None:
        return "sin_dato"
    try:
        pi = float(pi)
    except (TypeError, ValueError):
        return "sin_dato"
    pi_min, pi_max = _pi_rango(arteria)
    if pi < pi_min:
        return "bajo"
    if pi <= pi_max:
        return "normal"
    if pi <= pi_max * PI_LIMITE_FACTOR:
        return "limite"
    return "alto"


def _pi_estado(pi, arteria="ACM", pal=None):
    """(texto, color) del PI en la paleta indicada (por defecto, pantalla)."""
    z = _pi_zona(pi, arteria)
    pal = pal or _PALETTE_SCREEN
    return _PI_TEXTO[z], pal[_PI_PAL_KEY[z]]


def _pi_tag(pi, arteria="ACM"):
    """Tag tk.Text correspondiente al PI (dim/bajo/verde/amarillo/rojo)."""
    return _PI_TAG[_pi_zona(pi, arteria)]


# ── ALERTAS HEMODINÁMICAS DEL MONITOR PRINCIPAL ──────────────────────────────
#   Vasoespasmo: umbrales de VM (TAMV = campo "TAmax") de la pestaña
#   Referencia: (leve, moderado, severo). Con IL disponible se discrimina:
#   VM alta + IL < 3 → predomina hiperemia (ámbar, no rojo).
#   Diástole muy baja: ED < 80 % del límite inferior normal de la arteria
#   (mismo corte que usa _validar para marcar "✗ Anormal"; ACM → ED < 20).
VASOESPASMO_VM = {"ACM": (120, 150, 200), "ACA": (90, 120, 150),
                  "ACP": (70, 90, 110),   "AB":  (80, 100, 140),
                  "AV":  (60, 80, 110)}
ED_TOLERANCIA = 0.80


def _alertas_tcd(data):
    """Lista de (nivel, texto) con nivel 'rojo' | 'amarillo'. [] si no hay."""
    if not data:
        return []
    art = data.get("arteria", "ACM")
    out = []
    vm = _vm_de(data)
    il = data.get("IL")
    umb = VASOESPASMO_VM.get(art)
    if vm is not None and umb and vm >= umb[0]:
        leve, mod, sev = umb
        grado = "SEVERO" if vm > sev else "MODERADO" if vm > mod else "LEVE"
        vm_s = ("~" if data.get("VM_aproximada") else "") + f"{vm:.0f}"
        il_s = f"  IL {il:.1f}" if il is not None else ""
        if il is not None and il < 3 and vm <= sev:
            out.append(("amarillo",
                        f"VM ELEVADA {vm_s} cm/s  IL {il:.1f} (<3: sugiere hiperemia)"))
        elif il is None:
            out.append(("rojo",
                        f"POSIBLE VASOESPASMO {grado}  VM {vm_s} cm/s (sin IL)"))
        else:
            out.append(("rojo", f"VASOESPASMO {grado}  VM {vm_s} cm/s{il_s}"))
    ed = data.get("ED")
    r = DopplerACMApp.RANGOS_ARTERIA.get(art)
    if ed is not None and r and ed < r[2] * ED_TOLERANCIA:
        out.append(("rojo", f"DIÁSTOLE MUY BAJA  ED {ed:g} cm/s"))
    return out


def _color_alerta(data, rojo_col, amar_col):
    """Color de la alerta más grave (rojo > ámbar) o None si no hay alerta.
    Misma prioridad que la onda del monitor principal (_color_onda)."""
    al = _alertas_tcd(data)
    if not al:
        return None
    return rojo_col if any(n == "rojo" for n, _ in al) else amar_col


def _alertas_en_ax(ax, data, rojo_col, amar_col, printable,
                   x=0.01, y=0.97, fontsize=10):
    """Dibuja en un eje matplotlib el MISMO aviso que muestra el monitor
    (vasoespasmo / VM elevada con IL / diástole muy baja). Así las capturas
    y los PNG exportados no pierden la alerta. Devuelve el color o None."""
    al = _alertas_tcd(data)
    if not al:
        return None
    rojo = any(n == "rojo" for n, _ in al)
    col = rojo_col if rojo else amar_col
    if printable:
        fondo = "#FDEAEA" if rojo else "#FFF5D6"
    else:
        fondo = "#2A0000" if rojo else "#2A2A00"
    ax.text(x, y, "\n".join("⚠ " + t for _, t in al),
            transform=ax.transAxes, va="top", ha="left", color=col,
            fontsize=fontsize, fontweight="bold", zorder=8, clip_on=False,
            bbox=dict(boxstyle="round,pad=0.35", facecolor=fondo,
                      edgecolor=col, linewidth=1.2))
    return col


def _alerta_canvas_tk(cv, cw, ch, data):
    """Borde de color + barra de alarma al pie de un canvas Tk (monitor
    principal y panel «Otras arterias» usan esta misma función)."""
    al = _alertas_tcd(data)
    rojo = any(n == "rojo" for n, _ in al)
    col = RED if rojo else (YELLOW if al else BG)
    try:
        cv.config(highlightbackground=col, highlightcolor=col)
    except tk.TclError:
        pass
    if not al:
        return
    txt = "   ·   ".join("⚠ " + t for _, t in al)
    fs = 11 if len(txt) < 60 else 10 if len(txt) < 80 else 9
    t = cv.create_text(cw // 2, ch - 4, text=txt, fill=col, anchor="s",
                       width=max(cw - 24, 100), justify="center",
                       font=(_FAM_UI, fs, "bold"), tags="alerta")
    bb = cv.bbox(t)
    if bb:
        r = cv.create_rectangle(bb[0] - 8, bb[1] - 3, bb[2] + 8, bb[3] + 3,
                                fill="#2A0000" if rojo else "#2A2A00",
                                outline=col, tags="alerta")
        cv.tag_lower(r, t)


def _prellenar_tcd(tcd_app, prev):
    """Carga en el formulario TCD los valores del último análisis de esa
    arteria. Lo usan el monitor principal y el panel «Otras arterias» (antes
    el panel reabría el formulario vacío y había que volver a digitar todo).
    Solo carga valores MEDIDOS: TAmean aproximada y HR ausente quedan vacíos."""
    if not prev:
        return
    for clave in ("PS", "ED", "RI", "PI", "SD", "TAmax", "TAmean", "HR"):
        val = prev.get(clave)
        if val in (None, ""):
            continue
        if clave == "TAmean" and prev.get("TAmean_aproximada"):
            continue
        if clave in tcd_app.vars:
            tcd_app.vars[clave].set(str(val))
    for clave, ent in (("PAM", "pam_ent"), ("PaCO2", "co2_ent")):
        val = prev.get(clave)
        if val is None:
            continue
        try:
            w = getattr(tcd_app, ent)
            w.delete(0, "end")
            w.insert(0, f"{float(val):g}")
        except (AttributeError, ValueError, tk.TclError):
            pass


def _pic_color(pic_label):
    """Color por severidad de PIC: rojo=alta, amarillo=borderline, verde=normal."""
    if not pic_label or pic_label in ("—", "", "N/A"):
        return GRAY
    pl = pic_label.upper()
    if "ALTA" in pl or "MUY" in pl:
        return RED
    if "BORDER" in pl or "LEVE" in pl or "MODERADA" in pl:
        return YELLOW
    if "NORMAL" in pl:
        return GREEN
    import re as _re
    m = _re.search(r"(\d+(?:\.\d+)?)", pl)
    if m:
        v = float(m.group(1))
        if v >= 20: return RED
        if v >= 15: return YELLOW
        return GREEN
    return GRAY


# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 4B — PANEL COMPARATIVO "OTRAS ARTERIAS" (2, 4 o 6 simultáneas)
# ══════════════════════════════════════════════════════════════════════════════

ARTERIAS_TODAS = list(DopplerACMApp.RANGOS_ARTERIA.keys())


class OtrasArteriasApp:
    """
    Panel comparativo de arterias (2, 4 o 6 simultáneas, alternable).

    Modo 2 (1×2): comparación focal — típicamente un par bilateral
    (p. ej. ACP D/I o AV D/I) con el máximo de área de onda por panel.
    Modo 4 (2×2): default. Modo 6 (3×2): mapa vascular ampliado.
    Cambiar de modo NO borra las mediciones ya tomadas: los slots
    ocultos conservan sus datos y reaparecen al volver al modo mayor.
    Cada slot tiene una solapa (combobox arteria + lado) para elegir qué
    arteria/lado se muestra, botón para abrir el Doppler TCD completo
    (reutiliza DopplerACMApp) y muestra en vivo: onda, PS/ED/PI/RI/S-D,
    y — igual que en la ventana principal — PIC y PPC estimadas por
    Czosnyka (1998) y Bellner (2004). Ambas estimaciones solo son
    válidas para ACM (igual que en el módulo TCD principal); para el
    resto de arterias se indica "Solo aplicable en ACM".
    Permite generar un informe consolidado (imagen) con todos los slots
    visibles. Los rangos normales se toman de DopplerACMApp.RANGOS_ARTERIA.
    """

    N_SLOTS = 6
    DEFAULT_SLOTS = [
        ("ACA", "Derecho"),   ("ACA", "Izquierdo"),
        ("ACP", "Derecho"),   ("ACP", "Izquierdo"),
        ("AB",  "Único"),     ("AV",  "Derecho"),
    ]

    # Modos de panel disponibles → nº de filas del grid (siempre 2 columnas).
    # 2 → 1×2   ·   4 → 2×2   ·   6 → 3×2
    MODOS = {2: 1, 4: 2, 6: 3}

    def __init__(self, root, pac_nombre="", pac_hc="", foco=None,
                 on_change=None, datos_previos=None):
        self.root       = root
        self.on_change  = on_change        # cb(arteria, lado, data|None) → dashboard
        self.slot_key   = [None] * self.N_SLOTS
        self.pac_nombre = pac_nombre
        self.pac_hc     = pac_hc
        self.panel_mode = 4   # 2, 4 (default) o 6 slots visibles

        self.root.title("Otras arterias — Panel comparativo")
        self.root.geometry("1480x940")
        self.root.configure(bg=BG)
        self.root.minsize(1180, 720)
        _apply_ttk_dark_theme(self.root)

        self.slot_art    = []               # StringVar arteria, por slot
        self.slot_lado   = []               # StringVar lado, por slot
        self.slot_data   = [None] * self.N_SLOTS
        self.slot_win    = [None] * self.N_SLOTS
        self._panels     = []
        self._panel_frames = []
        self._slot_sync  = []               # callable por slot: sincroniza combo de lado

        base = list(self.DEFAULT_SLOTS)
        if foco and foco in ARTERIAS_TODAS:
            base[0] = (foco, "Derecho")

        self._build_header()
        self._build_grid(base)
        self._relayout_grid()
        self.root.update_idletasks()
        # Precarga de mediciones ya registradas en el dashboard (persisten
        # aunque se cierre y reabra este panel, y se guardan en el .ncu).
        for i in range(self.N_SLOTS):
            k = (self.slot_art[i].get(), self.slot_lado[i].get())
            d = (datos_previos or {}).get(f"{k[0]}|{k[1]}")
            if d:
                self._on_result(i, d, notificar=False)
        self._redraw_all()

    # ── Header ────────────────────────────────────────────────────────────────
    def _build_header(self):
        h = tk.Frame(self.root, bg=BG2, height=60,
                     highlightthickness=1, highlightbackground=BORDER)
        h.pack(fill="x")
        h.pack_propagate(False)
        tk.Label(h, text="🧠  OTRAS ARTERIAS  —  Panel comparativo",
                 font=F_TITLE, bg=BG2, fg=ACCENT).pack(side="left", padx=20, pady=12)
        tk.Label(h, text="Valores de referencia bibliográfica  ·  PIC/PPC (Czosnyka/Bellner) solo en ACM",
                 font=F_SMALL, bg=BG2, fg=TEXT2).pack(side="left", padx=(4, 0))

        tk.Button(h, text="📄  Generar reporte",
                  font=F_BODY, bg=BG3, fg=GREEN,
                  relief="flat", bd=0, padx=14, pady=6,
                  highlightthickness=1, highlightbackground=GREEN,
                  activebackground=BORDER, activeforeground=GREEN,
                  cursor="hand2", command=self._generar_reporte
                  ).pack(side="right", padx=10)

        # ── Selector de cantidad de paneles (2 / 4 / 6) ──
        toggle = tk.Frame(h, bg=BG3, highlightthickness=1, highlightbackground=BORDER)
        toggle.pack(side="right", padx=10)
        self._btn_modo = {}
        for _n in sorted(self.MODOS):
            b = tk.Button(toggle, text=f"{_n} arterias", font=F_SMALL,
                          relief="flat", bd=0, padx=10, pady=5, cursor="hand2",
                          command=lambda n=_n: self._set_panel_mode(n))
            b.pack(side="left")
            self._btn_modo[_n] = b
            if _n != max(self.MODOS):
                tk.Frame(toggle, bg=BORDER, width=1).pack(side="left", fill="y", pady=4)
        self._actualizar_botones_modo()

    def _actualizar_botones_modo(self):
        activo   = dict(bg=ACCENT, fg=BG)
        inactivo = dict(bg=BG3,    fg=TEXT2)
        for n, b in self._btn_modo.items():
            b.config(**(activo if self.panel_mode == n else inactivo))

    def _set_panel_mode(self, n):
        if n not in self.MODOS or n == self.panel_mode:
            return
        self.panel_mode = n
        self._actualizar_botones_modo()
        self._relayout_grid()
        self.root.after(30, self._redraw_all)

    # ── Grid de slots (1×2 en modo 2, 2×2 en modo 4, 3×2 en modo 6) ─────────────
    def _build_grid(self, base):
        self._grid = tk.Frame(self.root, bg=BG)
        self._grid.pack(fill="both", expand=True, padx=10, pady=10)
        for c in range(2):
            self._grid.columnconfigure(c, weight=1)
        for r in range(3):
            self._grid.rowconfigure(r, weight=1)

        for i in range(self.N_SLOTS):
            art0, lado0 = base[i]
            art_var  = tk.StringVar(value=art0)
            lado_var = tk.StringVar(value=lado0)
            self.slot_art.append(art_var)
            self.slot_lado.append(lado_var)

            panel = tk.Frame(self._grid, bg=BG2,
                              highlightthickness=1, highlightbackground=BORDER)
            self._panel_frames.append(panel)
            panel.columnconfigure(0, weight=1)
            panel.rowconfigure(2, weight=1)

            # ── Solapa de selección: arteria + lado ──
            sel = tk.Frame(panel, bg=BG2)
            sel.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 2))

            art_cb = ttk.Combobox(sel, textvariable=art_var, values=ARTERIAS_TODAS,
                                   width=15, state="readonly", font=F_SMALL)
            art_cb.pack(side="left")

            lado_cb = ttk.Combobox(sel, textvariable=lado_var, width=10,
                                    state="readonly", font=F_SMALL)
            lado_cb.pack(side="left", padx=(6, 0))

            def _sync_lado(art_v=art_var, lado_v=lado_var, cb=lado_cb):
                opciones = ["Único"] if art_v.get() == "AB" else ["Derecho", "Izquierdo"]
                cb.config(values=opciones)
                if lado_v.get() not in opciones:
                    lado_v.set(opciones[0])

            def _on_art_change(_e=None, idx=i, f=_sync_lado):
                f()
                self._reset_slot(idx)

            _sync_lado()
            # Guardamos el sincronizador: si la arteria se cambia por código
            # (submenú "Ir directo a una arteria") hay que reajustar el combo
            # de lado, o quedaría p.ej. "AB Derecho" — lado inexistente en AB.
            self._slot_sync.append(_sync_lado)
            art_cb.bind("<<ComboboxSelected>>", _on_art_change)
            lado_cb.bind("<<ComboboxSelected>>", lambda e, idx=i: self._reset_slot(idx))

            tk.Button(sel, text="✕", font=F_SMALL, bg=BG2, fg=TEXT3,
                      relief="flat", bd=0, cursor="hand2",
                      command=lambda idx=i: self._reset_slot(idx)
                      ).pack(side="right")

            # ── Etiqueta PIC/PPC (igual estilo que la ventana principal) ──
            pic_lbl = tk.Label(panel, text="", font=F_SMALL, bg=BG2, fg=TEXT2,
                                justify="right", anchor="e")
            pic_lbl.grid(row=1, column=0, sticky="ew", padx=10)

            # ── Canvas de la onda (vista rápida, en vivo) ──
            cv = tk.Canvas(panel, bg=BG, highlightthickness=0, height=150)
            cv.grid(row=2, column=0, sticky="nsew", padx=10, pady=(2, 2))
            cv.bind("<Configure>", lambda e, idx=i: self._redraw_slot(idx))

            info = tk.Text(panel, height=3, bg=BG2, fg=TEXT2, font=F_SMALL,
                            relief="flat", bd=0, highlightthickness=0,
                            wrap="word", state="disabled", cursor="arrow")
            info.grid(row=3, column=0, sticky="ew", padx=10)
            info.tag_configure("center", justify="center")
            info.tag_configure("normal",   foreground=TEXT2)
            info.tag_configure("bajo",     foreground=ACCENT)
            info.tag_configure("verde",    foreground=GREEN)
            info.tag_configure("amarillo", foreground=YELLOW)
            info.tag_configure("rojo",     foreground=RED)
            info.tag_configure("dim",      foreground=TEXT3)

            bf = tk.Frame(panel, bg=BG2)
            bf.grid(row=4, column=0, pady=(4, 8))
            tk.Button(bf, text="Abrir TCD", font=F_SMALL, bg=BG3, fg=TEXT2,
                      relief="flat", bd=0, padx=10, pady=4,
                      activebackground=BORDER, activeforeground=TEXT, cursor="hand2",
                      command=lambda idx=i: self._open_tcd(idx)
                      ).pack(side="left", padx=3)
            tk.Button(bf, text="🖼  PNG", font=F_SMALL, bg=BG3, fg=TEXT2,
                      relief="flat", bd=0, padx=10, pady=4,
                      activebackground=BORDER, activeforeground=TEXT, cursor="hand2",
                      command=lambda idx=i: self._exportar_slot(idx)
                      ).pack(side="left", padx=3)

            self._panels.append({"canvas": cv, "info": info, "pic_lbl": pic_lbl})

    # ── Alterna entre 2, 4 y 6 paneles visibles, sin perder datos medidos ──────
    def _relayout_grid(self):
        n    = self.panel_mode
        rows = self.MODOS.get(n, 2)
        for r in range(3):
            self._grid.rowconfigure(r, weight=(1 if r < rows else 0))
        for i, frame in enumerate(self._panel_frames):
            if i < n:
                r, c = divmod(i, 2)
                frame.grid(row=r, column=c, sticky="nsew", padx=6, pady=6)
            else:
                frame.grid_remove()

    # ── Fijar una arteria en un slot desde fuera (submenú "Ir directo a...") ────
    def set_foco(self, arteria, idx=0):
        """Coloca `arteria` en el slot `idx` dejando el combo de lado coherente.

        Necesario porque set() sobre el StringVar NO dispara
        <<ComboboxSelected>>: sin resincronizar quedaría, p.ej., "AB Derecho".
        """
        if arteria not in ARTERIAS_TODAS or not (0 <= idx < self.N_SLOTS):
            return
        self.slot_art[idx].set(arteria)
        if idx < len(self._slot_sync):
            self._slot_sync[idx]()
        self._reset_slot(idx)

    # ── Reset de un slot (cambio de arteria/lado sin medición previa) ───────────
    def _reset_slot(self, idx):
        win = self.slot_win[idx]
        if win and win.winfo_exists():
            win.destroy()
        self.slot_win[idx]  = None
        self.slot_data[idx] = None
        if self.slot_key[idx] and self.on_change:
            try:
                self.on_change(*self.slot_key[idx], None)
            except Exception:
                pass
        self.slot_key[idx] = None
        self._set_info(idx)
        self._set_pic_lbl(idx, None)
        self._redraw_slot(idx)

    # ── Abrir Doppler TCD completo para un slot (reutiliza DopplerACMApp) ───────
    def _open_tcd(self, idx):
        win = self.slot_win[idx]
        if win and win.winfo_exists():
            win.lift()
            return
        art  = self.slot_art[idx].get()
        lado = self.slot_lado[idx].get()
        w = tk.Toplevel(self.root)
        w.title(f"Doppler TCD — {art} {lado}")
        w.configure(bg=BG)
        self.slot_win[idx] = w
        tcd_app = DopplerACMApp(w, on_result=lambda d, i=idx: self._on_result(i, d))
        try:
            tcd_app.ent_nombre.insert(0, self.pac_nombre)
            tcd_app.ent_hc.insert(0, self.pac_hc)
            tcd_app.arteria_var.set(art)
            tcd_app._on_arteria_change()
            lado_map = {"Derecho": "Derecho", "Izquierdo": "Izquierdo", "Único": "Bilateral"}
            tcd_app.lado_var.set(lado_map.get(lado, "Derecho"))
            prev = self.slot_data[idx]
            if prev and prev.get("arteria", art) == art:
                _prellenar_tcd(tcd_app, prev)
        except Exception:
            pass
        self._set_info(idx, "EN USO...", "")

        def _on_close(i=idx, w=w):
            self.slot_win[i] = None
            w.destroy()
            if not self.slot_data[i]:
                self._set_info(i, "SIN DATOS", "")
            self._redraw_slot(i)

        w.protocol("WM_DELETE_WINDOW", _on_close)

    def _on_result(self, idx, data, notificar=True):
        self.slot_data[idx] = data
        art = data.get("arteria", self.slot_art[idx].get())
        lado_d = {"Bilateral": "Único"}.get(data.get("lado"), data.get("lado")) \
            or self.slot_lado[idx].get()
        self.slot_key[idx] = (art, lado_d)
        if notificar and self.on_change:
            try:
                self.on_change(art, lado_d, data)
            except Exception:
                pass

        pi = data.get("PI"); ri = data.get("RI"); sd = data.get("SD")
        pi_s = f"{pi:.2f}" if pi is not None else "—"
        ri_s = f"{ri:.2f}" if ri is not None else "—"
        sd_s = f"{sd:.2f}" if sd is not None else "—"
        art_s = f"{art} {data.get('lado', '')}"
        l1 = (f"{art_s}  ·  PS {_fnum(data.get('PS'), 2)}  ED {_fnum(data.get('ED'), 2)}  "
              f"TAmax {_fnum(data.get('TAmax'), 2)} cm/s")
        l2 = (f"PI {pi_s}  RI {ri_s}  S/D {sd_s}  PAM {_pam_txt(data)}  "
              f"CO₂ {_co2_txt(data)}  ·  {data.get('timestamp', '')}")

        # Color del PI: criterio único por arteria (ver _pi_zona)
        self._set_info(idx, l1, l2, _pi_tag(pi, art))
        self._set_pic_lbl(idx, data)
        self._redraw_slot(idx)

    def _set_pic_lbl(self, idx, data):
        """Muestra PIC y PPC (Czosnyka + Bellner) — igual formato que el panel
        principal (ACM D/I). Solo son válidas para ACM (Czosnyka 1998 y
        Bellner 2004 se validaron únicamente en esa arteria); para el resto
        simplemente se oculta el bloque en vez de mostrar un mensaje."""
        lbl = self._panels[idx]["pic_lbl"]
        if not data or data.get("arteria") != "ACM":
            lbl.config(text="", fg=TEXT2)
            return

        pic  = data.get("PIC_label",   "")
        bell = data.get("PIC_bellner", "")
        ppc_c = data.get("PPC_czosnyka_str", "")
        ppc_b = data.get("PPC_bellner_str",  "")

        czos_txt = f"Czos: {pic}"  if pic  else ""
        bell_txt = f"Bell: {bell}" if bell else ""
        pic_parts = [t for t in [czos_txt, bell_txt] if t]
        pic_inner = "  ·  ".join(pic_parts) if pic_parts else "—"

        ppc_czos_txt = f"Czos: {ppc_c}" if ppc_c else ""
        ppc_bell_txt = f"Bell: {ppc_b}" if ppc_b else ""
        ppc_parts = [t for t in [ppc_czos_txt, ppc_bell_txt] if t]
        ppc_inner = "  ·  ".join(ppc_parts) if ppc_parts else "—"

        texto = f"PIC  {pic_inner}\nPPC  {ppc_inner}"
        pc  = _pic_color(pic)
        pc2 = _pic_color(bell) if bell else GRAY
        color = RED if RED in (pc, pc2) else (YELLOW if YELLOW in (pc, pc2) else pc)
        lbl.config(text=texto, fg=color)

    def _set_info(self, idx, line1="", line2="", pi_color=None):
        it = self._panels[idx]["info"]
        it.config(state="normal")
        it.delete("1.0", "end")
        if not line1 and not line2:
            it.insert("end", "SIN DATOS", "dim")
        else:
            it.insert("end", line1, "normal")
            if line2:
                it.insert("end", "\n" + line2, pi_color or "dim")
        it.tag_add("center", "1.0", "end")
        it.config(state="disabled")

    # ── Dibujo rápido de la onda en el Canvas Tk (vista en vivo) ────────────────
    def _redraw_slot(self, idx):
        cv = self._panels[idx]["canvas"]
        cv.update_idletasks()
        w, h = cv.winfo_width(), cv.winfo_height()
        if w < 10 or h < 10:
            return
        cv.delete("all")
        art  = self.slot_art[idx].get()
        lado = self.slot_lado[idx].get()
        data = self.slot_data[idx]
        _r = DopplerACMApp.RANGOS_ARTERIA.get(art, DopplerACMApp.RANGOS_ARTERIA["ACM"])
        pi_min, pi_max = _r[6], _r[7]

        cv.create_text(12, 12, text=f"{art}  {lado.upper()}", fill=TEXT2,
                        font=(_FAM_UI, 10, "bold"), anchor="w")
        cv.create_text(w - 12, 12, text=f"Normal PI {pi_min}–{pi_max}", fill=TEXT3,
                        font=(_FAM_UI, 8), anchor="e")

        if not data:
            cv.create_text(w // 2, h // 2 + 8, text="Sin medición", fill=GRAY,
                            font=(_FAM_UI, 11))
            return

        ps = data.get("PS") or 0
        ed = data.get("ED") or 0
        pi = data.get("PI")
        pi_col = _pi_estado(pi, art)[1] if pi is not None else GRAY
        ed_f = min(ed / ps, 0.70) if ps > 0 else 0.12

        margin = 14
        top    = 26
        pts    = []
        n_cyc  = 3
        cyc_w  = (w - 2 * margin) / n_cyc
        ctrl = [(0.00, ed_f * 0.85), (0.06, 1.0), (0.30, ed_f + 0.15),
                (0.55, ed_f), (0.80, ed_f * 0.95), (1.0, ed_f * 0.85)]
        for c in range(n_cyc):
            for frac, amp in ctrl:
                x = margin + c * cyc_w + frac * cyc_w
                y = h - margin - amp * (h - top - margin)
                pts.append((x, y))
        flat = [coord for p in pts for coord in p]
        if len(flat) >= 4:
            onda_col = _color_alerta(data, RED, YELLOW) or pi_col
            cv.create_line(*flat, fill=onda_col, width=2, smooth=True)
        _alerta_canvas_tk(cv, w, h, data)

    def _redraw_all(self):
        for i in range(self.N_SLOTS):
            self._redraw_slot(i)

    # ── Dibujo matplotlib de un panel (compartido entre export individual y reporte) ─
    def _dibujar_ax(self, ax, art, lado, data, bg2, border, dim, dim2,
                    printable=False, texto_sin_dato="Sin medición"):
        import numpy as _np
        import matplotlib.patches as _mp

        _r = DopplerACMApp.RANGOS_ARTERIA.get(art, DopplerACMApp.RANGOS_ARTERIA["ACM"])
        pi_min, pi_max = _r[6], _r[7]

        ax.set_facecolor(bg2)
        for sp in ax.spines.values():
            sp.set_color(border)
        ax.set_xticks([])
        ax.tick_params(colors=dim2, labelsize=9)
        ax.text(0.01, 0.98, f"{art}  {lado.upper()}", transform=ax.transAxes,
                va="top", ha="left", color=dim, fontsize=13, fontweight="bold")
        ax.text(0.99, 0.98, f"Normal PI {pi_min}–{pi_max}", transform=ax.transAxes,
                va="top", ha="right", color=dim2, fontsize=10)

        if not data:
            ax.text(0.5, 0.5, texto_sin_dato, transform=ax.transAxes,
                    ha="center", va="center", color=border, fontsize=13)
            ax.set_yticks([])
            return

        ps = data.get("PS") or 0
        ed = data.get("ED") or 0
        pi = data.get("PI")
        # Color del PI con la paleta correcta: antes usaba las constantes de
        # PANTALLA incluso en el informe imprimible → onda verde neón (#00FF00)
        # sobre fondo blanco, prácticamente ilegible al imprimir.
        pi_col = (_pi_estado(pi, art, _paleta(printable))[1]
                  if pi is not None else dim2)
        ed_f = min(ed / ps, 0.70) if ps > 0 else 0.12
        ctrl = [(0.00, ed_f * 0.85), (0.06, 1.0), (0.30, ed_f + 0.15),
                (0.55, ed_f), (0.80, ed_f * 0.95), (1.0, ed_f * 0.85)]
        dt = 1.0 / 4
        t_one = _np.linspace(0, dt, 150)
        tc    = _np.array([c[0] * dt for c in ctrl])
        vc    = _np.array([c[1] * ps for c in ctrl])
        v_one = _np.interp(t_one, tc, vc)
        t_all = _np.concatenate([t_one + i * dt for i in range(4)])
        v_all = _np.tile(v_one, 4)
        ax.set_xlim(0, 1)
        ax.set_ylim(0, max(ps * 1.2, 1))
        _pal = _paleta(printable)
        onda_col = _color_alerta(data, _pal["RED"], _pal["YELLOW"]) or pi_col
        ax.plot(t_all, v_all, color=onda_col, linewidth=1.6)
        ax.set_ylabel("cm/s", color=dim2, fontsize=9)
        _alertas_en_ax(ax, data, _pal["RED"], _pal["YELLOW"], printable,
                       x=0.01, y=0.88, fontsize=10)

        pi_s = f"{pi:.2f}" if pi is not None else "—"
        ri   = data.get("RI"); ri_s = f"{ri:.2f}" if ri is not None else "—"
        sd   = data.get("SD"); sd_s = f"{sd:.2f}" if sd is not None else "—"
        rows = [
            # _fnum: antes un campo vacío salía como «None cm/s».
            ("PS",    f"{_fnum(data.get('PS'), 2)} cm/s"),
            ("ED",    f"{_fnum(data.get('ED'), 2)} cm/s"),
            ("TAmax", (f"{_fnum(data.get('TAmax'), 2)} cm/s"
                       if data.get("TAmax") is not None else "—")),
            ("PI",    pi_s),
            ("RI",    ri_s),
            ("S/D",   sd_s),
            ("PAM",   _pam_txt(data)),
            ("PaCO₂", _co2_txt(data).split(" (")[0]),
            ("Hora",  data.get("timestamp", "")),
        ]
        pic_c = data.get("PIC_label", "")
        pic_b = data.get("PIC_bellner", "")
        ppc_c = data.get("PPC_czosnyka_str", "")
        ppc_b = data.get("PPC_bellner_str", "")
        if art == "ACM" and (pic_c or pic_b or ppc_c or ppc_b):
            rows += [
                ("PIC Czos", pic_c or "—"),
                ("PIC Bell", pic_b or "—"),
                ("PPC Czos", ppc_c or "—"),
                ("PPC Bell", ppc_b or "—"),
            ]
        box = _mp.FancyBboxPatch((0.50, 0.02), 0.49, 0.60, boxstyle="round,pad=0.01",
                                  transform=ax.transAxes, facecolor=bg2,
                                  edgecolor=border, alpha=0.92, zorder=4)
        ax.add_patch(box)
        rh = 0.60 / len(rows)
        for i, (lbl, val) in enumerate(rows):
            ry = 0.02 + 0.60 - (i + 0.6) * rh
            ax.text(0.515, ry, lbl, transform=ax.transAxes, va="center", ha="left",
                    color=dim2, fontsize=9, zorder=5)
            _vcol = (pi_col if lbl == "PI" else
                     _paleta(printable)["YELLOW"] if (lbl == "PaCO₂" and _co2_fuera(data))
                     else dim)
            val = str(val).replace("⚠ PI desde VM aprox.", "⚠ PI aprox.")
            ax.text(0.975, ry, val, transform=ax.transAxes, va="center", ha="right",
                    color=_vcol, fontsize=(9 if len(val) <= 24 else 7.5),
                    fontweight="bold", zorder=5)

    # ── Exportar un solo panel como imagen legible ──────────────────────────────
    def _exportar_slot(self, idx):
        try:
            art  = self.slot_art[idx].get()
            lado = self.slot_lado[idx].get()
            data = self.slot_data[idx]

            fig, ax = _nuevos_subplots(figsize=(7.5, 4.2), facecolor=BG)
            self._dibujar_ax(ax, art, lado, data, BG2, BORDER, TEXT, TEXT2)
            fig.tight_layout()

            path = filedialog.asksaveasfilename(
                parent=self.root,
                defaultextension=".png",
                filetypes=[("PNG", "*.png")],
                initialfile=(f"TCD_{art}_{lado}_"
                             f"{datetime.now().strftime('%Y%m%d_%H%M')}.png"),
                title="Exportar imagen")
            if path:
                fig.savefig(path, dpi=200, bbox_inches="tight", facecolor=BG)
                messagebox.showinfo("Guardado", f"Imagen guardada en:\n{path}",
                                     parent=self.root)
            _cerrar_figura(fig)
        except Exception:
            messagebox.showerror("Error al exportar", traceback.format_exc(),
                                 parent=self.root)

    # ── Informe consolidado con las arterias del modo activo (2 / 4 / 6) ─────────
    def _generar_reporte(self):
        try:

            printable = messagebox.askyesno(
                "Versión del informe",
                "¿Generar versión IMPRIMIBLE (fondo blanco, ahorra tinta)?\n\n"
                "Sí = imprimible (fondo blanco)\n"
                "No = versión pantalla (fondo negro, estilo monitor)",
                parent=self.root)

            if printable:
                bg, bg2, border, txt, txt2 = ("#FFFFFF", "#FFFFFF", "#999999",
                                               "#000000", "#3A3A3A")
            else:
                bg, bg2, border, txt, txt2 = BG, BG2, BORDER, TEXT, TEXT2

            nombre = self.pac_nombre or "Sin nombre"
            hc     = self.pac_hc or "—"
            ts     = datetime.now().strftime("%d/%m/%Y  %H:%M")

            n_slots = self.panel_mode          # 2, 4 o 6, según el modo activo
            n_rows  = self.MODOS.get(n_slots, 2)
            # 3.6" por fila. Con una sola fila el título necesita más margen
            # relativo, o se solapa con el borde superior de los ejes.
            fig_h   = 3.6 * n_rows + (0.9 if n_rows == 1 else 0.0)
            top     = 0.85 if n_rows == 1 else 0.92
            bottom  = 0.11 if n_rows == 1 else 0.06
            fig = _nueva_figura(figsize=(18, fig_h), facecolor=bg)
            gsp = gridspec.GridSpec(n_rows, 2, figure=fig, left=0.03, right=0.98,
                                top=top, bottom=bottom, wspace=0.14, hspace=0.34)

            fig.suptitle(
                f"Doppler Transcraneal — Otras arterias    ·    {nombre}"
                f"   HC {hc}   ·   {ts}",
                color=txt, fontsize=15, fontweight="bold")

            for i in range(n_slots):
                r, c = divmod(i, 2)
                ax   = fig.add_subplot(gsp[r, c])
                art  = self.slot_art[i].get()
                lado = self.slot_lado[i].get()
                data = self.slot_data[i]
                self._dibujar_ax(ax, art, lado, data, bg2, border, txt, txt2,
                                 printable=printable)

            fig.text(0.5, 0.012,
                     "Valores de referencia bibliográfica por arteria (ver pestaña "
                     "Referencia del módulo TCD)  ·  Estimación PIC no invasiva "
                     "requiere correlación clínica",
                     ha="center", color=txt2, fontsize=10)

            path = filedialog.asksaveasfilename(
                parent=self.root,
                defaultextension=".png",
                filetypes=[("PNG", "*.png")],
                initialfile=(f"Informe_OtrasArterias_{nombre.replace(' ', '_')}_"
                             f"{datetime.now().strftime('%Y%m%d_%H%M')}.png"),
                title="Guardar informe")
            if path:
                fig.savefig(path, dpi=200, bbox_inches="tight", facecolor=bg)
                messagebox.showinfo("Guardado", f"Informe guardado en:\n{path}",
                                    parent=self.root)
            _cerrar_figura(fig)
        except Exception:
            messagebox.showerror("Error al generar informe", traceback.format_exc(),
                                 parent=self.root)


class NeurocriticalApp:

    def __init__(self, root):
        self.root        = root
        self.vno_data    = None
        self.tcd_data    = {"D": None, "I": None}
        self.tcd_win     = {"D": None, "I": None}
        self.otras_arterias_win = None
        self._otras_arterias_app = None
        self.otras_data  = {}     # "ACA|Derecho" → data (panel Otras arterias)
        self.ventanas    = {}     # "ACP|D" → ventana acústica de esa arteria
        self.pac_nombre  = ""
        self.pac_hc      = ""
        self.pac_lbl     = None   # creado en _build_ui

        self.root.title("Ultrasonografía Neurocrítica  [" + _BUILD_ID + "]")
        self.root.geometry("1275x825")
        self.root.configure(bg=BG)
        self.root.resizable(True, True)
        self.root.minsize(1025, 625)

        _apply_ttk_dark_theme(self.root)
        self._build_ui()
        self.root.bind("<Configure>", self._on_resize)
        self._last_w = self.root.winfo_width()
        self._last_h = self.root.winfo_height()

        # Al iniciar no se piden datos directamente: primero se elige entre
        # "Nuevo paciente" (que entonces pedirá nombre/HC) o "Abrir caso".
        self._after_tick = None
        self.root.protocol("WM_DELETE_WINDOW", self._salir)
        self.root.after(60, self._startup_choice)

    # ── Cierre de la aplicación ───────────────────────────────────────────────
    def _salir(self):
        """Cierre ordenado: cancela el reloj, destruye las ventanas hijas
        (cada DopplerACMApp se limpia en su <Destroy>), mata procesos de
        audio que quedaran vivos y sale del bucle de Tk SIEMPRE con quit():
        mainloop() retorna aunque algún componente hubiera dejado otra
        ventana raíz oculta."""
        if self._after_tick:
            try:
                self.root.after_cancel(self._after_tick)
            except tk.TclError:
                pass
            self._after_tick = None
        for w in list(self.root.winfo_children()):
            if isinstance(w, tk.Toplevel):
                try:
                    w.destroy()
                except tk.TclError:
                    pass
        _apagado_final()
        try:
            self.root.quit()
            self.root.destroy()
        except tk.TclError:
            pass

    # ── Paciente ──────────────────────────────────────────────────────────────
    def _pedir_paciente(self, nuevo=True):
        dlg = tk.Toplevel(self.root)
        dlg.title("Datos del paciente")
        dlg.geometry("500x287")
        dlg.resizable(False, False)
        dlg.configure(bg=BG)
        dlg.grab_set()
        dlg.transient(self.root)

        tk.Label(dlg, text="Datos del Paciente",
                 font=F_TITLE, bg=BG, fg=ACCENT
                 ).pack(pady=(22, 6))

        form = tk.Frame(dlg, bg=BG)
        form.pack(padx=30, fill="x")

        tk.Label(form, text="Nombre:", font=F_BODY, bg=BG, fg=TEXT2,
                 anchor="w").grid(row=0, column=0, sticky="w", pady=8)
        # nuevo=True → campos EN BLANCO. Prellenar con el paciente anterior
        # invita a confirmar sin darse cuenta y atribuir el estudio a quien no es.
        nom_var = tk.StringVar(value="" if nuevo else self.pac_nombre)
        nom_e = tk.Entry(form, textvariable=nom_var, width=28,
                         font=F_MONO, bg=BG3, fg=TEXT,
                         insertbackground=TEXT, relief="flat",
                         highlightthickness=1, highlightbackground=BORDER,
                         highlightcolor=ACCENT)
        nom_e.grid(row=0, column=1, padx=10, pady=8)

        tk.Label(form, text="HC / ID:", font=F_BODY, bg=BG, fg=TEXT2,
                 anchor="w").grid(row=1, column=0, sticky="w", pady=4)
        hc_var = tk.StringVar(value="" if nuevo else self.pac_hc)
        hc_e = tk.Entry(form, textvariable=hc_var, width=28,
                        font=F_MONO, bg=BG3, fg=TEXT,
                        insertbackground=TEXT, relief="flat",
                        highlightthickness=1, highlightbackground=BORDER,
                        highlightcolor=ACCENT)
        hc_e.grid(row=1, column=1, padx=10, pady=4)

        confirmed = [False]

        def _ok():
            n = nom_var.get().strip()
            if not n:
                nom_e.config(highlightbackground=RED)
                return
            self.pac_nombre  = n
            self.pac_hc      = hc_var.get().strip()
            confirmed[0]     = True
            dlg.destroy()

        bf = tk.Frame(dlg, bg=BG)
        bf.pack(pady=14)
        tk.Button(bf, text="Confirmar",
                  font=F_BODY, bg=BG3, fg=ACCENT,
                  relief="flat", bd=0, padx=18, pady=7,
                  highlightthickness=1, highlightbackground=ACCENT,
                  activebackground=BORDER, activeforeground=ACCENT,
                  cursor="hand2",
                  command=_ok).pack(side="left", padx=8)
        if nuevo:
            tk.Button(bf, text="Cancelar",
                      font=F_BODY, bg=BG3, fg=TEXT2,
                      relief="flat", bd=0, padx=12, pady=7,
                      activebackground=BORDER, cursor="hand2",
                      command=dlg.destroy).pack(side="left", padx=8)

        nom_e.focus_set()
        dlg.bind("<Return>", lambda e: _ok())
        dlg.wait_window()
        return confirmed[0]

    def _startup_choice(self):
        """Pantalla inicial: elegir entre nuevo paciente o abrir un caso."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Ultrasonografía Neurocrítica")
        dlg.geometry("560x300")
        dlg.resizable(False, False)
        dlg.configure(bg=BG)
        dlg.transient(self.root)
        dlg.grab_set()

        tk.Label(dlg, text="Ultrasonografía Neurocrítica",
                 font=F_TITLE, bg=BG, fg=ACCENT
                 ).pack(pady=(36, 6))
        tk.Label(dlg, text="¿Qué deseas hacer?",
                 font=F_BODY, bg=BG, fg=TEXT).pack(pady=(0, 26))

        bf = tk.Frame(dlg, bg=BG)
        bf.pack()

        def _nuevo():
            dlg.destroy()
            self._nuevo_paciente()

        def _abrir():
            dlg.destroy()
            self._abrir_caso()

        tk.Button(bf, text="Nuevo paciente",
                  font=F_BODY, bg=BG3, fg=ACCENT,
                  relief="flat", bd=0, padx=22, pady=12,
                  highlightthickness=1, highlightbackground=ACCENT,
                  activebackground=BORDER, activeforeground=ACCENT,
                  cursor="hand2", command=_nuevo).pack(side="left", padx=12)
        tk.Button(bf, text="Abrir caso",
                  font=F_BODY, bg=BG3, fg=TEXT,
                  relief="flat", bd=0, padx=22, pady=12,
                  activebackground=BORDER, activeforeground=TEXT,
                  cursor="hand2", command=_abrir).pack(side="left", padx=12)

        # Centrar sobre la ventana principal
        dlg.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width()  - dlg.winfo_width())  // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dlg.winfo_height()) // 2
        dlg.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _nuevo_paciente(self):
        if self._pedir_paciente(nuevo=True):
            self.vno_data = None
            self.tcd_data = {"D": None, "I": None}
            self.otras_data = {}
            self.ventanas = {}
            self._cerrar_otras_arterias()
            self.vno_status_var.set("NO MEDIDO")
            self.vno_status_lbl.config(fg=TEXT2)
            self.vno_detail_var.set("")
            for k in ("D", "I"):
                p = self._tcd_panels[k]
                self._set_info_text(k)
                p["pic_lbl"].config(text="", fg=TEXT2)
            self._update_pac_lbl()
            self._redraw_all()
            self._set_status(f"Nuevo paciente: {self.pac_nombre}  HC {self.pac_hc}")

    def _update_pac_lbl(self):
        if self.pac_lbl is None:
            return
        nombre = self.pac_nombre or "—"
        hc     = self.pac_hc     or "—"
        self.pac_lbl.config(text=f"👤  {nombre}   HC {hc}")

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        hdr = tk.Frame(self.root, bg=BG2, height=70,
                       highlightthickness=1, highlightbackground=BORDER)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="🔬  Ultrasonografía Neurocrítica",
                 font=F_TITLE, bg=BG2, fg=TEXT
                 ).pack(side="left", padx=22, pady=8)

        # ── Menú desplegable "Inicio": nuevo paciente / abrir / guardar caso ─
        self._btn_inicio = tk.Menubutton(hdr, text="☰  Inicio  ▾",
                  font=F_BODY, bg=BG3, fg=TEXT2,
                  relief="flat", bd=0, padx=12, pady=5,
                  activebackground=BORDER, activeforeground=TEXT,
                  cursor="hand2")
        self._btn_inicio.pack(side="left", padx=(0, 6))
        # Igual que en "Otras arterias": el Menu debe ser hijo del Menubutton.
        inicio_menu = tk.Menu(self._btn_inicio, tearoff=0,
                               bg=BG3, fg=TEXT, activebackground=ACCENT,
                               activeforeground=BG, font=F_BODY,
                               bd=0, relief="flat")
        inicio_menu.add_command(label="👤  Nuevo paciente", command=self._nuevo_paciente)
        inicio_menu.add_command(label="📂  Abrir caso",     command=self._abrir_caso)
        inicio_menu.add_command(label="💾  Guardar caso",   command=self._guardar_caso)
        self._btn_inicio.config(menu=inicio_menu)
        self._btn_inicio.menu = inicio_menu  # referencia viva (evita garbage collection)

        self.clock_lbl = tk.Label(hdr, text="", font=F_STATUS,
                                   bg=BG2, fg=TEXT2)
        self.clock_lbl.pack(side="right", padx=18)
        tk.Button(hdr, text="ⓘ  Créditos",
                  font=F_SMALL, bg=BG2, fg=TEXT2,
                  relief="flat", bd=0, padx=8, pady=4,
                  activebackground=BG3, activeforeground=ACCENT,
                  cursor="hand2",
                  command=lambda: _mostrar_creditos(self.root)
                  ).pack(side="right", padx=6)

        # ── Botón ÚNICO de informe: Doppler TCD bilateral (sin VNO) ──────────
        self._btn_informe = tk.Button(hdr, text="📄  Informe Doppler TCD",
                  font=F_BODY, bg=BG3, fg=GOLD,
                  relief="flat", bd=0, padx=12, pady=5,
                  highlightthickness=1, highlightbackground=GOLD,
                  highlightcolor=GOLD,
                  activebackground=BORDER, activeforeground=GOLD,
                  cursor="hand2", command=self._generar_informe_tcd)
        self._btn_informe.pack(side="right", padx=6)

        # ── Captura PNG: VNO + ACM D + ACM I (pantalla o imprimible) ─────────
        tk.Button(hdr, text="🖼  Captura",
                  font=F_BODY, bg=BG3, fg=TEXT2,
                  relief="flat", bd=0, padx=12, pady=5,
                  activebackground=BORDER, activeforeground=TEXT,
                  cursor="hand2", command=self._generar_informe
                  ).pack(side="right", padx=6)

        # ── Botón "Otras arterias" (dividido) ────────────────────────────────
        # Clic en el texto  → abre el Panel comparativo (2 / 4 / 6, alternable)
        # Clic en la flecha → despliega el submenú para ir directo a una arteria
        otras_wrap = tk.Frame(hdr, bg=BG3,
                               highlightthickness=1, highlightbackground=ACCENT)
        otras_wrap.pack(side="right", padx=6)

        tk.Button(otras_wrap, text="🧠  Otras arterias",
                  font=F_BODY, bg=BG3, fg=ACCENT,
                  relief="flat", bd=0, padx=12, pady=5,
                  activebackground=BORDER, activeforeground=ACCENT,
                  cursor="hand2",
                  command=lambda: self._open_otras_arterias()
                  ).pack(side="left")

        tk.Frame(otras_wrap, bg=ACCENT, width=1).pack(side="left", fill="y", pady=6)

        self._btn_otras_arrow = tk.Menubutton(otras_wrap, text="▾",
                  font=F_BODY, bg=BG3, fg=ACCENT,
                  relief="flat", bd=0, padx=8, pady=5,
                  activebackground=BORDER, activeforeground=ACCENT,
                  cursor="hand2")
        self._btn_otras_arrow.pack(side="left")

        # El Menu DEBE ser hijo del propio Menubutton (no de self.root):
        # si se crea con otro master, en varios builds de Tk el menú no
        # llega a desplegarse al hacer clic (falla silenciosa).
        otras_menu = tk.Menu(self._btn_otras_arrow, tearoff=0,
                              bg=BG3, fg=TEXT, activebackground=ACCENT,
                              activeforeground=BG, font=F_BODY,
                              bd=0, relief="flat")
        otras_menu.add_command(
            label="Ir directo a una arteria:", state="disabled")
        otras_menu.add_separator()
        for _art in DopplerACMApp.RANGOS_ARTERIA.keys():
            if _art == "ACM":
                continue   # ACM ya tiene su panel dedicado D/I
            otras_menu.add_command(
                label=_art,
                command=lambda a=_art: self._open_otras_arterias(foco=a))

        self._btn_otras_arrow.config(menu=otras_menu)
        self._btn_otras_arrow.menu = otras_menu  # referencia viva (evita garbage collection)


        self.pac_lbl = tk.Label(hdr, text="", font=F_HEAD,
                                 bg=BG2, fg=ACCENT)
        self.pac_lbl.pack(side="right", padx=14)
        self._tick()
        self._update_pac_lbl()

        self.body = tk.Frame(self.root, bg=BG)
        self.body.pack(fill="both", expand=True)
        self.body.columnconfigure(0, weight=35)   # VNO más estrecho
        self.body.columnconfigure(1, weight=65)   # TCD más ancho
        self.body.rowconfigure(0, weight=1)

        # ── Panel VNO izquierda ───────────────────────────────────────────────
        self.vno_frame = tk.Frame(self.body, bg=BG)
        self.vno_frame.grid(row=0, column=0, sticky="nsew")

        tk.Label(self.vno_frame, text="VAINA DE N. ÓPTICO",
                 font=F_LABEL, bg=BG, fg=TEXT
                 ).pack(pady=(20, 6))

        self.vno_canvas = tk.Canvas(self.vno_frame, bg=BG, highlightthickness=0)
        self.vno_canvas.pack(fill="both", expand=True, padx=16)

        self.vno_status_var = tk.StringVar(value="NO MEDIDO")
        self.vno_status_lbl = tk.Label(self.vno_frame,
                                        textvariable=self.vno_status_var,
                                        font=F_CHANNEL,
                                        bg=BG, fg=TEXT2)
        self.vno_status_lbl.pack(pady=(6, 2))

        self.vno_detail_var = tk.StringVar(value="")
        tk.Label(self.vno_frame, textvariable=self.vno_detail_var,
                 font=F_DATA, bg=BG, fg=TEXT,
                 wraplength=360, justify="center").pack(pady=(0, 8))

        tk.Button(self.vno_frame, text="Medir VNO",
                  font=F_BODY, bg=BG3, fg=TEXT2,
                  relief="flat", bd=0, padx=22, pady=9,
                  highlightthickness=1, highlightbackground=BORDER,
                  activebackground=BORDER, activeforeground=TEXT,
                  cursor="hand2", command=self._open_vno
                  ).pack(pady=(0, 16))

        tk.Frame(self.body, bg=BORDER, width=1
                 ).grid(row=0, column=0, sticky="nse")

        # ── Panel TCD derecha — dos sub-filas ─────────────────────────────────
        tcd_outer = tk.Frame(self.body, bg=BG)
        tcd_outer.grid(row=0, column=1, sticky="nsew")
        tcd_outer.columnconfigure(0, weight=1)
        tcd_outer.rowconfigure(0, weight=1)
        tcd_outer.rowconfigure(1, weight=0)
        tcd_outer.rowconfigure(2, weight=1)

        tk.Frame(tcd_outer, bg=BORDER, height=2
                 ).grid(row=1, column=0, sticky="ew", pady=(8, 0))

        self._tcd_panels = {}
        for idx, (lado_key, lado_label, lado_full) in enumerate([
                ("D", "ACM  DERECHA",   "Derecha"),
                ("I", "ACM  IZQUIERDA", "Izquierda")]):
            grid_row = 0 if idx == 0 else 2
            panel = tk.Frame(tcd_outer, bg=BG)
            panel.grid(row=grid_row, column=0, sticky="nsew")
            panel.columnconfigure(0, weight=1)
            panel.rowconfigure(1, weight=1)

            hf = tk.Frame(panel, bg=BG)
            hf.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 0))
            tk.Label(hf, text=lado_label,
                     font=F_CHANNEL, bg=BG, fg=TEXT
                     ).pack(side="left")
            pic_lbl = tk.Label(hf, text="",
                               font=F_CHANNEL_S,
                               bg=BG, fg=TEXT2)
            pic_lbl.pack(side="right")

            # Borde de 2 px: toma el color de la alerta (rojo/ámbar) o BG.
            cv = tk.Canvas(panel, bg=BG, highlightthickness=2,
                           highlightbackground=BG, highlightcolor=BG)
            cv.grid(row=1, column=0, sticky="nsew", padx=12)

            # Text widget de 2 líneas para colorear PI inline
            info_text = tk.Text(panel, height=3, bg=BG, fg=TEXT,
                                font=F_BODY,
                                relief="flat", bd=0,
                                highlightthickness=0,
                                wrap="word", state="disabled",
                                cursor="arrow")
            info_text.grid(row=2, column=0, pady=(4, 4), sticky="ew")
            info_text.tag_configure("center", justify="center")
            info_text.tag_configure("normal",  foreground=TEXT2)
            info_text.tag_configure("bajo",    foreground=ACCENT)
            info_text.tag_configure("verde",   foreground=GREEN)
            info_text.tag_configure("amarillo",foreground=YELLOW)
            info_text.tag_configure("rojo",    foreground=RED)
            info_text.tag_configure("dim",     foreground=TEXT2)

            btn_row = tk.Frame(panel, bg=BG)
            btn_row.grid(row=3, column=0, pady=(2, 10))

            btn = tk.Button(btn_row,
                            text=f"Abrir TCD — ACM {lado_full}",
                            font=F_BODY, bg=BG3, fg=TEXT2,
                            relief="flat", bd=0, padx=16, pady=7,
                            activebackground=BORDER, activeforeground=TEXT,
                            cursor="hand2",
                            command=lambda k=lado_key: self._open_tcd(k))
            btn.pack(side="left", padx=(0, 6))

            self._tcd_panels[lado_key] = {
                "canvas": cv, "info_text": info_text, "pic_lbl": pic_lbl,
            }

        sb = tk.Frame(self.root, bg=BG3, height=40,
                      highlightthickness=1, highlightbackground=BORDER)
        sb.pack(fill="x", side="bottom")
        sb.pack_propagate(False)
        self.status_lbl = tk.Label(sb, text="Listo  ·  Seleccione un módulo",
                                    font=F_STATUS, bg=BG3, fg=TEXT2)
        self.status_lbl.pack(side="left", padx=12, pady=5)
        tk.Label(sb, text="App: " + _autor_app(),
                 font=F_SMALL, bg=BG3, fg=TEXT3
                 ).pack(side="right", padx=12)


        self.root.update_idletasks()
        self._redraw_all()

    # ── Resize ────────────────────────────────────────────────────────────────
    def _on_resize(self, event):
        w, h = self.root.winfo_width(), self.root.winfo_height()
        if abs(w - self._last_w) > 4 or abs(h - self._last_h) > 4:
            self._last_w, self._last_h = w, h
            self.root.after(60, self._redraw_all)

    def _redraw_all(self):
        self._redraw_vno()
        for k in ("D", "I"):
            self._redraw_tcd(k)

    # ── VNO ───────────────────────────────────────────────────────────────────
    def _redraw_vno(self):
        cv = self.vno_canvas
        cv.update_idletasks()
        cw, ch = cv.winfo_width(), cv.winfo_height()
        if cw < 10 or ch < 10:
            return
        cv.delete("all")
        sc = min(cw / 240, ch / 180) * 0.85
        cx = cw // 2
        if self.vno_data:
            od, oi = self.vno_data["OD_mm"], self.vno_data["OI_mm"]
            # Codificación por gravedad: verde (normal) / amarillo (gris) / rojo (elevada)
            col_od = _vno_cat(od)[2]
            col_oi = _vno_cat(oi)[2]
            _draw_nervio_optico(cv, cx - int(38*sc), ch//2 - 10, col_od, sc)
            _draw_nervio_optico(cv, cx + int(38*sc), ch//2 - 10, col_oi, sc)
            fy = ch//2 - 10 + int(36*sc) + 16
            fs = max(9, int(10*sc))
            cv.create_text(cx - int(38*sc), fy, text=f"OD {od:.1f}",
                           fill=col_od, font=(_FAM_DATA, fs, "bold"))
            cv.create_text(cx + int(38*sc), fy, text=f"OI {oi:.1f}",
                           fill=col_oi, font=(_FAM_DATA, fs, "bold"))
        else:
            _draw_nervio_optico(cv, cx - int(38*sc), ch//2 - 10, GRAY, sc)
            _draw_nervio_optico(cv, cx + int(38*sc), ch//2 - 10, GRAY, sc)

    # ── TCD ───────────────────────────────────────────────────────────────────
    def _color_onda(self, data, win):
        """Color de la onda: alerta roja/ámbar > color por PIC > en uso > gris."""
        if data:
            al = _alertas_tcd(data)
            if any(n == "rojo" for n, _ in al):
                return RED
            if al:
                return YELLOW
            return _pic_color(data.get("PIC_label", ""))
        if win and win.winfo_exists():
            return ACCENT
        return GRAY

    def _alerta_canvas(self, cv, cw, ch, data):
        """Borde de color + barra de alarma en la parte baja del canvas."""
        _alerta_canvas_tk(cv, cw, ch, data)

    def _redraw_tcd(self, lado_key):
        p  = self._tcd_panels[lado_key]
        cv = p["canvas"]
        cv.update_idletasks()
        cw, ch = cv.winfo_width(), cv.winfo_height()
        if cw < 10 or ch < 10:
            return
        cv.delete("all")
        win  = self.tcd_win[lado_key]
        data = self.tcd_data[lado_key]

        # Alerta (vasoespasmo / diástole muy baja) > color PIC > en uso > gris
        color = self._color_onda(data, win)

        mx     = int(cw * 0.09)   # margen para etiquetas Y
        H_max  = int(ch * 0.78)
        y_base = int(ch * 0.88)

        ps_self = data["PS"] if data and data.get("PS") else None
        other_key  = "I" if lado_key == "D" else "D"
        other_data = self.tcd_data.get(other_key)
        ps_other = other_data["PS"] if other_data and other_data.get("PS") else None

        if ps_self and ps_other:
            ps_ref = max(ps_self, ps_other)
            H = max(int(H_max * ps_self / ps_ref), 30)
        else:
            H = max(H_max, 30)

        ps_v = data.get("PS") if data else None
        ed_v = data.get("ED") if data else None
        _draw_doppler_wave(cv, mx, y_base, color,
                           w=max(cw - 2*mx, 40), h=H,
                           ps=ps_v, ed=ed_v)

        # ── Escala Y elegante con líneas guía ───────────────────────────────────
        if ps_v and ps_v > 0:
            import math as _m
            ps_r = max(ps_v, ps_other or 0) * 1.15
            raw  = ps_r / 4.0
            mag  = 10 ** int(_m.floor(_m.log10(raw))) if raw >= 1 else 10
            step = max(int(round(raw / mag) * mag), 10)
            fnt  = (_FAM_DATA, max(7, int(cw * 0.018)))
            for tv in range(0, int(ps_r) + 1, step):
                ty = y_base - int(tv / ps_r * H)
                if ty < y_base - H - 4 or ty > y_base + 2: continue
                cv.create_line(mx, ty, cw - 4, ty,
                               fill="#333333", width=1, dash=(2,8), tags="icon")
                cv.create_text(mx - 4, ty, text=str(int(tv)),
                               fill="#FFFFFF", anchor="e",
                               font=fnt, tags="icon")

        self._alerta_canvas(cv, cw, ch, data)

        if data and ps_other:
            self.root.after(10, lambda k=other_key: self._redraw_tcd_nopropagate(k))

    def _redraw_tcd_nopropagate(self, lado_key):
        p  = self._tcd_panels[lado_key]
        cv = p["canvas"]
        cv.update_idletasks()
        cw, ch = cv.winfo_width(), cv.winfo_height()
        if cw < 10 or ch < 10:
            return
        cv.delete("all")
        win  = self.tcd_win[lado_key]
        data = self.tcd_data[lado_key]
        color = self._color_onda(data, win)
        mx     = int(cw * 0.09)   # margen para etiquetas Y
        H_max  = int(ch * 0.78)
        y_base = int(ch * 0.88)
        ps_self = data["PS"] if data and data.get("PS") else None
        other_data = self.tcd_data.get("D" if lado_key == "I" else "I")
        ps_other = other_data["PS"] if other_data and other_data.get("PS") else None
        if ps_self and ps_other:
            H = max(int(H_max * ps_self / max(ps_self, ps_other)), 30)
        else:
            H = max(H_max, 30)
        ps_v2 = data.get("PS") if data else None
        _draw_doppler_wave(cv, mx, y_base, color,
                           w=max(cw - 2*mx, 40), h=H,
                           ps=ps_v2, ed=data.get("ED") if data else None)

        if ps_v2 and ps_v2 > 0:
            import math as _m2
            ps_r2 = max(ps_v2, ps_other or 0) * 1.15
            raw2  = ps_r2 / 4.0
            mag2  = 10 ** int(_m2.floor(_m2.log10(raw2))) if raw2 >= 1 else 10
            step2 = max(int(round(raw2 / mag2) * mag2), 10)
            fnt2  = (_FAM_DATA, max(7, int(cw * 0.018)))
            for tv2 in range(0, int(ps_r2) + 1, step2):
                ty2 = y_base - int(tv2 / ps_r2 * H)
                if ty2 < y_base - H - 4 or ty2 > y_base + 2: continue
                cv.create_line(mx, ty2, cw - 4, ty2,
                               fill="#333333", width=1, dash=(2,8), tags="icon")
                cv.create_text(mx - 4, ty2, text=str(int(tv2)),
                               fill="#FFFFFF", anchor="e",
                               font=fnt2, tags="icon")

        self._alerta_canvas(cv, cw, ch, data)

    # ── Reloj ─────────────────────────────────────────────────────────────────
    def _tick(self):
        try:
            self.clock_lbl.config(text=datetime.now().strftime("%d/%m/%Y  %H:%M:%S"))
            self._after_tick = self.root.after(1000, self._tick)
        except tk.TclError:
            self._after_tick = None

    def _set_status(self, msg, color=TEXT2):
        self.status_lbl.config(text=msg, fg=color)

    # ── Persistencia de casos (.ncu) ────────────────────────────────────────────
    def _guardar_caso(self):
        if not (self.vno_data or self.tcd_data["D"] or self.tcd_data["I"]
                or self.otras_data):
            messagebox.showwarning("Sin datos",
                                   "No hay mediciones que guardar todavía.",
                                   parent=self.root)
            return
        caso = {
            "paciente": self.pac_nombre,
            "hc":       self.pac_hc,
            "vno":      self.vno_data,
            "tcd":      self.tcd_data,
            "otras":    self.otras_data,
            "ventanas": self.ventanas,
            "guardado": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "_app":     _meta_app(),   # crédito del software, no del médico
        }
        base = (self.pac_nombre or "caso").replace(" ", "_")
        path = filedialog.asksaveasfilename(
            parent=self.root,
            initialdir=_base_dir(),
            defaultextension=".ncu",
            initialfile=f"{base}_{datetime.now().strftime('%Y%m%d_%H%M')}.ncu",
            filetypes=[("Caso Neurocrítico", "*.ncu")],
            title="Guardar caso")
        if not path:
            return
        try:
            guardar_caso(caso, path)
            messagebox.showinfo("Guardado", f"Caso guardado en:\n{path}",
                                parent=self.root)
            self._set_status(f"Caso guardado: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Error al guardar", str(e), parent=self.root)

    def _abrir_caso(self):
        path = filedialog.askopenfilename(
            parent=self.root,
            initialdir=_base_dir(),
            filetypes=[("Caso Neurocrítico", "*.ncu"), ("Todos", "*.*")],
            title="Abrir caso")
        if not path:
            return
        try:
            caso = cargar_caso(path)
        except Exception as e:
            messagebox.showerror("No se pudo abrir",
                                 f"El archivo no es un caso válido.\n\n{e}",
                                 parent=self.root)
            return

        # Reiniciar estado y repoblar reutilizando los manejadores existentes
        self.pac_nombre = caso.get("paciente", "")
        self.pac_hc     = caso.get("hc", "")
        self.vno_data   = None
        self.tcd_data   = {"D": None, "I": None}
        self.otras_data = dict(caso.get("otras") or {})
        # Solo formato por arteria ("ACM|D"); formatos antiguos se ignoran.
        self.ventanas   = {k: v for k, v in (caso.get("ventanas") or {}).items()
                           if "|" in str(k)}
        self._cerrar_otras_arterias()
        self.vno_status_var.set("NO MEDIDO")
        self.vno_status_lbl.config(fg=TEXT2)
        self.vno_detail_var.set("")
        for k in ("D", "I"):
            self._set_info_text(k)
            self._tcd_panels[k]["pic_lbl"].config(text="", fg=TEXT2)
        self._update_pac_lbl()

        vno = caso.get("vno")
        if vno:
            self._on_vno_result(vno)
        for k in ("D", "I"):
            d = (caso.get("tcd") or {}).get(k)
            if d:
                self._on_tcd_result(k, d)

        self._redraw_all()
        self._set_status(f"Caso cargado: {os.path.basename(path)}"
                         f"  ({caso.get('paciente','—')})")

    # ── VNO actions ───────────────────────────────────────────────────────────
    def _open_vno(self):
        VNOWindow(self.root, self._on_vno_result)
        self._set_status("Módulo VNO abierto")

    def _on_vno_result(self, data):
        self.vno_data = data
        od, oi = data["OD_mm"], data["OI_mm"]
        ts = data["timestamp"]
        clave = data.get("categoria",
                          "elevada" if data.get("elevada") else "normal")
        if clave == "elevada":
            self.vno_status_var.set("⬆  VAINA DILATADA")
            self.vno_status_lbl.config(fg=RED)
        elif clave == "gris":
            self.vno_status_var.set("≈  ZONA GRIS")
            self.vno_status_lbl.config(fg=YELLOW)
        else:
            self.vno_status_var.set("✓  NORMAL")
            self.vno_status_lbl.config(fg=GREEN)
        self.vno_detail_var.set(
            f"OD {od:.1f} mm  |  OI {oi:.1f} mm  |  Prom {data['promedio']:.2f} mm\n{ts}")
        self._set_status(f"VNO registrada — {ts}")
        self._redraw_vno()

    # ── TCD actions ───────────────────────────────────────────────────────────
    def _open_tcd(self, lado_key):
        win = self.tcd_win[lado_key]
        if win and win.winfo_exists():
            win.lift()
            return
        lbl = "Derecha" if lado_key == "D" else "Izquierda"
        w = tk.Toplevel(self.root)
        w.title(f"Doppler TCD — ACM {lbl}")
        w.configure(bg=BG)
        self.tcd_win[lado_key] = w
        tcd_app = DopplerACMApp(w, on_result=lambda d, k=lado_key: self._on_tcd_result(k, d))
        try:
            # Pre-llenar nombre, HC y lado siempre
            tcd_app.ent_nombre.delete(0, "end")
            tcd_app.ent_nombre.insert(0, self.pac_nombre)
            tcd_app.ent_hc.delete(0, "end")
            tcd_app.ent_hc.insert(0, self.pac_hc)
            tcd_app.lado_var.set("Derecho" if lado_key == "D" else "Izquierdo")

            # Pre-llenar parámetros del último análisis si existen
            _prellenar_tcd(tcd_app, self.tcd_data[lado_key])
        except Exception:
            pass
        self._set_info_text(lado_key, "EN USO...", "")
        self._redraw_tcd(lado_key)
        self._set_status(f"TCD ACM {lbl} abierto")

        def _on_close(k=lado_key, lb=lbl):
            self.tcd_win[k] = None
            w.destroy()          # destruir ANTES de redibujar para que winfo_exists() sea False
            if not self.tcd_data[k]:
                self._set_info_text(k, "SIN DATOS", "")
            self._set_status(f"TCD ACM {lb} cerrado")
            self._redraw_tcd(k)

        w.protocol("WM_DELETE_WINDOW", _on_close)

    # ── Mediciones del panel "Otras arterias" (persisten en el dashboard) ────
    def _on_otras_result(self, arteria, lado, data):
        key = f"{arteria}|{lado}"
        if data:
            self.otras_data[key] = data
            self._set_status(f"Registrado {arteria} {lado} para el informe")
        else:
            self.otras_data.pop(key, None)

    def _cerrar_otras_arterias(self):
        w = self.otras_arterias_win
        if w and w.winfo_exists():
            w.destroy()
        self.otras_arterias_win = None
        self._otras_arterias_app = None

    def _mediciones_informe(self):
        """Reúne las mediciones de TODA la app, identificando cada una por la
        arteria y el lado que trae el propio análisis (no por el botón desde
        el que se abrió):
          · Paneles ACM D/I del dashboard (self.tcd_data).
          · Panel "Otras arterias" (self.otras_data, persiste aunque se cierre).
          · Slots del panel "Otras arterias" abierto en este momento.
        Prioridad: ACM → panel del dashboard; ACA/ACP → panel Otras arterias.
        Si falta, se usa la otra fuente (p. ej. ACM medida en un slot, o ACA
        medida cambiando la arteria dentro de la ventana TCD del dashboard)."""
        lado_k = {"Derecho": "D", "Izquierdo": "I"}
        dash, otras = {}, {}
        for k in ("D", "I"):
            d = self.tcd_data.get(k)
            if d:
                art = d.get("arteria", "ACM")
                kk = lado_k.get(d.get("lado"), k)
                dash[(art, kk)] = d
        for key, d in (self.otras_data or {}).items():
            if not d or "|" not in key:
                continue
            art, lado = key.split("|", 1)
            if lado in lado_k:
                otras[(art, lado_k[lado])] = d
        app = getattr(self, "_otras_arterias_app", None)
        if app is not None:
            for i, d in enumerate(app.slot_data):
                if d and app.slot_key[i] and app.slot_key[i][1] in lado_k:
                    otras.setdefault((app.slot_key[i][0], lado_k[app.slot_key[i][1]]), d)
        med = {}
        for art, _ in ARTERIAS_INFORME:
            for k, *_ in LADOS_INFORME:
                if art == "ACM":
                    med[(art, k)] = dash.get((art, k)) or otras.get((art, k))
                else:
                    med[(art, k)] = otras.get((art, k)) or dash.get((art, k))
        return med

    # ── Diálogo: ventana acústica + motivo de arterias no visualizadas ──────
    def _dialogo_ventanas(self, med):
        """Ventana acústica POR ARTERIA. Opciones restringidas por coherencia:
        medida → Adecuada/Inadecuada; no visualizada → Inadecuada/Ausente."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Informe Doppler TCD — ventana acústica por arteria")
        dlg.configure(bg=BG2)
        dlg.transient(self.root)
        dlg.grab_set()
        res = {"ok": False}

        tk.Label(dlg, text="VENTANA ACÚSTICA POR ARTERIA", font=F_HEAD,
                 bg=BG2, fg=ACCENT).grid(row=0, column=0, columnspan=3,
                                         sticky="w", padx=16, pady=(14, 2))
        tk.Label(dlg, text="Arteria sin medición = no visualizada → ventana Inadecuada o Ausente",
                 font=F_SMALL, bg=BG2, fg=TEXT2).grid(row=1, column=0, columnspan=3,
                                                     sticky="w", padx=16, pady=(0, 8))
        v_vars = {}
        fila = 2
        for art, _nom in ARTERIAS_INFORME:
            for k, lado_full, _, _ in LADOS_INFORME:
                medida = bool(med.get((art, k)))
                tk.Label(dlg, text=f"{art} {lado_full}", font=F_BODY, bg=BG2,
                         fg=TEXT).grid(row=fila, column=0, sticky="w", padx=16, pady=2)
                tk.Label(dlg, text="✓ Medida" if medida else "✗ No visualizada",
                         font=F_BODY, bg=BG2, fg=GREEN if medida else YELLOW
                         ).grid(row=fila, column=1, sticky="w", padx=(0, 12))
                v = tk.StringVar(value=_ventana_coherente(
                    self.ventanas.get(f"{art}|{k}"), medida))
                ttk.Combobox(dlg, textvariable=v, state="readonly", width=20,
                             values=VENTANA_OPC_MEDIDA if medida else VENTANA_OPC_NO_VISTA
                             ).grid(row=fila, column=2, sticky="w", padx=(0, 16), pady=2)
                v_vars[(art, k)] = v
                fila += 1

        def _ok():
            res["ok"] = True
            dlg.destroy()
        bf = tk.Frame(dlg, bg=BG2)
        bf.grid(row=fila, column=0, columnspan=3, pady=14)
        tk.Button(bf, text="Generar informe", font=F_BODY, bg=BG3, fg=GOLD,
                  relief="flat", bd=0, padx=16, pady=6,
                  highlightthickness=1, highlightbackground=GOLD,
                  activebackground=BORDER, activeforeground=GOLD,
                  cursor="hand2", command=_ok).pack(side="left", padx=6)
        tk.Button(bf, text="Cancelar", font=F_BODY, bg=BG3, fg=TEXT2,
                  relief="flat", padx=16, pady=6, command=dlg.destroy
                  ).pack(side="left", padx=6)
        self.root.wait_window(dlg)
        if not res["ok"]:
            return False
        for (art, k), v in v_vars.items():
            self.ventanas[f"{art}|{k}"] = v.get()
        return True

    # ── Botón ÚNICO: informe DOCX Doppler TCD bilateral (sin VNO) ───────────
    def _generar_informe_tcd(self):
        med = self._mediciones_informe()
        if not self._dialogo_ventanas(med):
            return
        if not _asegurar_docx_disponible(self.root):
            return
        ventanas = {(a, k): self.ventanas.get(f"{a}|{k}") for a, _ in ARTERIAS_INFORME
                    for k, *_ in LADOS_INFORME}
        nombre = (self.pac_nombre or "paciente").strip() or "paciente"
        path = filedialog.asksaveasfilename(
            parent=self.root, defaultextension=".docx",
            filetypes=[("Word", "*.docx"), ("Todos", "*.*")],
            initialfile=nombre.replace(" ", "_") + "_TCD_" +
                        datetime.now().strftime("%Y%m%d_%H%M%S") + ".docx",
            title="Guardar informe Doppler TCD")
        if not path:
            return
        imagen = None
        try:
            fig = _renderizar_ambas_acm_imprimible(med[("ACM", "D")], med[("ACM", "I")])
            imagen = os.path.join(tempfile.gettempdir(), "tcd_bilat_" +
                                  datetime.now().strftime("%Y%m%d_%H%M%S") + ".png")
            fig.savefig(imagen, dpi=200, bbox_inches="tight",
                        facecolor=fig.get_facecolor())
            _cerrar_figura(fig)
            doc = _construir_informe_tcd_bilateral_docx(
                self.pac_nombre, self.pac_hc, med, ventanas, imagen)
            doc.save(path)
        except Exception:
            messagebox.showerror("Error al generar informe", traceback.format_exc(),
                                 parent=self.root)
            return
        finally:
            if imagen:
                try:
                    os.remove(imagen)
                except OSError:
                    pass
        self._set_status("✓ Informe TCD guardado: " + path)
        messagebox.showinfo("Informe generado",
                            "Informe Doppler TCD guardado en:\n\n" + path,
                            parent=self.root)

    def _open_otras_arterias(self, foco=None):
        """Abre (o trae al frente) el panel comparativo (2 / 4 / 6 arterias)."""
        win = self.otras_arterias_win
        if win and win.winfo_exists():
            win.lift()
            if foco and getattr(self, "_otras_arterias_app", None):
                # set_foco resincroniza el combo de lado; set() directo no lo hace
                self._otras_arterias_app.set_foco(foco, 0)
            return
        w = tk.Toplevel(self.root)
        self.otras_arterias_win = w
        app = OtrasArteriasApp(w, pac_nombre=self.pac_nombre,
                                pac_hc=self.pac_hc, foco=foco,
                                on_change=self._on_otras_result,
                                datos_previos=self.otras_data)
        self._otras_arterias_app = app

        def _on_close():
            self.otras_arterias_win = None
            self._otras_arterias_app = None
            w.destroy()
        w.protocol("WM_DELETE_WINDOW", _on_close)

    def _generar_informe(self):
        """Informe matplotlib. VNO proporcional, datos como leyenda interna."""
        try:
            import matplotlib.gridspec as _gs
            import matplotlib.patches as _mp
            import numpy as _np

            nombre = self.pac_nombre or "Sin nombre"
            hc     = self.pac_hc     or "—"
            ts     = datetime.now().strftime("%d/%m/%Y  %H:%M")

            # Preguntar versión antes de dibujar: pantalla (fondo negro, como
            # el monitor) o imprimible (fondo blanco, bajo consumo de tinta
            # al pegar/imprimir el informe).
            printable = messagebox.askyesno(
                "Versión del informe",
                "¿Generar versión IMPRIMIBLE (fondo blanco, ahorra tinta)?\n\n"
                "Sí = imprimible (fondo blanco)\n"
                "No = versión pantalla (fondo negro, estilo monitor)",
                parent=self.root)

            if printable:
                C = {
                    "bg":  "#FFFFFF", "bg2": "#FFFFFF", "bg3": "#F0F0F0",
                    "border":"#999999","text":"#000000","dim": "#000000",
                    "accent":"#0072B2","green":"#1A7A1A","yellow":"#B8860B",
                    "red":  "#B00000","purple":"#9A6300",
                }
            else:
                C = {
                    "bg":  "#000000", "bg2": "#111111", "bg3": "#1C1C1C",
                    "border":"#2A2A2A","text":"#FFFFFF","dim": "#FFFFFF",
                    "accent":"#00FFFF","green":"#00FF00","yellow":"#FFFF00",
                    "red":  "#FF2020","purple":"#FFB300",
                }

            # ── Figura ────────────────────────────────────────────────────────
            # Layout: col izq = VNO, col der = TCD (2 ondas apiladas)
            # Cada onda ocupa toda su celda; los datos van DENTRO como texto
            fig = _nueva_figura(figsize=(17, 10), facecolor=C["bg"])
            gs  = _gs.GridSpec(2, 2, figure=fig,
                               left=0.02, right=0.99,
                               top=0.96, bottom=0.04,
                               wspace=0.10, hspace=0.06,
                               width_ratios=[1, 2.4],
                               height_ratios=[1, 1])

            # VNO ocupa ambas filas
            # Caja exterior: llena toda la celda, da el fondo y borde
            ax_vno = fig.add_subplot(gs[:, 0])
            ax_vno.set_facecolor(C["bg2"])
            for sp in ax_vno.spines.values():
                sp.set_color(C["border"])
            ax_vno.set_xticks([]); ax_vno.set_yticks([])

            # TCD dos paneles independientes sin datos flotantes
            ax_d = fig.add_subplot(gs[0, 1])
            ax_i = fig.add_subplot(gs[1, 1])

            # ── HEADER — solo línea de paciente, dentro del panel VNO ──────────
            # (El título "Ultrasonografía Neurocrítica" se elimina: redundante)
            fig.add_artist(_mp.FancyArrowPatch(
                (0.02, 0.958), (0.99, 0.958),
                arrowstyle="-", color=C["border"],
                transform=fig.transFigure, clip_on=False, linewidth=1))

            # ══════════════════════════════════════════════════════════════════
            # VNO — todo dibujado directamente en ax_vno, sin inset cuadrado.
            # El panel es angosto y alto; compensamos el aspect en el radio de
            # las elipses para que se vean redondas, y distribuimos el contenido
            # en toda la altura (sin espacio muerto ni solapes).
            # ══════════════════════════════════════════════════════════════════
            fig.canvas.draw()
            pos = ax_vno.get_position()
            fig_w, fig_h = fig.get_size_inches()
            # Relación px/py del panel: ancho_in / alto_in (en coords de datos 0–10)
            panel_w_in = pos.width  * fig_w
            panel_h_in = pos.height * fig_h
            aspect = panel_w_in / panel_h_in   # <1 porque es más alto que ancho

            ax_vno.set_xlim(0, 10)
            ax_vno.set_ylim(0, 10)

            # Distribución vertical (coords de datos 0–10, de arriba a abajo):
            #   9.6  título
            #   9.1  paciente
            #   6.3  centro de los globos
            #   3.7  etiquetas OD / OI
            #   1.7  estado
            #   0.8  timestamp
            ax_vno.text(5.0, 9.6, "VAINA DE NERVIO ÓPTICO",
                        color=C["dim"], fontsize=14, fontweight="bold",
                        ha="center", va="top")

            # Datos del paciente: nombre en su línea (auto-reduce si es muy
            # largo) y HC + fecha en otra, para no desbordar el panel angosto.
            panel_px = pos.width * fig_w * fig.dpi
            _r = fig.canvas.get_renderer()
            t_nom = ax_vno.text(5.0, 9.12, nombre,
                                color=C["accent"], fontsize=11,
                                ha="center", va="top")
            # Reducir la fuente del nombre hasta que quepa (~94% del ancho)
            fs_nom = 11
            while fs_nom > 7:
                fig.canvas.draw()
                if t_nom.get_window_extent(_r).width <= panel_px * 0.94:
                    break
                fs_nom -= 0.5
                t_nom.set_fontsize(fs_nom)
            ax_vno.text(5.0, 8.62, f"HC: {hc}    {ts}",
                        color=C["accent"], fontsize=10, ha="center", va="top")

            def _nervio(cx, cy, color):
                # Radio en X (datos). El radio en Y se corrige por aspect para
                # que el resultado sea un círculo perfecto en píxeles.
                rx = 1.9
                ry = rx * aspect          # compensa el panel angosto (verificado 1:1)
                sw = 0.55                 # ancho del tallo (datos X)
                sh = ry * 1.05            # alto del tallo (datos Y)
                ax_vno.add_patch(_mp.FancyBboxPatch(
                    (cx - sw/2, cy - ry - sh), sw, sh,
                    boxstyle="square,pad=0",
                    facecolor=color, edgecolor="none", zorder=3))
                ax_vno.add_patch(_mp.Ellipse(
                    (cx, cy), width=rx*2, height=ry*2,
                    color=color, zorder=3))

            if self.vno_data:
                od     = self.vno_data["OD_mm"]
                oi     = self.vno_data["OI_mm"]
                prom   = self.vno_data["promedio"]
                clave  = self.vno_data.get(
                             "categoria",
                             "elevada" if self.vno_data.get("elevada") else "normal")
                vno_ts = self.vno_data.get("timestamp", "")
                _cmap  = {"normal": C["green"], "gris": C["yellow"], "elevada": C["red"]}
                col_od = _cmap[_vno_cat(od)[0]]
                col_oi = _cmap[_vno_cat(oi)[0]]

                # Globos ocupando el ancho del panel (cx=2.9 y 7.1)
                _nervio(2.9, 6.4, col_od)
                _nervio(7.1, 6.4, col_oi)
                ax_vno.text(2.9, 3.7, f"OD  {od:.1f}",
                            ha="center", color=col_od,
                            fontsize=22, fontweight="bold")
                ax_vno.text(7.1, 3.7, f"OI  {oi:.1f}",
                            ha="center", color=col_oi,
                            fontsize=22, fontweight="bold")

                _emap = {"normal":  (C["green"],  "Normal"),
                         "gris":    (C["yellow"], "Zona gris (limítrofe)"),
                         "elevada": (C["red"],    "↑ VAINA DILATADA")}
                est_col, est_txt = _emap[clave]
                ax_vno.text(5.0, 1.9, f"{est_txt}",
                            ha="center", color=est_col,
                            fontsize=15, fontweight="bold")
                ax_vno.text(5.0, 1.25, f"Promedio {prom:.2f} mm",
                            ha="center", color=C["dim"], fontsize=12)
                ax_vno.text(5.0, 0.7, vno_ts,
                            ha="center", color=C["dim"], fontsize=10)
            else:
                ax_vno.text(5, 5, "Sin medición VNO",
                            ha="center", va="center",
                            color=C["border"], fontsize=17)

            # ══════════════════════════════════════════════════════════════════
            # TCD — onda + leyenda de datos DENTRO del propio ax
            # clip_on=True en todo → nada sale del panel
            # ══════════════════════════════════════════════════════════════════
            ps_vals = [self.tcd_data[k]["PS"]
                       for k in ("D","I")
                       if self.tcd_data[k] and self.tcd_data[k].get("PS")]
            ps_ref  = max(ps_vals) * 1.18 if ps_vals else 150

            def _pic_color_str(s):
                u = s.upper() if s else ""
                if "ALTA" in u or "MUY" in u: return C["red"]
                if "BORDER" in u or "ELEV" in u: return C["yellow"]
                if "NORMAL" in u: return C["green"]
                return C["dim"]

            def _draw_tcd(ax, key, label):
                data = self.tcd_data[key]

                # Estilo del eje
                ax.set_facecolor(C["bg2"])
                for sp in ax.spines.values():
                    sp.set_color(C["border"])
                ax.tick_params(colors=C["dim"], labelsize=12, length=3)
                ax.set_xlim(0, 1)
                ax.set_ylim(0, ps_ref)

                # PIC label (esquina superior derecha del ax) — solo ACM
                es_acm = bool(data) and data.get("arteria") == "ACM"
                czos = (data.get("PIC_label","")   if es_acm else "")
                bell = (data.get("PIC_bellner","") if es_acm else "")
                pc   = _pic_color_str(czos + " " + bell)
                parts_pic = []
                if czos and czos != "—": parts_pic.append(f"Czos: {czos}")
                if bell:                  parts_pic.append(f"Bell: {bell}")
                pic_txt = ("PIC  " + "  ·  ".join(parts_pic)) if parts_pic else ""

                ppc_czos = (data.get("PPC_czosnyka_str","") if es_acm else "")
                ppc_bell = (data.get("PPC_bellner_str","")  if es_acm else "")
                parts_ppc = []
                if ppc_czos and ppc_czos != "—": parts_ppc.append(f"Czos: {ppc_czos}")
                if ppc_bell and ppc_bell != "—": parts_ppc.append(f"Bell: {ppc_bell}")
                ppc_txt = ("PPC  " + "  ·  ".join(parts_ppc)) if parts_ppc else ""

                # Título ACM y PIC en el interior del ax (no transform=fig)
                ax.text(0.01, 0.98, label,
                        transform=ax.transAxes, va="top", ha="left",
                        color=C["dim"], fontsize=14, fontweight="bold",
                        clip_on=True)
                if pic_txt:
                    ax.text(0.99, 0.98, pic_txt,
                            transform=ax.transAxes, va="top", ha="right",
                            color=pc, fontsize=13, fontweight="bold",
                            clip_on=True)
                if ppc_txt:
                    ax.text(0.99, 0.90, ppc_txt,
                            transform=ax.transAxes, va="top", ha="right",
                            color=pc, fontsize=12, fontweight="bold",
                            clip_on=True)

                if not data:
                    ax.text(0.5, 0.5, "Sin medición",
                            transform=ax.transAxes,
                            ha="center", va="center",
                            color=C["border"], fontsize=16, clip_on=True)
                    ax.set_xticks([]); ax.set_yticks([])
                    return

                # ── Onda ──────────────────────────────────────────────────────
                ps    = data.get("PS") or 0
                ed    = data.get("ED") or 0
                tamax = data.get("TAmax") or 0
                ed_f  = min(ed/ps, 0.70) if ps > 0 else 0.12
                has_n = ed_f > 0.22
                t_nad = 0.28 + ed_f * 0.25
                n_h   = ed_f*0.55+0.04 if has_n else ed_f+0.02
                s2_h  = ed_f*0.65+0.03 if has_n else ed_f+0.01
                ctrl  = [
                    (0.00,ed_f*0.85),(0.04,ed_f*0.85),(0.08,1.0),
                    (0.13,ed_f+0.18),(t_nad,n_h),
                    (t_nad+0.06,s2_h),(t_nad+0.14,ed_f+0.04),
                    (0.55,ed_f+0.01),(0.75,ed_f),(0.92,ed_f*0.9),(1.0,ed_f*0.85),
                ]
                dt    = 1.0 / 4   # 4 ciclos en xlim 0-1
                t_one = _np.linspace(0, dt, 200)
                tc    = _np.array([c[0]*dt for c in ctrl])
                vc    = _np.array([max(0,min(c[1],1.0))*ps for c in ctrl])
                v_one = _np.interp(t_one, tc, vc)
                t_all = _np.concatenate([t_one + i*dt for i in range(4)])
                v_all = _np.clip(_np.tile(v_one, 4), 0, ps*1.05)

                # Sin relleno bajo la curva, trazo fino. Con alerta, la onda
                # toma el color de la alerta, igual que en el monitor.
                pc_onda = _color_alerta(data, C["red"], C["yellow"]) or pc
                ax.plot(t_all, v_all, color=pc_onda, linewidth=1.5,
                        zorder=3, clip_on=True)
                ax.axhline(0, color=pc_onda, linewidth=0.8, alpha=0.4, clip_on=True)
                _alertas_en_ax(ax, data, C["red"], C["yellow"], printable,
                               x=0.01, y=0.88, fontsize=12)

                ax.set_xticks([])
                ax.set_ylabel("cm/s", color=C["dim"], fontsize=11)

                # ── Leyenda de datos (esquina inferior derecha, dentro del ax) ─
                pi_v  = data.get("PI")
                ri_v  = data.get("RI")
                sd_v  = data.get("SD")
                pam_v = data.get("PAM")
                pi_s  = f"{pi_v:.2f}" if pi_v is not None else "—"
                ri_s  = f"{ri_v:.2f}" if ri_v is not None else "—"
                sd_s  = f"{sd_v:.2f}" if sd_v is not None else "—"
                pam_s = (str(pam_v)+" mmHg") if pam_v is not None else "—"

                # Mismo criterio por arteria que el resto de la app.
                _zc = {"sin_dato": "dim", "bajo": "accent", "normal": "green",
                       "limite": "yellow", "alto": "red"}
                pi_col = C[_zc[_pi_zona(pi_v, data.get("arteria", "ACM"))]]

                # Caja de leyenda translúcida
                legend_x = 0.62
                legend_y = 0.02
                lbox = _mp.FancyBboxPatch(
                    (legend_x, legend_y), 0.37, 0.56,
                    boxstyle="round,pad=0.01",
                    transform=ax.transAxes,
                    facecolor=C["bg3"], edgecolor=C["border"],
                    alpha=0.92, zorder=4, clip_on=True)
                ax.add_patch(lbox)

                arteria_s = str(data.get("arteria","ACM"))
                lado_s    = str(data.get("lado",""))
                hr_s      = str(data.get("HR","—")) + " lpm" if data.get("HR") else "—"
                tam_s2    = str(tamax)+" cm/s" if tamax else "—"
                ts_s      = str(data.get("timestamp",""))

                rows = [
                    ("Arteria", f"{arteria_s} {lado_s}",  C["accent"]),
                    ("PS",      str(ps)+" cm/s",           C["dim"]),
                    ("ED",      str(ed)+" cm/s",           C["dim"]),
                    ("TAmax",   tam_s2,                    C["dim"]),
                    ("TAmean",  _fnum(data.get("TAmean"), 2)+" cm/s" if data.get("TAmean") else "—", C["dim"]),
                    ("PI",      pi_s,                      pi_col),
                    ("RI",      ri_s,                      C["dim"]),
                    ("S/D",     sd_s,                      C["dim"]),
                    ("PAM",     pam_s,                     C["dim"]),
                    ("PaCO₂",   _co2_txt(data).split(" (")[0],
                                C["yellow"] if _co2_fuera(data) else C["dim"]),
                    ("HR",      hr_s,                      C["dim"]),
                    ("Hora",    ts_s,                      C["dim"]),
                ]
                box_h     = 0.56
                lbox.set_height(box_h)
                lbox.set_y(legend_y)
                row_h = box_h / len(rows)
                for i, (lbl, val, vc) in enumerate(rows):
                    ry = legend_y + box_h - (i+0.6)*row_h
                    ax.text(legend_x + 0.01, ry, lbl,
                            transform=ax.transAxes,
                            va="center", ha="left",
                            color=C["dim"], fontsize=14, zorder=5,
                            clip_on=True)
                    ax.text(legend_x + 0.36, ry, val,
                            transform=ax.transAxes,
                            va="center", ha="right",
                            color=vc, fontsize=14,
                            fontweight="bold", zorder=5, clip_on=True)

            _draw_tcd(ax_d, "D", "ACM  DERECHA")
            _draw_tcd(ax_i, "I", "ACM  IZQUIERDA")

            # ── FOOTER ────────────────────────────────────────────────────────
            fig.text(0.5, 0.008,
                     "Estimación PIC no invasiva  ·  Czosnyka 1998 (requiere PAM+TAMV)"
                     "  ·  Bellner 2004 (requiere PI)  ·  Correlacionar con clínica",
                     ha="center", color=C["dim"], fontsize=11)

            # ── Guardar ───────────────────────────────────────────────────────
            path = filedialog.asksaveasfilename(
                parent=self.root,
                defaultextension=".png",
                filetypes=[("PNG","*.png")],
                initialfile=(f"Informe_{nombre.replace(' ','_')}_"
                             f"{datetime.now().strftime('%Y%m%d_%H%M')}.png"),
                title="Guardar informe")
            if path:
                fig.savefig(path, dpi=200, bbox_inches="tight",
                            facecolor=C["bg"])
                messagebox.showinfo("Guardado",
                                    f"Informe guardado en:\n{path}",
                                    parent=self.root)
            _cerrar_figura(fig)

        except Exception:
            messagebox.showerror("Error al generar informe",
                                 traceback.format_exc(), parent=self.root)
    def _set_info_text(self, lado_key, line1="", line2="", pi_val=None, pi_color=None):
        """Escribe en el info_text del panel TCD coloreando el valor de PI."""
        it = self._tcd_panels[lado_key]["info_text"]
        it.config(state="normal")
        it.delete("1.0", "end")
        if not line1 and not line2:
            it.insert("end", "NO MEDIDO", "dim")
            it.config(state="disabled")
            return

        def _insert_coloreando_pi(texto, base_tag):
            """Inserta `texto` con `base_tag`, coloreando 'PI X.XX' si aparece."""
            if pi_val is not None and pi_color:
                pi_str = f"PI {pi_val}"
                idx = texto.find(pi_str)
                if idx >= 0:
                    it.insert("end", texto[:idx],                   base_tag)
                    it.insert("end", pi_str,                        pi_color)
                    it.insert("end", texto[idx + len(pi_str):],     base_tag)
                    return
            it.insert("end", texto, base_tag)

        # Línea 1 (velocidades) en "normal"; el PI puede estar aquí o en línea 2
        if line1:
            _insert_coloreando_pi(line1, "normal")
        # Línea 2 (índices, donde vive PI) en "dim", coloreando PI si está
        if line2:
            it.insert("end", "\n")
            _insert_coloreando_pi(line2, "dim")

        it.tag_add("center", "1.0", "end")
        it.config(state="disabled")

    def _on_tcd_result(self, lado_key, data):
        if isinstance(data, dict) and data.get("Estenosis_texto"):
            data = dict(data)
            data["Estenosis_texto"] = _limpiar_cita_guia_estenosis(data.get("Estenosis_texto"))
        self.tcd_data[lado_key] = data
        p       = self._tcd_panels[lado_key]
        es_acm  = data.get("arteria", "ACM") == "ACM"
        pic     = data.get("PIC_label",   "—") if es_acm else ""
        bell    = data.get("PIC_bellner", "")  if es_acm else ""
        pc      = _pic_color(pic)
        # .get() en todo: un .ncu guardado por una versión anterior puede no
        # traer todas las claves, y un KeyError aquí aborta _abrir_caso entero.
        pi_v0 = data.get("PI")
        ri_v0 = data.get("RI")
        pi_s  = f"{pi_v0:.2f}" if pi_v0 is not None else "—"
        ri_s  = f"{ri_v0:.2f}" if ri_v0 is not None else "—"
        tam_s = f"{data.get('TAmax')}" if data.get("TAmax") else "—"

        if not es_acm:
            # Czosnyka/Bellner solo son válidos en ACM: se oculta el bloque
            # en vez de mostrar "No aplicable" en otras arterias.
            p["pic_lbl"].config(text="", fg=TEXT2)
        else:
            czos_txt = f"Czos: {pic}"  if pic  and pic  != "—" else ""
            bell_txt = f"Bell: {bell}" if bell and bell != "—" else ""
            parts    = [t for t in [czos_txt, bell_txt] if t]
            inner    = "  ·  ".join(parts) if parts else "sin datos"

            ppc_czos = data.get("PPC_czosnyka_str", "")
            ppc_bell = data.get("PPC_bellner_str",  "")
            ppc_czos_txt = f"Czos: {ppc_czos}" if ppc_czos else ""
            ppc_bell_txt = f"Bell: {ppc_bell}" if ppc_bell else ""
            ppc_parts = [t for t in [ppc_czos_txt, ppc_bell_txt] if t]
            ppc_inner = "  ·  ".join(ppc_parts) if ppc_parts else "sin datos"

            pic_txt  = f"PIC  {inner}\nPPC  {ppc_inner}"

            pc2      = _pic_color(bell) if bell else GRAY
            pc_final = RED if RED in (pc, pc2) else (YELLOW if YELLOW in (pc, pc2) else pc)
            p["pic_lbl"].config(text=pic_txt, fg=pc_final, justify="right")

        # Línea 1: velocidades
        sd_s2  = f"{data['SD']:.2f}"   if data.get("SD")     is not None else "—"
        tam2_s = _fnum(data.get("TAmean"), 2) if data.get("TAmean") else "—"
        art_s  = f"{data.get('arteria','ACM')} {data.get('lado','')}"
        hr_s   = f"{data.get('HR','—')} lpm" if data.get('HR') else "—"
        l1 = (f"{art_s}  ·  PS {data.get('PS', '—')}  ED {data.get('ED', '—')}  "
              f"TAmax {tam_s}  TAmean {tam2_s}  cm/s")
        # Línea 2: índices + PAM + hora
        l2 = (f"PI {pi_s}  RI {ri_s}  S/D {sd_s2}  "
              f"PAM {_pam_txt(data)}  CO₂ {_co2_txt(data)}  HR {hr_s}  ·  "
              f"{data.get('timestamp', '')}")

        # Color del PI: MISMO criterio por arteria que el resto de la app
        # (antes aquí eran 1.0 / 1.4 fijos, distintos del panel de otras
        # arterias y de la tabla de índices de la propia figura).
        pi_color = _pi_tag(pi_v0, data.get("arteria", "ACM"))

        self._set_info_text(lado_key, l1, l2, pi_val=pi_s, pi_color=pi_color)
        self._redraw_tcd(lado_key)
        al = _alertas_tcd(data)
        lado_txt = 'Derecha' if lado_key == 'D' else 'Izquierda'
        if al:
            rojo = any(n == "rojo" for n, _ in al)
            self._set_status(
                f"⚠ ACM {lado_txt}: " + "  ·  ".join(t for _, t in al),
                RED if rojo else YELLOW)
        else:
            self._set_status(f"TCD ACM {lado_txt} — {data.get('timestamp', '')}")


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════
def main():
    _configurar_log()
    _log.info("Inicio build %s (frozen=%s)", _BUILD_ID, getattr(sys, "frozen", False))
    root = tk.Tk()

    def _error_callback(tipo, valor, tb):
        _log.error("Error en la interfaz:\n%s",
                   "".join(traceback.format_exception(tipo, valor, tb)))
        try:
            messagebox.showerror(
                "Error inesperado",
                str(valor)[:300] + "\n\nDetalle guardado en:\n" +
                _ruta_config("errores.log"), parent=root)
        except Exception:
            pass
    root.report_callback_exception = _error_callback

    app = NeurocriticalApp(root)
    if os.environ.get("NCU_SMOKE_TEST") == "1":
        # Usado por el compilador para verificar que el .exe abre Y CIERRA.
        root.after(4000, app._salir)
    try:
        root.mainloop()
    finally:
        _apagado_final()
        _log.info("Cierre normal")
        logging.shutdown()
    if getattr(sys, "frozen", False):
        # Garantía final en el .exe: ya se liberó todo de forma ordenada;
        # os._exit evita que un hilo nativo de terceros (SDL, asyncio de
        # edge-tts, cliente HTTP) retenga el proceso durante el apagado del
        # intérprete. En desarrollo se sale normal para no ocultar errores.
        os._exit(0)


if __name__ == "__main__":
    main()
