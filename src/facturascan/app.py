import sys, os, ctypes, threading, winreg
import customtkinter as ctk
from tkinter import messagebox

from utils.log_utils import (
    set_debug, is_debug, registrar_log,
    registrar_link_documento, obtener_link_documento
)
from core.monitor_core import aplicar_nueva_config
from gui.apariencia_gui import cargar_tamano_log, guardar_tamano_log, abrir_modal_apariencia

import re

# === assets e icono ===
# getattr = pregunta si el atributo sys.frozen existe y es verdadero, si el programa es un .exe estara en True si es Python normal .py estara en False
if getattr(sys, "frozen", False):  
    # BASE_DIR = ajusta la ubicación de la carpeta del proyecto al lado del .exe por ejemplo c:/FacturaScan
    BASE_DIR = os.path.dirname(sys.executable)
else: 
    # si se ejecuta en python .py buscara el archivo en la ruta absoluta donde esta el .py
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Carpeta assets
ASSETS_DIR = os.path.join(BASE_DIR, "assets")

# Funcion auxiliar para no estar armando rutas cada vez
def asset_path(nombre: str) -> str:
    """Devuelve la ruta absoluta dentro de /assets."""
    return os.path.join(ASSETS_DIR, nombre)

# Fijar AppUserModelID Para anclar el programa a la barra de tareas
def anclar_programa(app_id: str) -> None:
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)  
    except Exception:
        pass

anclar_programa("FacturaScan.App")

# Fijar icono sin parpadeo con API nativa (WM_SETICON)
def _get_icon_ico_path() -> str:
    # Siempre desde assets al lado del exe / o del .py
    p = os.path.join(BASE_DIR, "assets", "iconoScan.ico")
    return p

def aplicar_icono(win):
    ico_path = _get_icon_ico_path()
    try:
        if os.path.exists(ico_path):
            win.iconbitmap(ico_path)
        else:
            registrar_log(f"Icono no encontrado en: {ico_path}")
    except Exception as e:
        print("No se pudo aplicar ícono:", e)
        try:
            registrar_log(f"No se pudo aplicar ícono: {e}")
        except Exception:
            pass

def abrir_panel_debug(root):
    from debug import debugapp as dbg

    win = ctk.CTkToplevel(root)
    win.title("Panel Debug OCR")
    aplicar_icono(win)
    win.after(200, lambda: aplicar_icono(win))

    ancho, alto = 300, 200
    x = (win.winfo_screenwidth() - ancho) // 2
    y = (win.winfo_screenheight() - alto) // 2
    win.geometry(f"{ancho}x{alto}+{x}+{y}")
    win.resizable(False, False)

    win.transient(root)
    win.grab_set()

    # Variables UI (solo OCR)
    var_rut = ctk.BooleanVar(value=bool(dbg.DEBUG.mostrar_ocr_rut))
    var_fact = ctk.BooleanVar(value=bool(dbg.DEBUG.mostrar_ocr_factura))

    titulo = ctk.CTkLabel(
        win, text="Opciones OCR (Debug)",
        font=ctk.CTkFont(size=16, weight="bold")
    )
    titulo.pack(pady=(16, 10))

    frame = ctk.CTkFrame(win)
    frame.pack(fill="both", expand=True, padx=16, pady=10)

    def aplicar_estado():
        # Solo flags OCR (independientes del debug general)
        dbg.set_debug_flags(
            mostrar_ocr_rut=var_rut.get(),
            mostrar_ocr_factura=var_fact.get()
        )

    sw_rut = ctk.CTkSwitch(frame, text="Ver OCR RUT", variable=var_rut, command=aplicar_estado)
    sw_rut.pack(anchor="w", padx=14, pady=(14, 8))

    sw_fact = ctk.CTkSwitch(frame, text="Ver OCR Factura", variable=var_fact, command=aplicar_estado)
    sw_fact.pack(anchor="w", padx=14, pady=8)

    btn_frame = ctk.CTkFrame(win, fg_color="transparent")
    btn_frame.pack(fill="x", padx=16, pady=(0, 12))
    ctk.CTkButton(btn_frame, text="Cerrar", command=win.destroy).pack(side="right")

    # Aplica estado inicial (por si el módulo cambió)
    aplicar_estado()

# funcipon para mostrar mensajes al inicio ya sea por import críticos o fallas de módulos.
def show_startup_error(msg: str):
    try:
        import tkinter as tk
        from tkinter import messagebox as mb
        r = tk.Tk(); r.withdraw()
        try:
            aplicar_icono(r)
        except Exception:
            pass
        mb.showerror("Error al iniciar FacturaScan", msg)
        r.destroy()
    except Exception:
        print(msg)

import tempfile, msvcrt

_lock_file_handle = None
def instanciaUnica():
    global _lock_file_handle
    lock_path = os.path.join(tempfile.gettempdir(), "FacturaScan.lock")
    try:
        _lock_file_handle = os.open(lock_path, os.O_CREAT | os.O_RDWR)
        # intenta bloquear el archivo
        msvcrt.locking(_lock_file_handle, msvcrt.LK_NBLCK, 1)
    except OSError:
        show_startup_error("FacturaScan ya está en ejecución.")
        sys.exit(0)

# Solo 1 instancia en ejecución
instanciaUnica()

# Imports críticos
try:
    from gui.config_gui import cargar_o_configurar, actualizar_rutas, seleccionar_razon_sucursal_grid
    from core.monitor_core import  procesar_archivo, procesar_entrada_una_vez
except Exception as e:
    show_startup_error(f"No se pudo importar un módulo crítico:\n\n{e}")
    sys.exit(1)

import importlib.util
faltan = []
for _mod in ["PIL", "pdf2image", "easyocr", "win32com.client", "reportlab"]:
    if importlib.util.find_spec(_mod) is None:
        faltan.append(f"- {_mod}: no encontrado")

if faltan:
    show_startup_error("Módulos opcionales no disponibles:\n\n" + "\n".join(faltan))

# Helper para terminar con mensaje
def fatal(origen: str, e: Exception):
    show_startup_error(f"{origen}:\n\n{e}")
    sys.exit(1)

def _cancelar_todos_after(win):
    """Cancela todos los callbacks 'after' pendientes del root."""
    try:
        ids = win.tk.eval('after info').split()
        for cb in ids:
            try:
                win.after_cancel(cb)
            except Exception:
                pass
    except Exception:
        pass

# CONFIGURACIÓN INICIAL
variables = None
try:
    variables = cargar_o_configurar()
except Exception as e:
    fatal("CONFIG", e)

if variables is None:
    fatal("CONFIG", Exception("No se obtuvo configuración"))

aplicar_nueva_config(variables)

from inicial import __version__, MOSTRAR_BT_CAMBIAR_SUCURSAL_OF, ACTUALIZAR_PROGRAMA
VERSION = __version__

# ================== UTILIDADES ==================
def Valida_PopplerPath():
    ruta_poppler = r"C:\poppler\Library\bin"
    ruta_normalizada = os.path.normcase(os.path.normpath(ruta_poppler))

    # 1) Si la carpeta no existe, solo lo anotamos en el log y no tocamos nada
    if not os.path.isdir(ruta_poppler):
        try:
            registrar_log(f"Poppler no encontrado en {ruta_poppler}.")
        except Exception:
            pass
        return

    # 2) Leer PATH del usuario (NO el de sistema)
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment", 0, winreg.KEY_READ) as clave:
            valor_actual, _ = winreg.QueryValueEx(clave, "Path")
    except FileNotFoundError:
        valor_actual = ""
    except Exception as e:
        # Si algo raro pasa leyendo el registro, seguimos solo con el PATH del proceso
        try:
            registrar_log(f"No se pudo leer PATH de usuario: {e}")
        except Exception:
            pass
        valor_actual = ""

    segmentos_norm = [
        os.path.normcase(os.path.normpath(p.strip()))
        for p in valor_actual.split(";")
        if p.strip()
    ]

    ya_en_path_usuario = ruta_normalizada in segmentos_norm

    # 3) Si no está en PATH de usuario, lo agregamos silenciosamente
    if not ya_en_path_usuario:
        nuevo_valor = f"{valor_actual};{ruta_poppler}" if valor_actual else ruta_poppler
        try:
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Environment",
                0,
                winreg.KEY_SET_VALUE
            ) as clave:
                winreg.SetValueEx(clave, "Path", 0, winreg.REG_EXPAND_SZ, nuevo_valor)
            try:
                registrar_log(f"Ruta Poppler añadida al PATH del usuario: {ruta_poppler}")
            except Exception:
                pass
        except PermissionError:
            # Sin permisos para tocar el registro → seguimos igual, pero lo dejamos en el log
            try:
                registrar_log("Sin permisos para actualizar PATH del usuario. Se continuará sin modificar el registro.")
            except Exception:
                pass
        except Exception as e:
            try:
                registrar_log(f"Error al actualizar PATH del usuario: {e}")
            except Exception:
                pass

    # 4) Asegurar que el proceso actual también vea la ruta, aunque el registro falle
    try:
        path_proceso = os.environ.get("PATH", "")
        if ruta_poppler not in path_proceso:
            os.environ["PATH"] = path_proceso + (";" if path_proceso else "") + ruta_poppler
    except Exception:
        pass

# Al intentar cerrar FacturaScan mostrará un mensaje de confirmación
def cerrar_aplicacion(ventana, modales_abiertos=None):
    # Si hay un modal abierto, no cerrar aún
    if modales_abiertos and (modales_abiertos.get("config") or modales_abiertos.get("rutas")or modales_abiertos.get("sucursal")):
        messagebox.showwarning("Ventana abierta", "Cierra primero la ventana de configuración.")
        return

    if not messagebox.askyesno("Cerrar", "¿Deseas cerrar FacturaScan?"):
        return
    try:
        registrar_log('🔴 FacturaScan cerrado por el usuario')
    except Exception:
        pass

    try:
        # Evita reentradas visuales
        try:
            ventana.withdraw()
        except Exception:
            pass

        # Libera posibles grabs de modales
        try:
            ventana.grab_release()
        except Exception:
            pass

        # Cancela todos los after (como el actualizador del textbox)
        _cancelar_todos_after(ventana)

        # Cierra limpio el loop y el root
        try:
            ventana.quit()
        except Exception:
            pass
        try:
            ventana.destroy()
        except Exception:
            pass
    finally:
        os._exit(0)

# Apariencia botones:
CUSTOMCOLOR = "#a6a6a6"
# hober anterior #8c8c8c
HOBERCOLOR = "#00C9B1"
TEXCOLOR = "#000000"
BORDERCOLOR = "#000000"


# ================== INTERFAZ PRINCIPAL ==================
def menu_Principal():
    from PIL import Image
    from datetime import datetime
    from core.scanner import escanear_y_guardar_pdf, seleccionar_scanner_predeterminado

    registrar_log("🟢 FacturaScan iniciado correctamente")

    en_proceso = {"activo": False}
    modales_abiertos = {"config": False, "rutas": False, "sucursal": False}

    # ===== Apariencia y ventana =====
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    ventana = ctk.CTk()
    ventana.title(f"Control documental {VERSION}")
    aplicar_icono(ventana)
    ventana.after(150, lambda: aplicar_icono(ventana))

    # Actualizaciones por Github (todo vive en updater.py)
    from update.updater import schedule_update_prompt
    if ACTUALIZAR_PROGRAMA:
        schedule_update_prompt(ventana, current_version=VERSION, apply_icono_fn=aplicar_icono)
        
    try:
        from ocr.ocr_utils import warmup_ocr
        ventana.after(200, lambda: threading.Thread(target=warmup_ocr, daemon=True).start())
    except Exception:
        pass

    # Centro y tamaño
    ancho, alto = 720, 600
    x = (ventana.winfo_screenwidth() - ancho) // 2
    y = (ventana.winfo_screenheight() - alto) // 2
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")
    ventana.resizable(False, False)

    fuente_titulo = ctk.CTkFont(size=40, weight="bold")
    fuente_texto = ctk.CTkFont(family="Segoe UI", size=15)

    # ===== Barra superior: Izquierda (Sucursal) | Centro (Título) | Derecha (Buscar) =====
    topbar = ctk.CTkFrame(ventana, fg_color="transparent")
    topbar.pack(fill="x", padx=12, pady=(10, 0))

    SIDE_W = 170  # ancho fijo de los laterales (igual a ambos lados)

    # 3 columnas: izquierda fija | centro expandible | derecha fija
    topbar.grid_columnconfigure(0, minsize=SIDE_W, weight=0)
    topbar.grid_columnconfigure(1, weight=1)
    topbar.grid_columnconfigure(2, minsize=SIDE_W, weight=0)

    # Fila con altura mínima para alinear verticalmente botones vs título
    topbar.grid_rowconfigure(0, minsize=56)

    ctk.CTkFrame(topbar, width=SIDE_W, height=32, fg_color="transparent").grid(row=0, column=0, sticky="w")

    # --- Centro (título centrado REAL) ---
    lbl_titulo = ctk.CTkLabel(
        topbar,
        text="FacturaScan",
        font=fuente_titulo,
        anchor="center"
    )
    lbl_titulo.grid(row=0, column=1, sticky="ew")

    # --- Derecha: contenedor para botones (Buscar + Debug) ---
    rightbar = ctk.CTkFrame(topbar, fg_color="transparent")
    rightbar.grid(row=0, column=2, sticky="e")

    btn_historial = ctk.CTkButton(
        rightbar, text="Buscar",
        width=SIDE_W, height=32, corner_radius=16,
        fg_color=CUSTOMCOLOR, text_color=TEXCOLOR,
        hover_color=HOBERCOLOR,
        border_color=BORDERCOLOR, border_width=1,
        command=lambda: _abrir_ventana_historial()
    )
    btn_historial.pack(side="right")

    frame_botones = ctk.CTkFrame(ventana, fg_color="transparent")
    frame_botones.pack(pady=10)

    # Iconos desde assets
    try:
        icono_escaneo = ctk.CTkImage(
            light_image=Image.open(asset_path("icono_escanear.png")),
            size=(26, 26)
        )
    except Exception:
        registrar_log("no se encontro imagen icono_escanear.png")
        icono_escaneo = None

    try:
        icono_cambiar_scanner = ctk.CTkImage(
            light_image=Image.open(asset_path("icono_cambiar_scanner.png")),
            size=(22, 22)
        )
    except Exception:
        registrar_log("no se encontro imagen icono_cambiar_scanner.png")
        icono_cambiar_scanner = None

    try:
        icono_carpeta = ctk.CTkImage(
            light_image=Image.open(asset_path("icono_carpeta.png")),
            size=(26, 26)
        )
    except Exception:
        registrar_log("no se encontro imagen icono_carpeta.png")
        icono_carpeta = None

    # ===== Textbox de log y mensaje inferior =====
    # Fuente del log (ajustable desde "Ajustes")
    size_inicial = cargar_tamano_log(BASE_DIR, default=12, log_fn=registrar_log)

    log_font_state = {"size": size_inicial}
    font_log = ctk.CTkFont(family="Consolas", size=log_font_state["size"])

    # ===== Contenedor del log + botón limpiar =====
    frame_log = ctk.CTkFrame(ventana, fg_color="transparent")
    frame_log.pack(pady=15, padx=15, fill="x")

    frame_log.grid_columnconfigure(0, weight=1)

    btn_limpiar_log = ctk.CTkButton(
        frame_log,
        text="🧹",
        width=35,
        height=35,
        fg_color=CUSTOMCOLOR,
        text_color=TEXCOLOR,
        corner_radius=6,
        # hover_color="#8c8c8c",
        hover_color=HOBERCOLOR,
        border_color=BORDERCOLOR, border_width=1,

        command=lambda: limpiar_log_desde_opcion()
    )
    btn_limpiar_log.grid(row=0, column=1, sticky="e", pady=(0, 6))

    texto_log = ctk.CTkTextbox(
        frame_log, width=680, height=290,
        font=font_log, wrap="word",
        corner_radius=6, fg_color="white", text_color="black"
    )
    texto_log.grid(row=1, column=0, columnspan=2, sticky="nsew")

    def abrir_documento_desde_log(event):
        """
        Abre el PDF asociado a la línea donde se hizo doble click.
        La línea debe terminar con el nombre del archivo .pdf,
        por ejemplo: '1/1 ✅ Procesado: Lo Blanco_93178000-K_factura_19211105_2025_3.pdf'
        """
        try:
            # Coordenadas del click dentro del textbox
            idx = texto_log.index(f"@{event.x},{event.y}")
            linea_num = int(idx.split(".")[0])

            # Contenido completo de esa línea
            linea = texto_log.get(f"{linea_num}.0", f"{linea_num}.end")

            # Buscamos la parte después de los dos puntos (:)
            if ":" not in linea:
                return
            nombre_archivo = linea.split(":", 1)[1].strip()
            if not nombre_archivo.lower().endswith(".pdf"):
                return

            # Preguntar a log_utils por la ruta real
            ruta = obtener_link_documento(nombre_archivo)
            if not ruta:
                print(f"⚠️ No se encontró ruta para: {nombre_archivo}")
                return

            if not os.path.exists(ruta):
                print(f"⚠️ El archivo ya no existe: {ruta}")
                return

            os.startfile(ruta)
        except Exception as e:
            print(f"❗ Error al abrir documento desde log: {e}")

    # Doble click con el botón izquierdo
    texto_log.bind("<Double-Button-1>", abrir_documento_desde_log)

    mensaje_espera = ctk.CTkLabel(ventana, text="", font=fuente_texto, text_color="gray")
    mensaje_espera.pack(pady=(0, 10))

    # ===== HISTORIAL / BUSCAR DOCUMENTOS =====
    def _abrir_ventana_historial():
        """
        Muestra el panel de historial sin cargar documentos al abrir.
        Solo carga (y filtra) cuando el usuario presiona 'Buscar'.
        Incluye spinner/estado de carga para no dar sensación de cuelgue.
        """
        from calendar import monthrange
        from datetime import datetime as _dt_now
        import threading

        # Ocultar ventana principal
        ventana.withdraw()
        hist = ctk.CTkToplevel(ventana)
        hist.withdraw() 
        hist.title("Historial de documentos")

        # (opcional) setear geometry antes de mostrar
        ancho, alto = 1000, 600
        x = (ventana.winfo_screenwidth() - ancho) // 2
        y = (ventana.winfo_screenheight() - alto) // 2
        hist.geometry(f"{ancho}x{alto}+{x}+{y}")
        hist.resizable(True, True)

        # Fuerza a que Tk cree los handles internos
        hist.update_idletasks()

        hist.after(200, lambda: aplicar_icono(hist))

        hist.deiconify() 

        fuente_titulo_hist = ctk.CTkFont(size=20, weight="bold")
        fuente_filtro = ctk.CTkFont(size=13)
        fuente_row = ctk.CTkFont(family="Consolas", size=12)

        ui_alive = {"value": True}
        searching = {"value": False}
        last_search_token = {"value": 0}

        def _cerrar_historial():
            ui_alive["value"] = False
            try:
                hist.destroy()
            except Exception:
                pass
            ventana.deiconify()
            ventana.lift()
            ventana.focus_force()

        hist.protocol("WM_DELETE_WINDOW", _cerrar_historial)

        # ---- Encabezado ----
        top = ctk.CTkFrame(hist, fg_color="transparent")
        top.pack(fill="x", padx=16, pady=(10, 4))

        ctk.CTkLabel(
            top,
            text="Historial de documentos (Procesados)",
            font=fuente_titulo_hist
        ).pack(anchor="w")

        def _limpiar_filtros():
            var_anio.set("Todos")
            var_mes.set("Todos")
            var_dia.set("Todos")
            var_tipo_doc.set("Todos")
            var_buscar.set("")
            _actualizar_dias()
            # Si estaba cargando/buscando, oculta el spinner
            _set_loading(False)

            # Limpia resultados y muestra placeholder
            for w in cont_lista.winfo_children():
                w.destroy()

            ctk.CTkLabel(
                cont_lista,
                text="✅ Filtros limpiados. Ajusta los filtros y presiona “Buscar” para listar documentos.",
                text_color="#16a34a"  # verde suave
            ).pack(pady=24)

            # (opcional) enfocar el campo de búsqueda para seguir rápido
            try:
                entry_buscar.focus_set()
            except Exception:
                pass

        # ---- Filtros ----
        filtros = ctk.CTkFrame(hist, fg_color="transparent")
        filtros.pack(fill="x", padx=16, pady=(0, 6))

        current_year = _dt_now.now().year
        anios_disp = [current_year - i for i in range(5)]
        meses_disp = list(range(1, 13))

        var_anio = ctk.StringVar(value="Todos")
        var_mes = ctk.StringVar(value="Todos")
        var_dia = ctk.StringVar(value="Todos")
        var_buscar = ctk.StringVar(value="")
        var_tipo_doc = ctk.StringVar(value="Todos")

        # Fila 1: combos MES / DÍA / AÑO
        fila1 = ctk.CTkFrame(filtros, fg_color="transparent")
        fila1.pack(fill="x", pady=(4, 2))

        ctk.CTkLabel(fila1, text="Mes:", font=fuente_filtro).pack(side="left", padx=(0, 4))
        cb_mes = ctk.CTkComboBox(
            fila1,
            variable=var_mes,
            values=["Todos"] + [str(m) for m in meses_disp],
            width=90
        )
        cb_mes.pack(side="left", padx=(0, 16))

        ctk.CTkLabel(fila1, text="Día:", font=fuente_filtro).pack(side="left", padx=(0, 4))
        cb_dia = ctk.CTkComboBox(
            fila1,
            variable=var_dia,
            values=["Todos"],
            width=90
        )
        cb_dia.pack(side="left", padx=(0, 16))

        ctk.CTkLabel(fila1, text="Año:", font=fuente_filtro).pack(side="left", padx=(0, 4))
        cb_anio = ctk.CTkComboBox(
            fila1,
            variable=var_anio,
            values=["Todos"] + [str(a) for a in anios_disp],
            width=90
        )
        cb_anio.pack(side="left", padx=(0, 16))

        # Fila 2: buscador + tipo + Buscar + Cerrar
        fila2 = ctk.CTkFrame(filtros, fg_color="transparent")
        fila2.pack(fill="x", pady=(4, 8))

        ctk.CTkLabel(
            fila2,
            text="Buscar (RUT o N° factura):",
            font=fuente_filtro
        ).pack(side="left", padx=(0, 6))

        entry_buscar = ctk.CTkEntry(
            fila2,
            textvariable=var_buscar,
            width=220,
            font=fuente_filtro
        )
        entry_buscar.pack(side="left", padx=(0, 8))

        ctk.CTkLabel(fila2, text="Tipo:", font=fuente_filtro).pack(side="left", padx=(10, 4))
        cb_tipo_doc = ctk.CTkComboBox(
            fila2,
            variable=var_tipo_doc,
            values=["Todos", "Factura", "Guía de despacho", "CHEP", "Otros"],
            width=160
        )
        cb_tipo_doc.pack(side="left", padx=(0, 8))

        # Botón Buscar
        btn_buscar = ctk.CTkButton(
            fila2, text="Buscar", width=110, height=30,
            fg_color=CUSTOMCOLOR, 
            # hover_color="#8c8c8c",
            hover_color=HOBERCOLOR,
            text_color=TEXCOLOR,
            border_color=BORDERCOLOR, border_width=1,
            command=lambda: _iniciar_busqueda()
        )
        btn_buscar.pack(side="right")

        # Botón Limpiar
        btn_limpiar = ctk.CTkButton(
            fila2, text="Limpiar", width=90, height=30,
            fg_color=CUSTOMCOLOR, 
            hover_color=HOBERCOLOR,
            text_color=TEXCOLOR,
            border_color=BORDERCOLOR, border_width=1,
            command=_limpiar_filtros
        )
        btn_limpiar.pack(side="right", padx=(8, 8))

        # ---- Zona de lista (scrollable) ----
        cont_lista = ctk.CTkScrollableFrame(hist, fg_color="#f9fafb")
        cont_lista.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        # Placeholder inicial
        placeholder = ctk.CTkLabel(
            cont_lista,
            text="Ajusta los filtros y presiona “Buscar” para listar documentos.",
            text_color=HOBERCOLOR
        )
        placeholder.pack(pady=24)

        # ---- Overlay de carga (spinner) ----
        loading = ctk.CTkFrame(hist, fg_color="white", corner_radius=12)
        loading_label = ctk.CTkLabel(loading, text="🔎 Buscando documentos…", font=ctk.CTkFont(size=14, weight="bold"))
        loading_label.pack(padx=18, pady=(14, 8))
        pb = ctk.CTkProgressBar(loading, mode="indeterminate", width=240)
        pb.pack(padx=18, pady=(0, 14))
        loading.place_forget()

        def _set_loading(activo: bool, texto: str = "🔎 Buscando documentos…"):
            if not ui_alive["value"]:
                return
            try:
                loading_label.configure(text=texto)
            except Exception:
                pass

            if activo:
                try:
                    loading.place(relx=0.5, rely=0.5, anchor="center")
                    pb.start()
                except Exception:
                    pass
                # deshabilitar controles mientras busca
                for w in (cb_mes, cb_dia, cb_anio, entry_buscar, cb_tipo_doc, btn_buscar):
                    try:
                        w.configure(state="disabled")
                    except Exception:
                        pass
            else:
                try:
                    pb.stop()
                    loading.place_forget()
                except Exception:
                    pass
                for w in (cb_mes, cb_dia, cb_anio, entry_buscar, cb_tipo_doc, btn_buscar):
                    try:
                        w.configure(state="normal")
                    except Exception:
                        pass

        # --- Actualizar días según mes/año ---
        def _actualizar_dias():
            sel_mes = var_mes.get()
            sel_anio = var_anio.get()
            try:
                if sel_mes != "Todos":
                    m = int(sel_mes)
                    y = int(sel_anio) if sel_anio != "Todos" else current_year
                    _, ultimo_dia = monthrange(y, m)
                    dias_vals = [str(d) for d in range(1, ultimo_dia + 1)]
                else:
                    dias_vals = [str(d) for d in range(1, 32)]
            except Exception:
                dias_vals = [str(d) for d in range(1, 32)]

            valores = ["Todos"] + dias_vals
            cb_dia.configure(values=valores)

            if var_dia.get() not in valores:
                var_dia.set("Todos")

        _actualizar_dias()

        # Solo actualizar días al cambiar mes/año
        def _on_cambio_mes_anio(*_):
            _actualizar_dias()

        var_mes.trace_add("write", _on_cambio_mes_anio)
        var_anio.trace_add("write", _on_cambio_mes_anio)

        def _abrir_pdf(ruta):
            try:
                if os.path.exists(ruta):
                    os.startfile(ruta)
                else:
                    messagebox.showwarning("Archivo no encontrado", f"El archivo ya no existe:\n{ruta}")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")

        # Mapa año -> info {frame, visible}
        secciones_ano = {}

        def _pintar_por_ano(lista):
            # Limpia
            for w in cont_lista.winfo_children():
                w.destroy()

            if not lista:
                ctk.CTkLabel(
                    cont_lista,
                    text="Sin resultados para los filtros ingresados.",
                    text_color="#6b7280"
                ).pack(pady=20)
                return

            # --- Agrupar: año -> mes -> registros ---
            from datetime import datetime as _dt_min
            agrup = {}  # {anio: {mes: [regs]}}

            for r in lista:
                anio = r.get("anio")
                mes = r.get("mes")
                if not anio or not mes:
                    continue
                agrup.setdefault(anio, {}).setdefault(mes, []).append(r)

            # Orden: años desc, meses desc
            anios_ordenados = sorted(agrup.keys(), reverse=True)

            # Estado UI: año abierto/cerrado, mes abierto/cerrado
            ui_state = {
                "anio": {},   # {anio: {"visible": bool, "frame": frame, "header": header}}
                "mes": {}     # {(anio, mes): {"visible": bool, "frame": frame, "header": header}}
            }

            def _nombre_mes(m):
                nombres = {
                    1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
                    7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"
                }
                return nombres.get(m, str(m))

            def _toggle_mes(anio, mes):
                key = (anio, mes)
                info = ui_state["mes"].get(key)
                if not info:
                    return

                if info["visible"]:
                    info["frame"].pack_forget()
                    info["visible"] = False
                    return

                # Si se va a abrir: crear/repintar lista de docs del mes
                frame_mes = info["frame"]

                # Limpiar contenido anterior para evitar duplicados
                for w in frame_mes.winfo_children():
                    w.destroy()

                regs = sorted(
                    agrup[anio][mes],
                    key=lambda x: (x["fecha"] or _dt_min.min, x.get("archivo") or ""),
                    reverse=True
                )

                # Pintar documentos del mes (solo aquí, lazy)
                for r in regs:
                    fecha_txt = r["fecha"].strftime("%Y-%m-%d") if r.get("fecha") else "-"
                    rut_txt = r.get("rut") or "-"
                    num_txt = r.get("numero") or "-"
                    nom_txt = r.get("archivo") or "-"

                    texto_row = (
                        f"{fecha_txt}   |   RUT: {rut_txt:>11}   |   "
                        f"Número: {num_txt:>10}   |   {nom_txt}"
                    )

                    btn_row = ctk.CTkButton(
                        frame_mes,
                        text=texto_row,
                        anchor="w",
                        height=28,
                        fg_color="#f3f4f6",
                        hover_color="#e5e7eb",
                        text_color="#111827",
                        font=fuente_row,
                        command=lambda p=r["ruta"]: _abrir_pdf(p)
                    )
                    btn_row.pack(fill="x", padx=8, pady=2)

                # Mostrar el frame justo debajo del header de mes
                frame_mes.pack(fill="x", padx=26, pady=(0, 6), after=info["header"])
                info["visible"] = True

            def _toggle_anio(anio):
                info = ui_state["anio"].get(anio)
                if not info:
                    return

                if info["visible"]:
                    # Colapsa todo el año (y sus meses si estaban)
                    info["frame"].pack_forget()
                    info["visible"] = False
                    return

                # Expandir: pintar SOLO los headers de meses (no documentos)
                frame_anio = info["frame"]
                for w in frame_anio.winfo_children():
                    w.destroy()

                meses = sorted(agrup[anio].keys(), reverse=True)

                for mes in meses:
                    header_mes = ctk.CTkButton(
                        frame_anio,
                        text=f"▸ {_nombre_mes(mes)}",
                        anchor="w",
                        height=30,
                        fg_color="#eef2ff",
                        hover_color="#e0e7ff",
                        text_color="#111827",
                        font=ctk.CTkFont(size=13, weight="bold")
                    )
                    header_mes.pack(fill="x", padx=8, pady=(6, 0))

                    frame_mes = ctk.CTkFrame(frame_anio, fg_color="white", corner_radius=8)
                    # no lo packeamos hasta que se abra

                    ui_state["mes"][(anio, mes)] = {
                        "visible": False,
                        "frame": frame_mes,
                        "header": header_mes
                    }

                    header_mes.configure(command=lambda a=anio, m=mes: _toggle_mes(a, m))

                # Mostrar el frame del año debajo del header del año
                frame_anio.pack(fill="x", padx=16, pady=(0, 6), after=info["header"])
                info["visible"] = True

            # Pintar AÑOS (solo headers)
            for anio in anios_ordenados:
                header_anio = ctk.CTkButton(
                    cont_lista,
                    text=f"Año {anio}",
                    anchor="w",
                    height=34,
                    fg_color="#e5e7eb",
                    hover_color="#d1d5db",
                    text_color="#111827",
                    font=ctk.CTkFont(size=14, weight="bold")
                )
                header_anio.pack(fill="x", pady=(6, 0))

                frame_anio = ctk.CTkFrame(cont_lista, fg_color="white", corner_radius=8)
                # no se muestra hasta abrir

                ui_state["anio"][anio] = {
                    "visible": False,
                    "frame": frame_anio,
                    "header": header_anio
                }

                header_anio.configure(command=lambda a=anio: _toggle_anio(a))

        def _filtrar(registros, sel_anio, sel_mes, sel_dia, sel_tipo, texto):
            texto = (texto or "").strip().lower()

            def coincide(r):
                if sel_anio != "Todos":
                    if str(r.get("anio")) != sel_anio:
                        return False

                if sel_mes != "Todos":
                    try:
                        m_int = int(sel_mes)
                    except Exception:
                        return False
                    if r.get("mes") != m_int:
                        return False

                if sel_dia != "Todos":
                    try:
                        d_int = int(sel_dia)
                    except Exception:
                        return False
                    if r.get("dia") != d_int:
                        return False

                if sel_tipo != "Todos":
                    if r.get("tipo") != sel_tipo:
                        return False

                if texto:
                    cad = " ".join([
                        r.get("rut", ""),
                        r.get("numero", ""),
                        r.get("archivo", ""),
                        r.get("tipo", ""),
                    ]).lower()
                    if texto not in cad:
                        return False

                return True

            return [r for r in registros if coincide(r)]

        def _iniciar_busqueda():
            # Evitar doble click
            if searching["value"]:
                return

            searching["value"] = True
            last_search_token["value"] += 1
            token = last_search_token["value"]

            # snapshot de filtros al momento del click
            sel_anio = var_anio.get()
            sel_mes  = var_mes.get()
            sel_dia  = var_dia.get()
            sel_tipo = var_tipo_doc.get()
            texto    = var_buscar.get()

            _set_loading(True, "🔎 Buscando documentos…")

            def worker():
                try:
                    # Carga “pesada”
                    regs = _construir_historial_desde_salida()
                    filtrados = _filtrar(regs, sel_anio, sel_mes, sel_dia, sel_tipo, texto)

                    def _ui():
                        # si se cerró la ventana o ya hubo otra búsqueda, ignorar
                        if not ui_alive["value"]:
                            return
                        if token != last_search_token["value"]:
                            return

                        _pintar_por_ano(filtrados)
                        _set_loading(False)
                        searching["value"] = False

                    hist.after(0, _ui)

                except Exception as e:
                    def _ui_err():
                        if not ui_alive["value"]:
                            return
                        _set_loading(False)
                        searching["value"] = False
                        messagebox.showerror("Historial", f"Error al buscar documentos:\n{e}")

                    hist.after(0, _ui_err)

            threading.Thread(target=worker, daemon=True).start()

        # Enter en el buscador = Buscar
        entry_buscar.bind("<Return>", lambda e: _iniciar_busqueda())

    # Redirección de stdout/stderr al textbox (cola + after)
    import queue as _q
    log_queue = _q.Queue()

    class _ConsoleRedirect:
        def __init__(self, queue_): self.queue = queue_
        def write(self, text): self.queue.put(text)
        def flush(self): pass

    sys.stdout = _ConsoleRedirect(log_queue)
    sys.stderr = _ConsoleRedirect(log_queue)

    def limpiar_log():
        """Borra el textbox y drena la cola, así no reaparecen mensajes viejos."""
        try:
            texto_log.delete("1.0", "end")
        except Exception:
            pass
        try:
            while not log_queue.empty():
                log_queue.get_nowait()
        except Exception:
            pass
        try:
            texto_log.update_idletasks()
        except Exception:
            pass

    def limpiar_log_desde_opcion():
        """
        Limpia SOLO desde la línea 'Seleccione una opción:' hacia abajo,
        manteniendo los datos de sucursal visibles.
        Termina en 'Seleccione una opción:' y deja un salto extra.
        """
        try:
            idx = texto_log.search("Seleccione una opción:", "1.0", stopindex="end")
            if idx:
                # Borrar todo lo que viene después de esa línea
                fin_linea = texto_log.index(f"{idx} lineend")
                texto_log.delete(f"{fin_linea}+1c", "end")

                # Asegurar que quede con un salto extra hacia abajo
                # (si ya hay saltos, no pasa nada grave, pero lo normaliza)
                texto_log.insert("end", "\n")
            else:
                # Si no encuentra el marcador, limpia todo y deja el marcador
                texto_log.delete("1.0", "end")
                texto_log.insert("end", "Seleccione una opción:\n")
        except Exception:
            pass

        # Drena la cola para que no reaparezcan mensajes viejos
        try:
            while not log_queue.empty():
                log_queue.get_nowait()
        except Exception:
            pass

        try:
            texto_log.see("end")
            texto_log.update_idletasks()
        except Exception:
            pass

    # === HISTORIAL DESDE CARPETA DE SALIDA ============================
    def _construir_historial_desde_salida():
        """
        Lee TODOS los PDFs que existan en la carpeta de salida y subcarpetas
        y construye una lista de registros con:
        - ruta, archivo
        - fecha (desde fecha de modificación del archivo)
        - anio, mes, dia
        - rut
        - numero (factura o guía)
        - tipo: 'Factura', 'Guía de despacho', 'CHEP' u 'Otros'
        """
        carpeta_salida = variables.get("CarpSalida")
        if not carpeta_salida or not os.path.isdir(carpeta_salida):
            return []

        registros = []

        for root, dirs, files in os.walk(carpeta_salida):
            for name in files:
                if not name.lower().endswith(".pdf"):
                    continue

                ruta = os.path.join(root, name)

                # --- Fecha desde el sistema de archivos ---
                try:
                    ts = os.path.getmtime(ruta)
                    dt = datetime.fromtimestamp(ts)
                except Exception:
                    dt = datetime.min


                # --- Nombre base y versiones en mayúscula/minúscula ---
                base_no_ext = os.path.splitext(name)[0]
                lower = base_no_ext.lower()
                upper = base_no_ext.upper()

                # --- RUT ---
                m_rut = re.search(r"(\d{7,8}-[0-9kK])", base_no_ext)
                rut = m_rut.group(1) if m_rut else ""

                # --- Tipo de documento y número ---
                tipo = "Factura"
                numero = ""

                if "_guia_" in lower:
                    # Guías de despacho: Lo Valledor_76505519-9_guia_252346_2025.pdf
                    tipo = "Guía de despacho"
                    m_num = re.search(r"_guia_([0-9]+)", lower)
                    if m_num:
                        numero = m_num.group(1)

                elif "_chep_" in lower:
                    # CHEP: JJ Perez_CHEP_20251118_150925.pdf
                    tipo = "CHEP"
                    m_num = re.search(r"_chep_([0-9]{8})_([0-9]{6})", lower)
                    if m_num:
                        # Ej: 20251118-150925
                        numero = f"{m_num.group(1)}-{m_num.group(2)}"

                else:
                    # Facturas: Lo Blanco_93178000-K_factura_19211105_2025_3.pdf
                    if "_factura_" in lower:
                        tipo = "Factura"
                        m_num = re.search(r"_factura_([0-9]+)", lower)
                        if m_num:
                            numero = m_num.group(1)
                    else:
                        tipo = "Otros"


                registros.append({
                    "ruta": ruta,
                    "archivo": name,
                    "rut": rut,
                    "numero": numero,
                    "tipo": tipo,
                    "fecha": dt,
                    "anio": dt.year if dt else None,
                    "mes": dt.month if dt else None,
                    "dia": dt.day if dt else None,
                })

        # Orden por fecha (más nuevo primero) y luego por nombre
        registros.sort(
            key=lambda r: ((r["fecha"] or datetime.min), r["archivo"]),
            reverse=True
        )
        return registros

    def imprimir_config_actual():
        """Imprime la configuración actual en el log (resumen estándar)."""
        print(f"Razón social: {variables.get('RazonSocial')}")
        print(f"RUT empresa: {variables.get('RutEmpresa')}")
        print(f"Sucursal: {variables.get('NomSucursal')}")
        print(f"Dirección: {variables.get('DirSucursal')}\n")
        print("Seleccione una opción:")

    def repintar_config():
        # Limpia y agenda el pintado en el siguiente ciclo del loop Tk
        limpiar_log()
        ventana.after(0, imprimir_config_actual)

    # Mostrar datos actuales de la configuración
    print(f"Razón social: {variables.get('RazonSocial')}")
    print(f"RUT empresa: {variables.get('RutEmpresa')}")
    print(f"Sucursal: {variables.get('NomSucursal')}")
    print(f"Dirección: {variables.get('DirSucursal')}\n")
    print("Seleccione una opción:")

    def _cambiar_config():

        try:
            modales_abiertos["config"] = True
            ventana.configure(cursor="wait")
            mensaje_espera.configure(text="⚙️ Abriendo configuración…")
            for b in (btn_escanear, btn_procesar,btn_historial):
                if b is None:
                    continue
                b.configure(state="disabled")


            nuevas = cargar_o_configurar(force_selector=True)
            if not nuevas:
                return

            aplicar_nueva_config(nuevas)
            variables.clear(); variables.update(nuevas)

            repintar_config()

            messagebox.showinfo("Configuración", "La configuración se actualizó correctamente.")
        except Exception as e:
            messagebox.showerror("Configuración", f"No se pudo actualizar la configuración:\n{e}")
        finally:
            modales_abiertos["config"] = False
            mensaje_espera.configure(text=""); ventana.configure(cursor="")
            for b in (btn_escanear, btn_procesar, btn_historial):
                if b is None:
                    continue
                b.configure(state="normal")
            try: ventana.after(0, actualizar_texto)
            except Exception: pass

#     # --- Cambiar rutas (solo CarEntrada/CarpSalida) ---
    def _cambiar_rutas():
        try:
            modales_abiertos["rutas"] = True
            ventana.configure(cursor="wait")
            mensaje_espera.configure(text="🗂️ Abriendo cambio de rutas…")
            for b in (btn_escanear, btn_procesar,btn_historial):
                if b is None:
                    continue
                b.configure(state="disabled")

            nuevas = actualizar_rutas(variables, parent=ventana) 
            if not nuevas:
                return

            aplicar_nueva_config(nuevas)

            variables.clear()
            variables.update(nuevas)

            print("\n📁 Rutas actualizadas:")
            print(f"Carpeta entrada: {variables.get('CarEntrada')}")
            print(f"Carpeta salida : {variables.get('CarpSalida')}\n")

            messagebox.showinfo("Rutas", "Rutas de entrada/salida actualizadas correctamente.")
        except Exception as e:
            messagebox.showerror("Rutas", f"No se pudieron actualizar las rutas:\n{e}")
        finally:
            modales_abiertos["rutas"] = False
            mensaje_espera.configure(text=""); ventana.configure(cursor="")
            for b in (btn_escanear, btn_procesar,btn_historial):
                if b is None:
                    continue
                b.configure(state="normal")

            # <- rearmar el refresco del log
            try: ventana.after(0, actualizar_texto)
            except Exception: pass

# Seleccionar sucursal (Solución versión oficina)
    def cambiar_Sucursales():
        import traceback
        try:
            modales_abiertos["sucursal"] = True
            ventana.configure(cursor="wait")
            mensaje_espera.configure(text="🏷️ Seleccionando sucursal…")
            for b in (btn_escanear, btn_procesar,btn_historial):
                if b is None:
                    continue
                b.configure(state="disabled")

            nuevas = seleccionar_razon_sucursal_grid(variables, parent=ventana)
            if not nuevas:
                print("ℹ️ Cambio de sucursal cancelado por el usuario.")
                return

            aplicar_nueva_config(nuevas)
            variables.clear(); variables.update(nuevas)

            # refresco de log
            try:
                texto_log.delete("1.0", "end")
            except Exception:
                pass
            print("\n⚙️ Configuración actualizada")
            print(f"Razón social: {variables.get('RazonSocial')}")
            print(f"RUT empresa: {variables.get('RutEmpresa')}")
            print(f"Sucursal: {variables.get('NomSucursal')}")
            print(f"Dirección: {variables.get('DirSucursal')}\n")
            print("Seleccione una opción:")

            messagebox.showinfo("Configuración", "Sucursal cambiada correctamente.")
        except Exception as err:
            # Si por algún motivo 'err' no es una excepción normal, mostrar tipo + repr
            try:
                detalle = f"{err!s}"
            except Exception:
                detalle = f"{type(err).__name__}: {repr(err)}"
            traceback.print_exc()
            messagebox.showerror("Configuración", f"No se pudo cambiar la sucursal:\n{detalle}")
        finally:
            modales_abiertos["sucursal"] = False
            mensaje_espera.configure(text=""); ventana.configure(cursor="")
            for b in (btn_escanear, btn_procesar,btn_historial):
                if b is None:
                    continue
                try: b.configure(state="normal")
                except Exception: pass
            try: ventana.after(0, actualizar_texto)
            except Exception: pass

    # === Actualización del textbox ===
    def actualizar_texto():
        while not log_queue.empty():
            texto_log.insert("end", log_queue.get())
            texto_log.see("end")
        ventana.after(100, actualizar_texto)

    # ====== Hilos de acciones principales ======
    def hilo_escanear():
        try:
            en_proceso["activo"] = True
            btn_escanear.configure(state="disabled")
            btn_procesar.configure(state="disabled")
            mensaje_espera.configure(text="🔄 Escaneando...")
            ventana.configure(cursor="wait")

            nombre_pdf = f"DocEscaneado_{datetime.now():%Y%m%d_%H%M%S}.pdf"
            rutas = escanear_y_guardar_pdf(nombre_pdf, variables["CarEntrada"])

            if not rutas:
                print("⚠️ Escaneo cancelado o sin páginas")
                return

            # Normalizamos a lista por si en algún momento devuelve solo un string
            if isinstance(rutas, str):
                rutas = [rutas]

            for ruta in rutas:
                msg = f"Documento escaneado: {os.path.basename(ruta)}"
                print(msg)
                registrar_log(msg)

                resultado = procesar_archivo(ruta)
                if resultado:
                    nombre_out = os.path.basename(resultado)

                    # Guardar ruta para poder abrirla desde el log
                    registrar_link_documento(nombre_out, resultado)

                    # Log a archivo con ruta clickeable
                    uri = "file:///" + resultado.replace("\\", "/")

                    if "No_Reconocidos" in resultado:
                        aviso = f"⚠️ Documento movido a No_Reconocidos: {nombre_out}"
                        print(aviso)
                        registrar_log(aviso)
                    else:
                        msg_ok = f"✅ Procesado: {nombre_out}"
                        print(msg_ok)
                        registrar_log(f"✅ Procesado: {uri}")

        except Exception as e:
            print(f"❗ Error en escaneo: {e}")
        finally:
            en_proceso["activo"] = False
            mensaje_espera.configure(text="")
            btn_escanear.configure(state="normal")
            btn_procesar.configure(state="normal")
            ventana.configure(cursor="")

    def hilo_procesar():
        try:
            en_proceso["activo"] = True
            btn_escanear.configure(state="disabled")
            btn_procesar.configure(state="disabled")
            mensaje_espera.configure(text="🗂️ Procesando carpeta...")
            ventana.configure(cursor="wait")
            procesar_entrada_una_vez()
        finally:
            en_proceso["activo"] = False
            mensaje_espera.configure(text="")
            btn_escanear.configure(state="normal")
            btn_procesar.configure(state="normal")
            ventana.configure(cursor="")

    def iniciar_escanear():
        # Evita doble inicio por doble click/Enter
        if en_proceso.get("activo"):
            print("⏳ Ya hay una tarea en curso; se ignora el click duplicado.")
            return
        en_proceso["activo"] = True

        try:
            btn_escanear.configure(state="disabled")
            btn_procesar.configure(state="disabled")
        except Exception:
            pass

        threading.Thread(target=hilo_escanear, daemon=True).start()

    def cambiar_scanner():
        if en_proceso.get("activo"):
            messagebox.showwarning("Proceso en curso", "Espera a que termine el proceso actual antes de cambiar el escáner.")
            return

        # deshabilitamos botones mientras se abre el diálogo
        try:
            btn_escanear.configure(state="disabled")
            btn_procesar.configure(state="disabled")
            btn_cambiar_scanner.configure(state="disabled")
            mensaje_espera.configure(text="🖨️ Selecciona el escáner predeterminado…")
            ventana.configure(cursor="wait")

            info = seleccionar_scanner_predeterminado()
            if info and info.get("device_id"):
                nombre = info.get("name") or info["device_id"]
                messagebox.showinfo("Escáner", f"Escáner predeterminado actualizado:\n{nombre}")
                print(f"🖨️ Escáner predeterminado: {nombre}")
            else:
                # cancelado o no seleccionado
                pass
        finally:
            mensaje_espera.configure(text="")
            ventana.configure(cursor="")
            try:
                btn_escanear.configure(state="normal")
                btn_procesar.configure(state="normal")
                btn_cambiar_scanner.configure(state="normal")
            except Exception:
                pass

    def iniciar_procesar():
        if en_proceso.get("activo"):
            print("⏳ Ya hay una tarea en curso; se ignora el click duplicado.")
            return
        en_proceso["activo"] = True

        try:
            btn_escanear.configure(state="disabled")
            btn_procesar.configure(state="disabled")
        except Exception:
            pass

        threading.Thread(target=hilo_procesar, daemon=True).start()

    # Botones principales
    BTN_W = 300
    BTN_H = 60
    SW_W  = 30
    SW_H  = 40
    GAP_X = 8

    frame_botones.grid_columnconfigure(0, weight=0)
    frame_botones.grid_columnconfigure(1, weight=0)

    # ===== Fila 1: Escanear + (opcional) Cambiar escáner =====
    btn_escanear = ctk.CTkButton(
        frame_botones, text="ESCANEAR DOCUMENTO", image=icono_escaneo,
        compound="left", width=BTN_W, height=BTN_H, font=fuente_texto,
        fg_color=CUSTOMCOLOR, hover_color=HOBERCOLOR, text_color="black",
        border_color=BORDERCOLOR, border_width=1,
        command=iniciar_escanear
    )
    btn_escanear.grid(row=0, column=0, pady=(0, 8))

    btn_cambiar_scanner = ctk.CTkButton(
        frame_botones,
        text="",
        image=icono_cambiar_scanner,
        width=SW_W,
        height=SW_H,
        corner_radius=SW_W // 2,
        fg_color="#a6a6a6",
        hover_color="#8c8c8c",
        text_color="black",
        command=cambiar_scanner
    )
    # Si lo quieres visible, descomenta esta línea:
    # btn_cambiar_scanner.grid(row=0, column=1, pady=(0, 8), sticky="w")

    # Tooltip simple
    def _tip_on(_=None): mensaje_espera.configure(text="🖨️ Cambiar escáner predeterminado")
    def _tip_off(_=None): mensaje_espera.configure(text="")
    btn_cambiar_scanner.bind("<Enter>", _tip_on)
    btn_cambiar_scanner.bind("<Leave>", _tip_off)

    # ===== Fila 2: Procesar =====
    btn_procesar = ctk.CTkButton(
        frame_botones, text="PROCESAR CARPETA", image=icono_carpeta,
        compound="left", width=BTN_W, height=BTN_H, font=fuente_texto,
        fg_color=CUSTOMCOLOR, hover_color=HOBERCOLOR, text_color="black",
        border_color=BORDERCOLOR, border_width=1,
        command=iniciar_procesar
    )
    btn_procesar.grid(row=1, column=0, pady=(0, 0))

    # ================== BARRA SUPERIOR (MENUBAR) ==================
    import tkinter as tk

    def _menu_escanear():
        if en_proceso.get("activo"):
            messagebox.showwarning("Proceso en curso", "Espera a que termine el proceso actual.")
            return
        iniciar_escanear()

    def _menu_procesar():
        if en_proceso.get("activo"):
            messagebox.showwarning("Proceso en curso", "Espera a que termine el proceso actual.")
            return
        iniciar_procesar()

    menubar = tk.Menu(ventana)

    # --- Menu ---
    menu_app = tk.Menu(menubar, tearoff=0)
    menu_app.add_command(label="Escanear documento", command=_menu_escanear)
    menu_app.add_separator()
    menu_app.add_command(label="Procesar carpeta", command=_menu_procesar)
    menubar.add_cascade(label="Menu", menu=menu_app)
    
    # --- Ajustes ---
    menu_ajustes = tk.Menu(menubar, tearoff=0)
    if MOSTRAR_BT_CAMBIAR_SUCURSAL_OF:
        menu_ajustes.add_command(label="Cambiar sucursal", command=cambiar_Sucursales)
        menu_ajustes.add_separator()

    menu_ajustes.add_command(label="Cambiar escáner", command=cambiar_scanner)  # <-- NUEVO
    menubar.add_cascade(label="Gestión", menu=menu_ajustes)

    # --- Apariencia ---
    menu_apariencia = tk.Menu(menubar, tearoff=0)
    menu_apariencia.add_command(
        label="Tamaño letra del log…",
        command=lambda: abrir_modal_apariencia(
            parent=ventana,
            base_dir=BASE_DIR,
            font_log=font_log,
            log_font_state=log_font_state,
            mensaje_label=mensaje_espera,
            aplicar_icono_fn=aplicar_icono,
            log_fn=registrar_log
        )
    )
    menubar.add_cascade(label="Apariencia", menu=menu_apariencia)

    # --- (Opcional) Ayuda ---
    menu_ayuda = tk.Menu(menubar, tearoff=0)
    menu_ayuda.add_command(label="Acerca de", command=lambda: messagebox.showinfo("FacturaScan", f"FacturaScan {VERSION}"))
    menubar.add_cascade(label="Ayuda", menu=menu_ayuda)

        # ================== MENÚ ADMIN (OCULTO / VISIBLE CON CTRL+F) ==================
    admin_menu = tk.Menu(menubar, tearoff=0)
    admin_menu.add_command(label="Cambiar sucursal", command=_cambiar_config)
    admin_menu.add_command(label="Cambiar rutas", command=_cambiar_rutas)

    admin_visible = {"value": False}

    # --- Debug (oculto / visible con Ctrl+F junto a Admin) ---
    debug_menu = tk.Menu(menubar, tearoff=0)
    debug_menu.add_command(label="Abrir panel debug", command=lambda: abrir_panel_debug(ventana))

    def _show_admin_menu():
        if admin_visible["value"]:
            return

        menubar.add_cascade(label="Admin", menu=admin_menu)
        menubar.add_cascade(label="Debug", menu=debug_menu)  # <-- NUEVO

        admin_visible["value"] = True
        ventana.config(menu=menubar)

    def _hide_admin_menu():
        if not admin_visible["value"]:
            return
        try:
            end = menubar.index("end")
            if end is None:
                return

            # borrar "Admin" y "Debug"
            labels_a_borrar = {"Admin", "Debug"}
            for i in range(end, -1, -1):  # iterar al revés para no desordenar índices
                try:
                    if menubar.type(i) == "cascade" and menubar.entrycget(i, "label") in labels_a_borrar:
                        menubar.delete(i)
                except Exception:
                    pass
        finally:
            admin_visible["value"] = False
            ventana.config(menu=menubar)

    ventana.config(menu=menubar)

    def _toggle_admin_visibility(event=None):
        if modales_abiertos["config"] or modales_abiertos["rutas"] or modales_abiertos["sucursal"]:
            return
        if admin_visible["value"]:
            _hide_admin_menu()
        else:
            _show_admin_menu()

    ventana.bind_all("<Control-f>", _toggle_admin_visibility)
    ventana.bind_all("<Control-F>", _toggle_admin_visibility)


    # Cierre seguro
    def intento_cerrar():
        if en_proceso["activo"]:
            messagebox.showwarning("Proceso en curso", "No puedes cerrar la aplicación mientras se ejecuta una tarea.")
            return
        try:
            guardar_tamano_log(BASE_DIR, log_font_state.get("size", 12), log_fn=registrar_log)
        except Exception:
            pass

        cerrar_aplicacion(ventana, modales_abiertos)

    ventana.protocol("WM_DELETE_WINDOW", intento_cerrar)

    # Loop UI
    actualizar_texto()
    ventana.mainloop()

# ================== EJECUCIÓN DEL PROGRAMA ==================
if __name__ == "__main__":
    try:
        Valida_PopplerPath()
        menu_Principal()
    except Exception as e:
        show_startup_error(f"Error al iniciar FacturaScan:\n\n{e}")
