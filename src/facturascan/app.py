import sys, os, ctypes, threading, winreg
from importlib import resources

import customtkinter as ctk
from tkinter import messagebox

from utils.log_utils import set_debug, is_debug, registrar_log
from core.monitor_core import aplicar_nueva_config

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
    """
    Devuelve la ruta del icono .ico tanto en desarrollo como compilado (Nuitka).
    """
    try:
        p = resources.files("facturascan.resources.icons") / "iconoScan.ico"
        with resources.as_file(p) as real_path:
            return str(real_path)
    except Exception:
        here = os.path.dirname(__file__)
        return os.path.join(here, "assets", "iconoScan.ico")
    
def aplicar_icono(win):
    ico_path = _get_icon_ico_path()
    ico_path = ico_path.replace("\\", "/")
    try:
        win.iconbitmap(default=ico_path)
    except Exception as e:
        # No revienta la app si el icono falla
        print("No se pudo aplicar ícono:", e)

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

# 20-11-2025
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

from __init__ import __version__, MOSTRAR_BT_CAMBIAR_SUCURSAL_OF, ACTUALIZAR_PROGRAMA
VERSION = __version__

# ====== BLOQUE ACTUALIZACIONES (forzado, sin cancelar) ======
from tkinter import messagebox as _mb
from update.updater import (
    is_update_available, download_with_progress, verify_sha256, run_installer,
    cleanup_temp_dir
)

def _mostrar_dialogo_update(ventana):
    """
    Actualización forzada al inicio:
    - No pregunta al usuario.
    - Muestra aviso + barra de progreso sin botón de cancelar.
    - Descarga el instalador, verifica (si hay SHA), ejecuta con progreso
      y cierra la app actual para que el instalador reinicie FacturaScan.
    """
    try:
        info = is_update_available(VERSION)
        if info.get("error") or not info.get("update_available"):
            return  # No hay update → salimos silenciosamente

        latest = info.get("latest", "")     # p.ej. v1.9.4
        url    = info.get("installer_url")  # .exe del Release
        sha    = info.get("sha256")         # puede ser None
        if not url:
            return

        # ---------- helpers ----------
        def _centrar(child, parent, w, h):
            parent.update_idletasks()
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            pw, ph = parent.winfo_width(), parent.winfo_height()
            x = px + max(0, (pw - w) // 2)
            y = py + max(0, (ph - h) // 2)
            child.geometry(f"{w}x{h}+{x}+{y}")

        # ---------- diálogo sin cancelar ----------
        prog = ctk.CTkToplevel(ventana)
        prog.withdraw()
        prog.title("Actualizando FacturaScan…")
        prog.resizable(False, False)
        try:
            aplicar_icono(prog)
        except Exception:
            pass
        prog.after(200, lambda: aplicar_icono(prog))

        W, H = 460, 190
        _centrar(prog, ventana, W, H)
        prog.deiconify()
        prog.transient(ventana)
        prog.grab_set()
        prog.focus_force()
        prog.attributes("-topmost", True)
        prog.after(60, lambda: prog.attributes("-topmost", False))

        ctk.CTkLabel(
            prog,
            text=f"Hay una nueva versión {latest}.\nSe descargará e instalará automáticamente."
        ).pack(pady=(14, 8))

        pct_var = ctk.StringVar(value="0 %")
        ctk.CTkLabel(prog, textvariable=pct_var).pack(pady=(0, 6))

        bar = ctk.CTkProgressBar(prog); bar.set(0.0); bar.pack(fill="x", padx=16, pady=8)

        status_var = ctk.StringVar(value="Descargando instalador…")
        ctk.CTkLabel(prog, textvariable=status_var).pack(pady=(2, 10))

        # Deshabilitar cierre del modal y tecla Escape (no se puede cancelar)
        prog.protocol("WM_DELETE_WINDOW", lambda: None)
        prog.bind("<Escape>", lambda e: "break")

        import tempfile, threading, os
        tmpdir   = os.path.join(tempfile.gettempdir(), "FacturaScan_Update")
        ui_alive = {"value": True}

        def _safe_ui(fn):
            if not ui_alive["value"]:
                return
            try:
                if prog.winfo_exists():
                    fn()
            except Exception:
                pass

        def _set_progress(read, total):
            if not ui_alive["value"]:
                return
            if total and total > 0:
                frac = max(0.0, min(1.0, read / total))
                prog.after(0, lambda: _safe_ui(lambda: (bar.set(frac), pct_var.set(f"{int(frac*100)} %"))))
            else:
                kb = read // 1024
                prog.after(0, lambda: _safe_ui(lambda: pct_var.set(f"{kb} KB")))

        def _worker():
            try:
                # 👇 SIN cancel_event → imposible cancelar desde UI
                exe_path = download_with_progress(url, tmpdir, progress_cb=_set_progress, cancel_event=None)

                if sha:
                    prog.after(0, lambda: _safe_ui(lambda: status_var.set("Verificando integridad…")))
                    if not verify_sha256(exe_path, sha):
                        prog.after(0, lambda: _mb.showerror("Actualización", "El archivo no pasó la verificación SHA-256."))
                        try: cleanup_temp_dir(tmpdir)
                        except Exception: pass
                        ui_alive["value"] = False
                        try: prog.destroy()
                        except Exception: pass
                        return

                prog.after(0, lambda: _safe_ui(lambda: status_var.set("Instalando actualización…")))
                ui_alive["value"] = False
                try:
                    prog.destroy()
                except Exception:
                    pass

                # Ejecuta instalador con ventana de progreso y reinicio automático
                # (cierra la app actual mientras Inno instala y la relanza).
                ventana.after(200, lambda: run_installer(exe_path, mode="progress", cleanup_dir=tmpdir))

            except Exception as e:
                # Errores de red u otros
                try:
                    prog.after(0, lambda: _mb.showerror("Actualización", f"Error al actualizar:\n{e}"))
                finally:
                    ui_alive["value"] = False
                    try: prog.destroy()
                    except Exception: pass
                    try: cleanup_temp_dir(tmpdir)
                    except Exception: pass

        threading.Thread(target=_worker, daemon=True).start()

    except Exception:
        # No rompemos la app si algo falla en el update
        pass

def validarActualización(ventana):
    # Espera a que la UI esté estable y lanza el diálogo
    ventana.after(800, lambda: _mostrar_dialogo_update(ventana))

# ================== UTILIDADES ==================
# 20-11-2025
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

# ================== INTERFAZ PRINCIPAL ==================
def menu_Principal():
    from PIL import Image
    from datetime import datetime
    from core.scanner import escanear_y_guardar_pdf

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

    # Actualizaciones por Github
    if ACTUALIZAR_PROGRAMA:    
        validarActualización(ventana)
        
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
    ventana.resizable(True, True)

    fuente_titulo = ctk.CTkFont(size=40, weight="bold")
    fuente_texto = ctk.CTkFont(family="Segoe UI", size=15)

    # Título y zona de botones principales
    ctk.CTkLabel(ventana, text="FacturaScan", font=fuente_titulo).pack(pady=15)
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
        icono_carpeta = ctk.CTkImage(
            light_image=Image.open(asset_path("icono_carpeta.png")),
            size=(26, 26)
        )
    except Exception:
        registrar_log("no se encontro imagen icono_carpeta.png")
        icono_carpeta = None

    # ===== Textbox de log y mensaje inferior =====
    texto_log = ctk.CTkTextbox(
        ventana, width=650, height=260,
        font=("Consolas", 12), wrap="word",
        corner_radius=6, fg_color="white", text_color="black"
    )
    texto_log.pack(pady=15, padx=15)

    mensaje_espera = ctk.CTkLabel(ventana, text="", font=fuente_texto, text_color="gray")
    mensaje_espera.pack(pady=(0, 10))

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

    # ===== Chip de debug y botones de administración ocultos =====
    debug_ui_visible = {"value": False}
    chip_pad = 12
    chip_width, chip_height = 120, 30
    btn_small_h = 30
    GAP = 6

    def _actualizar_chip_estilo():
        if is_debug():
            debug_chip.configure(text="DEBUG ON", fg_color="#10B981", text_color="white", hover_color="#059669")
            print("Modo DEBUG ACTIVADO")
        else:
            print("Modo DEBUG DESACTIVADO")
            debug_chip.configure(text="DEBUG OFF", fg_color="#E5E7EB", text_color="#111827", hover_color="#D1D5DB")
                
    def _toggle_debug_state():
        nuevo = not is_debug()
        set_debug(nuevo)
        _actualizar_chip_estilo()

    debug_chip = ctk.CTkButton(
        ventana, text="DEBUG OFF",
        width=chip_width, height=chip_height, corner_radius=16,
        fg_color="#E5E7EB", text_color="#111827", hover_color="#D1D5DB",
        command=_toggle_debug_state
    )
    debug_chip.place_forget()

    # --- Cambiar sucursal (selector completo de config) ---
    def _cambiar_config():
        try:
            modales_abiertos["config"] = True
            ventana.configure(cursor="wait")
            mensaje_espera.configure(text="⚙️ Abriendo configuración…")
            for b in (btn_escanear, btn_procesar, btn_config, btn_rutas, debug_chip, btn_sucursal_rapida):
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
            for b in (btn_escanear, btn_procesar, btn_config, btn_rutas, debug_chip,btn_sucursal_rapida):
                if b is None:
                    continue
                b.configure(state="normal")
            try: ventana.after(0, actualizar_texto)
            except Exception: pass

    btn_config = ctk.CTkButton(
        ventana, text="Cambiar sucursal",
        width=140, height=btn_small_h, corner_radius=16,
        fg_color="#E5E7EB", text_color="#111827", hover_color="#D1D5DB",
        command=_cambiar_config
    )
    btn_config.place_forget()

    # --- Cambiar rutas (solo CarEntrada/CarpSalida) ---
    def _cambiar_rutas():
        try:
            modales_abiertos["rutas"] = True
            ventana.configure(cursor="wait")
            mensaje_espera.configure(text="🗂️ Abriendo cambio de rutas…")
            for b in (btn_escanear, btn_procesar, btn_config, btn_rutas, debug_chip, btn_sucursal_rapida):
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
            for b in (btn_escanear, btn_procesar, btn_config, btn_rutas, debug_chip, btn_sucursal_rapida):
                if b is None:
                    continue
                b.configure(state="normal")

            # <- rearmar el refresco del log
            try: ventana.after(0, actualizar_texto)
            except Exception: pass

    btn_rutas = ctk.CTkButton(
        ventana, text="Cambiar rutas",
        width=140, height=btn_small_h, corner_radius=16,
        fg_color="#E5E7EB", text_color="#111827", hover_color="#D1D5DB",
        command=_cambiar_rutas
    )
    btn_rutas.place_forget()

# Seleccionar sucursal (Solución versión oficina)
    def cambiar_Sucursales():
        import traceback
        try:
            modales_abiertos["sucursal"] = True
            ventana.configure(cursor="wait")
            mensaje_espera.configure(text="🏷️ Seleccionando sucursal…")
            for b in (btn_escanear, btn_procesar, btn_config, btn_rutas, debug_chip, btn_sucursal_rapida):
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
            for b in (btn_escanear, btn_procesar, btn_config, btn_rutas, debug_chip,btn_sucursal_rapida):
                if b is None:
                    continue
                try: b.configure(state="normal")
                except Exception: pass
            try: ventana.after(0, actualizar_texto)
            except Exception: pass

    # Botón visible arriba-izquierda cambiar sucursal "oficina"
    btn_sucursal_rapida = None
    if MOSTRAR_BT_CAMBIAR_SUCURSAL_OF:
        btn_sucursal_rapida = ctk.CTkButton(
            ventana, text="Seleccionar sucursal",
            width=160, height=32, corner_radius=16,
            fg_color="#E5E7EB", text_color="#111827", hover_color="#D1D5DB",
            command=cambiar_Sucursales
        )
        btn_sucursal_rapida.place(relx=0.0, rely=0.0, x=12, y=12, anchor="nw")

    # Mostrar/Ocultar chip y botones de admin
    def _mostrar_chip():
        debug_chip.place(relx=1.0, rely=0.0, x=-chip_pad, y=chip_pad, anchor="ne")
        _actualizar_chip_estilo()
        base_y = chip_pad + chip_height + GAP
        btn_config.place(relx=1.0, rely=0.0, x=-chip_pad, y=base_y, anchor="ne")
        btn_rutas.place(relx=1.0, rely=0.0, x=-chip_pad, y=base_y + btn_small_h + GAP, anchor="ne")
        debug_ui_visible["value"] = True

    def _ocultar_chip():
        debug_chip.place_forget()
        btn_config.place_forget()
        btn_rutas.place_forget()
        debug_ui_visible["value"] = False

    def _toggle_chip_visibility(event=None):
        # No permitir mostrar/ocultar mientras haya modales abiertos
        if modales_abiertos["config"] or modales_abiertos["rutas"] or modales_abiertos["sucursal"]:
            return
        _ocultar_chip() if debug_ui_visible["value"] else _mostrar_chip()

    ventana.bind_all("<Control-f>", _toggle_chip_visibility)
    ventana.bind_all("<Control-F>", _toggle_chip_visibility)

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
            ruta = escanear_y_guardar_pdf(nombre_pdf, variables["CarEntrada"])

            if ruta:
                msg = f"Documento escaneado: {os.path.basename(ruta)}"
                print(msg); registrar_log(msg)

                resultado = procesar_archivo(ruta)
                if resultado:
                    if "No_Reconocidos" in resultado:
                        aviso = f"⚠️ Documento movido a No_Reconocidos: {os.path.basename(resultado)}"
                        print(aviso); registrar_log(aviso)
                    else:
                        print(f"✅ Procesado: {os.path.basename(resultado)}")
                        registrar_log(f"✅ Procesado: {os.path.basename(resultado)}")
            else:
                print("⚠️ Escaneo cancelado por el usuario")
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
    btn_escanear = ctk.CTkButton(
        frame_botones, text="ESCANEAR DOCUMENTO", image=icono_escaneo,
        compound="left", width=300, height=60, font=fuente_texto,
        fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black",
        command=iniciar_escanear
    ); btn_escanear.pack(pady=6)

    btn_procesar = ctk.CTkButton(
        frame_botones, text="PROCESAR CARPETA", image=icono_carpeta,
        compound="left", width=300, height=60, font=fuente_texto,
        fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black",
        command=iniciar_procesar
    ); btn_procesar.pack(pady=6)

    # Cierre seguro
    def intento_cerrar():
        if en_proceso["activo"]:
            messagebox.showwarning("Proceso en curso", "No puedes cerrar la aplicación mientras se ejecuta una tarea.")
            return
        cerrar_aplicacion(ventana, modales_abiertos)

    ventana.protocol("WM_DELETE_WINDOW", intento_cerrar)

    # Loop UI
    actualizar_texto()
    ventana.mainloop()

# ================== EJECUCIÓN DEL PROGRAMA ==================
# 20-11-2025
if __name__ == "__main__":
    try:
        Valida_PopplerPath()
        menu_Principal()
    except Exception as e:
        show_startup_error(f"Error al iniciar FacturaScan:\n\n{e}")
