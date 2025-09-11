import sys, os, ctypes
import queue
import winreg
import customtkinter as ctk
from tkinter import messagebox

from log_utils import set_debug, is_debug
from monitor_core import aplicar_nueva_config

# === Helpers de assets e icono ===
if getattr(sys, "frozen", False):  
    BASE_DIR = os.path.dirname(sys.executable)
else: 
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ASSETS_DIR = os.path.join(BASE_DIR, "assets")
ICON_BIG   = os.path.join(ASSETS_DIR, "iconoScan.ico")
ICON_SMALL = os.path.join(ASSETS_DIR, "iconoScan16.ico")

def asset_path(nombre: str) -> str:
    """Devuelve la ruta absoluta dentro de /assets."""
    return os.path.join(ASSETS_DIR, nombre)


# Fijar AppUserModelID (sirve para anclar en la barra de tareas correctamente)
def set_appusermodel_id(app_id: str) -> None:
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)  
    except Exception:
        pass

# Fijar icono sin parpadeo con API nativa (WM_SETICON)
from ctypes import wintypes
def aplicar_icono(win) -> bool:
    """
    Fija el icono de la ventana en Windows:
    - default y actual (fallback Tk)
    - WM_SETICON (small y big) con WinAPI
    - re-aplica tras idle para evitar que CTk lo pise
    """
    ok = False

    # 1) Fijar como default para que nuevos Toplevels hereden
    try:
        if os.path.exists(ICON_BIG):
            win.iconbitmap(default=ICON_BIG)
            win.iconbitmap(ICON_BIG)  
            ok = True
    except Exception:
        pass

    # 2) Forzar small/big con WinAPI (estable, sin parpadeos)
    try:
        user32 = ctypes.windll.user32
        IMAGE_ICON, LR_LOADFROMFILE = 1, 0x0010
        WM_SETICON, ICON_SMALL_W, ICON_BIG_W = 0x0080, 0, 1

        # tamaños del sistema
        cx_s = user32.GetSystemMetrics(49)  # SM_CXSMICON
        cy_s = user32.GetSystemMetrics(50)  # SM_CYSMICON
        cx_b = user32.GetSystemMetrics(11)  # SM_CXICON
        cy_b = user32.GetSystemMetrics(12)  # SM_CYICON

        LoadImageW = user32.LoadImageW
        LoadImageW.restype = wintypes.HANDLE

        # Small icon: usa el 16x16 si existe; si no, escala el grande
        small_src = ICON_SMALL if os.path.exists(ICON_SMALL) else ICON_BIG

        h_small = LoadImageW(None, small_src, IMAGE_ICON, cx_s, cy_s, LR_LOADFROMFILE)
        h_big   = LoadImageW(None, ICON_BIG,  IMAGE_ICON, cx_b, cy_b, LR_LOADFROMFILE)

        hwnd = win.winfo_id()
        if h_small:
            user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL_W, h_small)
            ok = True
        if h_big:
            user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG_W, h_big)
            ok = True

        # evitar GC de los handles
        win._hicon_small = h_small  # type: ignore[attr-defined]
        win._hicon_big   = h_big    # type: ignore[attr-defined]
    except Exception:
        pass

    # 3) Re-aplicar después de que CTk termina de montar estilos
    try:
        win.after_idle(lambda: win.iconbitmap(default=ICON_BIG))
        win.after(150,  lambda: win.iconbitmap(default=ICON_BIG))
    except Exception:
        pass

    return ok

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

# Evita que se abran dos FacturaScan a la vez.
def _ensure_single_instance():
    if os.name != "nt":
        return

    import atexit
    import ctypes
    from ctypes import wintypes

    MUTEX_NAME = "Local\\FacturaScanSingleton"

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    ERROR_ALREADY_EXISTS = 183
    # —— tipos/retornos correctos ——
    kernel32.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
    kernel32.CreateMutexW.restype  = wintypes.HANDLE
    kernel32.CloseHandle.argtypes  = [wintypes.HANDLE]
    kernel32.CloseHandle.restype   = wintypes.BOOL
    # ————————————————

    handle = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if not handle:
        return  # no bloquear si falla

    if ctypes.get_last_error() == ERROR_ALREADY_EXISTS:
        try:
            show_startup_error("FacturaScan ya está en ejecución.")
        except Exception:
            pass
        sys.exit(0)

    atexit.register(lambda: kernel32.CloseHandle(handle))

set_appusermodel_id("FacturaScan.App")
# Ejecutar el guardián de instancia única cuanto antes:
_ensure_single_instance()

# ----------------- Imports críticos -----------------
try:
    from config_gui import cargar_o_configurar
    from monitor_core import registrar_log, procesar_archivo, procesar_entrada_una_vez
except Exception as e:
    show_startup_error(f"No se pudo importar un módulo crítico:\n\n{e}")
    sys.exit(1)

faltan = []
for _mod in ["PIL", "pdf2image", "easyocr", "win32com.client", "reportlab"]:
    try:
        __import__(_mod)
    except Exception as _e:
        faltan.append(f"- {_mod}: {_e}")

if faltan:
    show_startup_error("Módulos opcionales no disponibles:\n\n" + "\n".join(faltan))

# === Helper para terminar con mensaje ===
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


# ================== CONFIGURACIÓN INICIAL ==================

variables = None
try:
    variables = cargar_o_configurar()
except Exception as e:
    fatal("CONFIG", e)

if variables is None:
    fatal("CONFIG", Exception("No se obtuvo configuración"))

aplicar_nueva_config(variables)

VERSION = "1.7.0"

# ====== BLOQUE NUEVO: UI de actualización al inicio ======
import threading
from tkinter import messagebox as _mb
from updater import (
    is_update_available, download_with_progress, verify_sha256, run_installer,
    DownloadCancelled, cleanup_temp_dir
)

def _mostrar_dialogo_update(ventana):
    """
    Diálogo de actualización centrado, con icono y cancelación segura.
    Muestra mensaje cuando el usuario cancela.
    """
    try:
        info = is_update_available(VERSION)
        if info.get("error") or not info.get("update_available"):
            return

        latest = info.get("latest", "")
        url    = info.get("installer_url")
        sha    = info.get("sha256")
        if not url:
            return

        if not _mb.askyesno("Actualización disponible",
                            f"Hay una versión nueva {latest}.\n\n¿Quieres actualizar ahora?"):
            return

        # ---------- helpers ----------
        def _centrar(child, parent, w, h):
            parent.update_idletasks()
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            pw, ph = parent.winfo_width(), parent.winfo_height()
            x = px + max(0, (pw - w) // 2)
            y = py + max(0, (ph - h) // 2)
            child.geometry(f"{w}x{h}+{x}+{y}")

        # ---------- diálogo ----------
        prog = ctk.CTkToplevel(ventana)
        prog.withdraw()
        prog.title("Descargando actualización…")
        prog.resizable(False, False)

        # Icono: varias pasadas para que no lo “pise” CTk
        try:
            if os.path.exists(ICON_BIG):
                prog.iconbitmap(default=ICON_BIG)
                prog.iconbitmap(ICON_BIG)
        except Exception:
            pass
        try:
            aplicar_icono(prog)
        except Exception:
            pass
        prog.after_idle(lambda: aplicar_icono(prog))
        prog.after(200,  lambda: aplicar_icono(prog))
        ventana.after(200, lambda: aplicar_icono(prog))
        prog.after(800,  lambda: aplicar_icono(prog))

        # Tamaño y centrado
        W, H = 420, 180
        _centrar(prog, ventana, W, H)
        prog.deiconify()
        prog.transient(ventana)
        prog.lift()
        prog.attributes("-topmost", True)
        prog.after(60, lambda: prog.attributes("-topmost", False))

        # Contenido
        ctk.CTkLabel(prog, text=f"Descargando FacturaScan {latest}").pack(pady=(14, 6))
        pct_var = ctk.StringVar(value="0 %")
        ctk.CTkLabel(prog, textvariable=pct_var).pack(pady=(0, 6))

        bar = ctk.CTkProgressBar(prog)
        bar.set(0.0)
        bar.pack(fill="x", padx=16, pady=8)

        status_var = ctk.StringVar(value="Descargando instalador…")
        ctk.CTkLabel(prog, textvariable=status_var).pack(pady=(2, 6))

        # Cancelación segura
        import tempfile
        cancel_event = threading.Event()
        cancelled = {"value": False}
        ui_alive = {"value": True}

        def _on_cancel():
            if cancelled["value"]:
                return
            cancelled["value"] = True
            cancel_event.set()
            ventana.after(0, lambda: _mb.showinfo("Actualización", "Descarga cancelada por el usuario."))
            try:
                ui_alive["value"] = False
                prog.destroy()
            except Exception:
                pass

        prog.protocol("WM_DELETE_WINDOW", _on_cancel)

        # Botón Cancelar grande
        BTN_W = W - 32  # a todo lo ancho (con márgenes)
        btn_cancel = ctk.CTkButton(
            prog, text="Cancelar",
            width=BTN_W, height=40, corner_radius=18,
            fg_color="#E5E7EB", text_color="#111827",
            hover_color="#D1D5DB",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=_on_cancel,
        )
        btn_cancel.pack(padx=16, pady=(4, 14))

        # Directorio temporal
        tmpdir = os.path.join(tempfile.gettempdir(), "FacturaScan_Update")

        # Actualizar UI solo si sigue viva
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
                prog.after(0, lambda: _safe_ui(
                    lambda: (bar.set(frac), pct_var.set(f"{int(frac*100)} %"))
                ))
            else:
                kb = read // 1024
                prog.after(0, lambda: _safe_ui(lambda: pct_var.set(f"{kb} KB")))

        def _worker():
            try:
                from updater import DownloadCancelled
                exe_path = download_with_progress(
                    url, tmpdir, progress_cb=_set_progress, cancel_event=cancel_event
                )

                if sha:
                    prog.after(0, lambda: _safe_ui(lambda: status_var.set("Verificando integridad…")))
                    if not verify_sha256(exe_path, sha):
                        prog.after(0, lambda: (
                            _mb.showerror("Actualización",
                                          "El archivo descargado no pasó la verificación SHA-256."),
                            _on_cancel()
                        ))
                        return

                prog.after(0, lambda: _safe_ui(lambda: status_var.set("Abriendo instalador…")))
                ui_alive["value"] = False
                try:
                    prog.destroy()
                except Exception:
                    pass
                ventana.after(200, lambda: run_installer(exe_path, silent=False))

            except DownloadCancelled:
                # Mensaje ya mostrado en _on_cancel
                pass
            except Exception as e:
                if cancel_event.is_set():
                    return
                try:
                    prog.after(0, lambda: _mb.showerror("Actualización", f"Error al actualizar:\n{e}"))
                finally:
                    ui_alive["value"] = False
                    try:
                        prog.destroy()
                    except Exception:
                        pass

        threading.Thread(target=_worker, daemon=True).start()

    except Exception:
        pass




def _schedule_update_prompt(ventana):
    # Espera a que la UI esté estable y lanza el diálogo
    ventana.after(800, lambda: _mostrar_dialogo_update(ventana))


# ================== UTILIDADES ==================

# Verifica que Poppler (necesario para convertir PDFs a imágenes) esté en el PATH.
# Si no lo está, lo añade en el registro de Windows y reinicia el programa.
def Valida_PopplerPath():
    ruta_poppler = r"C:\poppler\Library\bin"
    ruta_normalizada = os.path.normcase(os.path.normpath(ruta_poppler))
    path_modificado = False

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment", 0, winreg.KEY_READ) as clave:
            valor_actual, _ = winreg.QueryValueEx(clave, "Path")
    except FileNotFoundError:
        valor_actual = ""

    paths = [os.path.normcase(os.path.normpath(p.strip())) for p in valor_actual.split(";") if p.strip()]
    if ruta_normalizada not in paths:
        nuevo_valor = f"{valor_actual};{ruta_poppler}" if valor_actual else ruta_poppler
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment", 0, winreg.KEY_SET_VALUE) as clave:
                winreg.SetValueEx(clave, "Path", 0, winreg.REG_EXPAND_SZ, nuevo_valor)
            path_modificado = True

            # Notificar a Windows que se cambió el PATH
            ctypes.windll.user32.SendMessageTimeoutW(
                0xFFFF, 0x001A, 0, "Environment", 0x0002, 5000, None
            )
        except PermissionError:
            # Mostrar aviso y seguir sin reiniciar
            try:
                import tkinter as _tk
                from tkinter import messagebox as _mb
                r = _tk.Tk(); r.withdraw()
                _mb.showwarning("Poppler", "No se pudo añadir Poppler al PATH. Ejecuta como administrador.")
                r.destroy()
            except Exception:
                pass

    if path_modificado:
        try:
            import tkinter as _tk
            from tkinter import messagebox as _mb
            r = _tk.Tk(); r.withdraw()
            _mb.showinfo("FacturaScan", "Se añadió Poppler al PATH. La aplicación se reiniciará.")
            r.destroy()
        except Exception:
            pass
        os.execv(sys.executable, [sys.executable] + sys.argv)
# Al intentar cerrar FacturaScan mostrará un mensaje de confirmación
def cerrar_aplicacion(ventana, modales_abiertos=None):
    # Si hay un modal abierto, no cerrar aún
    if modales_abiertos and (modales_abiertos.get("config") or modales_abiertos.get("rutas")or modales_abiertos.get("sucursal")):
        messagebox.showwarning("Ventana abierta", "Cierra primero la ventana de configuración.")
        return

    if not messagebox.askyesno("Cerrar", "¿Deseas cerrar FacturaScan?"):
        return

    try:
        registrar_log('FacturaScan cerrado por el usuario')
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
        registrar_log("FacturaScan Cerrado correctamente")
        # Último recurso para que el proceso termine siempre
        os._exit(0)

# ================== INTERFAZ PRINCIPAL ==================
def mostrar_menu_principal():
    from PIL import Image
    import threading
    from datetime import datetime
    from scanner import escanear_y_guardar_pdf

    registrar_log("FacturaScan iniciado correctamente")

    en_proceso = {"activo": False}
    modales_abiertos = {"config": False, "rutas": False, "sucursal": False}

    # ===== Apariencia y ventana =====
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    ventana = ctk.CTk()
    ventana.title(f"Control documental - FacturaScan {VERSION}")
    aplicar_icono(ventana)
    ventana.after(150, lambda: aplicar_icono(ventana))
    _schedule_update_prompt(ventana)

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
        icono_escaneo = None

    try:
        icono_carpeta = ctk.CTkImage(
            light_image=Image.open(asset_path("icono_carpeta.png")),
            size=(26, 26)
        )
    except Exception:
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
            for b in (btn_escanear, btn_procesar, btn_config, btn_rutas, debug_chip, 
                    #   btn_sucursal_rapida
                      ):
                b.configure(state="disabled")

            from config_gui import cargar_o_configurar
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
            for b in (btn_escanear, btn_procesar, btn_config, btn_rutas, debug_chip, 
                    #   btn_sucursal_rapida
                      ):
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
            for b in (btn_escanear, btn_procesar, btn_config, btn_rutas, debug_chip, 
                    #   btn_sucursal_rapida
                      ):
                b.configure(state="disabled")

            from config_gui import actualizar_rutas
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
            for b in (btn_escanear, btn_procesar, btn_config, btn_rutas, debug_chip, 
                    #   btn_sucursal_rapida
                      ):
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
# --- Seleccionar sucursal (Solución versión oficina) ---
    def _seleccionar_sucursal_rapida():
        try:
            modales_abiertos["sucursal"] = True
            ventana.configure(cursor="wait")
            mensaje_espera.configure(text="🏷️ Seleccionando sucursal…")
            for b in (btn_escanear, btn_procesar, btn_config, btn_rutas, debug_chip, 
                    #   btn_sucursal_rapida
                      ):
                b.configure(state="disabled")

            from config_gui import seleccionar_sucursal_simple
            nuevas = seleccionar_sucursal_simple(variables, parent=ventana)
            if not nuevas:
                print("ℹ️ Operación cancelada por el usuario.")
                return

            aplicar_nueva_config(nuevas)
            variables.clear(); variables.update(nuevas)

            # Limpiar log y volver a imprimir cabecera
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
        except Exception as e:
            messagebox.showerror("Configuración", f"No se pudo cambiar la sucursal:\n{e}")
        finally:
            modales_abiertos["sucursal"] = False
            mensaje_espera.configure(text=""); ventana.configure(cursor="")
            for b in (btn_escanear, btn_procesar, btn_config, btn_rutas, debug_chip, 
                    #   btn_sucursal_rapida
                      ):
                try: b.configure(state="normal")
                except Exception: pass
            try: ventana.after(0, actualizar_texto)
            except Exception: pass

    # # Botón visible arriba-izquierda SIEMPRE
    # btn_sucursal_rapida = ctk.CTkButton(
    #     ventana, text="Seleccionar sucursal",
    #     width=160, height=32, corner_radius=16,
    #     fg_color="#E5E7EB", text_color="#111827", hover_color="#D1D5DB",
    #     command=_seleccionar_sucursal_rapida
    # )
    # btn_sucursal_rapida.place(relx=0.0, rely=0.0, x=12, y=12, anchor="nw")

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
                        print(f"✅ Procesado: {os.path.basename(resultado)}"); registrar_log(f"✅ Procesado: {os.path.basename(resultado)}")
            else:
                print("⚠️ Escaneo cancelado o sin páginas.")
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
if __name__ == "__main__":
    if os.name == 'nt':
        kernel32 = ctypes.WinDLL('kernel32')
        user32 = ctypes.WinDLL('user32')
        whnd = kernel32.GetConsoleWindow()
        if whnd != 0:
            user32.ShowWindow(whnd, 0)
    try:
        if os.name == "nt":
            Valida_PopplerPath()  # <- solo Windows
        mostrar_menu_principal()
    except Exception as e:
        show_startup_error(f"Error al iniciar FacturaScan:\n\n{e}")

