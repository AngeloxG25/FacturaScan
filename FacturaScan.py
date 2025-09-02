import sys, os, ctypes
import queue
import winreg
import customtkinter as ctk
from tkinter import messagebox

from log_utils import set_debug, is_debug
from monitor_core import aplicar_nueva_config

# === Helpers de assets e icono ===
if getattr(sys, "frozen", False):  # si está compilado (exe con Nuitka/PyInstaller)
    BASE_DIR = os.path.dirname(sys.executable)
else:  # si está en modo script normal
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
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)  # type: ignore[attr-defined]
    except Exception:
        pass

# Fijar icono sin parpadeo con API nativa (WM_SETICON)
from ctypes import wintypes
def _set_icon_win32(window, path: str) -> bool:
    try:
        user32 = ctypes.windll.user32  # type: ignore[attr-defined]
        IMAGE_ICON      = 1
        LR_LOADFROMFILE = 0x0010
        WM_SETICON      = 0x0080
        ICON_SMALL      = 0
        ICON_BIG        = 1
        SM_CXSMICON, SM_CYSMICON = 49, 50
        SM_CXICON,  SM_CYICON  = 11, 12

        hwnd = window.winfo_id()
        cx_small = user32.GetSystemMetrics(SM_CXSMICON)
        cy_small = user32.GetSystemMetrics(SM_CYSMICON)
        cx_big   = user32.GetSystemMetrics(SM_CXICON)
        cy_big   = user32.GetSystemMetrics(SM_CYICON)

        LoadImageW = user32.LoadImageW
        LoadImageW.restype = wintypes.HANDLE

        hicon_small = LoadImageW(None, path, IMAGE_ICON, cx_small, cy_small, LR_LOADFROMFILE)
        hicon_big   = LoadImageW(None, path, IMAGE_ICON, cx_big,   cy_big,   LR_LOADFROMFILE)

        if hicon_small:
            user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hicon_small)
        if hicon_big:
            user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG,   hicon_big)

        # Evita que Python libere los handles
        window._hicon_small = hicon_small  # type: ignore[attr-defined]
        window._hicon_big   = hicon_big    # type: ignore[attr-defined]
        return bool(hicon_small or hicon_big)
    except Exception:
        return False

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
            win.iconbitmap(ICON_BIG)  # también el actual
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
# === FIJAR AppUserModelID ANTES DE CUALQUIER TK/CTK ===

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

# === Helper para terminar con mensaje (sin logs ruidosos) ===
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

# Cola para manejar mensajes de log que luego se muestran en la interfaz.
log_queue = queue.Queue()

# Versión de la aplicación
version = "v1.6"

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
    if modales_abiertos and (modales_abiertos.get("config") or modales_abiertos.get("rutas")):
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
    modales_abiertos = {"config": False, "rutas": False}

    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    ventana = ctk.CTk()
    ventana.title(f"Control documental - FacturaScan {version}")
    aplicar_icono(ventana)
    ventana.after(150, lambda: aplicar_icono(ventana))


    ancho, alto = 720, 600
    x = (ventana.winfo_screenwidth() - ancho) // 2
    y = (ventana.winfo_screenheight() - alto) // 2
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")
    ventana.resizable(True, True)

    fuente_titulo = ctk.CTkFont(size=40, weight="bold")
    fuente_texto = ctk.CTkFont(family="Segoe UI", size=15)

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

    texto_log = ctk.CTkTextbox(
        ventana, width=650, height=260,
        font=("Consolas", 12), wrap="word",
        corner_radius=6, fg_color="white", text_color="black"
    )
    texto_log.pack(pady=15, padx=15)

    mensaje_espera = ctk.CTkLabel(ventana, text="", font=fuente_texto, text_color="gray")
    mensaje_espera.pack(pady=(0, 10))

    class _ConsoleRedirect:
        def __init__(self, queue_): self.queue = queue_
        def write(self, text): self.queue.put(text)
        def flush(self): pass
    sys.stdout = _ConsoleRedirect(log_queue)
    sys.stderr = _ConsoleRedirect(log_queue)

    print(f"Razón social: {variables.get('RazonSocial')}")
    print(f"RUT empresa: {variables.get('RutEmpresa')}")
    print(f"Sucursal: {variables.get('NomSucursal')}")
    print(f"Dirección: {variables.get('DirSucursal')}\n")
    print("Seleccione una opción:")

    # === Debug Chip + botones ocultos (Ctrl+F) ===
    debug_ui_visible = {"value": False}
    chip_pad = 12
    chip_width, chip_height = 120, 30
    btn_small_h = 30
    GAP = 6

    def _actualizar_chip_estilo():
        if is_debug():
            debug_chip.configure(text="DEBUG ON", fg_color="#10B981", text_color="white", hover_color="#059669")
        else:
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

    # --- Cambiar sucursal ---
    def _cambiar_config():
        try:
            modales_abiertos["config"] = True
            ventana.configure(cursor="wait")
            mensaje_espera.configure(text="⚙️ Cambiando razón social / sucursal…")
            for b in (btn_escanear, btn_procesar, btn_config, btn_rutas, debug_chip):
                b.configure(state="disabled")

            # Nuevo flujo: solo razón/sucursal desde archivo de razones
            from config_gui import cambiar_razon_sucursal
            nuevas = cambiar_razon_sucursal(variables, parent=ventana)  # modal on-top
            if not nuevas:
                return

            aplicar_nueva_config(nuevas)

            # Actualiza dict sin reasignar
            variables.clear()
            variables.update(nuevas)

            print("\n⚙️ Configuración actualizada (Razón/Sucursal):")
            print(f"Razón social: {variables.get('RazonSocial')}")
            print(f"RUT empresa: {variables.get('RutEmpresa')}")
            print(f"Sucursal: {variables.get('NomSucursal')}")
            print(f"Dirección: {variables.get('DirSucursal')}\n")

            messagebox.showinfo("Configuración", "Se actualizó la razón social y la sucursal.")
        except Exception as e:
            messagebox.showerror("Configuración", f"No se pudo actualizar la configuración:\n{e}")
        finally:
            modales_abiertos["config"] = False
            mensaje_espera.configure(text=""); ventana.configure(cursor="")
            for b in (btn_escanear, btn_procesar, btn_config, btn_rutas, debug_chip):
                b.configure(state="normal")


    btn_config = ctk.CTkButton(
        ventana, text="Cambiar sucursal",
        width=140, height=btn_small_h, corner_radius=16,
        fg_color="#E5E7EB", text_color="#111827", hover_color="#D1D5DB",
        command=_cambiar_config
    )
    btn_config.place_forget()

    # --- Cambiar rutas (solo entrada/salida) ---
    def _cambiar_rutas():
        try:
            modales_abiertos["rutas"] = True
            ventana.configure(cursor="wait")
            mensaje_espera.configure(text="📁 Cambiando rutas…")
            for b in (btn_escanear, btn_procesar, btn_config, btn_rutas, debug_chip):
                b.configure(state="disabled")

            from config_gui import actualizar_rutas
            nuevas = actualizar_rutas(variables, parent=ventana)  # <- parent para modal on-top
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
            for b in (btn_escanear, btn_procesar, btn_config, btn_rutas, debug_chip):
                b.configure(state="normal")

    btn_rutas = ctk.CTkButton(
        ventana, text="Cambiar rutas",
        width=140, height=btn_small_h, corner_radius=16,
        fg_color="#E5E7EB", text_color="#111827", hover_color="#D1D5DB",
        command=_cambiar_rutas
    )
    btn_rutas.place_forget()

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
        if modales_abiertos["config"] or modales_abiertos["rutas"]:
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

    # -------- Hilos --------
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
                        ok = f"✅ Documento procesado: {os.path.basename(resultado)}"
                        print(ok); registrar_log(ok)
                else:
                    registrar_log("⚠️ El documento no pudo ser procesado.")
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

    def iniciar_escanear(): threading.Thread(target=hilo_escanear, daemon=True).start()
    def iniciar_procesar(): threading.Thread(target=hilo_procesar, daemon=True).start()

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

    def intento_cerrar():
        if en_proceso["activo"]:
            messagebox.showwarning("Proceso en curso", "No puedes cerrar la aplicación mientras se ejecuta una tarea.")
            return
        cerrar_aplicacion(ventana, modales_abiertos)

    ventana.protocol("WM_DELETE_WINDOW", intento_cerrar)

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
        Valida_PopplerPath()
        mostrar_menu_principal()
    except Exception as e:
        show_startup_error(f"Error al iniciar FacturaScan:\n\n{e}")
