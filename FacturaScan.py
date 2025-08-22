import sys, os, ctypes
import queue
import winreg
import customtkinter as ctk
from tkinter import messagebox
from log_utils import set_debug, is_debug

# ---- Mostrar error de inicio directamente en popup ----
def show_startup_error(msg: str):
    try:
        import tkinter as tk
        from tkinter import messagebox as mb
        r = tk.Tk(); r.withdraw()
        mb.showerror("Error al iniciar FacturaScan", msg)
        r.destroy()
    except Exception:
        # Último recurso si ni Tk está disponible
        print(msg)

# ==== Instancia única (Windows) ====
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
            show_startup_error(
                "FacturaScan ya está en ejecución."
            )
        except Exception:
            pass
        sys.exit(0)

    atexit.register(lambda: kernel32.CloseHandle(handle))

# Ejecutar el guardián de instancia única cuanto antes:
_ensure_single_instance()


# ----------------- Imports críticos -----------------
try:
    from config_gui import cargar_o_configurar
    from monitor_core import registrar_log, procesar_archivo, procesar_entrada_una_vez
except Exception as e:
    show_startup_error(f"No se pudo importar un módulo crítico:\n\n{e}")
    sys.exit(1)

# (Opcionales; si faltan, solo avisamos en un popup y seguimos)
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

# ================== CONFIGURACIÓN INICIAL ==================

variables = None
try:
    variables = cargar_o_configurar()
except Exception as e:
    fatal("CONFIG", e)

if variables is None:
    fatal("CONFIG", Exception("No se obtuvo configuración"))

# Cola para manejar mensajes de log que luego se muestran en la interfaz.
log_queue = queue.Queue()

# Versión de la aplicación
version = "v1.5"

# ================== REDIRECCIÓN DE CONSOLA ==================

# Clase que redirige la salida estándar (print) a una cola,
# permitiendo mostrar logs en la interfaz gráfica (Textbox).
class ConsoleRedirect:
    def __init__(self, queue): 
        self.queue = queue
    def write(self, text): 
        self.queue.put(text)
    def flush(self): 
        pass

# ================== UTILIDADES ==================

# Devuelve la ruta absoluta de un recurso (ícono, imágenes, etc).
# Si la aplicación está empaquetada con PyInstaller (sys._MEIPASS),
# busca los recursos dentro de ese directorio temporal.
def obtener_ruta_recurso(ruta_relativa):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, ruta_relativa)
    return os.path.join(os.path.dirname(__file__), ruta_relativa)

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
        os.execv(sys.executable, [sys.executable] + sys.argv)

def cerrar_aplicacion(ventana):
    if messagebox.askyesno("Cerrar", "¿Deseas cerrar FacturaScan?"):
        registrar_log('FacturaScan cerrado por el usuario')
        ventana.destroy()
        sys.exit(0)


# ================== INTERFAZ PRINCIPAL ==================
def mostrar_menu_principal():
    from PIL import Image
    import threading
    from datetime import datetime
    from scanner import escanear_y_guardar_pdf
    from log_utils import set_debug, is_debug  # asegúrate de tener este import arriba también

    registrar_log("FacturaScan iniciado correctamente")

    en_proceso = {"activo": False}
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    ventana = ctk.CTk()
    ventana.title(f"Control documental - FacturaScan {version}")
    ventana.iconbitmap(obtener_ruta_recurso("iconoScan.ico"))

    ancho, alto = 720, 600
    x = (ventana.winfo_screenwidth() - ancho) // 2
    y = (ventana.winfo_screenheight() - alto) // 2
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")
    ventana.resizable(False, False)

    fuente_titulo = ctk.CTkFont(size=40, weight="bold")
    fuente_texto = ctk.CTkFont(family="Segoe UI", size=15)

    ctk.CTkLabel(ventana, text="FacturaScan", font=fuente_titulo).pack(pady=15)
    frame_botones = ctk.CTkFrame(ventana, fg_color="transparent")
    frame_botones.pack(pady=10)

    icono_escaneo = ctk.CTkImage(
        light_image=Image.open(obtener_ruta_recurso("images/icono_escanear.png")),
        size=(26, 26))
    icono_carpeta = ctk.CTkImage(
        light_image=Image.open(obtener_ruta_recurso("images/icono_carpeta.png")),
        size=(26, 26))

    texto_log = ctk.CTkTextbox(
        ventana, width=650, height=260,
        font=("Consolas", 12), wrap="word",
        corner_radius=6, fg_color="white", text_color="black")
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

    # === Debug Chip (superior derecha) ===
    debug_ui_visible = {"value": False}  # visibilidad del chip
    chip_pad = 12
    chip_width, chip_height = 120, 30

    def _actualizar_chip_estilo():
        if is_debug():
            debug_chip.configure(
                text="DEBUG ON",
                fg_color="#10B981", text_color="white", hover_color="#059669"
            )
        else:
            debug_chip.configure(
                text="DEBUG OFF",
                fg_color="#E5E7EB", text_color="#111827", hover_color="#D1D5DB"
            )

    def _toggle_debug_state():
        nuevo = not is_debug()
        set_debug(nuevo)
        _actualizar_chip_estilo()
        # print(f"🔧 Modo debug {'ACTIVADO' if nuevo else 'DESACTIVADO'}")

    debug_chip = ctk.CTkButton(
        ventana, text="DEBUG OFF",
        width=chip_width, height=chip_height, corner_radius=16,
        fg_color="#E5E7EB", text_color="#111827", hover_color="#D1D5DB",
        command=_toggle_debug_state
    )
    # Oculto inicialmente; se muestra con Ctrl+F

    def _mostrar_chip():
        # ⬅️ Superior derecha
        debug_chip.place(relx=1.0, rely=0.0, x=-chip_pad, y=chip_pad, anchor="ne")
        debug_ui_visible["value"] = True
        _actualizar_chip_estilo()

    def _ocultar_chip():
        debug_chip.place_forget()
        debug_ui_visible["value"] = False

    def _toggle_chip_visibility(event=None):
        if debug_ui_visible["value"]:
            _ocultar_chip()
        else:
            _mostrar_chip()

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
        command=iniciar_escanear)
    btn_escanear.pack(pady=6)

    btn_procesar = ctk.CTkButton(
        frame_botones, text="PROCESAR CARPETA", image=icono_carpeta,
        compound="left", width=300, height=60, font=fuente_texto,
        fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black",
        command=iniciar_procesar)
    btn_procesar.pack(pady=6)

    def intento_cerrar():
        if en_proceso["activo"]:
            messagebox.showwarning("Proceso en curso", "No puedes cerrar la aplicación mientras se ejecuta una tarea.")
        else:
            cerrar_aplicacion(ventana)
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

        # <- elimina el bloque extra de Tk/messagebox aquí
