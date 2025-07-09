import os
import sys
import queue
import ctypes
import winreg
import customtkinter as ctk
from tkinter import messagebox
from config_gui import cargar_o_configurar
from monitor_core import registrar_log, procesar_archivo, procesar_entrada_una_vez
from log_utils import registrar_log_proceso

# Obtener configuración
variables = cargar_o_configurar()
if variables is None:
    print("❌ No se obtuvo configuración. Cerrando...")
    exit()

log_queue = queue.Queue()
version = "v1.2"

# Redirigir salida a consola para el GUI
class ConsoleRedirect:
    def __init__(self, queue): self.queue = queue
    def write(self, text): self.queue.put(text)
    def flush(self): pass

# Ruta de recursos como íconos
def obtener_ruta_recurso(ruta_relativa):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, ruta_relativa)
    return os.path.join(os.path.dirname(__file__), ruta_relativa)

# Validar y agregar Poppler al PATH
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

            ctypes.windll.user32.SendMessageTimeoutW(
                0xFFFF, 0x001A, 0, "Environment", 0x0002, 5000, None
            )
            registrar_log_proceso("🛠️ Poppler añadido al PATH. Reiniciando...")
        except PermissionError:
            registrar_log_proceso("❌ No se pudo modificar el PATH. Ejecuta como administrador.")

    if path_modificado:
        os.execv(sys.executable, [sys.executable] + sys.argv)

# Confirmación para cerrar
def cerrar_aplicacion(ventana):
    if messagebox.askyesno("Cerrar", "¿Deseas cerrar FacturaScan?"):
        registrar_log_proceso("⛔ FacturaScan cerrado por el usuario.")
        ventana.destroy()
        sys.exit(0)

# Menú principal
def mostrar_menu_principal():
    from PIL import Image
    import threading
    from datetime import datetime
    from scanner import escanear_y_guardar_pdf

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

    icono_escaneo = ctk.CTkImage(light_image=Image.open(obtener_ruta_recurso("images/icono_escanear.png")), size=(26, 26))
    icono_carpeta = ctk.CTkImage(light_image=Image.open(obtener_ruta_recurso("images/icono_carpeta.png")), size=(26, 26))

    texto_log = ctk.CTkTextbox(
        ventana, width=650, height=260,
        font=("Consolas", 12), wrap="word",
        corner_radius=6, fg_color="white", text_color="black"
    )
    texto_log.pack(pady=15, padx=15)

    mensaje_espera = ctk.CTkLabel(ventana, text="", font=fuente_texto, text_color="gray")
    mensaje_espera.pack(pady=(0, 10))

    sys.stdout = ConsoleRedirect(log_queue)
    sys.stderr = ConsoleRedirect(log_queue)

    print(f"Razón social: {variables.get('RazonSocial')}")
    print(f"RUT empresa: {variables.get('RutEmpresa')}")
    print(f"Sucursal: {variables.get('NomSucursal')}")
    print(f"Dirección: {variables.get('DirSucursal')}\n")
    print("Seleccione una opción:")
    registrar_log_proceso("🟢 Sistema FacturaScan Iniciado.")

    def actualizar_texto():
        while not log_queue.empty():
            texto_log.insert("end", log_queue.get())
            texto_log.see("end")
        ventana.after(100, actualizar_texto)

    def hilo_escanear():
        try:
            en_proceso["activo"] = True
            btn_escanear.configure(state="disabled")
            btn_procesar.configure(state="disabled")
            mensaje_espera.configure(text="🔄 Escaneando...")

            ventana.configure(cursor="wait")

            nombre_pdf = f"escaneo_{datetime.now():%Y%m%d_%H%M%S}.pdf"
            ruta = escanear_y_guardar_pdf(nombre_pdf, variables["CarEntrada"])
            if ruta:
                mensaje_escaneado = f"📥 Documento escaneado: {os.path.basename(ruta)}"
                print(mensaje_escaneado)
                registrar_log(mensaje_escaneado)

                resultado = procesar_archivo(ruta)
                if resultado:
                    mensaje_procesado = f"✅ Documento procesado: {resultado}"
                    print(mensaje_procesado)
                    registrar_log(mensaje_procesado)
            else:
                registrar_log_proceso("⚠️ No se detectó escáner.")
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

    btn_escanear = ctk.CTkButton(frame_botones, text="ESCANEAR DOCUMENTO", image=icono_escaneo,
        compound="left", width=300, height=60, font=fuente_texto,
        fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black", command=iniciar_escanear)
    btn_escanear.pack(pady=6)

    btn_procesar = ctk.CTkButton(frame_botones, text="PROCESAR CARPETA", image=icono_carpeta,
        compound="left", width=300, height=60, font=fuente_texto,
        fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black", command=iniciar_procesar)
    btn_procesar.pack(pady=6)

    def intento_cerrar():
        if en_proceso["activo"]:
            messagebox.showwarning("Proceso en curso", "No puedes cerrar la aplicación mientras se ejecuta una tarea.")
        else:
            cerrar_aplicacion(ventana)

    ventana.protocol("WM_DELETE_WINDOW", intento_cerrar)
    actualizar_texto()
    ventana.mainloop()

# Ocultar consola negra al inicio en Windows
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
        registrar_log_proceso(f"❌ Error al iniciar FacturaScan: {e}")
