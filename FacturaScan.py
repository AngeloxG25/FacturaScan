# FacturaScan.py (main)
import os
import sys
import queue
import ctypes
import traceback
import winreg
from tkinter import messagebox
from config_gui import cargar_o_configurar
from monitor_core import registrar_log

variables = cargar_o_configurar()
log_queue = queue.Queue()


class ConsoleRedirect:
    def __init__(self, queue):
        self.queue = queue

    def write(self, text):
        self.queue.put(text)

    def flush(self):
        pass


def obtener_ruta_recurso(ruta_relativa):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, ruta_relativa)
    return os.path.join(os.path.dirname(__file__), ruta_relativa)


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
        nuevo_valor = valor_actual + ";" + ruta_poppler if valor_actual else ruta_poppler
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment", 0, winreg.KEY_SET_VALUE) as clave:
                winreg.SetValueEx(clave, "Path", 0, winreg.REG_EXPAND_SZ, nuevo_valor)
            path_modificado = True

            HWND_BROADCAST = 0xFFFF
            WM_SETTINGCHANGE = 0x001A
            SMTO_ABORTIFHUNG = 0x0002
            ctypes.windll.user32.SendMessageTimeoutW(HWND_BROADCAST, WM_SETTINGCHANGE, 0,
                                                      "Environment", SMTO_ABORTIFHUNG, 5000, None)
            print("🛠️ Poppler añadido al PATH del usuario. Reiniciando FacturaScan...")
        except PermissionError:
            print("❌ No se pudo modificar el PATH. Ejecuta como administrador.")

    if path_modificado:
        ruta_exe = sys.executable
        if ruta_exe.lower().endswith(".exe"):
            os.execv(ruta_exe, [ruta_exe] + sys.argv)
        else:
            os.execl(sys.executable, sys.executable, *sys.argv)


def cerrar_aplicacion(ventana):
    if messagebox.askyesno("Cerrar", "¿Deseas cerrar FacturaScan?"):
        registrar_log("\n⛔ Programa cerrado por el usuario.")
        ventana.destroy()
        sys.exit(0)


def mostrar_menu_principal():
    import tkinter as tk
    from tkinter import ttk, scrolledtext
    from PIL import Image, ImageTk
    import threading
    from datetime import datetime
    from scanner import escanear_y_guardar_pdf
    from monitor_core import procesar_archivo, registrar_log, procesar_entrada_una_vez

    ventana = tk.Tk()
    ventana.title("Control documental - FacturaScan")

    ancho = 600
    alto = 480
    x = (ventana.winfo_screenwidth() // 2) - (ancho // 2)
    y = (ventana.winfo_screenheight() // 2) - (alto // 2)
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")
    ventana.resizable(False, False)

    style = ttk.Style()
    style.theme_use("default")
    style.configure("TButton", background="#6f7e8e", foreground="white", font=("Arial", 11, "bold"))
    style.map("TButton", background=[("active", "#34495e")])
    style.configure("TLabel", background="#f8f9fa", foreground="#2c3e50", font=("Arial", 12))

    ttk.Label(ventana, text="FacturaScan", font=("Arial", 16, "bold")).pack(pady=10)

    marco_botones = ttk.Frame(ventana)
    marco_botones.pack(pady=10)

    icono_escaneo = Image.open(obtener_ruta_recurso("images/icono_escanear.png")).resize((24, 24), Image.Resampling.LANCZOS)
    icono_carpeta = Image.open(obtener_ruta_recurso("images/icono_carpeta.png")).resize((24, 24), Image.Resampling.LANCZOS)

    img_escaneo = ImageTk.PhotoImage(icono_escaneo)
    img_carpeta = ImageTk.PhotoImage(icono_carpeta)

    ventana.iconos = [img_escaneo, img_carpeta]  # evitar recolección de basura

    texto_log = scrolledtext.ScrolledText(
        ventana, wrap=tk.WORD, font=("Consolas", 10),
        background="#ffffff", foreground="#000000",
        insertbackground="white", borderwidth=1,
        relief="solid", height=12
    )
    texto_log.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

    sys.stdout = ConsoleRedirect(log_queue)
    sys.stderr = ConsoleRedirect(log_queue)

    print(f"Razón social: {variables.get('RazonSocial', 'desconocida')}")
    print(f"RUT empresa: {variables.get('RutEmpresa', 'desconocido')}")
    print(f"Sucursal: {variables.get('NomSucursal', 'sucursal_default')}")
    print(f"Dirección: {variables.get('DirSucursal', 'direccion_no_definida')}\n")
    print('SELECCIONE UNA OPCIÓN:')
    registrar_log("🟢 Monitor iniciado correctamente.")

    def actualizar_texto():
        while not log_queue.empty():
            texto = log_queue.get()
            texto_log.insert(tk.END, texto)
            texto_log.see(tk.END)
        ventana.after(100, actualizar_texto)

    def hilo_escanear():
        nombre_pdf = "escaneo_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".pdf"
        # ruta del archivo escaneado
        ruta = escanear_y_guardar_pdf(nombre_pdf, variables["CarEntrada"], r"C:\FacturaScan\debug")
        if ruta:
            print(f"📥 Documento escaneado: {os.path.basename(ruta)}")
            registrar_log(f"📥 Documento escaneado: {os.path.basename(ruta)}")
            procesar_archivo(ruta)
        else:
            registrar_log("⚠️ No se detectó escáner.")

    def hilo_procesar():
        procesar_entrada_una_vez()

    def iniciar_escanear():
        threading.Thread(target=hilo_escanear, daemon=True).start()

    def iniciar_procesar():
        threading.Thread(target=hilo_procesar, daemon=True).start()

    ttk.Button(
        marco_botones, text="  ESCANEAR DOCUMENTO", image=img_escaneo,
        compound="left", width=30, command=iniciar_escanear
    ).pack(pady=5)

    ttk.Button(
        marco_botones, text="  PROCESAR CARPETA", image=img_carpeta,
        compound="left", width=30, command=iniciar_procesar
    ).pack(pady=5)

    ventana.protocol("WM_DELETE_WINDOW", lambda: cerrar_aplicacion(ventana))
    actualizar_texto()
    ventana.mainloop()


if __name__ == "__main__":
    try:
        Valida_PopplerPath()
        mostrar_menu_principal()
    except Exception:
        print("\n❌ Error inesperado:")
        traceback.print_exc()
        try:
            input("\nPresiona ENTER para cerrar...")
        except:
            pass
        sys.exit(1)
