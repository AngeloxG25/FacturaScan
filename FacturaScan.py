from log_utils import registrar_log_proceso
import os
import sys
import queue
import ctypes
import winreg
import customtkinter as ctk
from tkinter import messagebox
from config_gui import cargar_o_configurar
from monitor_core import registrar_log, procesar_archivo, procesar_entrada_una_vez

# ================== CONFIGURACIÓN INICIAL ==================

# Se carga la configuración de la empresa/sucursal desde config_gui.py.
# Si no hay configuración válida, el programa se cierra.
variables = cargar_o_configurar()
if variables is None:
    print("❌ No se obtuvo configuración. Cerrando...")
    exit()

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

    # Leer el valor actual del PATH en registro de usuario
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment", 0, winreg.KEY_READ) as clave:
            valor_actual, _ = winreg.QueryValueEx(clave, "Path")
    except FileNotFoundError:
        valor_actual = ""

    # Normalizar rutas ya existentes en el PATH
    paths = [os.path.normcase(os.path.normpath(p.strip())) for p in valor_actual.split(";") if p.strip()]
    if ruta_normalizada not in paths:
        # Si Poppler no está, lo añadimos
        nuevo_valor = f"{valor_actual};{ruta_poppler}" if valor_actual else ruta_poppler
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment", 0, winreg.KEY_SET_VALUE) as clave:
                winreg.SetValueEx(clave, "Path", 0, winreg.REG_EXPAND_SZ, nuevo_valor)
            path_modificado = True

            # Notificar a Windows que se cambió el PATH
            ctypes.windll.user32.SendMessageTimeoutW(
                0xFFFF, 0x001A, 0, "Environment", 0x0002, 5000, None
            )
            registrar_log_proceso("🛠️ Poppler añadido al PATH. Reiniciando...")
        except PermissionError:
            registrar_log_proceso("❌ No se pudo modificar el PATH. Ejecuta como administrador.")

    # Si se modificó, reinicia el proceso con el nuevo PATH
    if path_modificado:
        os.execv(sys.executable, [sys.executable] + sys.argv)


# Muestra un cuadro de confirmación al cerrar la aplicación.
def cerrar_aplicacion(ventana):
    if messagebox.askyesno("Cerrar", "¿Deseas cerrar FacturaScan?"):
        registrar_log_proceso("⛔ FacturaScan cerrado por el usuario.")
        ventana.destroy()
        sys.exit(0)


# ================== INTERFAZ PRINCIPAL ==================

def mostrar_menu_principal():
    from PIL import Image
    import threading
    from datetime import datetime
    from scanner import escanear_y_guardar_pdf
    from log_utils import set_debug, is_debug  # bandera global de modo debug

    en_proceso = {"activo": False}  # Controla si hay un proceso en ejecución
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    # Ventana principal
    ventana = ctk.CTk()
    ventana.title(f"Control documental - FacturaScan {version}")
    ventana.iconbitmap(obtener_ruta_recurso("iconoScan.ico"))

    # Tamaño fijo centrado en pantalla
    ancho, alto = 720, 600
    x = (ventana.winfo_screenwidth() - ancho) // 2
    y = (ventana.winfo_screenheight() - alto) // 2
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")
    ventana.resizable(False, False)

    fuente_titulo = ctk.CTkFont(size=40, weight="bold")
    fuente_texto = ctk.CTkFont(family="Segoe UI", size=15)

    # Título
    ctk.CTkLabel(ventana, text="FacturaScan", font=fuente_titulo).pack(pady=15)

    # Contenedor para botones principales
    frame_botones = ctk.CTkFrame(ventana, fg_color="transparent")
    frame_botones.pack(pady=10)

    # Iconos
    icono_escaneo = ctk.CTkImage(
        light_image=Image.open(obtener_ruta_recurso("images/icono_escanear.png")),
        size=(26, 26))
    icono_carpeta = ctk.CTkImage(
        light_image=Image.open(obtener_ruta_recurso("images/icono_carpeta.png")),
        size=(26, 26))

    # Textbox donde se muestran logs de ejecución
    texto_log = ctk.CTkTextbox(
        ventana, width=650, height=260,
        font=("Consolas", 12), wrap="word",
        corner_radius=6, fg_color="white", text_color="black")
    texto_log.pack(pady=15, padx=15)

    # Mensaje de estado (ej: "Escaneando...", "Procesando...")
    mensaje_espera = ctk.CTkLabel(ventana, text="", font=fuente_texto, text_color="gray")
    mensaje_espera.pack(pady=(0, 10))

    # Redirigir stdout/stderr hacia el textbox
    class ConsoleRedirect:
        def __init__(self, queue_): self.queue = queue_
        def write(self, text): self.queue.put(text)
        def flush(self): pass
    sys.stdout = ConsoleRedirect(log_queue)
    sys.stderr = ConsoleRedirect(log_queue)

    # Mostrar datos cargados desde configuración
    print(f"Razón social: {variables.get('RazonSocial')}")
    print(f"RUT empresa: {variables.get('RutEmpresa')}")
    print(f"Sucursal: {variables.get('NomSucursal')}")
    print(f"Dirección: {variables.get('DirSucursal')}\n")
    print("Seleccione una opción:")
    registrar_log_proceso("🟢 Sistema FacturaScan Iniciado.")

    # Función que actualiza periódicamente el textbox con logs nuevos
    def actualizar_texto():
        while not log_queue.empty():
            texto_log.insert("end", log_queue.get())
            texto_log.see("end")
        ventana.after(100, actualizar_texto)


    # ========== MODO DEBUG ==========
    # Botón oculto que se activa con Ctrl+F para mostrar/ocultar el estado de debug
    debug_visible = {"show": False}
    debug_state = ctk.BooleanVar(value=is_debug())

    # Mini notificación emergente dentro de la ventana
    def toast(msg: str):
        t = ctk.CTkLabel(
            ventana, text=msg, fg_color="#000000", text_color="white",
            corner_radius=12, font=ctk.CTkFont(family="Segoe UI", size=12))
        t.place(relx=1.0, rely=1.0, x=-16, y=-16, anchor="se")
        ventana.after(1600, t.destroy)

    # Cambia el estilo del botón debug según estado
    def apply_debug_button_style():
        if debug_state.get():
            debug_btn.configure(
                text="DEBUG • ON", fg_color="#1db954",
                hover_color="#179945", text_color="white")
        else:
            debug_btn.configure(
                text="DEBUG • OFF", fg_color="#e5e5e5",
                hover_color="#d4d4d4", text_color="black")

    # Alterna estado de debug y muestra aviso
    def toggle_debug_state():
        current = not debug_state.get()
        debug_state.set(current)
        set_debug(current)
        apply_debug_button_style()
        toast("Modo debug ACTIVADO" if current else "Modo debug DESACTIVADO")

    # Botón de debug (se muestra/oculta con Ctrl+F)
    debug_btn = ctk.CTkButton(
        ventana, text="DEBUG • OFF", width=110, height=28, corner_radius=14,
        command=toggle_debug_state, fg_color="#e5e5e5",
        hover_color="#d4d4d4", text_color="black")
    apply_debug_button_style()

    def toggle_debug_widget(event=None):
        if debug_visible["show"]:
            debug_btn.place_forget()
            debug_visible["show"] = False
        else:
            debug_btn.place(relx=1.0, x=-12, y=12, anchor="ne")
            debug_visible["show"] = True
    ventana.bind_all("<Control-f>", toggle_debug_widget)


    # ================== HILOS DE TRABAJO ==================

    # Hilo para escanear un documento
    def hilo_escanear():
        try:
            en_proceso["activo"] = True
            btn_escanear.configure(state="disabled")
            btn_procesar.configure(state="disabled")
            mensaje_espera.configure(text="🔄 Escaneando...")
            ventana.configure(cursor="wait")

            # Nombre del PDF generado con timestamp
            nombre_pdf = f"DocEscaneado_{datetime.now():%Y%m%d_%H%M%S}.pdf"
            ruta = escanear_y_guardar_pdf(nombre_pdf, variables["CarEntrada"])

            if ruta:
                # Log del documento escaneado
                mensaje_escaneado = f"Documento escaneado: {os.path.basename(ruta)}"
                print(mensaje_escaneado)
                registrar_log(mensaje_escaneado)

                # Procesar el archivo escaneado
                resultado = procesar_archivo(ruta)
                if resultado:
                    if "No_Reconocidos" in resultado:
                        mensaje_fallido = f"⚠️ Documento movido a No_Reconocidos: {os.path.basename(resultado)}"
                        print(mensaje_fallido)
                        registrar_log(mensaje_fallido)
                    else:
                        mensaje_procesado = f"✅ Documento procesado: {os.path.basename(resultado)}"
                        print(mensaje_procesado)
                        registrar_log(mensaje_procesado)
                else:
                    registrar_log("⚠️ El documento no pudo ser procesado.")
            else:
                registrar_log_proceso("⚠️ No se detectó escáner.")
        finally:
            # Restaurar interfaz
            en_proceso["activo"] = False
            mensaje_espera.configure(text="")
            btn_escanear.configure(state="normal")
            btn_procesar.configure(state="normal")
            ventana.configure(cursor="")

    # Hilo para procesar todos los archivos de la carpeta de entrada
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

    # Lanzadores de hilos
    def iniciar_escanear(): threading.Thread(target=hilo_escanear, daemon=True).start()
    def iniciar_procesar(): threading.Thread(target=hilo_procesar, daemon=True).start()


    # ================== BOTONES PRINCIPALES ==================

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


    # ================== MANEJO DE CIERRE ==================

    # Evita cerrar mientras hay procesos en ejecución
    def intento_cerrar():
        if en_proceso["activo"]:
            messagebox.showwarning("Proceso en curso", "No puedes cerrar la aplicación mientras se ejecuta una tarea.")
        else:
            cerrar_aplicacion(ventana)
    ventana.protocol("WM_DELETE_WINDOW", intento_cerrar)

    # Iniciar loop principal
    actualizar_texto()
    ventana.mainloop()


# ================== EJECUCIÓN DEL PROGRAMA ==================

if __name__ == "__main__":
    # Oculta la consola negra de Windows cuando se ejecuta como .exe
    if os.name == 'nt':
        kernel32 = ctypes.WinDLL('kernel32')
        user32 = ctypes.WinDLL('user32')
        whnd = kernel32.GetConsoleWindow()
        if whnd != 0:
            user32.ShowWindow(whnd, 0)

    try:
        Valida_PopplerPath()        # Verifica que Poppler esté accesible
        mostrar_menu_principal()    # Lanza la interfaz principal
    except Exception as e:
        registrar_log_proceso(f"❌ Error al iniciar FacturaScan: {e}")
