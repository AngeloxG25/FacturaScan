import os
import sys
import json
import ctypes
import customtkinter as ctk
from tkinter import filedialog, messagebox
from log_utils import registrar_log_proceso
import tkinter

# Redirigir stderr a null para silenciar errores Tcl/Tk
if os.name == "nt":
    sys.stderr = open(os.devnull, 'w')

# Evitar problemas de DPI si se usa escalado en Windows
ctk.deactivate_automatic_dpi_awareness()
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

def cargar_o_configurar():
    base_config_dir = "C:\\FacturaScan"
    os.makedirs(base_config_dir, exist_ok=True)

    # Si ya existe un archivo de configuración, cargarlo directamente
    for archivo in os.listdir(base_config_dir):
        if archivo.startswith("config_") and archivo.endswith(".txt"):
            config_path = os.path.join(base_config_dir, archivo)
            with open(config_path, "r", encoding="utf-8") as f:
                datos = {}
                for line in f:
                    if "=" in line:
                        key, val = line.strip().split("=", 1)
                        datos[key] = val.strip('"')
                return datos

    razones_sociales = {}
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    ventana = ctk.CTk()
    ventana.title("Configuración inicial")


    ventana.geometry("450x500")
    ventana.resizable(False, False)
    
    def tcl_error_handler(exc_type, exc_value, exc_traceback):
        if exc_type.__name__ == "TclError":
            return
        sys.__excepthook__(exc_type, exc_value, exc_traceback)

    ventana.report_callback_exception = tcl_error_handler

    # Centrar ventana
    ventana.update_idletasks()
    ancho_ventana = ventana.winfo_width()
    alto_ventana = ventana.winfo_height()
    pantalla_ancho = ventana.winfo_screenwidth()
    pantalla_alto = ventana.winfo_screenheight()
    x = int((pantalla_ancho - ancho_ventana) / 2)
    y = int((pantalla_alto - alto_ventana) / 2)
    ventana.geometry(f"{ancho_ventana}x{alto_ventana}+{x}+{y}")

    fuente = ctk.CTkFont(family="Segoe UI", size=12)

    # Variables
    razon_var = ctk.StringVar()
    sucursal_var = ctk.StringVar()

    ctk.CTkButton(ventana, text="Cargar Datos", command=lambda: cargar_datos(),
                  fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black").pack(pady=(20, 10))

    ctk.CTkLabel(ventana, text="Selecciona la razón social:", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(0, 5))
    razon_combo = ctk.CTkComboBox(ventana, variable=razon_var, values=[], state="disabled", font=fuente, width=350)
    razon_combo.pack()

    ctk.CTkLabel(ventana, text="Selecciona la sucursal:", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(20, 5))
    sucursal_combo = ctk.CTkComboBox(ventana, variable=sucursal_var, values=[], state="disabled", font=fuente, width=350)
    sucursal_combo.pack()

    def actualizar_sucursales(*_):
        razon = razon_var.get().strip()
        if razon and razon in razones_sociales:
            sucursales = list(razones_sociales[razon].get("sucursales", {}).keys())
            if sucursales:
                sucursal_combo.configure(values=sucursales, state="readonly")
                sucursal_var.set(sucursales[0] if len(sucursales) == 1 else "")
            else:
                sucursal_combo.configure(values=[], state="disabled")
                sucursal_var.set("")
        else:
            sucursal_combo.configure(values=[], state="disabled")
            sucursal_var.set("")

    razon_var.trace_add("write", actualizar_sucursales)

    def cargar_datos():
        nonlocal razones_sociales
        path = filedialog.askopenfilename(
            title="Selecciona archivo de datos (.json o .txt)",
            filetypes=[("Archivos JSON o TXT", "*.json *.txt")],
            initialdir="C:\\"
        )

        if not path:
            return

        try:
            if path.endswith(".json"):
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                for razon, datos in data.items():
                    if "rut" not in datos or "sucursales" not in datos:
                        raise ValueError(f"Falta 'rut' o 'sucursales' en: {razon}")
                razones_sociales = data

            elif path.endswith(".txt"):
                razones_sociales = {}
                with open(path, "r", encoding="utf-8") as f:
                    for linea in f:
                        partes = linea.strip().split(";")
                        if len(partes) >= 3:
                            razon = partes[0]
                            rut = partes[1]
                            sucursales_raw = partes[2]
                            sucursales = {}
                            for item in sucursales_raw.split("|"):
                                if "=" in item:
                                    nombre, direccion = item.split("=", 1)
                                    sucursales[nombre.strip()] = direccion.strip()
                            razones_sociales[razon] = {
                                "rut": rut,
                                "sucursales": sucursales
                            }
            else:
                raise ValueError("Formato no soportado. Usa un archivo .json o .txt")

            razones = list(razones_sociales.keys())
            razon_combo.configure(values=razones, state="readonly")
            razon_var.set("")
            sucursal_var.set("")
            sucursal_combo.configure(values=[], state="disabled")
            messagebox.showinfo("Cargado", "Datos cargados correctamente.\nSelecciona una razón social para continuar.")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el archivo:\n{e}")

    entrada_var = ctk.StringVar()
    salida_var = ctk.StringVar()

    def elegir_entrada():
        escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
        carpeta = filedialog.askdirectory(title="Selecciona carpeta de ENTRADA", initialdir=escritorio)
        if carpeta:
            entrada_var.set(carpeta)

    def elegir_salida():
        escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
        carpeta = filedialog.askdirectory(title="Selecciona carpeta de SALIDA", initialdir=escritorio)
        if carpeta:
            salida_var.set(carpeta)

    ctk.CTkLabel(ventana, text="Carpeta de entrada:", font=fuente).pack(pady=(10, 5))
    frame_entrada = ctk.CTkFrame(ventana, fg_color="transparent")
    frame_entrada.pack(pady=(0, 15))
    ctk.CTkButton(frame_entrada, text="Buscar...", command=elegir_entrada,
                  fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black", width=80).pack(side="left", padx=(0, 5))
    ctk.CTkEntry(frame_entrada, textvariable=entrada_var, width=260, font=fuente).pack(side="left")

    ctk.CTkLabel(ventana, text="Carpeta de salida:", font=fuente).pack()
    frame_salida = ctk.CTkFrame(ventana, fg_color="transparent")
    frame_salida.pack(pady=(0, 25))
    ctk.CTkButton(frame_salida, text="Buscar...", command=elegir_salida,
                  fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black", width=80).pack(side="left", padx=(0, 5))
    ctk.CTkEntry(frame_salida, textvariable=salida_var, width=260, font=fuente).pack(side="left")

    def cerrar_ventana_seguro():
        try:
            # Cancelar todos los callbacks pendientes
            for callback_id in ventana.tk.eval('after info').split():
                ventana.after_cancel(callback_id)
            ventana.destroy()
        except tkinter.TclError:
            pass
        except Exception:
            pass

    def guardar_y_cerrar():
        razon = razon_var.get().strip()
        sucursal = sucursal_var.get().strip()
        entrada = entrada_var.get().strip()
        salida = salida_var.get().strip()

        if not all([razon, sucursal, entrada, salida]):
            messagebox.showwarning("Falta información", "Completa todos los campos antes de continuar.")
            return

        ventana.config_data = {
            "RazonSocial": razon,
            "RutEmpresa": razones_sociales[razon]["rut"],
            "NomSucursal": sucursal,
            "DirSucursal": razones_sociales[razon]["sucursales"][sucursal],
            "CarEntrada": entrada,
            "CarpSalida": salida
        }
        # Asegurarse de que la ventana se cierra correctamente
        ventana.after(100, lambda: cerrar_ventana_seguro())

    ctk.CTkButton(ventana, text="Guardar configuración", command=guardar_y_cerrar,
                  fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black").pack(pady=(0, 10))

    ventana.config_data = None
    ventana.mainloop()

    if ventana.config_data:
        sucursal_nombre = ventana.config_data["NomSucursal"].lower().replace(" ", "_")
        config_filename = f"config_{sucursal_nombre}.txt"
        config_path = os.path.join(base_config_dir, config_filename)

        with open(config_path, "w", encoding="utf-8") as file:
            for key, value in ventana.config_data.items():
                file.write(f'{key}="{value}"\n')

        FILE_ATTRIBUTE_HIDDEN = 0x02
        ctypes.windll.kernel32.SetFileAttributesW(config_path, FILE_ATTRIBUTE_HIDDEN)

        return ventana.config_data

    registrar_log_proceso("❌ Configuración cancelada por el usuario en la ventana inicial.")
    exit()
    
