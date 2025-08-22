import os
import sys
import json
import ctypes
import customtkinter as ctk
from tkinter import filedialog, messagebox

if os.name == "nt":
    sys.stderr = open(os.devnull, 'w')

import contextlib

@contextlib.contextmanager
def ocultar_stderr():
    original_stderr = sys.stderr
    sys.stderr = open(os.devnull, 'w')
    try:
        yield
    finally:
        sys.stderr = original_stderr

ctk.deactivate_automatic_dpi_awareness()
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

base_config_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
os.makedirs(base_config_dir, exist_ok=True)

def limpiar_callbacks(ventana):
    if ventana and ventana.winfo_exists():
        try:
            for callback_id in ventana.tk.eval('after info').split():
                try:
                    ventana.after_cancel(callback_id)
                except:
                    pass
        except:
            pass

def cargar_o_configurar():
    archivos_config = [f for f in os.listdir(base_config_dir) if f.startswith("config_") and f.endswith(".txt")]
    if len(archivos_config) == 1:
        config_path = os.path.join(base_config_dir, archivos_config[0])
        with open(config_path, "r", encoding="utf-8") as f:
            datos = {}
            for line in f:
                if "=" in line:
                    key, val = line.strip().split("=", 1)
                    datos[key] = val.strip('"')
            return datos

    razones_sociales = {}

    def mostrar_configuracion_completa():
        ventana = ctk.CTk()
        ventana.title("Configuración de FacturaScan")
        ventana.resizable(False, False)
        ancho, alto = 500, 530
        x = (ventana.winfo_screenwidth() // 2) - (ancho // 2)
        y = (ventana.winfo_screenheight() // 2) - (alto // 2)
        ventana.geometry(f"{ancho}x{alto}+{x}+{y}")
        ventana.protocol("WM_DELETE_WINDOW", lambda: [limpiar_callbacks(ventana), ventana.quit()])

        fuente = ctk.CTkFont(family="Segoe UI", size=15)

        razon_var = ctk.StringVar()
        sucursal_var = ctk.StringVar()
        entrada_var = ctk.StringVar()
        salida_var = ctk.StringVar()

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
                "CarpSalida": salida,
                "CarpSalidaUsoAtm": ""}

            sucursal_nombre = sucursal.lower().replace(" ", "_")
            config_filename = f"config_{sucursal_nombre}.txt"
            config_path = os.path.join(base_config_dir, config_filename)

            with open(config_path, "w", encoding="utf-8") as file:
                for key, value in ventana.config_data.items():
                    file.write(f'{key}="{value}"\n')

            FILE_ATTRIBUTE_HIDDEN = 0x02
            ctypes.windll.kernel32.SetFileAttributesW(config_path, FILE_ATTRIBUTE_HIDDEN)

            with open(config_path, "r", encoding="utf-8") as f:
                datos = {}
                for line in f:
                    if "=" in line:
                        key, val = line.strip().split("=", 1)
                        datos[key] = val.strip('"')

            ventana.resultado = datos
            ventana.quit()

        frame = ctk.CTkFrame(ventana, fg_color="transparent")
        frame.pack(pady=20)

        ctk.CTkLabel(frame, text="Selecciona la razón social:", font=fuente).pack(pady=(10, 5))
        razon_combo = ctk.CTkComboBox(frame, variable=razon_var, values=list(razones_sociales.keys()),
                                    state="readonly", font=fuente, width=350)
        razon_combo.pack()

        ctk.CTkLabel(frame, text="Selecciona la sucursal:", font=fuente).pack(pady=(20, 5))
        sucursal_combo = ctk.CTkComboBox(frame, variable=sucursal_var, values=[],
                                        state="disabled", font=fuente, width=350)
        sucursal_combo.pack()

        razon_var.trace_add("write", actualizar_sucursales)

        ctk.CTkLabel(frame, text="Carpeta de entrada:", font=fuente).pack(pady=(15, 5))
        entrada_frame = ctk.CTkFrame(frame, fg_color="transparent")
        entrada_frame.pack(pady=(0, 10))
        ctk.CTkButton(entrada_frame, text="Buscar...", command=elegir_entrada,
                    fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black").pack(side="left", padx=(0, 5))
        ctk.CTkEntry(entrada_frame, textvariable=entrada_var, width=260, font=fuente).pack(side="left")

        ctk.CTkLabel(frame, text="Carpeta de salida:", font=fuente).pack(pady=(15, 5))
        salida_frame = ctk.CTkFrame(frame, fg_color="transparent")
        salida_frame.pack(pady=(0, 20))
        ctk.CTkButton(salida_frame, text="Buscar...", command=elegir_salida,
                    fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black").pack(side="left", padx=(0, 5))
        ctk.CTkEntry(salida_frame, textvariable=salida_var, width=260, font=fuente).pack(side="left")

        ctk.CTkButton(ventana, text="Guardar configuración", command=guardar_y_cerrar,
                    width=250, height=40, font=fuente,
                    fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black").pack(pady=(0, 20))

        ventana.resultado = None
        with ocultar_stderr():
            ventana.mainloop()

        resultado = getattr(ventana, "resultado", None)
        ventana.destroy()
        return resultado

    ventana_inicial = ctk.CTk()
    ventana_inicial.title("Configuración inicial")
    ventana_inicial.geometry("400x200")
    ventana_inicial.resizable(False, False)
    ventana_inicial.update_idletasks()
    x = (ventana_inicial.winfo_screenwidth() // 2) - 200
    y = (ventana_inicial.winfo_screenheight() // 2) - 100
    ventana_inicial.geometry(f"400x200+{x}+{y}")

    def cerrar_configuracion():
        if messagebox.askyesno("Salir", "¿Deseas cerrar FacturaScan sin configurar?"):
            limpiar_callbacks(ventana_inicial)
            sys.exit(0)

    ventana_inicial.protocol("WM_DELETE_WINDOW", cerrar_configuracion)

    ctk.CTkLabel(ventana_inicial, text="Bienvenido a FacturaScan", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=20)
    ctk.CTkLabel(ventana_inicial, text="Cargue los datos de empresas para comenzar.").pack(pady=(0, 15))

    resultado = {}

    def cargar_datos_y_continuar():
        nonlocal razones_sociales, resultado
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
                                "sucursales": sucursales}
            else:
                raise ValueError("Formato no soportado. Usa un archivo .json o .txt")

            ventana_inicial.destroy()
            resultado = mostrar_configuracion_completa()

        except Exception as e:
            messagebox.showerror("Error", f"Datos incorrectos:\n{e}")

    ctk.CTkButton(ventana_inicial, text="Cargar Datos", command=cargar_datos_y_continuar, fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black").pack(pady=10)
    ventana_inicial.mainloop()

    if resultado:
        return resultado

    if messagebox.askyesno("Cancelar", "No se completó la configuración.\n¿Deseas salir de FacturaScan?"):
        sys.exit(0)  # Cierre limpio, no error
    else:
        # Volver a mostrar la ventana inicial si el usuario se arrepiente
        return cargar_o_configurar()
