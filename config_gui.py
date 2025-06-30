import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import glob

def cargar_o_configurar():
    if os.name == "nt":
        base_config_dir = "C:\\FacturaScan"
    else:
        base_config_dir = os.path.join(os.path.expanduser("~/.config"), "FacturaScan")

    os.makedirs(base_config_dir, exist_ok=True)

    # Buscar archivos de configuración existentes
    config_files = glob.glob(os.path.join(base_config_dir, "config_*.txt"))
    if config_files:
        config_path = config_files[0]
        variables = {}
        with open(config_path, "r", encoding="utf-8") as file:
            for line in file:
                if '=' in line:
                    key, value = line.strip().split('=', 1)
                    variables[key.strip()] = value.strip().strip('"').strip("'")
        return variables

    # Datos predefinidos
    razones_sociales = {
        "COMERCIAL TEBA SPA": {
            "rut": "76.466.343-8",
            "sucursales": {
                "Gran Avenida": "Avenida Jose Miguel Carrera #13365, San Bernardo",
                "Lo Valledor": "Avenida General Velázquez #3409, Cerrillos",
                "Rancagua": "Avenida Federico Koke #250, Rancagua",
                "JJ Perez": "Avenida Jose Joaquin Perez #6142, Cerro Navia",
                "Pinto": "Pinto 8, San Bernardo",
                "Lo Blanco": "Avenida Lo Blanco #2561, La Pintana"
            }
        },
        "COMERCIAL NABEK LIMITADA": {
            "rut": "78.767.200-0",
            "sucursales": {
                "Nabek Lo Blanco": "Avenida Lo Blanco #2561, La Pintana"
            }
        },
        "RECURSOS HUMANOS A TIEMPO SPA": {
            "rut": "77.076.847-0",
            "sucursales": {
                "Lo Blanco": "Avenida Lo Blanco #2561, La Pintana"
            }
        },
        "TRANSPORTE LS SPA": {
            "rut": "76.704.181-0",
            "sucursales": {
                "Lo Blanco": "Avenida Lo Blanco #2561, La Pintana"
            }
        },
        "INMOVILIARIA NABEK SPA": {
            "rut": "77.963.143-5",
            "sucursales": {
                "Lo Blanco": "Avenida Lo Blanco #2561, La Pintana"
            }
        },
        "TEAM WORK NABEK SPA": {
            "rut": "78.075.668-3",
            "sucursales": {
                "Lo Blanco": "Avenida Lo Blanco #2561, La Pintana"
            }
        }
    }

    def actualizar_sucursales(event):
        razon = razon_combo.get()
        sucursales = list(razones_sociales[razon]["sucursales"].keys())
        sucursal_combo["values"] = sucursales
        sucursal_combo.set("")

    def guardar_y_cerrar():
        razon = razon_combo.get()
        sucursal = sucursal_combo.get()
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
        ventana.destroy()

    # Interfaz gráfica
    ventana = tk.Tk()
    ventana.title("Configuración inicial")
    ventana.geometry("620x480")
    ventana.resizable(False, False)

    x = (ventana.winfo_screenwidth() // 2) - 310
    y = (ventana.winfo_screenheight() // 2) - 240
    ventana.geometry(f"+{x}+{y}")

    fuente = ('Segoe UI', 11)

    ttk.Label(ventana, text="Selecciona la razón social:", font=('Segoe UI', 12, 'bold')).pack(pady=(20, 5))
    razon_combo = ttk.Combobox(ventana, values=list(razones_sociales.keys()), state="readonly", font=fuente, width=45)
    razon_combo.pack()

    ttk.Label(ventana, text="Selecciona la sucursal:", font=('Segoe UI', 12, 'bold')).pack(pady=(20, 5))
    sucursal_combo = ttk.Combobox(ventana, state="readonly", font=fuente, width=45)
    sucursal_combo.pack()

    entrada_var = tk.StringVar()
    salida_var = tk.StringVar()

    def elegir_entrada():
        carpeta = filedialog.askdirectory(title="Selecciona carpeta de ENTRADA", initialdir=os.path.expanduser("~"))
        if carpeta:
            entrada_var.set(carpeta)

    def elegir_salida():
        carpeta = filedialog.askdirectory(title="Selecciona carpeta de SALIDA", initialdir=os.path.expanduser("~"))
        if carpeta:
            salida_var.set(carpeta)

    ttk.Label(ventana, text="Carpeta de entrada:", font=('Segoe UI', 11)).pack(pady=(25, 5))
    ttk.Entry(ventana, textvariable=entrada_var, width=58, font=fuente).pack()
    ttk.Button(ventana, text="Buscar...", command=elegir_entrada).pack(pady=(2, 15))

    ttk.Label(ventana, text="Carpeta de salida:", font=('Segoe UI', 11)).pack()
    ttk.Entry(ventana, textvariable=salida_var, width=58, font=fuente).pack()
    ttk.Button(ventana, text="Buscar...", command=elegir_salida).pack(pady=(2, 25))

    ttk.Button(ventana, text="Guardar configuración", command=guardar_y_cerrar).pack(pady=(0, 10))

    razon_combo.bind("<<ComboboxSelected>>", actualizar_sucursales)
    ventana.config_data = None
    ventana.mainloop()

    if ventana.config_data:
        sucursal_nombre = ventana.config_data["NomSucursal"].lower().replace(" ", "_")
        config_filename = f"config_{sucursal_nombre}.txt"
        config_path = os.path.join(base_config_dir, config_filename)

        # Guardar archivo
        with open(config_path, "w", encoding="utf-8") as file:
            for key, value in ventana.config_data.items():
                file.write(f'{key}="{value}"\n')

        # Ocultar en Windows
        if os.name == "nt":
            import ctypes
            FILE_ATTRIBUTE_HIDDEN = 0x02
            ctypes.windll.kernel32.SetFileAttributesW(config_path, FILE_ATTRIBUTE_HIDDEN)

        return ventana.config_data

    print("❌ Configuración cancelada.")
    exit()
