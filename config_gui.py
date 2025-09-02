import os
import sys
import json
import ctypes
import contextlib
import customtkinter as ctk
from tkinter import filedialog, messagebox

# Silenciar stderr de algunos drivers en Windows
if os.name == "nt":
    sys.stderr = open(os.devnull, 'w')
import platform
import ctypes
from ctypes import wintypes

# === Ruta dinámica a /assets ===
if getattr(sys, "frozen", False):  # Si está compilado (exe con Nuitka/PyInstaller)
    BASE_DIR = os.path.dirname(sys.executable)
else:  # Si está en modo script normal
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ASSETS_DIR = os.path.join(BASE_DIR, "assets")

ICON_BIG   = os.path.join(ASSETS_DIR, "iconoScan.ico")
ICON_SMALL = os.path.join(ASSETS_DIR, "iconoScan16.ico")

def asset_path(nombre: str) -> str:
    """Devuelve la ruta absoluta dentro de /assets."""
    return os.path.join(ASSETS_DIR, nombre)


def aplicar_icono(win) -> bool:
    """Fija el icono .ico (small/big) y lo deja como default para Toplevels.
    Incluye re-aplicación tras idle para evitar que CTk lo pise.
    """
    if not os.path.exists(ICON_BIG):
        return False

    ok = False
    try:
        # default → nuevos Toplevels heredan
        win.iconbitmap(default=ICON_BIG)
        win.iconbitmap(ICON_BIG)
        ok = True
    except Exception:
        pass

    # Forzar small/big con WinAPI
    try:
        user32 = ctypes.windll.user32
        IMAGE_ICON, LR_LOADFROMFILE = 1, 0x0010
        WM_SETICON, ICON_SMALL_W, ICON_BIG_W = 0x0080, 0, 1
        SM_CXSMICON, SM_CYSMICON = 49, 50
        SM_CXICON,  SM_CYICON  = 11, 12

        hwnd = win.winfo_id()

        LoadImageW = user32.LoadImageW
        LoadImageW.restype = wintypes.HANDLE

        # small
        if os.path.exists(ICON_SMALL):
            h_small = LoadImageW(None, ICON_SMALL, IMAGE_ICON,
                                 user32.GetSystemMetrics(SM_CXSMICON),
                                 user32.GetSystemMetrics(SM_CYSMICON),
                                 LR_LOADFROMFILE)
        else:
            h_small = LoadImageW(None, ICON_BIG, IMAGE_ICON,
                                 user32.GetSystemMetrics(SM_CXSMICON),
                                 user32.GetSystemMetrics(SM_CYSMICON),
                                 LR_LOADFROMFILE)

        # big
        h_big = LoadImageW(None, ICON_BIG, IMAGE_ICON,
                           user32.GetSystemMetrics(SM_CXICON),
                           user32.GetSystemMetrics(SM_CYICON),
                           LR_LOADFROMFILE)

        if h_small:
            user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL_W, h_small)
            win._hicon_small = h_small
            ok = True
        if h_big:
            user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG_W, h_big)
            win._hicon_big = h_big
            ok = True
    except Exception:
        pass

    # Re-aplicar tras idle
    win.after_idle(lambda: win.iconbitmap(default=ICON_BIG))
    return ok

# ===== Utilidades básicas =====
@contextlib.contextmanager
def ocultar_stderr():
    original_stderr = sys.stderr
    sys.stderr = open(os.devnull, 'w')
    try:
        yield
    finally:
        sys.stderr = original_stderr

def limpiar_callbacks(ventana):
    """Cancela callbacks pendientes 'after' para evitar warnings al destruir ventanas."""
    if ventana and ventana.winfo_exists():
        try:
            for callback_id in ventana.tk.eval('after info').split():
                try:
                    ventana.after_cancel(callback_id)
                except:
                    pass
        except:
            pass

def _parse_config_txt(path: str) -> dict:
    datos = {}
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            if "=" in line:
                key, val = line.strip().split("=", 1)
                datos[key] = val.strip('"')
    return datos

def _askdir_above(win, title, initialdir):
    """Muestra un askdirectory por encima del modal 'win'."""
    was_top = False
    try:
        was_top = bool(win.attributes("-topmost"))
    except Exception:
        pass
    try:
        if was_top:
            win.attributes("-topmost", False)
        win.update_idletasks()
        carpeta = filedialog.askdirectory(title=title, initialdir=initialdir, parent=win)
        return carpeta
    finally:
        try:
            if was_top:
                win.attributes("-topmost", True)
                win.lift()
                win.focus_force()
        except Exception:
            pass

def _askopen_above(win, **kwargs):
    """Muestra un askopenfilename por encima del modal 'win'."""
    was_top = False
    try:
        was_top = bool(win.attributes("-topmost"))
    except Exception:
        pass
    try:
        if was_top:
            win.attributes("-topmost", False)
        win.update_idletasks()
        kwargs.setdefault("parent", win)
        ruta = filedialog.askopenfilename(**kwargs)
        return ruta
    finally:
        try:
            if was_top:
                win.attributes("-topmost", True)
                win.lift()
                win.focus_force()
        except Exception:
            pass

# ===== Apariencia / DPI =====
ctk.deactivate_automatic_dpi_awareness()
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# ===== Directorio de trabajo =====
base_config_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
os.makedirs(base_config_dir, exist_ok=True)

# ===== Puntero de configuración activa y purga opcional =====
ACTIVE_POINTER = os.path.join(base_config_dir, "config_actual.txt")
PURGE_OTHERS  = True   # True => al guardar, borra cualquier otro config_*.txt

# Atributos de Windows
FILE_ATTRIBUTE_HIDDEN = 0x02
FILE_ATTRIBUTE_NORMAL = 0x80

def _win_make_writable(path: str):
    """Quita atributos que puedan bloquear escritura y asegura permisos."""
    if not os.path.exists(path):
        return
    try:
        if os.name == "nt":
            ctypes.windll.kernel32.SetFileAttributesW(path, FILE_ATTRIBUTE_NORMAL)
    except Exception:
        pass
    try:
        os.chmod(path, 0o666)
    except Exception:
        pass

def _safe_write_text(path: str, content: str, make_hidden: bool = False):
    """Escritura atómica: a .tmp y luego replace. Previo, fuerza writable."""
    try:
        _win_make_writable(path)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            fh.write(content)
        os.replace(tmp, path)  # atómico en Windows
        if make_hidden and os.name == "nt":
            try:
                ctypes.windll.kernel32.SetFileAttributesW(path, FILE_ATTRIBUTE_HIDDEN)
            except Exception:
                pass
    except Exception as e:
        print(f"No se pudo escribir '{os.path.basename(path)}': {e}")

def _purge_other_configs(current_filename: str):
    """Elimina otras config_*.txt dejando solo la actual."""
    for fname in os.listdir(base_config_dir):
        if fname.startswith("config_") and fname.endswith(".txt") and fname != current_filename:
            fpath = os.path.join(base_config_dir, fname)
            try:
                _win_make_writable(fpath)
                os.remove(fpath)
            except Exception as ex:
                print(f"No se pudo borrar {fname}: {ex}")

def cargar_o_configurar(force_selector: bool = False, parent=None):
    """
    Retorna un dict con la configuración seleccionada o, si force_selector=True y el usuario cancela,
    retorna None (para que la app principal re-habilite botones).

    Inicio normal (force_selector=False):
      1) Si existe 'config_actual.txt', se usa esa.
      2) Si hay varias 'config_*.txt', se toma la más reciente.
      3) Si no hay ninguna, se abre el asistente (cargar JSON/TXT y seleccionar).
    """

    # 0) Si hay un puntero y NO se está forzando el selector → úsalo
    if not force_selector:
        try:
            if os.path.exists(ACTIVE_POINTER):
                name = open(ACTIVE_POINTER, "r", encoding="utf-8").read().strip()
                if name:
                    path = name if os.path.isabs(name) else os.path.join(base_config_dir, name)
                    if os.path.exists(path):
                        return _parse_config_txt(path)
        except Exception:
            pass

    # 1) Si ya hay configs guardadas y NO se fuerza selector → usar la más reciente
    archivos_config = [f for f in os.listdir(base_config_dir)
                       if f.startswith("config_") and f.endswith(".txt") and f != os.path.basename(ACTIVE_POINTER)]
    if not force_selector and len(archivos_config) >= 1:
        archivos_config.sort(
            key=lambda fn: os.path.getmtime(os.path.join(base_config_dir, fn)),
            reverse=True
        )
        config_path = os.path.join(base_config_dir, archivos_config[0])
        return _parse_config_txt(config_path)

    # 2) Preparar estructuras (cuando no hay configs o se fuerza selector)
    razones_sociales = {}

    # 3) Ventana modal de selección completa (Toplevel)
    def mostrar_configuracion_completa(parent=None):
        ventana = ctk.CTkToplevel(parent)
        ventana.title("Configuración de FacturaScan")
        aplicar_icono(ventana) 
        ventana.after(200, lambda: aplicar_icono(ventana))
 
        ventana.resizable(False, False)

        ancho, alto = 500, 530
        x = (ventana.winfo_screenwidth() // 2) - (ancho // 2)
        y = (ventana.winfo_screenheight() // 2) - (alto // 2)
        ventana.geometry(f"{ancho}x{alto}+{x}+{y}")

        def on_close():
            limpiar_callbacks(ventana)   # Cancela todos los after pendientes
            try:
                ventana.grab_release()   # Libera el control del modal
            except:
                pass
            ventana.destroy()            # Cierra la ventana

        ventana.protocol("WM_DELETE_WINDOW", on_close)
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
                    sucursal_var.set("")  # no auto-seleccionamos
                else:
                    sucursal_combo.configure(values=[], state="disabled")
                    sucursal_var.set("")
            else:
                sucursal_combo.configure(values=[], state="disabled")
                sucursal_var.set("")

        def elegir_entrada():
            carpeta = _askdir_above(
                ventana,
                "Selecciona carpeta de ENTRADA",
                os.path.join(os.path.expanduser("~"), "Desktop")
            )
            if carpeta:
                entrada_var.set(carpeta)

        def elegir_salida():
            carpeta = _askdir_above(
                ventana,
                "Selecciona carpeta de SALIDA",
                os.path.join(os.path.expanduser("~"), "Desktop")
            )
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

            config_data = {
                "RazonSocial": razon,
                "RutEmpresa": razones_sociales[razon]["rut"],
                "NomSucursal": sucursal,
                "DirSucursal": razones_sociales[razon]["sucursales"][sucursal],
                "CarEntrada": entrada,
                "CarpSalida": salida,
                "CarpSalidaUsoAtm": "",
            }

            sucursal_nombre = sucursal.lower().replace(" ", "_")
            config_filename = f"config_{sucursal_nombre}.txt"
            config_path = os.path.join(base_config_dir, config_filename)

            # Guardar config
            contenido = "".join(f'{k}="{v}"\n' for k, v in config_data.items())
            _safe_write_text(config_path, contenido, make_hidden=True)

            # Actualiza puntero a la config activa (escritura atómica y oculta)
            _safe_write_text(ACTIVE_POINTER, config_filename, make_hidden=True)

            # (Opcional) purgar otras configs
            if PURGE_OTHERS:
                _purge_other_configs(config_filename)

            ventana.resultado = _parse_config_txt(config_path)
            try:
                ventana.grab_release()
            except:
                pass
            ventana.destroy()

        # ---- UI ----
        frame = ctk.CTkFrame(ventana, fg_color="transparent")
        frame.pack(pady=20)

        ctk.CTkLabel(frame, text="Selecciona la razón social:", font=fuente).pack(pady=(10, 5))
        razon_combo = ctk.CTkComboBox(
            frame, variable=razon_var, values=list(razones_sociales.keys()),
            state="readonly", font=fuente, width=350
        )
        razon_combo.pack()

        ctk.CTkLabel(frame, text="Selecciona la sucursal:", font=fuente).pack(pady=(20, 5))
        sucursal_combo = ctk.CTkComboBox(
            frame, variable=sucursal_var, values=[],
            state="disabled", font=fuente, width=350
        )
        sucursal_combo.pack()

        razon_var.trace_add("write", actualizar_sucursales)

        ctk.CTkLabel(frame, text="Carpeta de entrada:", font=fuente).pack(pady=(15, 5))
        entrada_frame = ctk.CTkFrame(frame, fg_color="transparent")
        entrada_frame.pack(pady=(0, 10))
        ctk.CTkButton(
            entrada_frame, text="Buscar...", command=elegir_entrada,
            fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black"
        ).pack(side="left", padx=(0, 5))
        ctk.CTkEntry(entrada_frame, textvariable=entrada_var, width=260, font=fuente).pack(side="left")

        ctk.CTkLabel(frame, text="Carpeta de salida:", font=fuente).pack(pady=(15, 5))
        salida_frame = ctk.CTkFrame(frame, fg_color="transparent")
        salida_frame.pack(pady=(0, 20))
        ctk.CTkButton(
            salida_frame, text="Buscar...", command=elegir_salida,
            fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black"
        ).pack(side="left", padx=(0, 5))
        ctk.CTkEntry(salida_frame, textvariable=salida_var, width=260, font=fuente).pack(side="left")

        ctk.CTkButton(
            ventana, text="Guardar configuración", command=guardar_y_cerrar,
            width=250, height=40, font=fuente,
            fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black"
        ).pack(pady=(0, 20))

        # Estado inicial y modalidad
        ventana.resultado = None
        ventana.attributes("-topmost", True)
        ventana.grab_set()
        ventana.focus_force()
        ventana.wait_window()  # Modal
        return getattr(ventana, "resultado", None)

    # 4) Ventana inicial para cargar datos (json/txt)
    if force_selector:
        ventana_inicial = ctk.CTkToplevel(parent if force_selector else None)
    else:
        ventana_inicial = ctk.CTk()                # root normal en arranque

    ventana_inicial.title("Configuración inicial")
    aplicar_icono(ventana_inicial)
    ventana_inicial.after(200, lambda: aplicar_icono(ventana_inicial))
    ventana_inicial.geometry("400x200")
    ventana_inicial.resizable(False, False)
    ventana_inicial.update_idletasks()
    x = (ventana_inicial.winfo_screenwidth() // 2) - 200
    y = (ventana_inicial.winfo_screenheight() // 2) - 100
    ventana_inicial.geometry(f"400x200+{x}+{y}")

    def cerrar_configuracion():
        if force_selector:
            # Llamado desde el botón del menú → cancelar y volver (no cerrar app)
            limpiar_callbacks(ventana_inicial)
            ventana_inicial.destroy()
        else:
            # Arranque inicial → puedes permitir cerrar la app
            if messagebox.askyesno("Salir", "¿Deseas cerrar FacturaScan sin configurar?"):
                limpiar_callbacks(ventana_inicial)
                sys.exit(0)

    ventana_inicial.protocol("WM_DELETE_WINDOW", cerrar_configuracion)

    ctk.CTkLabel(
        ventana_inicial, text="Bienvenido a FacturaScan",
        font=ctk.CTkFont(size=18, weight="bold")
    ).pack(pady=20)
    ctk.CTkLabel(
        ventana_inicial, text="Cargue los datos de empresas para comenzar."
    ).pack(pady=(0, 15))

    resultado = None

    def cargar_datos_y_continuar():
        nonlocal razones_sociales, resultado
        path = _askopen_above(
            ventana_inicial,
            title="Selecciona archivo de datos (.json o .txt)",
            filetypes=[("Archivos JSON o TXT", "*.json *.txt")],
            initialdir="C:\\"
        )
        if not path:
            return

        try:
            if path.lower().endswith(".json"):
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                # Validación mínima
                for razon, datos in data.items():
                    if "rut" not in datos or "sucursales" not in datos:
                        raise ValueError(f"Falta 'rut' o 'sucursales' en: {razon}")
                razones_sociales = data

            elif path.lower().endswith(".txt"):
                razones_sociales = {}
                with open(path, "r", encoding="utf-8") as f:
                    for linea in f:
                        partes = linea.strip().split(";")
                        if len(partes) >= 3:
                            razon = partes[0].strip()
                            rut = partes[1].strip()
                            sucursales_raw = partes[2]
                            sucursales = {}
                            for item in sucursales_raw.split("|"):
                                if "=" in item:
                                    nombre, direccion = item.split("=", 1)
                                    sucursales[nombre.strip()] = direccion.strip()
                            if razon:
                                razones_sociales[razon] = {
                                    "rut": rut,
                                    "sucursales": sucursales
                                }
            else:
                raise ValueError("Formato no soportado. Usa un archivo .json o .txt")

            if not razones_sociales:
                raise ValueError("No se cargaron razones sociales válidas.")

            # Pasar a la selección detallada
            ventana_inicial.withdraw()
            resultado = mostrar_configuracion_completa(parent=ventana_inicial)
            ventana_inicial.destroy()

        except Exception as e:
            messagebox.showerror("Error", f"Datos incorrectos:\n{e}")

    ctk.CTkButton(
        ventana_inicial, text="Cargar Datos",
        command=cargar_datos_y_continuar,
        fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black"
    ).pack(pady=10)

    if force_selector:
        ventana_inicial.attributes("-topmost", True)
        ventana_inicial.grab_set()
        ventana_inicial.focus_force()
        with ocultar_stderr():
            ventana_inicial.wait_window()   # <- en lugar de mainloop()
    else:
        with ocultar_stderr():
            ventana_inicial.mainloop()

    # 5) Salida de la función según flujo
    if resultado:
        return resultado

    if force_selector:
        return None
    else:
        if messagebox.askyesno("Cancelar", "No se completó la configuración.\n¿Deseas salir de FacturaScan?"):
            sys.exit(0)
        else:
            return cargar_o_configurar()

# === CAMBIAR SOLO RUTAS DE ENTRADA/SALIDA ==========================
def actualizar_rutas(config_actual: dict | None = None, parent=None):
    """
    Abre un modal para cambiar SOLO CarEntrada y CarpSalida de la
    configuración actualmente activa. Devuelve el dict completo
    actualizado si se guardó, o None si el usuario canceló.
    """
    # 1) Ubicar el archivo config_*.txt activo
    cfg_path = None
    try:
        if os.path.exists(ACTIVE_POINTER):
            name = open(ACTIVE_POINTER, "r", encoding="utf-8").read().strip()
            path = name if os.path.isabs(name) else os.path.join(base_config_dir, name)
            if os.path.exists(path):
                cfg_path = path
    except Exception:
        pass

    if not cfg_path:
        # fallback por nombre de sucursal si nos lo entregan
        if config_actual and config_actual.get("NomSucursal"):
            suc = config_actual.get("NomSucursal", "").lower().replace(" ", "_")
            cand = os.path.join(base_config_dir, f"config_{suc}.txt")
            if os.path.exists(cand):
                cfg_path = cand

    if not cfg_path:
        # última config_*.txt por fecha
        candidatos = [f for f in os.listdir(base_config_dir) if f.startswith("config_") and f.endswith(".txt")]
        if candidatos:
            candidatos.sort(key=lambda fn: os.path.getmtime(os.path.join(base_config_dir, fn)), reverse=True)
            cfg_path = os.path.join(base_config_dir, candidatos[0])

    if not cfg_path or not os.path.exists(cfg_path):
        messagebox.showerror("Config", "No se encontró una configuración activa para actualizar.")
        return None

    # 2) Cargar la config actual y pre-rellenar
    datos = _parse_config_txt(cfg_path)
    entrada_ini = datos.get("CarEntrada", "")
    salida_ini  = datos.get("CarpSalida", "")

    # 3) Modal con dos campos (solo rutas)
    win = ctk.CTkToplevel(parent)
    win.title("Cambiar rutas de entrada/salida")
    aplicar_icono(win)  # o aplicar_icono(win)
    win.after(200, lambda: aplicar_icono(win))
    win.resizable(False, False)
    w, h = 520, 220
    x = (win.winfo_screenwidth() // 2) - (w // 2)
    y = (win.winfo_screenheight() // 2) - (h // 2)
    win.geometry(f"{w}x{h}+{x}+{y}")

    def on_close():
        limpiar_callbacks(win)
        try:
            win.grab_release()
        except:
            pass
        win.destroy()

    win.protocol("WM_DELETE_WINDOW", on_close)
    fuente = ctk.CTkFont(family="Segoe UI", size=14)

    var_in = ctk.StringVar(value=entrada_ini)
    var_out = ctk.StringVar(value=salida_ini)

    def pick_in():
        carpeta = _askdir_above(
            win,
            "Selecciona carpeta de ENTRADA",
            entrada_ini or os.path.expanduser("~")
        )
        if carpeta:
            var_in.set(carpeta)

    def pick_out():
        carpeta = _askdir_above(
            win,
            "Selecciona carpeta de SALIDA",
            salida_ini or os.path.expanduser("~")
        )
        if carpeta:
            var_out.set(carpeta)

    frm = ctk.CTkFrame(win, fg_color="transparent")
    frm.pack(padx=16, pady=16, fill="x")

    ctk.CTkLabel(frm, text="Carpeta de entrada:", font=fuente).pack(anchor="w")
    f1 = ctk.CTkFrame(frm, fg_color="transparent")
    f1.pack(fill="x", pady=(0,10))
    ctk.CTkButton(f1, text="Buscar...", command=pick_in, width=100,
                  fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black").pack(side="left")
    ctk.CTkEntry(f1, textvariable=var_in, width=360, font=fuente).pack(side="left", padx=8)

    ctk.CTkLabel(frm, text="Carpeta de salida:", font=fuente).pack(anchor="w")
    f2 = ctk.CTkFrame(frm, fg_color="transparent")
    f2.pack(fill="x", pady=(0,14))
    ctk.CTkButton(f2, text="Buscar...", command=pick_out, width=100,
                  fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black").pack(side="left")
    ctk.CTkEntry(f2, textvariable=var_out, width=360, font=fuente).pack(side="left", padx=8)

    def guardar():
        entrada = var_in.get().strip()
        salida  = var_out.get().strip()
        if not entrada or not salida:
            messagebox.showwarning("Falta información", "Debes seleccionar ambas carpetas.")
            return

        # Reescribir SOLO esas dos claves en el mismo archivo
        nuevos = dict(datos)
        nuevos["CarEntrada"] = entrada
        nuevos["CarpSalida"] = salida

        # composición del txt
        contenido = "".join(f'{k}="{v}"\n' for k, v in nuevos.items())

        try:
            # por si está oculto o con atributos que bloquean
            _win_make_writable(cfg_path)
            _safe_write_text(cfg_path, contenido, make_hidden=True)
        except Exception as e:
            messagebox.showerror("Config", f"No se pudo guardar:\n{e}")
            return

        # Mantener el puntero tal cual (sigue apuntando a la misma config)
        win.resultado = nuevos
        try: win.grab_release()
        except: pass
        win.destroy()

    ctk.CTkButton(win, text="Guardar", command=guardar, width=180, height=36,
                  fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black").pack(pady=(0,14))

    # Modalidad
    win.attributes("-topmost", True)
    win.grab_set()
    win.focus_force()
    win.wait_window()

    return getattr(win, "resultado", None)
# ==================================================================

# === CAMBIAR SOLO RAZÓN SOCIAL / SUCURSAL ==========================
def cambiar_razon_sucursal(config_actual: dict | None = None, parent=None):
    """
    Actualiza SOLO:
      - RazonSocial
      - RutEmpresa  (derivado de la razón social)
      - NomSucursal
      - DirSucursal (derivada de la sucursal)
    usando un archivo de razones sociales (.json o .txt).

    Mantiene intactos el resto de campos (CarEntrada, CarpSalida, CarpSalidaUsoAtm, etc.).
    Reescribe el MISMO config_*.txt activo (no cambia el puntero).
    Devuelve el dict completo actualizado si se guardó, o None si se canceló.
    """
    # 1) Ubicar el archivo config_*.txt activo
    cfg_path = None
    try:
        if os.path.exists(ACTIVE_POINTER):
            name = open(ACTIVE_POINTER, "r", encoding="utf-8").read().strip()
            path = name if os.path.isabs(name) else os.path.join(base_config_dir, name)
            if os.path.exists(path):
                cfg_path = path
    except Exception:
        pass

    if not cfg_path and config_actual and config_actual.get("NomSucursal"):
        suc = config_actual.get("NomSucursal", "").lower().replace(" ", "_")
        cand = os.path.join(base_config_dir, f"config_{suc}.txt")
        if os.path.exists(cand):
            cfg_path = cand

    if not cfg_path:
        candidatos = [f for f in os.listdir(base_config_dir) if f.startswith("config_") and f.endswith(".txt")]
        if candidatos:
            candidatos.sort(key=lambda fn: os.path.getmtime(os.path.join(base_config_dir, fn)), reverse=True)
            cfg_path = os.path.join(base_config_dir, candidatos[0])

    if not cfg_path or not os.path.exists(cfg_path):
        messagebox.showerror("Config", "No se encontró una configuración activa para actualizar.")
        return None

    # 2) Cargar config actual para preservar rutas/otros campos
    datos = _parse_config_txt(cfg_path)

    # 3) Modal UI
    win = ctk.CTkToplevel(parent)
    win.title("Cambiar Razón Social / Sucursal")
    aplicar_icono(win)
    win.after(200, lambda: aplicar_icono(win))
    win.resizable(False, False)
    w, h = 560, 320
    x = (win.winfo_screenwidth() // 2) - (w // 2)
    y = (win.winfo_screenheight() // 2) - (h // 2)
    win.geometry(f"{w}x{h}+{x}+{y}")

    def on_close():
        limpiar_callbacks(win)
        try: win.grab_release()
        except: pass
        win.destroy()
    win.protocol("WM_DELETE_WINDOW", on_close)

    fuente = ctk.CTkFont(family="Segoe UI", size=14)

    razones_sociales: dict = {}
    razon_var = ctk.StringVar(value="")
    sucursal_var = ctk.StringVar(value="")
    archivo_var = ctk.StringVar(value="(ninguno)")

    frm = ctk.CTkFrame(win, fg_color="transparent")
    frm.pack(padx=16, pady=16, fill="x")

    ctk.CTkLabel(frm, text="Archivo de razones sociales (.json o .txt):", font=fuente).pack(anchor="w")
    ffile = ctk.CTkFrame(frm, fg_color="transparent"); ffile.pack(fill="x", pady=(0, 10))

    def cargar_archivo():
        path = _askopen_above(
            win,
            title="Selecciona archivo de datos (.json o .txt)",
            filetypes=[("Archivos JSON o TXT", "*.json *.txt")],
            initialdir=r"C:\\"
        )
        if not path:
            return
        try:
            local = {}
            if path.lower().endswith(".json"):
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                # Formato esperado: { "Razon A": {"rut": "76.123.456-7", "sucursales": {"Sucursal 1":"Dir 1", ...}}, ... }
                for razon, datos_r in data.items():
                    if "rut" not in datos_r or "sucursales" not in datos_r:
                        raise ValueError(f"Falta 'rut' o 'sucursales' en: {razon}")
                local = data
            elif path.lower().endswith(".txt"):
                # Formato esperado por línea:
                # RazonSocial;76.123.456-7;Sucursal 1=Dirección 1|Sucursal 2=Dirección 2
                with open(path, "r", encoding="utf-8") as f:
                    for linea in f:
                        partes = linea.strip().split(";")
                        if len(partes) >= 3:
                            razon = partes[0].strip()
                            rut = partes[1].strip()
                            sucursales = {}
                            for item in partes[2].split("|"):
                                if "=" in item:
                                    nombre, direccion = item.split("=", 1)
                                    sucursales[nombre.strip()] = direccion.strip()
                            if razon:
                                local[razon] = {"rut": rut, "sucursales": sucursales}
            else:
                raise ValueError("Formato no soportado. Usa un archivo .json o .txt")

            if not local:
                raise ValueError("No se cargaron razones sociales válidas.")

            razones_sociales.clear()
            razones_sociales.update(local)
            archivo_var.set(os.path.basename(path))

            razon_combo.configure(values=list(razones_sociales.keys()), state="readonly")
            sucursal_combo.configure(values=[], state="disabled")
            razon_var.set(""); sucursal_var.set("")
        except Exception as e:
            messagebox.showerror("Archivo inválido", f"No se pudo cargar el archivo:\n{e}")

    ctk.CTkButton(ffile, text="Buscar...", command=cargar_archivo, width=120,
                  fg_color="#a6a6a6", hover_color="#8c8c8c", text_color="black").pack(side="left")
    ctk.CTkLabel(ffile, textvariable=archivo_var, font=fuente).pack(side="left", padx=8)

    ctk.CTkLabel(frm, text="Razón social:", font=fuente).pack(anchor="w", pady=(8, 4))
    razon_combo = ctk.CTkComboBox(frm, variable=razon_var, values=[], state="disabled", width=400, font=fuente)
    razon_combo.pack()

    ctk.CTkLabel(frm, text="Sucursal:", font=fuente).pack(anchor="w", pady=(12, 4))
    sucursal_combo = ctk.CTkComboBox(frm, variable=sucursal_var, values=[], state="disabled", width=400, font=fuente)
    sucursal_combo.pack()

    def actualizar_sucursales(*_):
        razon = razon_var.get().strip()
        if razon and razon in razones_sociales:
            sucs = list(razones_sociales[razon].get("sucursales", {}).keys())
            if sucs:
                sucursal_combo.configure(values=sucs, state="readonly")
                sucursal_var.set("")
            else:
                sucursal_combo.configure(values=[], state="disabled"); sucursal_var.set("")
        else:
            sucursal_combo.configure(values=[], state="disabled"); sucursal_var.set("")
    razon_var.trace_add("write", actualizar_sucursales)

    def guardar():
        razon = razon_var.get().strip()
        sucursal = sucursal_var.get().strip()
        if not razones_sociales:
            messagebox.showwarning("Falta archivo", "Primero carga el archivo de razones sociales.")
            return
        if not razon or not sucursal:
            messagebox.showwarning("Faltan datos", "Selecciona la razón social y la sucursal.")
            return

        nuevos = dict(datos)  # preserva rutas y demás
        try:
            nuevos["RazonSocial"] = razon
            nuevos["RutEmpresa"]  = razones_sociales[razon]["rut"]
            nuevos["NomSucursal"] = sucursal
            nuevos["DirSucursal"] = razones_sociales[razon]["sucursales"][sucursal]
        except Exception as e:
            messagebox.showerror("Datos", f"Error al armar la configuración:\n{e}")
            return

        contenido = "".join(f'{k}="{v}"\n' for k, v in nuevos.items())
        try:
            _win_make_writable(cfg_path)
            _safe_write_text(cfg_path, contenido, make_hidden=True)
        except Exception as e:
            messagebox.showerror("Config", f"No se pudo guardar:\n{e}")
            return

        win.resultado = nuevos
        try: win.grab_release()
        except: pass
        win.destroy()

    ctk.CTkButton(win, text="Guardar cambios", command=guardar,
                  width=200, height=36, fg_color="#a6a6a6",
                  hover_color="#8c8c8c", text_color="black").pack(pady=(0, 12))

    # Modal
    win.attributes("-topmost", True)
    win.grab_set()
    win.focus_force()
    win.wait_window()

    return getattr(win, "resultado", None)
# ==================================================================
