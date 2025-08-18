import hide_subprocess  # Esto parcha subprocess.run/call/Popen globalmente
import os
import time
import shutil
from datetime import datetime
from pdf2image import convert_from_path
from PIL import Image
from concurrent.futures import ThreadPoolExecutor, as_completed
import tkinter as tk
from tkinter import messagebox
import threading
import subprocess
import sys

from config_gui import cargar_o_configurar
from ocr_utils import ocr_zona_factura_desde_png, extraer_rut, extraer_numero_factura
from pdf_tools import comprimir_pdf
from log_utils import registrar_log_proceso, registrar_log, is_debug

# ===================== Helpers: crear carpetas una sola vez por ejecución =====================
_dir_cache = set()
_dir_lock = threading.Lock()

def _canon(p: str) -> str:
    return os.path.normcase(os.path.abspath(os.path.normpath(p)))

def ensure_dir(path: str) -> str:
    if not path:
        return path
    p = _canon(path)
    with _dir_lock:
        if p in _dir_cache:
            return p
        os.makedirs(p, exist_ok=True)
        _dir_cache.add(p)
    return p
# =============================================================================================

nombre_lock = threading.Lock()
variables = cargar_o_configurar()

RAZON_SOCIAL   = variables.get("RazonSocial", "desconocida")
RUT_EMPRESA    = variables.get("RutEmpresa", "desconocido")
SUCURSAL       = variables.get("NomSucursal", "sucursal_default")
DIRECCION      = variables.get("DirSucursal", "direccion_no_definida")
CARPETA_ENTRADA = variables.get("CarEntrada", "entrada_default")
CARPETA_SALIDA  = variables.get("CarpSalida", "salida_default")

INTERVALO = 1
# Ajustes de compresión global
CALIDAD_PDF   = "default"   # screen, ebook, printer, prepress, default
DPI_PDF       = 100
COMPRIMIR_PDF = True

# Inicializar estructura de carpetas una sola vez
def obtener_carpeta_salida_anual(base_path):
    año_actual = datetime.now().strftime("%Y")
    return ensure_dir(os.path.join(base_path, año_actual))

def inicializar_estructura_directorios():
    ensure_dir(CARPETA_ENTRADA)
    ensure_dir(CARPETA_SALIDA)
    ensure_dir(os.path.join(CARPETA_SALIDA, "No_Reconocidos"))
    # Precrear año actual para ambas clasificaciones
    for sub in ("Cliente", "Proveedores"):
        obtener_carpeta_salida_anual(os.path.join(CARPETA_SALIDA, sub))

inicializar_estructura_directorios()

GS_PATH = next((
    ruta for ruta in [
        r"C:\\Program Files\\gs\\gs10.05.1\\bin\\gswin64c.exe",
        r"C:\\Program Files (x86)\\gs\\gs10.05.1\\bin\\gswin32c.exe"
    ] if os.path.exists(ruta)), None)

def generar_nombre_incremental(base_path, nombre_base, extension):
    """
    Genera un nombre de archivo único evitando colisiones en un entorno multi-hilo.
    Si ya existe un archivo con el nombre, le agrega _1, _2, etc.
    """
    base_path = ensure_dir(base_path)
    with nombre_lock:  # protege en entornos concurrentes
        contador = 0
        while True:
            if contador == 0:
                nombre_final = f"{nombre_base}{extension}"
            else:
                nombre_final = f"{nombre_base}_{contador}{extension}"
            ruta_completa = os.path.join(base_path, nombre_final)
            if not os.path.exists(ruta_completa):
                return nombre_final
            contador += 1

def procesar_archivo(pdf_path):
    from PIL import ImageFilter
    from pdf2image.exceptions import PDFInfoNotInstalledError, PDFPageCountError, PDFSyntaxError
    import traceback

    modo_debug = is_debug()
    nombre = os.path.basename(pdf_path)
    registrar_log_proceso(f"📄 Iniciando procesamiento de: {nombre}")

    nombre_base = os.path.splitext(nombre)[0]
    base_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)
    ruta_debug_dir = ensure_dir(os.path.join(base_dir, "debug"))

    ruta_png = os.path.join(ruta_debug_dir, nombre_base + ".png") if modo_debug else None

    try:
        imagenes = convert_from_path(
            pdf_path,
            dpi=300,
            fmt="jpeg",
            thread_count=1,
            first_page=1,
            last_page=1,
            poppler_path=r"C:\poppler\Library\bin")

        if not imagenes:
            registrar_log_proceso(f"❌ No se pudo convertir {nombre} a imagen.")
            return

        imagen_temporal = imagenes[0].filter(ImageFilter.SHARPEN).filter(ImageFilter.DETAIL)

        if modo_debug:
            # Guardar una vista completa de la página para inspección
            imagen_temporal.save(ruta_png, "PNG", optimize=True, compress_level=7)
            registrar_log_proceso(f"📸 Imagen completa guardada en: {ruta_png}")

        # Copia independiente para OCR
        imagen_temporal = imagen_temporal.copy()

    except Exception as e:
        registrar_log_proceso(f"❌ Error al procesar imagen de {nombre}:\n{traceback.format_exc()}")
        return

    try:
        if modo_debug and ruta_png and os.path.exists(ruta_png):
            ruta_recorte = os.path.join(ruta_debug_dir, nombre_base + "_recorte.png")
            texto = ocr_zona_factura_desde_png(ruta_png, ruta_debug=ruta_recorte)
            # El propio ocr_zona_factura_desde_png guarda recortes SOLO si debug está ON
        else:
            texto = ocr_zona_factura_desde_png(imagen_temporal, ruta_debug=None)
    except Exception as e:
        registrar_log_proceso(f"⚠️ Error durante OCR ({nombre}): {e}")
        return

    if not texto.strip():
        no_reconocidos_path = ensure_dir(os.path.join(CARPETA_SALIDA, "No_Reconocidos"))

        base_error_name = f"Documento_NoReconocido_{SUCURSAL}_{datetime.now().strftime('%Y%m%d_%H%M%S%f')}"
        nombre_final = generar_nombre_incremental(no_reconocidos_path, base_error_name, ".pdf")
        ruta_destino = os.path.join(no_reconocidos_path, nombre_final)

        shutil.move(pdf_path, ruta_destino)
        registrar_log_proceso(f"⚠️ Documento sin texto OCR. Movido a No_Reconocidos como: {nombre_final}")
        return ruta_destino

    rut_proveedor = extraer_rut(texto)
    numero_factura = extraer_numero_factura(texto)

    hoy = datetime.now()
    anio = hoy.strftime("%Y")

    rut_valido = rut_proveedor and rut_proveedor != "desconocido"
    factura_valida = numero_factura and numero_factura != ""

    rut_nombre = rut_proveedor if rut_valido else "noreconocido"
    factura_nombre = numero_factura if factura_valida else "noreconocido"
    base_name = f"{SUCURSAL}_{rut_nombre}_factura_{factura_nombre}_{anio}"

    # Si alguno no fue reconocido → guardar en No_Reconocidos con nombre completo
    if not rut_valido or not factura_valida:
        no_reconocidos_path = ensure_dir(os.path.join(CARPETA_SALIDA, "No_Reconocidos"))

        nombre_final = generar_nombre_incremental(no_reconocidos_path, base_name, ".pdf")
        ruta_destino = os.path.join(no_reconocidos_path, nombre_final)

        shutil.move(pdf_path, ruta_destino)

        # Comprimir incluso si es No_Reconocido
        if COMPRIMIR_PDF and GS_PATH:
            try:
                comprimir_pdf(GS_PATH, ruta_destino, calidad=CALIDAD_PDF, dpi=DPI_PDF, tamano_pagina='a4')
            except Exception as e:
                registrar_log_proceso(f"⚠️ Fallo al comprimir {ruta_destino}. Detalle: {e}")

        motivo = []
        if not rut_valido:
            motivo.append("RUT no reconocido")
        if not factura_valida:
            motivo.append("N° factura no reconocido")

        registrar_log_proceso(f"⚠️ Documento movido a No_Reconocidos como: {nombre_final} | Motivo: {', '.join(motivo)}")
        return ruta_destino

    # Clasificar en Cliente o Proveedores
    subcarpeta = "Cliente" if rut_proveedor.replace(".", "").replace("-", "") == RUT_EMPRESA.replace(".", "").replace("-", "") else "Proveedores"
    carpeta_clasificada = os.path.join(CARPETA_SALIDA, subcarpeta)
    carpeta_anual = obtener_carpeta_salida_anual(carpeta_clasificada)

    temp_nombre = base_name + "_" + datetime.now().strftime('%H%M%S%f')
    temp_ruta = os.path.join(carpeta_anual, f"{temp_nombre}.pdf")

    try:
        shutil.move(pdf_path, temp_ruta)
    except Exception as e:
        registrar_log_proceso(f"❗ Error al mover archivo original: {e}")
        return

    if COMPRIMIR_PDF and GS_PATH:
        try:
            comprimir_pdf(GS_PATH, temp_ruta, calidad=CALIDAD_PDF, dpi=DPI_PDF, tamano_pagina='a4')
        except Exception as e:
            registrar_log_proceso(f"⚠️ Fallo al comprimir {temp_ruta}. Se guarda sin comprimir. Detalle: {e}")

    try:
        for intento in range(5):
            nombre_final = generar_nombre_incremental(carpeta_anual, base_name, ".pdf")
            ruta_destino = os.path.join(carpeta_anual, nombre_final)
            if not os.path.exists(ruta_destino):
                os.rename(temp_ruta, ruta_destino)
                registrar_log(f"✅ Procesado archivo: {os.path.basename(ruta_destino)}")
                return os.path.basename(ruta_destino)
            time.sleep(0.2)
    except Exception as e:
        fallback_name = f"{base_name}_backup_{datetime.now().strftime('%H%M%S%f')}.pdf"
        fallback_path = os.path.join(carpeta_anual, fallback_name)
        shutil.move(temp_ruta, fallback_path)
        registrar_log_proceso(f"❗ Error al renombrar archivo. Guardado como fallback: {fallback_name} | Detalle: {e}")
        return fallback_name


def procesar_entrada_una_vez():
    inicio = time.time()
    archivos_pdf = sorted(
        [f for f in os.listdir(CARPETA_ENTRADA) if f.lower().endswith(".pdf")],
        key=lambda f: os.path.getmtime(os.path.join(CARPETA_ENTRADA, f)))

    if not archivos_pdf:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Sin documentos", "No se encontraron documentos pendientes en la carpeta de entrada.")
        root.destroy()
        return

    total = len(archivos_pdf)
    nucleos = os.cpu_count()
    max_hilos = min(nucleos or 1, 8)

    registrar_log_proceso(f"🧠 Núcleos detectados: {nucleos} | Hilos usados: {max_hilos}")
    print('Iniciando procesamiento...')
    root = tk.Tk()
    root.withdraw()

    with ThreadPoolExecutor(max_workers=max_hilos) as executor:
        tareas = {
            executor.submit(procesar_archivo, os.path.join(CARPETA_ENTRADA, archivo)): archivo
            for archivo in archivos_pdf
        }

        for i, tarea in enumerate(as_completed(tareas), 1):
            archivo = tareas[tarea]
            try:
                resultado = tarea.result()
                print(f"{i}/{total} Entrada: {archivo}")
                if resultado:
                    print(f"✅ Procesado: {resultado}")
                else:
                    print(f"⚠️ {i}/{total} Procesado con errores: {archivo}")

            except Exception as e:
                registrar_log_proceso(f"❌ Error procesando archivo {archivo}: {e}")
    
    duracion = time.time() - inicio
    minutos = int(duracion // 60)
    segundos = int(duracion % 60)
    messagebox.showinfo("Finalizado", f"✅ Procesamiento completado.\nTiempo total: {minutos} min {segundos} seg.")
    root.destroy()
