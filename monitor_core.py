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
from config_gui import cargar_o_configurar
from ocr_utils import ocr_zona_factura_desde_png, extraer_rut, extraer_numero_factura
from pdf_tools import comprimir_pdf
from log_utils import registrar_log_proceso, registrar_log
  # o ajusta según el nombre del archivo

# Obtener configuración
variables = cargar_o_configurar()
nombre_lock = threading.Lock()
RAZON_SOCIAL = variables.get("RazonSocial", "desconocida")
RUT_EMPRESA = variables.get("RutEmpresa", "desconocido")
SUCURSAL = variables.get("NomSucursal", "sucursal_default")
DIRECCION = variables.get("DirSucursal", "direccion_no_definida")
CARPETA_ENTRADA = variables.get("CarEntrada", "entrada_default")
CARPETA_SALIDA = variables.get("CarpSalida", "salida_default")

INTERVALO = 1

# Ajustes de compresión global
CALIDAD_PDF = "default"  # Puedes cambiar a: screen, ebook, printer, prepress, default
DPI_PDF = 100
COMPRIMIR_PDF = True  # Cambia a True si deseas activar la compresión


os.makedirs(CARPETA_ENTRADA, exist_ok=True)
os.makedirs(CARPETA_SALIDA, exist_ok=True)

def obtener_carpeta_salida_anual(base_path):
    # Crea (si no existe) y retorna la carpeta del año actual dentro de la salida
    año_actual = datetime.now().strftime("%Y")
    carpeta_anual = os.path.join(base_path, año_actual)
    os.makedirs(carpeta_anual, exist_ok=True)
    return carpeta_anual

GS_PATH = next((
    ruta for ruta in [
        r"C:\\Program Files\\gs\\gs10.05.1\\bin\\gswin64c.exe",
        r"C:\\Program Files (x86)\\gs\\gs10.05.1\\bin\\gswin32c.exe"
    ] if os.path.exists(ruta)
), None)

def generar_nombre_incremental(base_path, nombre_base, extension):
    with nombre_lock:
        nombre_sin_sufijo = f"{nombre_base}{extension}"
        ruta_sin_sufijo = os.path.join(base_path, nombre_sin_sufijo)

        if not os.path.exists(ruta_sin_sufijo):
            return nombre_sin_sufijo

        contador = 1
        while True:
            nuevo_nombre = f"{nombre_base}_{contador}{extension}"
            ruta_con_sufijo = os.path.join(base_path, nuevo_nombre)
            if not os.path.exists(ruta_con_sufijo):
                return nuevo_nombre
            contador += 1

# Modo debug: guarda PNG preprocesados y recortes

def procesar_archivo(pdf_path):
    from PIL import ImageFilter

    modo_debug = False  # Cambia a True si deseas guardar los PNG preprocesados
    nombre = os.path.basename(pdf_path)
    nombre_base = os.path.splitext(nombre)[0]

    # Ruta PNG si está el modo debug
    ruta_debug_dir = r"C:\FacturaScan\debug"
    ruta_png = os.path.join(ruta_debug_dir, nombre_base + ".png") if modo_debug else None

    try:
        imagenes = convert_from_path(pdf_path, dpi=300, first_page=1, last_page=1)
        if imagenes:
            imagen_filtrada = imagenes[0].filter(ImageFilter.SHARPEN).filter(ImageFilter.DETAIL)
            if modo_debug:
                os.makedirs(ruta_debug_dir, exist_ok=True)
                imagen_filtrada.save(ruta_png, "PNG", optimize=True, compress_level=7)
            else:
                imagen_temporal = imagen_filtrada.convert("RGB")
    except Exception as e:
        # print(f"❌ Error al convertir PDF a imagen: {e}")
        registrar_log_proceso(f"❌ Error al convertir PDF a imagen ({nombre}): {e}")
        return

    # Validación imagen
    try:
        if modo_debug and ruta_png and os.path.exists(ruta_png):
            with Image.open(ruta_png) as img:
                img.verify()
        elif not modo_debug:
            imagen_temporal.verify()
    except Exception as e:
        # print(f"⚠️ Imagen inválida: {e}")
        registrar_log_proceso(f"⚠️ Imagen inválida: {e}")
        return

    # OCR
    try:
        if modo_debug and ruta_png:
            ruta_debug = os.path.join(ruta_debug_dir, nombre_base + "_recorte.png")
            texto = ocr_zona_factura_desde_png(ruta_png, ruta_debug=ruta_debug)
        else:
            texto = ocr_zona_factura_desde_png(imagen_temporal, ruta_debug=None)
    except Exception as e:
        # print(f"⚠️ Error durante OCR: {e}")
        registrar_log_proceso(f"⚠️ Error durante OCR ({nombre}): {e}")
        return

    if not texto.strip():
        registrar_log_proceso(f"⚠️ No se extrajo texto OCR de {nombre}")
        documentos_path = os.path.join(os.path.expanduser("~"), "Documents")
        hoy = datetime.now()
        nombre_final = f"documento_escaneado_{hoy.strftime('%Y%m%d_%H%M')}.pdf"
        ruta_destino = os.path.join(documentos_path, nombre_final)
        shutil.move(pdf_path, ruta_destino)
        mensaje = f"⚠️ Documento sin RUT ni número de factura detectado: {nombre}.\n📤 Se movió a Documentos como: {nombre_final}"
        print(mensaje)  # ✅ Se envía al CTkTextbox en tiempo real
        registrar_log_proceso(mensaje)  # También lo deja en logs históricos
        registrar_log(mensaje)  # ✅ Guarda en logs por si necesitas revisar
        return

    rut_proveedor = extraer_rut(texto)
    numero_factura = extraer_numero_factura(texto)

    hoy = datetime.now()
    anio = hoy.strftime("%Y")

    if rut_proveedor == "desconocido" and not numero_factura:
        documentos_path = os.path.join(os.path.expanduser("~"), "Documents")
        nombre_final = f"documento_escaneado_{hoy.strftime('%Y%m%d_%H%M')}.pdf"
        ruta_destino = os.path.join(documentos_path, nombre_final)
        print("⚠️ Documento sin RUT ni número. Moviendo a Documentos.")
        registrar_log(f"⚠️ Documento sin RUT ni número. Movido como: {nombre_final}")
        shutil.move(pdf_path, ruta_destino)
        return
# revisar
    carpeta_anual = obtener_carpeta_salida_anual(CARPETA_SALIDA)

    if not rut_proveedor or rut_proveedor == "desconocido":
        rut_proveedor = "rut_desconocido"
    if not numero_factura:
        numero_factura = "factura_desconocida"

    base_name = f"{SUCURSAL}_{rut_proveedor}_factura_{numero_factura}_{anio}"
    temp_nombre = base_name + "_" + datetime.now().strftime("%H%M%S%f")
    temp_ruta = os.path.join(carpeta_anual, f"{temp_nombre}.pdf")
    shutil.move(pdf_path, temp_ruta)

    # Solo comprimir si la variable global lo indica
    try:
        if COMPRIMIR_PDF and (rut_proveedor != "desconocido" or numero_factura):
            comprimir_pdf(GS_PATH, temp_ruta, calidad=CALIDAD_PDF, dpi=DPI_PDF, tamano_pagina='a4')
    except Exception as e:
        # print(f"⚠️ Error al comprimir: {e}")
        registrar_log_proceso(f"⚠️ Error al comprimir PDF: {e}")

    nombre_final = generar_nombre_incremental(carpeta_anual, base_name, ".pdf")
    ruta_destino = os.path.join(carpeta_anual, nombre_final)
    os.rename(temp_ruta, ruta_destino)

    print(f"✅ Archivo procesado: {os.path.basename(ruta_destino)}")
    registrar_log(f"✅ Procesado archivo: {os.path.basename(ruta_destino)}")

from concurrent.futures import ThreadPoolExecutor, as_completed
from multiprocessing import cpu_count
import os
import time
import tkinter as tk
from tkinter import messagebox

def procesar_entrada_una_vez():
    inicio = time.time()

    archivos_pdf = sorted(
        [f for f in os.listdir(CARPETA_ENTRADA) if f.lower().endswith(".pdf")],
        key=lambda f: os.path.getmtime(os.path.join(CARPETA_ENTRADA, f))
    )

    if not archivos_pdf:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Sin documentos", "📭 No se encontraron documentos pendientes en la carpeta de entrada.")
        root.destroy()
        return

    total = len(archivos_pdf)

    # Detectar núcleos y definir cantidad óptima de hilos
    nucleos = cpu_count()
    max_hilos = min(nucleos, 8)  # Máximo hasta 8 para evitar saturación

    registrar_log_proceso(f"🧠 Núcleos detectados: {nucleos} | Hilos usados: {max_hilos}")
    print(f"🧠 Núcleos detectados: {nucleos} | Hilos usados: {max_hilos}")

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
                tarea.result()
                print(f"✅ {i}/{total} procesado: {archivo}")
            except Exception as e:
                registrar_log_proceso(f"❌ Error procesando archivo {archivo}: {e}")

    duracion = time.time() - inicio
    minutos = int(duracion // 60)
    segundos = int(duracion % 60)

    messagebox.showinfo("Finalizado", f"✅ Procesamiento completado.\nTiempo total: {minutos} min {segundos} seg.")
    root.destroy()
