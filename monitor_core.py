import os
import time
import shutil
from datetime import datetime
from pdf2image import convert_from_path
from PIL import Image, UnidentifiedImageError
from concurrent.futures import ThreadPoolExecutor, as_completed
import tkinter as tk
from tkinter import messagebox
import threading
from config_gui import cargar_o_configurar
from ocr_utils import ocr_zona_factura_desde_png, extraer_rut, extraer_numero_factura
from pdf_tools import comprimir_pdf
from scanner import escanear_y_guardar_pdf

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

os.makedirs(CARPETA_ENTRADA, exist_ok=True)
os.makedirs(CARPETA_SALIDA, exist_ok=True)

def obtener_carpeta_salida_anual(base_path):
    """Crea (si no existe) y retorna la carpeta del año actual dentro de la salida."""
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

def registrar_log(mensaje):
    ahora = datetime.now()
    carpeta_logs = r"C:\\FacturaScan\\logs"
    os.makedirs(carpeta_logs, exist_ok=True)

    nombre_log = f"log_{ahora.strftime('%Y_%m')}_{ahora.strftime('%d')}.txt"
    ruta_log = os.path.join(carpeta_logs, nombre_log)

    timestamp = ahora.strftime("[%Y-%m-%d %H:%M:%S]")
    with open(ruta_log, "a", encoding="utf-8") as f:
        f.write(f"{timestamp} {mensaje}\n")

def procesar_archivo(pdf_path):
    from PIL import ImageFilter

    nombre = os.path.basename(pdf_path)
    nombre_base = os.path.splitext(nombre)[0]
    ruta_debug_dir = r"C:\FacturaScan\debug"
    os.makedirs(ruta_debug_dir, exist_ok=True)

    ruta_png = os.path.join(ruta_debug_dir, nombre_base + ".png")

    if not os.path.exists(ruta_png):
        try:
            imagenes = convert_from_path(pdf_path, dpi=300, first_page=1, last_page=1)
            if imagenes:
                imagen_filtrada = imagenes[0].filter(ImageFilter.SHARPEN).filter(ImageFilter.DETAIL)
                imagen_filtrada.save(ruta_png, "PNG", optimize=True, compress_level=5)



        except Exception as e:
            print(f"❌ Error al convertir PDF a imagen: {e}")
            registrar_log(f"❌ Error al convertir PDF a imagen ({nombre}): {e}")
            return

    try:
        with Image.open(ruta_png) as img:
            img.verify()
    except (UnidentifiedImageError, OSError):
        print(f"⚠️ Imagen inválida o corrupta: {ruta_png}")
        registrar_log(f"⚠️ Imagen inválida o corrupta: {ruta_png}")
        return

    try:
        # recorte al debug
        nombre_recorte = nombre_base + "_recorte.png"
        ruta_debug_dir = r"C:\FacturaScan\debug"
        os.makedirs(ruta_debug_dir, exist_ok=True)  # Crea la carpeta si no existe
        ruta_debug = os.path.join(ruta_debug_dir, nombre_recorte)

        texto = ocr_zona_factura_desde_png(ruta_png, ruta_debug=ruta_debug)


    except Exception as e:
        print(f"⚠️ Error durante OCR: {e}")
        registrar_log(f"⚠️ Error durante OCR ({nombre}): {e}")
        return

    if not texto.strip():
        registrar_log(f"⚠️ No se extrajo texto OCR de {nombre}")
        documentos_path = os.path.join(os.path.expanduser("~"), "Documents")
        hoy = datetime.now()
        nombre_final = f"documento_escaneado_{hoy.strftime('%Y%m%d_%H%M')}.pdf"
        ruta_destino = os.path.join(documentos_path, nombre_final)
        shutil.move(pdf_path, ruta_destino)
        print(f"📤 Documento sin texto OCR movido a Documentos como: {nombre_final}")
        registrar_log(f"📤 Documento sin texto OCR movido a Documentos como: {nombre_final}")
        return

    rut_proveedor = extraer_rut(texto)
    numero_factura = extraer_numero_factura(texto)

    try:
        comprimir_pdf(GS_PATH, pdf_path, dpi=100)
    except Exception as e:
        print(f"⚠️ Error al comprimir {nombre}: {e}")
        registrar_log(f"⚠️ Error al comprimir PDF ({nombre}): {e}")
        return

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

    # Carpeta anual destino
    carpeta_anual = obtener_carpeta_salida_anual(CARPETA_SALIDA)

    # Valores por defecto si falta info
    if not rut_proveedor or rut_proveedor == "desconocido":
        rut_proveedor = "rut_desconocido"
    if not numero_factura:
        numero_factura = "factura_desconocida"

    base_name = f"{SUCURSAL}_{rut_proveedor}_factura_{numero_factura}_{anio}"

    # 1. Nombre temporal único
    temp_nombre = base_name + "_" + datetime.now().strftime("%H%M%S%f")
    temp_ruta = os.path.join(carpeta_anual, f"{temp_nombre}.pdf")
    shutil.move(pdf_path, temp_ruta)

    # 2. Generar nombre incremental seguro
    nombre_final = generar_nombre_incremental(carpeta_anual, base_name, ".pdf")
    ruta_destino = os.path.join(carpeta_anual, nombre_final)

    # 3. Renombrar
    os.rename(temp_ruta, ruta_destino)

    print(f"✅ Archivo procesado: {os.path.basename(ruta_destino)}")
    registrar_log(f"✅ Procesado archivo: {os.path.basename(ruta_destino)}")


def ejecutar_monitor():
    print(f"Razón social: {RAZON_SOCIAL}")
    print(f"RUT: {RUT_EMPRESA}")
    print(f"Sucursal: {SUCURSAL}")
    print(f"Dirección: {DIRECCION}")

    try:
        while True:
            nombre_pdf = "escaneo_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".pdf"
            ruta_pdf = escanear_y_guardar_pdf(nombre_pdf, CARPETA_ENTRADA, r"C:\FacturaScan\debug")

            if ruta_pdf:
                registrar_log(f"📥 Documento escaneado: {os.path.basename(ruta_pdf)}")
                procesar_archivo(ruta_pdf)
            else:
                print("🔍 Procesando archivos pendientes en carpeta de entrada...")
                procesar_entrada_una_vez()

            print("🕒 Esperando nuevos escaneos...")
            time.sleep(INTERVALO)

    except KeyboardInterrupt:
        print("\n⛔ Programa interrumpido por el usuario. Cerrando...")
        registrar_log("⛔ Programa interrumpido manualmente por el usuario.")

def procesar_entrada_una_vez():
    print("🔍 Procesando archivos pendientes en carpeta de entrada...")
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

    root = tk.Tk()
    root.withdraw()

    with ThreadPoolExecutor(max_workers=4) as executor:
        tareas = {
            executor.submit(procesar_archivo, os.path.join(CARPETA_ENTRADA, archivo)): archivo
            for archivo in archivos_pdf
        }

        for tarea in as_completed(tareas):
            archivo = tareas[tarea]
            try:
                tarea.result()
            except Exception as e:
                registrar_log(f"❌ Error procesando archivo {archivo}: {e}")
                print(f"❌ Error procesando archivo {archivo}: {e}")

    duracion = time.time() - inicio
    minutos = int(duracion // 60)
    segundos = int(duracion % 60)
    print(f"✅ Procesamiento completado.\nTiempo total: {minutos} min {segundos} segundos.")
    messagebox.showinfo("Finalizado", f"✅ Procesamiento completado.\nTiempo total: {minutos} min {segundos} seg.")
    root.destroy()
