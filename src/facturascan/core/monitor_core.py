import utils.hide_subprocess as hide_subprocess  # Parchea subprocess.run/call/Popen para ocultar ventanas en Windows
import os, re
# import time
# import shutil
from datetime import datetime
from pdf2image import convert_from_path
# from PIL import Image
# from concurrent.futures import ThreadPoolExecutor, as_completed
# import tkinter as tk
# from tkinter import messagebox
import threading
import subprocess
import sys

from ocr.ocr_utils import ocr_zona_factura_desde_png, extraer_rut, extraer_numero_factura
from pdf.pdf_tools import comprimir_pdf
from utils.log_utils import registrar_log_proceso, registrar_log, is_debug

# ===================== Helpers de carpetas (idempotentes por ejecución) =====================
_dir_cache = set()
_dir_lock = threading.Lock()

def _canon(p: str) -> str:
    """Normaliza rutas (absoluta + normcase + normpath) para comparar y cachear."""
    return os.path.normcase(os.path.abspath(os.path.normpath(p)))

def ensure_dir(path: str) -> str:
    """
    Crea la carpeta si no existe (una sola vez por ejecución).
    Seguro para multi-hilo gracias a _dir_lock y _dir_cache.
    """
    if not path:
        return path
    p = _canon(path)
    with _dir_lock:
        if p in _dir_cache:
            return p
        os.makedirs(p, exist_ok=True)
        _dir_cache.add(p)
    return p

def mkdir(path: str) -> str:
    """Crea la carpeta (sin cache) cada vez que se llama."""
    if path:
        os.makedirs(path, exist_ok=True)
    return path

# =============================================================================================

# Lock para generar nombres únicos de archivos en entorno multi-hilo
nombre_lock = threading.Lock()

# Cargar configuración de la app (razón social, rutas, etc.)
# variables = cargar_o_configurar()
variables = {}
RAZON_SOCIAL    = variables.get("RazonSocial", "desconocida")

# Variables de contexto (con defaults por si falta algo en la config)
RAZON_SOCIAL    = variables.get("RazonSocial", "desconocida")
RUT_EMPRESA     = variables.get("RutEmpresa", "desconocido")
SUCURSAL        = variables.get("NomSucursal", "sucursal_default")
DIRECCION       = variables.get("DirSucursal", "direccion_no_definida")
CARPETA_ENTRADA = variables.get("CarEntrada", "entrada_default")
CARPETA_SALIDA  = variables.get("CarpSalida", "salida_default")
# NUEVO: carpeta opcional desde config para documentos con "USO ATM"
CARPETA_SALIDA_USO_ATM = variables.get("CarpSalidaUsoAtm", "").strip()
# Fallback solicitado por el usuario: SIEMPRE C:\ si no hay ruta válida en config
FALLBACK_USO_ATM_DIR = r"C:\ATM"

# Intervalo para escaneos temporizados (si se implementa watch loop)
INTERVALO = 1

# ===== Ajustes globales de compresión de PDF (Ghostscript) =====
CALIDAD_PDF   = "default"   # screen, ebook, printer, prepress, default
DPI_PDF       = 200
COMPRIMIR_PDF = True

def aplicar_nueva_config(nuevas: dict):
    global variables, RAZON_SOCIAL, RUT_EMPRESA, SUCURSAL, DIRECCION
    global CARPETA_ENTRADA, CARPETA_SALIDA, CARPETA_SALIDA_USO_ATM

    variables = nuevas or {}
    RAZON_SOCIAL    = variables.get("RazonSocial",    RAZON_SOCIAL)
    RUT_EMPRESA     = variables.get("RutEmpresa",     RUT_EMPRESA)
    SUCURSAL        = variables.get("NomSucursal",    SUCURSAL)
    DIRECCION       = variables.get("DirSucursal",    DIRECCION)
    CARPETA_ENTRADA = variables.get("CarEntrada",     CARPETA_ENTRADA)
    CARPETA_SALIDA  = variables.get("CarpSalida",     CARPETA_SALIDA)
    CARPETA_SALIDA_USO_ATM = variables.get("CarpSalidaUsoAtm", CARPETA_SALIDA_USO_ATM)

    # Asegura que existan las rutas nuevas (idempotente)
    ensure_dir(CARPETA_ENTRADA)
    ensure_dir(CARPETA_SALIDA)

# ===================== Estructura de salida (año/cliente/proveedores) =====================

def obtener_carpeta_salida_anual(base_path):
    """Devuelve la carpeta del año actual dentro de base_path, creándola si no existe."""
    año_actual = datetime.now().strftime("%Y")
    return ensure_dir(os.path.join(base_path, año_actual))

def _find_gs():
    bases = [r"C:\Program Files\gs", r"C:\Program Files (x86)\gs"]
    for base in bases:
        if not os.path.isdir(base):
            continue
        for d in sorted(os.listdir(base), reverse=True):  # toma la más nueva
            cand64 = os.path.join(base, d, "bin", "gswin64c.exe")
            cand32 = os.path.join(base, d, "bin", "gswin32c.exe")
            if os.path.exists(cand64): return cand64
            if os.path.exists(cand32): return cand32
    return None

GS_PATH = _find_gs()

# ===================== Nombres de archivo únicos (thread-safe) =====================
def generar_nombre_incremental(base_path, nombre_base, extension):
    """
    Genera un nombre único dentro de base_path:
    nombre_base + ( _1 | _2 | ... ) + extension
    Protegido por lock para evitar colisiones entre hilos.
    """
    base_path = ensure_dir(base_path)
    with nombre_lock:
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

def _es_guia_despacho(texto: str) -> bool:
    """
    Clasificador robusto de 'Guía de Despacho' con puntaje:
    + Señales fuertes (GUIA~DESPACH, 'GUIA DE TRASLADO', 'GDE', DTE 52, 'DE DESPACHO ELECTRONICA', 'SOLO TRASLADO')
    - Señales de que NO es guía (FACTURA, DTE 33/34, NOTAS, 'DIRECCION/LUGAR/FECHA DE DESPACHO')
    Devuelve True si el puntaje >= 3.
    """
  
    t = texto.upper()
    t = re.sub(r'[^A-Z0-9\s\.]', ' ', t)   # conserva letras/números/espacios/puntos (para G.D.E.)
    t = re.sub(r'\s+', ' ', t).strip()

    score = 0

    # -------- POSITIVOS --------
    # 'GUIA' o 'GUA' (OCR) cerca de DESPACH (24 chars máx entre medio)
    if re.search(r'\bGUI?A\b.{0,24}\bDESPACH', t) or re.search(r'\bDESPACH.{0,24}\bGUI?A\b', t):
        score += 3

    # 'GUIA DE TRASLADO'
    if re.search(r'\bGUI?A\s+DE\s+TRASLADO\b', t):
        score += 3

    # Abreviaturas GDE / G.D.E. (Guía de Despacho Electrónica)
    if re.search(r'\bG\.?\s*D\.?\s*E\.?\b', t):
        score += 2

    # DTE 52 = Guía de Despacho
    if re.search(r'\bDTE\s*52\b', t) or re.search(r'\bTIPO\s*D(OC(UMENTO)?)?\s*:?\s*52\b', t):
        score += 4

    # 'GUIA ... ELECTRONICA'
    if re.search(r'\bGUI?A\b.*\bELECTRONIC', t):
        score += 1

    # **NUEVOS POSITIVOS** para tus casos
    # 'DE DESPACHO ELECTRONICA' (a veces se corta la palabra 'GUIA' en OCR)
    if re.search(r'\bDE\s+DESPACHO\s+ELECTRONIC', t):
        score += 3

    # 'SOLO TRASLADO' es muy característico de guías
    if re.search(r'\bSOLO\s+TRASLADO\b', t):
        score += 3

    # 'TRASLADO' a secas (apoyo)
    if re.search(r'\bTRASLADO\b', t):
        score += 1

    # -------- NEGATIVOS --------
    if re.search(r'\bFACTURA\b', t):
        score -= 3
    if re.search(r'\bDTE\s*(33|34)\b', t) or re.search(r'\bTIPO\s*D(OC(UMENTO)?)?\s*:?\s*(33|34)\b', t):
        score -= 4
    if re.search(r'\bNOTA\s+DE\s+CREDITO\b', t):
        score -= 3
    if re.search(r'\bNOTA\s+DE\s+DEBITO\b', t):
        score -= 3

    # Frases típicas de factura que no deben gatillar guía
    if re.search(r'\b(DIRECCION|LUGAR|FECHA)\s+DE\s+DESPACH', t):
        score -= 2

    return score >= 3


# ===================== Pipeline principal por archivo =====================
def procesar_archivo(pdf_path):
    """
    Pipeline de 1 PDF (rápido/robusto):
      1) Espera breve si el archivo aún se está escribiendo.
      2) PDF -> Imagen (solo pág.1, DPI ajustable).
      3) OCR header (con auto-rotación y recorte interno).
      4) Reglas: USO ATM / GUÍA DESPACHO.
      5) Extracción RUT/folio y clasificación Cliente/Proveedores o No_Reconocidos.
      6) Compresión opcional (Ghostscript).
      7) Renombrado final (con reintentos).
    """
    import os, re, time, shutil, traceback
    from datetime import datetime
    from pdf2image import convert_from_path
    # from PIL import Image

    modo_debug = is_debug()
    nombre     = os.path.basename(pdf_path)
    registrar_log(f"📄 Entrada: {nombre}")

    # ---------------- helpers rápidos ----------------
    def _norm_rut(s: str) -> str:
        return re.sub(r'[^0-9Kk]', '', s or '').upper()

    def _fast_move(src: str, dst: str):
        try:
            os.replace(src, dst)   # más rápido si es mismo volumen
        except Exception:
            shutil.move(src, dst)

    def _wait_until_stable(path: str, timeout=3.0, step=0.15):
        """Evita leer PDFs aún en escritura (scanner/copias de red)."""
        end = time.time() + timeout
        try:
            last = (os.path.getsize(path), os.path.getmtime(path))
            while time.time() < end:
                time.sleep(step)
                cur = (os.path.getsize(path), os.path.getmtime(path))
                if cur == last:
                    return True
                last = cur
        except Exception:
            pass
        return True

    # Prepara rutas DEBUG (solo para recortes/rotadas del header)
    nombre_base   = os.path.splitext(nombre)[0]
    base_dir      = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)
    ruta_debug_dir= ensure_dir(os.path.join(base_dir, "debug"))
    ruta_recorte  = os.path.join(ruta_debug_dir, f"{nombre_base}_recorte.png") if modo_debug else None

    # Normaliza RUT empresa una sola vez
    RUT_EMP_NORM = _norm_rut(RUT_EMPRESA)

    # ------------- 0) esperar si el archivo aún se vuelca -------------
    _wait_until_stable(pdf_path)

    # ------------- 1) PDF → Imagen (pág.1, DPI ajustable) -------------
    # 👉 Ajusta este DPI si quieres más/menos velocidad/calidad del header:
    OCR_DPI = 300
    try:
        # Nota: ya añadiste Poppler al PATH; no hace falta poppler_path=...
        imagenes = convert_from_path(
            pdf_path,
            dpi=OCR_DPI,
            fmt="jpeg",
            thread_count=1,
            first_page=1,
            last_page=1
        )
        if not imagenes:
            registrar_log_proceso(f"❌ No se pudo rasterizar {nombre}.")
            return

        # Sin filtros pesados: el preprocesado lo hace el OCR (crop+gris+autocontraste)
        imagen = imagenes[0]
        # Copia independiente (por si `convert_from_path` devuelve objeto con recursos compartidos)
        imagen = imagen.copy()
        del imagenes
    except Exception as e:
        registrar_log_proceso(f"❌ Error rasterizando {nombre}:\n{traceback.format_exc()}")
        return

    # -------- 2) OCR header (usa recorte interno + auto-rotación) --------
    try:
        # Pasamos el PIL.Image directamente y, si DEBUG, solo guardamos el recorte de cabecera:
        texto = ocr_zona_factura_desde_png(imagen, ruta_debug=ruta_recorte)
    except Exception as e:
        registrar_log_proceso(f"⚠️ Error OCR ({nombre}): {e}")
        return
    finally:
        try:
            imagen.close()
        except Exception:
            pass

    from ocr.ocr_utils import looks_like_chep
        # -------- 2.5) Regla especial: CHEP --------
    try:
        if looks_like_chep(texto):
            # Subcarpeta fija "chep" dentro de la Carpeta de Salida
            destino_dir = ensure_dir(os.path.join(CARPETA_SALIDA, "chep"))

            # nombre simple; si prefieres, puedes reutilizar tu patrón base_name
            base_name    = f"{SUCURSAL}_CHEP_{datetime.now():%Y%m%d_%H%M%S}"
            nombre_final = generar_nombre_incremental(destino_dir, base_name, ".pdf")
            ruta_destino = os.path.join(destino_dir, nombre_final)

            _fast_move(pdf_path, ruta_destino)

            # Compresión opcional (igual que el resto del flujo)
            if COMPRIMIR_PDF and GS_PATH:
                try:
                    comprimir_pdf(GS_PATH, ruta_destino, calidad=CALIDAD_PDF, dpi=DPI_PDF, tamano_pagina='a4')
                except Exception as e:
                    registrar_log_proceso(f"⚠️ Compresión CHEP fallida: {ruta_destino} | {e}")

            registrar_log(f"🏷️ CHEP detectado → '{destino_dir}'. Guardado: {nombre_final}")
            return ruta_destino
    except Exception as e:
        registrar_log_proceso(f"❗ Error en regla CHEP: {e}")

    # -------- 3) Regla especial: USO ATM --------
    texto_upper = " ".join((texto or "").upper().split())
    if "USO ATM" in texto_upper:
        try:
            if CARPETA_SALIDA_USO_ATM and os.path.isabs(CARPETA_SALIDA_USO_ATM) and os.path.isdir(CARPETA_SALIDA_USO_ATM):
                destino_dir = CARPETA_SALIDA_USO_ATM
                origen      = "config"
            else:
                destino_dir = ensure_dir(FALLBACK_USO_ATM_DIR)
                origen      = "fallback C:"

            base_name    = f"Recibo_Valores_{datetime.now():%Y%m%d_%H%M%S}"
            nombre_final = generar_nombre_incremental(destino_dir, base_name, ".pdf")
            ruta_destino = os.path.join(destino_dir, nombre_final)
            _fast_move(pdf_path, ruta_destino)

            if COMPRIMIR_PDF and GS_PATH:
                try:
                    comprimir_pdf(GS_PATH, ruta_destino, calidad=CALIDAD_PDF, dpi=DPI_PDF, tamano_pagina='a4')
                except Exception as e:
                    registrar_log_proceso(f"⚠️ Compresión fallida: {ruta_destino} | {e}")

            registrar_log(f"📥 'USO ATM' → {destino_dir} ({origen}). Guardado: {nombre_final}")
            return ruta_destino
        except Exception as e:
            registrar_log_proceso(f"❗ Error moviendo 'USO ATM': {e}")
            try:
                no_rec = ensure_dir(os.path.join(CARPETA_SALIDA, "No_Reconocidos"))
                base_error = f"Recibo_Valores_{datetime.now():%Y%m%d_%H%M%S}"
                nombre_fallo = generar_nombre_incremental(no_rec, base_error, ".pdf")
                ruta_fallo   = os.path.join(no_rec, nombre_fallo)
                _fast_move(pdf_path, ruta_fallo)
                registrar_log_proceso(f"⚠️ 'USO ATM' → No_Reconocidos. Guardado: {nombre_fallo}")
                return ruta_fallo
            except Exception as e2:
                registrar_log_proceso(f"❌ Falla secundaria moviendo a No_Reconocidos: {e2}")
                return

    # -------- 3.6) Regla especial: Guía de despacho --------
    if _es_guia_despacho(texto):
        try:
            destino_dir = obtener_carpeta_salida_anual(os.path.join(CARPETA_SALIDA, "guias de despachos"))
            mkdir(destino_dir)

            rut_proveedor    = extraer_rut(texto) or "desconocido"
            numero_documento = extraer_numero_factura(texto) or ""
            anio             = datetime.now().strftime("%Y")

            rut_nombre   = rut_proveedor if rut_proveedor != "desconocido" else "noreconocido"
            folio_nombre = numero_documento if numero_documento else "noreconocido"
            base_name    = f"{SUCURSAL}_{rut_nombre}_guia_{folio_nombre}_{anio}"

            nombre_final = generar_nombre_incremental(destino_dir, base_name, ".pdf")
            ruta_destino = os.path.join(destino_dir, nombre_final)
            _fast_move(pdf_path, ruta_destino)

            if COMPRIMIR_PDF and GS_PATH:
                try:
                    comprimir_pdf(GS_PATH, ruta_destino, calidad=CALIDAD_PDF, dpi=DPI_PDF, tamano_pagina='a4')
                except Exception as e:
                    registrar_log_proceso(f"⚠️ Compresión fallida guía: {ruta_destino} | {e}")

            registrar_log(f"📦 Guía detectada → '{destino_dir}' como: {nombre_final}")
            return ruta_destino
        except Exception as e:
            registrar_log_proceso(f"❗ Error moviendo guía de despacho: {e}")
            # si falla, seguimos con flujo normal

    # -------- 4) Extraer RUT/Folio y armar nombre base --------
    rut_proveedor  = extraer_rut(texto)
    numero_factura = extraer_numero_factura(texto)
    anio           = datetime.now().strftime("%Y")

    rut_valido     = bool(rut_proveedor and rut_proveedor != "desconocido")
    folio_valido   = bool(numero_factura)
    rut_nombre     = rut_proveedor if rut_valido else "noreconocido"
    folio_nombre   = numero_factura if folio_valido else "noreconocido"
    base_name      = f"{SUCURSAL}_{rut_nombre}_factura_{folio_nombre}_{anio}"

    # -------- 5) No_Reconocidos si falta dato clave --------
    if not (rut_valido and folio_valido):
        no_rec       = ensure_dir(os.path.join(CARPETA_SALIDA, "No_Reconocidos"))
        nombre_final = generar_nombre_incremental(no_rec, base_name, ".pdf")
        ruta_destino = os.path.join(no_rec, nombre_final)
        _fast_move(pdf_path, ruta_destino)

        if COMPRIMIR_PDF and GS_PATH:
            try:
                comprimir_pdf(GS_PATH, ruta_destino, calidad=CALIDAD_PDF, dpi=DPI_PDF, tamano_pagina='a4')
                
                #LOG no reconocido
                registrar_log_proceso(f"Doc No reconocido Calidad: {CALIDAD_PDF} y dpi: {DPI_PDF}")
            except Exception as e:
                registrar_log_proceso(f"⚠️ Compresión fallida (No_Reconocidos): {ruta_destino} | {e}")

        motivo = []
        if not rut_valido:   motivo.append("RUT no reconocido")
        if not folio_valido: motivo.append("N° factura no reconocido")
        registrar_log(f"⚠️ No_Reconocidos: {nombre_final} | Motivo: {', '.join(motivo)}")
        return ruta_destino

    # -------- 6) Clasificación Cliente / Proveedores --------
    subcarpeta       = "Cliente" if _norm_rut(rut_proveedor) == RUT_EMP_NORM else "Proveedores"
    carpeta_clase    = os.path.join(CARPETA_SALIDA, subcarpeta)
    carpeta_anual    = obtener_carpeta_salida_anual(carpeta_clase)
    mkdir(carpeta_anual)

    # nombre temporal para evitar colisiones durante compresión
    temp_nombre = f"{base_name}_{datetime.now():%H%M%S%f}"
    temp_ruta   = os.path.join(carpeta_anual, f"{temp_nombre}.pdf")

    try:
        _fast_move(pdf_path, temp_ruta)
    except Exception as e:
        registrar_log_proceso(f"❗ Error moviendo original: {e}")
        return

    # -------- 7) Compresión opcional --------
    if COMPRIMIR_PDF and GS_PATH:
        try:
            comprimir_pdf(GS_PATH, temp_ruta, calidad=CALIDAD_PDF, dpi=DPI_PDF, tamano_pagina='a4')
            #LOG no reconocido
            registrar_log_proceso(f"Doc Calidad: {CALIDAD_PDF} y dpi: {DPI_PDF}")
        except Exception as e:
            registrar_log_proceso(f"⚠️ Compresión fallida: {temp_ruta}. Se deja sin comprimir. Detalle: {e}")

    # -------- 8) Renombrado final seguro --------
    try:
        for _ in range(6):
            nombre_final = generar_nombre_incremental(carpeta_anual, base_name, ".pdf")
            ruta_destino = os.path.join(carpeta_anual, nombre_final)
            if not os.path.exists(ruta_destino):
                os.rename(temp_ruta, ruta_destino)
                return ruta_destino
            time.sleep(0.15)
    except Exception as e:
        fallback_name = f"{base_name}_backup_{datetime.now():%H%M%S%f}.pdf"
        fallback_path = os.path.join(carpeta_anual, fallback_name)
        try:
            _fast_move(temp_ruta, fallback_path)
        except Exception:
            pass
        registrar_log_proceso(f"❗ Renombrado fallido. Guardado como: {fallback_name} | {e}")
        return fallback_path

# ===================== Procesamiento por carpeta (multi-hilo) =====================

def procesar_entrada_una_vez():
    """
    Procesa TODOS los PDFs de CARPETA_ENTRADA una sola vez con mejor tiempo de arranque:
      - Arranca ya con un burst inicial (sin ordenar) para dar feedback inmediato.
      - Ordena el resto por fecha de modificación (antiguos primero).
      - Usa ThreadPoolExecutor con hasta 8 hilos (o núcleos de CPU, lo que sea menor).
      - Muestra un messagebox al finalizar con el tiempo total.
    """
    import time, itertools, os, tkinter as tk
    from concurrent.futures import ThreadPoolExecutor, as_completed
    from tkinter import messagebox

    inicio = time.perf_counter()

    # (Opcional pero útil) Asegura que el OCR esté cargado antes de lanzar hilos
    try:
        from ocr.ocr_utils import inicializar_ocr
        inicializar_ocr()
    except Exception:
        pass

    # Config de concurrencia
    nucleos = os.cpu_count() or 1
    max_hilos = min(nucleos, 8)
    burst = max_hilos * 2  # primeros N archivos "ya" sin ordenar

    registrar_log_proceso(f"🧠 Núcleos detectados: {nucleos} | Hilos usados: {max_hilos}")
    print("🔍 Buscando documentos en la carpeta de entrada...")

    # Generador rápido con os.scandir (más veloz que listdir + joins)
    def _iter_pdf_entries(dirname):
        with os.scandir(dirname) as it:
            for e in it:
                try:
                    if e.is_file() and e.name.lower().endswith(".pdf"):
                        yield e
                except Exception:
                    # Si no podemos stat/leer una entry, seguimos
                    continue

    entries_iter = _iter_pdf_entries(CARPETA_ENTRADA)
    primeros = list(itertools.islice(entries_iter, burst))
    resto = list(entries_iter)

    if not primeros and not resto:
        # Sin documentos
        try:
            root = tk.Tk(); root.withdraw()
            messagebox.showinfo("Sin documentos", "No se encontraron documentos pendientes en la carpeta de entrada.")
            print("Sin documentos pendientes")
            root.destroy()
        except Exception:
            print("Sin documentos pendientes")
        return

    total = len(primeros) + len(resto)
    print(f"🗂️ Encontrados: {total} documento(s) PDF.")

    # Ordena el resto por mtime (antiguos primero)
    if resto:
        try:
            resto.sort(key=lambda de: (de.stat().st_mtime, de.name))
        except Exception:
            # Si stat falla para alguno, caemos a ordenar sólo por nombre
            resto.sort(key=lambda de: de.name)

    # Pool de hilos para procesar en paralelo
    procesados = 0
    with ThreadPoolExecutor(max_workers=max_hilos) as executor:
        futures = {}

        # Lanza inmediatamente el burst inicial (sin ordenar) para feedback rápido
        for e in primeros:
            futures[executor.submit(procesar_archivo, e.path)] = e.path

        # Luego lanza el resto ya ordenado cronológicamente
        for e in resto:
            futures[executor.submit(procesar_archivo, e.path)] = e.path

        # Consume a medida que terminen (no en orden de envío)
        for fut in as_completed(futures):
            path = futures[fut]
            nombre = os.path.basename(path)
            procesados += 1
            try:
                resultado = fut.result()
                if resultado:
                    print(f"{procesados}/{total} ✅ Procesado: {os.path.basename(resultado)}")
                    registrar_log(f"✅ Procesado: {os.path.basename(resultado)}")
                else:
                    print(f"{procesados}/{total} ⚠️ Procesado con advertencias: {nombre}")
            except Exception as e:
                registrar_log_proceso(f"❌ Error procesando archivo {nombre}: {e}")

    # Informe final de duración total
    duracion = time.perf_counter() - inicio
    minutos = int(duracion // 60)
    segundos = int(duracion % 60)
    try:
        root = tk.Tk(); root.withdraw()
        messagebox.showinfo("Finalizado", f"✅ Procesamiento completado.\nTiempo total: {minutos} min {segundos} seg.")
        root.destroy()
    except Exception:
        print(f"✅ Procesamiento completado en {minutos} min {segundos} seg.")
