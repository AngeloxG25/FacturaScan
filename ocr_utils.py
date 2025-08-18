import hide_subprocess
import re
import os,sys
import io
import logging
import contextlib
from datetime import datetime
import itertools
# --- Sube estos imports a nivel de módulo (arriba del archivo) ---
import os, sys
from datetime import datetime
from PIL import Image, ImageOps
import numpy as np

from log_utils import registrar_log_proceso, is_debug
# Palabras clave precompiladas (mayúsculas)
_PALABRAS_CLAVE = {"RUT", "FACTURA", "ELECTRONICA", "NRO", "SII"}

# Mapea ángulos a transposes rápidos (counter-clockwise, igual que PIL)
_TRANSPOSE_POR_ANGULO = {
    0:  None,  # sin cambio
    90: Image.ROTATE_90,
    180: Image.ROTATE_180,
    270: Image.ROTATE_270,
}

# contador global para nombres únicos de debug
DEBUG_COUNTER = itertools.count(1)

def _unique_path(candidate: str) -> str:
    """Devuelve una ruta única (agrega _1, _2, ... si ya existe)."""
    base, ext = os.path.splitext(candidate)
    if not os.path.exists(candidate):
        return candidate
    i = 1
    while True:
        alt = f"{base}_{i}{ext}"
        if not os.path.exists(alt):
            return alt
        i += 1

def _is_dir_like(p: str) -> bool:
    """True si p es carpeta existente o termina con separador o no tiene extensión ."""
    if not p:
        return False
    if os.path.isdir(p):
        return True
    if p.endswith(os.sep):
        return True
    root, ext = os.path.splitext(p)
    return ext == ""  # lo tratamos como carpeta si no tiene extensión


# Silenciar advertencias GPU de torch/easyocr
os.environ["CUDA_VISIBLE_DEVICES"] = "-1"
logging.getLogger("torch").setLevel(logging.ERROR)

# evitar mensaje de Pytorch y EasyOCR
import warnings
warnings.filterwarnings("ignore")
logging.getLogger("torch").setLevel(logging.ERROR)

# Cargar EasyOCR sin mostrar advertencias
with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
    import easyocr

# Precarga global del modelo EasyOCR (una sola vez)
import threading
reader = None
reader_lock = threading.Lock()

def inicializar_ocr():
    global reader
    if reader is None:
        with reader_lock:
            if reader is None:
                import sys
                sys.stdout = open(os.devnull, 'w')
                sys.stderr = open(os.devnull, 'w')
                # reader = easyocr.Reader(['es'], gpu=False)
                reader = easyocr.Reader(['es'], gpu=False, verbose=False)
                sys.stdout = sys.__stdout__
                sys.stderr = sys.__stderr__

inicializar_ocr()

# def ocr_zona_factura_desde_png(imagen_entrada, ruta_debug=None):
#     """
#     Detecta automáticamente la orientación del documento y realiza OCR en la zona superior derecha.
#     - Si el modo debug (is_debug()) está ACTIVO:
#         * Guarda el recorte elegido en 'ruta_debug' (si se entrega) o en una ruta auto-generada ./debug/
#         * Si hubo rotación automática, también guarda una copia de la imagen rotada.
#     - Si el modo debug está INACTIVO:
#         * Procesa normalmente sin guardar archivos de depuración.
#     """
#     from PIL import Image
#     import numpy as np
#     import os
#     from datetime import datetime

#     # Cargar imagen entrada (ruta o PIL.Image)
#     if isinstance(imagen_entrada, str):
#         imagen_original = Image.open(imagen_entrada)
#         nombre_base = os.path.splitext(os.path.basename(imagen_entrada))[0]
#     elif hasattr(imagen_entrada, "crop"):
#         imagen_original = imagen_entrada
#         nombre_base = "imagen_en_memoria"
#     else:
#         raise ValueError("imagen_entrada debe ser una ruta o un objeto PIL.Image")

# # --- Configuración de rutas solo si estamos en DEBUG ---
#     debug_activo = is_debug()
#     ruta_debug_final = None

#     if debug_activo:
#         base_app = os.path.dirname(os.path.abspath(sys.argv[0]))
#         carpeta_por_defecto = os.path.join(base_app, "debug")
#         os.makedirs(carpeta_por_defecto, exist_ok=True)

#         # timestamp + contador para evitar colisiones dentro del mismo segundo
#         ts = datetime.now().strftime("%Y%m%d_%H%M%S")
#         seq = next(DEBUG_COUNTER)

#         if not ruta_debug:
#             # ruta automática
#             ruta_debug_final = os.path.join(
#                 carpeta_por_defecto,
#                 f"{nombre_base}_recorte_{ts}_{seq}.png")
#         else:
#             if _is_dir_like(ruta_debug):
#                 # si nos pasaron una carpeta o algo "tipo carpeta", armamos nombre dentro
#                 carpeta_dest = ruta_debug if os.path.isabs(ruta_debug) else os.path.join(base_app, ruta_debug)
#                 os.makedirs(carpeta_dest, exist_ok=True)
#                 ruta_debug_final = os.path.join(
#                     carpeta_dest,
#                     f"{nombre_base}_recorte_{ts}_{seq}.png")
#             else:
#                 # es una ruta de archivo -> garantizar unicidad
#                 # si es relativa, resolverla respecto a la app
#                 ruta_absoluta = ruta_debug if os.path.isabs(ruta_debug) else os.path.join(base_app, ruta_debug)
#                 os.makedirs(os.path.dirname(ruta_absoluta), exist_ok=True)
#                 # si el llamador puso un nombre fijo, lo hacemos único
#                 ruta_debug_final = _unique_path(ruta_absoluta)

#     # --- Búsqueda del mejor ángulo y recorte ---
#     mejor_texto = ""
#     mejor_puntaje = 0
#     mejor_recorte = None
#     mejor_angulo = 0

#     for angulo in [0, 90, 180, 270]:
#         imagen_rotada = imagen_original.rotate(angulo, expand=True)
#         ancho, alto = imagen_rotada.size

#         # Zona superior derecha (30% vertical, 35% horizontal final)
#         recorte = imagen_rotada.crop((
#             int(ancho * 0.65),
#             int(alto * 0.01),
#             int(ancho * 1.00),
#             int(alto * 0.30)))

#         # Reducción ligera para favorecer EasyOCR en ruido/escala
#         zona_reducida = recorte.resize(
#             (max(1, recorte.width // 2), max(1, recorte.height // 2)),
#             resample=Image.BICUBIC)
#         zona_np = np.array(zona_reducida)

#         texto = reader.readtext(zona_np, detail=0, batch_size=1)
#         texto_completo = " ".join(texto).strip()

#         # Heurística simple por palabras clave
#         palabras_clave = sum(
#             1 for palabra in texto_completo.upper().split()
#             if palabra in ["RUT", "FACTURA", "ELECTRONICA", "NRO", "SII"])

#         if palabras_clave > mejor_puntaje:
#             mejor_puntaje = palabras_clave
#             mejor_texto = texto_completo
#             mejor_recorte = recorte
#             mejor_angulo = angulo

#     # --- Guardados de depuración SOLO si el modo debug está activo ---
#     if debug_activo:
#         try:
#             if mejor_angulo != 0:
#                 registrar_log_proceso(f"🔁 Imagen rotada automáticamente {mejor_angulo}°")
#                 # Guardar la versión rotada (solo si tenemos ruta base para recorte)
#                 if ruta_debug_final:
#                     root, ext = os.path.splitext(ruta_debug_final)
#                     ruta_rotada_base = f"{root.rsplit('_recorte', 1)[0]}_rotada{mejor_angulo}{ext}"
#                     ruta_rotada = _unique_path(ruta_rotada_base)
#                     imagen_original.rotate(mejor_angulo, expand=True).save(ruta_rotada)

#             if ruta_debug_final and mejor_recorte is not None:
#                 mejor_recorte.save(ruta_debug_final)
#                 registrar_log_proceso(f"📎 Recorte guardado en: {ruta_debug_final}")
#         except Exception as e:
#             registrar_log_proceso(f"⚠️ Error guardando recortes de debug: {e}")

#     # --- Retorno estándar: solo el texto detectado (sin importar el modo) ---
#     return mejor_texto

def ocr_zona_factura_desde_png(imagen_entrada, ruta_debug=None, early_threshold=3):
    """
    Detecta automáticamente la orientación del documento y realiza OCR en la zona superior derecha.
    Optimizada para velocidad:
      - Usa transpose en vez de rotate expand (más rápido).
      - Trabaja siempre en escala de grises sobre el recorte.
      - Auto-contraste local del recorte para mejorar OCR sin filtros costosos.
      - Heurística con 'salida temprana' si el puntaje supera early_threshold.
      - Evita conversiones a NumPy innecesarias (una sola por intento).

    Parámetros:
      imagen_entrada: str (ruta) o PIL.Image
      ruta_debug: str|None
      early_threshold: int  -> si alcanza este puntaje de palabras clave, corta el bucle.
    """
    # Cargar imagen
    if isinstance(imagen_entrada, str):
        imagen_original = Image.open(imagen_entrada)
        nombre_base = os.path.splitext(os.path.basename(imagen_entrada))[0]
    elif hasattr(imagen_entrada, "crop"):
        imagen_original = imagen_entrada
        nombre_base = "imagen_en_memoria"
    else:
        raise ValueError("imagen_entrada debe ser una ruta o un objeto PIL.Image")

    # Config debug (idéntico a tu versión, pero sin imports internos)
    debug_activo = is_debug()
    ruta_debug_final = None
    if debug_activo:
        base_app = os.path.dirname(os.path.abspath(sys.argv[0]))
        carpeta_por_defecto = os.path.join(base_app, "debug")
        os.makedirs(carpeta_por_defecto, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        seq = next(DEBUG_COUNTER)
        if not ruta_debug:
            ruta_debug_final = os.path.join(carpeta_por_defecto, f"{nombre_base}_recorte_{ts}_{seq}.png")
        else:
            if _is_dir_like(ruta_debug):
                carpeta_dest = ruta_debug if os.path.isabs(ruta_debug) else os.path.join(base_app, ruta_debug)
                os.makedirs(carpeta_dest, exist_ok=True)
                ruta_debug_final = os.path.join(carpeta_dest, f"{nombre_base}_recorte_{ts}_{seq}.png")
            else:
                ruta_absoluta = ruta_debug if os.path.isabs(ruta_debug) else os.path.join(base_app, ruta_debug)
                os.makedirs(os.path.dirname(ruta_absoluta), exist_ok=True)
                ruta_debug_final = _unique_path(ruta_absoluta)

    # Búsqueda del mejor ángulo
    mejor_texto = ""
    mejor_puntaje = -1
    mejor_recorte = None
    mejor_angulo = 0

    # Precalcular tamaño original
    ancho0, alto0 = imagen_original.size

    # Itera sobre ángulos usando transposes (mucho más rápidos que rotate+expand)
    for angulo in (0, 90, 180, 270):
        tr_op = _TRANSPOSE_POR_ANGULO[angulo]
        if tr_op is None:
            img = imagen_original
        else:
            # transpose no requiere expand y es más eficiente para giros múltiples de 90°
            img = imagen_original.transpose(tr_op)

        ancho, alto = img.size

        # Recorte superior derecho (igual a tu lógica)
        x0 = int(ancho * 0.65)
        y0 = int(alto * 0.01)
        x1 = int(ancho * 1.00)
        y1 = int(alto * 0.30)
        recorte = img.crop((x0, y0, x1, y1))

        # Convertir a escala de grises y auto-contraste local (rápido)
        recorte = ImageOps.grayscale(recorte)
        # Pequeña reducción (factor 0.5) con LANCZOS para reducir ruido sin perder nitidez
        if recorte.width > 2 and recorte.height > 2:
            recorte = recorte.resize((recorte.width // 2, recorte.height // 2), Image.LANCZOS)

        # Auto-contraste (limita recortes extremos)
        recorte = ImageOps.autocontrast(recorte, cutoff=1)

        # A NumPy (EasyOCR espera RGB/GRAYSCALE; uint8)
        zona_np = np.array(recorte, dtype=np.uint8)

        # OCR: limitar caracteres acelera el decoder (ajusta si afecta tus keywords)
        # Mantengo letras, números y separadores comunes en facturas chilenas.
        texto = reader.readtext(
            zona_np,
            detail=0,
            batch_size=1,
            allowlist="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-./#:() ",
            # parámetros que evitan trabajo adicional:
            mag_ratio=1.0,      # evita magnificación costosa
            width_ths=0.6,      # segmentación menos agresiva
            slope_ths=0.999     # desactiva corrección de inclinación (ya rotamos 90° c/caso)
        )
        texto_completo = " ".join(texto).strip()
        if texto_completo:
            # Heurística de score por palabras clave
            palabras = texto_completo.upper().split()
            puntaje = sum(1 for p in palabras if p in _PALABRAS_CLAVE)
        else:
            puntaje = 0

        if puntaje > mejor_puntaje:
            mejor_puntaje = puntaje
            mejor_texto = texto_completo
            mejor_recorte = recorte
            mejor_angulo = angulo

            # Salida temprana: si ya tenemos suficientes señales, no pruebes más ángulos
            if mejor_puntaje >= early_threshold:
                break

    # Debug (igual que tu versión, pero usando el recorte ya en gris)
    if debug_activo:
        try:
            if mejor_angulo != 0:
                registrar_log_proceso(f"🔁 Imagen rotada automáticamente {mejor_angulo}°")
                if ruta_debug_final:
                    root, ext = os.path.splitext(ruta_debug_final)
                    ruta_rotada_base = f"{root.rsplit('_recorte', 1)[0]}_rotada{mejor_angulo}{ext}"
                    ruta_rotada = _unique_path(ruta_rotada_base)
                    # Guarda la versión rotada del original (usar transpose rápido)
                    tr_op = _TRANSPOSE_POR_ANGULO[mejor_angulo]
                    img_rotada = imagen_original if tr_op is None else imagen_original.transpose(tr_op)
                    img_rotada.save(ruta_rotada)

            if ruta_debug_final and mejor_recorte is not None:
                mejor_recorte.save(ruta_debug_final)
                registrar_log_proceso(f"📎 Recorte guardado en: {ruta_debug_final}")
        except Exception as e:
            registrar_log_proceso(f"⚠️ Error guardando recortes de debug: {e}")

    return mejor_texto


# def extraer_rut(texto):
#     # print("🟡 Texto OCR original (RUT):\n", texto)
#     texto_original = texto

#     reemplazos = {
#         "RUT.": "RUT", "R.U.T.": "RUT", "R-U-T": "RUT", "RUT:": "RUT", "RUT;": "RUT",
#         "RUT=": "RUT", "RU.T": "RUT", "RU:T": "RUT", "R:UT": "RUT", "RU.T.": "RUT",
#         "RUI": "RUT", "RU1": "RUT", "R.UT.": "RUT", "RuT;": "RUT", "RUTTT;": "RUT",
#         "Ru:,n.": "RUT", "Ru.t:": "RUT", "RVT ;": "RUT", "RVT ": "RUT", "RVT": "RUT",
#         "RUT.:": "RUT","R.UT.:": "RUT",        
#     }

#     for k, v in reemplazos.items():
#         texto = texto.replace(k, v)

#     texto = texto.replace(',', '.')
#     texto = texto.replace('O', '0').replace('o', '0').replace('I', '1').replace('l', '1')
#     texto = texto.replace('B', '8').replace('Z', '2').replace('G', '6')
#     texto = texto.replace('–', '-').replace('—', '-').replace('‐', '-')
#     texto = texto.replace('+', '-')
#     # print("🟢 Texto tras limpieza final (RUT):\n", texto)

#     posibles = re.findall(r'\d{1,2}\s*[\.,]?\s*\d{3}\s*[\.,]?\s*\d{3}\s*[-‐–—]?\s*[\dkK]', texto)
#     if posibles:
#         rut = posibles[0]
#         rut = re.sub(r'[^0-9kK]', '', rut[:-1]) + '-' + rut[-1].upper()
#         return rut

#     posibles2 = re.findall(r'(\d{1,2})[^\d]{0,2}(\d{3})[^\d]{0,2}(\d{3})[^\dkK]{0,2}([\dkK])', texto_original)
#     if posibles2:
#         rut = f"{posibles2[0][0]}{posibles2[0][1]}{posibles2[0][2]}-{posibles2[0][3].upper()}"
#         return rut

#     registrar_log_proceso("⚠️ RUT no detectado.")
#     return "desconocido"

import re
from log_utils import registrar_log_proceso

def extraer_rut(texto):
    texto_original = texto
    # print('texto original: \n', texto)

    # Normalización de variaciones comunes de "RUT"
    reemplazos = {
        "RUT.": "RUT", "R.U.T.": "RUT", "R-U-T": "RUT", "RUT:": "RUT", "RUT;": "RUT",
        "RUT=": "RUT", "RU.T": "RUT", "RU:T": "RUT", "R:UT": "RUT", "RU.T.": "RUT",
        "RUI": "RUT", "RU1": "RUT", "R.UT.": "RUT", "RuT;": "RUT", "RUTTT;": "RUT",
        "Ru:,n.": "RUT", "Ru.t:": "RUT", "RVT ;": "RUT", "RVT ": "RUT", "RVT": "RUT",
        "RUT.:": "RUT","R.UT.:": "RUT",
    }

    for k, v in reemplazos.items():
        texto = texto.replace(k, v)

    # Limpieza de caracteres confusos
    texto = texto.replace(',', '.')
    texto = texto.replace('O', '0').replace('o', '0').replace('I', '1').replace('l', '1')
    texto = texto.replace('B', '8').replace('Z', '2').replace('G', '6')
    texto = texto.replace('–', '-').replace('—', '-').replace('‐', '-')
    texto = texto.replace('+', '-')

    # print('texto limpiado: \n', texto)

    # --- Función para calcular DV ---
    def calcular_dv(rut_sin_dv: str) -> str:
        try:
            rut = list(map(int, rut_sin_dv[::-1]))
        except ValueError:
            return ""
        factores = [2, 3, 4, 5, 6, 7]
        suma = 0
        for i, d in enumerate(rut):
            suma += d * factores[i % len(factores)]
        resto = 11 - (suma % 11)
        if resto == 11:
            return "0"
        if resto == 10:
            return "K"
        return str(resto)

    # Lista de candidatos
    candidatos = []

    # 1) Caso normal con DV explícito (tolerante a puntos y espacios en los bloques)
    posibles = re.findall(
        r'(\d{1,2}(?:[\s\.]?\d{3}){2})\s*[-‐–—]?\s*([\dkK])',
        texto
    )
    for cuerpo, dv in posibles:
        rut_sin = re.sub(r'\D', '', cuerpo)
        dv = dv.upper()
        dv_calc = calcular_dv(rut_sin)
        if dv == dv_calc:
            candidatos.append((rut_sin + "-" + dv, "DV válido"))
            # print(f"🟢 Detectado RUT con espacios/puntos: {cuerpo.strip()} → {rut_sin}-{dv}")
        else:
            registrar_log_proceso(
                f"⚠️ RUT detectado con DV incorrecto: {rut_sin}-{dv} (debería ser {dv_calc})"
            )

    # 2) Variante tolerante sobre el texto original (más flexible)
    posibles2 = re.findall(
        r'(\d{1,2})[^\d]{0,2}(\d{3})[^\d]{0,2}(\d{3})[^\dkK]{0,2}([\dkK])',
        texto_original
    )
    for p in posibles2:
        rut_sin = f"{p[0]}{p[1]}{p[2]}"
        dv = p[3].upper()
        dv_calc = calcular_dv(rut_sin)
        if dv == dv_calc:
            candidatos.append((rut_sin + "-" + dv, "DV válido (tolerante)"))
        else:
            registrar_log_proceso(
                f"⚠️ RUT tolerante con DV incorrecto: {rut_sin}-{dv} (debería ser {dv_calc})"
            )

    # 3) Caso RUT sin DV explícito (se calcula)
    posibles3 = re.findall(r'RUT\s*:?[\s]*([0-9\.]{7,12})(?![-0-9Kk])', texto)
    for p in posibles3:
        rut_sin = re.sub(r'\D', '', p)  # quitar puntos
        if 7 <= len(rut_sin) <= 8:      # rango válido
            dv_calc = calcular_dv(rut_sin)
            if dv_calc:
                rut_final = f"{rut_sin}-{dv_calc}"
                candidatos.append((rut_final, "DV calculado"))
                print(f"🟡 Detectado RUT sin DV: {rut_sin} → {rut_final}")

    # Selección final
    if candidatos:
        rut, origen = candidatos[0]
        registrar_log_proceso(f"✅ RUT validado: {rut} ({origen})")
        # print('Rut encontrado: \n',rut)
        return rut

    registrar_log_proceso("⚠️ RUT no detectado.")
    return "desconocido"


# def extraer_numero_factura(texto: str) -> str:
#     print("🟡 Texto OCR original (Número Factura):\n", texto)
#     def corregir_ocr_numero(numero: str) -> str:
#         traduccion = str.maketrans({
#             'O': '0', 'Q': '0', 'B': '8', 'I': '1', 'L': '1', 'S': '5',
#             'Z': '2', 'D': '0', 'E': '8', 'A': '4', 'U': '0'  # <- U como 0
#         })
#         return numero.translate(traduccion).replace('.', '').replace(' ', '')

#     reemplazos = {
#         "N°": "NRO ", "N'": "NRO ", 'N"': "NRO ", "N :": "NRO ", "N.": "NRO ",
#         "Nº": "NRO ", "N:": "NRO ", "NE": "NRO ", "N?": "NRO ", "FNLC": "NRO ",
#         "FNL": "NRO ", "FNLD": "NRO ", "FULD": "NRO ", "FOLIO": "NRO ",
#         "NC:": "NRO:", "NC ": "NRO ", "N C": "NRO ", '"NC': "NRO ", "'NC": "NRO ",
#         "NP ": "NRO ", "N°P": "NRO ", "N P": "NRO ", '"NP': "NRO ", "'NP": "NRO ",
#         "NP:": "NRO:", "Nro.:": "NRO:", "Nro. :": "NRO:", "Nro :": "NRO:",
#         "Nro.": "NRO", "NF": "NRO", "NiP": "NRO", "MP": "NRO", "NO": "NRO",
#         "Nro.  :": "NRO:", "N9": "NRO", "Ne": "NRO", "nE": "NRO", "Nro": "NRO",
#         "Nro ": "NRO", "Nro  ": "NRO", "Folio N?": "NRO", "FOLION?": "NRO ",
#         "FOLIO N°": "NRO ", "FCLIO": "NRO ", "NUMERO": "NRO ",
#         "NUMERO .": "NRO ", "NO": "NRO ", "No": "NRO ","Nra": "NRO ",
#         "Folio /": "NRO ","RUMERO :": "NRO ","Né": "NRO ","N '": "NRO ", 
#     }

#     for k, v in reemplazos.items():
#         texto = texto.replace(k, v)

#     texto = texto.replace('=', ':').replace('f.1', 'ELECTRONICA').replace('f 1', 'ELECTRONICA')
#     texto = texto.replace('Nro =', 'NRO:').replace('Nro :', 'NRO:').replace('Nro::', 'NRO:')
#     texto = texto.replace('N:', 'NRO:').replace('NRO:::', 'NRO:')
#     texto = texto.replace("NRO :", "NRO:").replace("NRO  :", "NRO:")
#     texto = re.sub(r'F[A4@][CÇ][\s\-]*U[R][A4@][\s\-]*E[L1I][E3]C[T7][R][O0][N][I1][C][A4@]', "FACTURA ELECTRONICA", texto, flags=re.IGNORECASE)
#     texto = re.sub(r'F[\s\-]*A[\s\-]*C[\s\-]*U[\s\-]*R[\s\-]*A[\s\-]*[E3][L1I][E3]C[T7][R][O0][N][I1][C][A4@]', "FACTURA ELECTRONICA", texto, flags=re.IGNORECASE)
#     texto = texto.replace("FACTURAELECTRONICA", "FACTURA ELECTRONICA")
#     texto = texto.replace("FACTURAELECTRÓNICA", "FACTURA ELECTRONICA")
#     texto = texto.replace("FACTURAELECTRONICP", "FACTURA ELECTRONICA")
#     texto = texto.replace("FACTVRAELECTRoNICA", "FACTURA ELECTRONICA")
#     texto = texto.replace("FACTURA ELECTRÓNICA", "FACTURA ELECTRONICA")
#     texto = texto.replace("FACTURA ELECTRNICA", "FACTURA ELECTRONICA")
#     texto = texto.replace("FACTURA FLECTRONICA", "FACTURA ELECTRONICA")
#     texto = texto.replace("FAC URA FLECTRONICA", "FACTURA ELECTRONICA")
#     texto = texto.replace("FACTURA ELECIRONICA", "FACTURA ELECTRONICA")
#     texto = texto.replace("FACTURLELECTRONICA", "FACTURA ELECTRONICA")
#     texto = texto.replace("FACTURA Electronica", "FACTURA ELECTRONICA")
#     texto = texto.replace("FACTURA ELECTRNICA FOLIO", "FACTURA ELECTRONICA")
#     texto = texto.replace("FACTURA ELECTRNICA FOLIO NRO", "FACTURA ELECTRONICA")
#     texto = texto.replace("FACTURA ELECTRONICA NO1", "FACTURA ELECTRONICA")
#     texto = texto.upper()
#     texto = re.sub(r'[^\x00-\x7F]+', '', texto)
#     texto = re.sub(
#         r'(NRO|NUMERO)[\s:=\.\-]+([0-9OQBILSZDEUA]{3,20})',
#         lambda m: f"NRO:{corregir_ocr_numero(m.group(2).upper())}",
#         texto, flags=re.IGNORECASE)

#     texto = re.sub(
#         r'(NRO)[\s]{1,5}([0-9OQBILSZDEUA]{1,5})[\s\.]{1,2}([0-9OQBILSZDEUA]{1,5})',
#         lambda m: f"NRO:{corregir_ocr_numero(m.group(2) + m.group(3))}",
#         texto)

#     print("🟢 Texto tras limpieza completa (Número Factura):\n", texto)
#     lineas = texto.splitlines()
#     candidatos = []

#     def es_posible_numero_factura(num: str) -> bool:
#         num = corregir_ocr_numero(num)
#         if re.match(r'^\+?56\s?\d{2}\s?\d{4,}$', num):  # teléfono
#             return False
#         if len(num) < 3 or len(num) > 12:
#             return False
#         if not num.isdigit():
#             return False
#         return True

#     for i, linea in enumerate(lineas):
#         match_exacto = re.search(r'NRO[\s:=\.\-]*([0-9OQBILSZDEUA]{3,20})', linea)
#         if match_exacto:
#             candidato = corregir_ocr_numero(match_exacto.group(1))
#             if es_posible_numero_factura(candidato):
#                 candidatos.append((candidato, "NRO: exacto"))
#             continue

#         match_tolerante = re.search(r'\bN[\s°ºOo0]?R?O?[\.:=\- ]{0,10}([0-9OQBILSZDEUA]{3,20})\b', linea)
#         if match_tolerante:
#             candidato = corregir_ocr_numero(match_tolerante.group(1))
#             if es_posible_numero_factura(candidato):
#                 candidatos.append((candidato, "Prefijo tolerante"))
#             continue

#         match_factura = re.search(r'FACTURA\s+(?:F\s*1\s+)?ELECTRONICA[^0-9A-Z]{0,10}([0-9OQBILSZDEUA]{3,20})\b', linea)
#         if match_factura:
#             candidato = corregir_ocr_numero(match_factura.group(1))
#             if es_posible_numero_factura(candidato):
#                 candidatos.append((candidato, "Factura + número en línea"))
#             continue

#         match_nopegado = re.search(r'\bNO([0-9OQBILSZDEUA]{3,20})\b', linea)
#         if match_nopegado:
#             candidato = corregir_ocr_numero(match_nopegado.group(1))
#             if es_posible_numero_factura(candidato):
#                 candidatos.append((candidato, "NO+Número sin espacio"))
#             continue

#         if "FACTURA" in linea and i + 1 < len(lineas):
#             siguiente = lineas[i + 1]
#             match_sig = re.search(r'\b([0-9OQBILSZDEUA]{3,20})\b', siguiente)
#             if match_sig:
#                 candidato = corregir_ocr_numero(match_sig.group(1))
#                 if es_posible_numero_factura(candidato):
#                     candidatos.append((candidato, "Número en línea inferior"))
#             continue

#         match_solo = re.fullmatch(r'\s*([0-9OQBILSZDEUA]{3,20})\s*', linea)
#         if match_solo:
#             candidato = corregir_ocr_numero(match_solo.group(1))
#             if es_posible_numero_factura(candidato):
#                 candidatos.append((candidato, "Línea numérica pura"))
#             continue

#     if not candidatos:
#         match_respaldo = re.findall(r'\b([0-9OQBILSZDEUA]{3,20})\b', texto)
#         for m in match_respaldo:
#             candidato = corregir_ocr_numero(m)
#             if es_posible_numero_factura(candidato):
#                 candidatos.append((candidato, "Respaldo: número general"))

#     if candidatos:
#         numero_crudo, origen = max(candidatos, key=lambda x: len(x[0]))
#         return numero_crudo

#     return ""

def extraer_numero_factura(texto: str) -> str:
    import re

    print("🟡 Texto OCR original (Número Factura):\n", texto)

    def corregir_ocr_numero(numero: str) -> str:
        traduccion = str.maketrans({
            'O': '0', 'Q': '0', 'B': '8', 'I': '1', 'L': '1', 'S': '5',
            'Z': '2', 'D': '0', 'E': '8', 'A': '4', 'U': '0', '/': '1'
        })
        return numero.translate(traduccion).replace('.', '').replace(' ', '')

    # ====== Normalizaciones existentes ======
    reemplazos = {
        "N°": "NRO ", "N'": "NRO ", 'N"': "NRO ", "N :": "NRO ", "N.": "NRO ",
        "Nº": "NRO ", "N:": "NRO ", "NE": "NRO ", "N?": "NRO ", "FNLC": "NRO ",
        "FNL": "NRO ", "FNLD": "NRO ", "FULD": "NRO ", "FOLIO": "NRO ",
        "NC:": "NRO:", "NC ": "NRO ", "N C": "NRO ", '"NC': "NRO ", "'NC": "NRO ",
        "NP ": "NRO ", "N°P": "NRO ", "N P": "NRO ", '"NP': "NRO ", "'NP": "NRO ",
        "NP:": "NRO:", "Nro.:": "NRO:", "Nro. :": "NRO:", "Nro :": "NRO:",
        "Nro.": "NRO", "NF": "NRO", "NiP": "NRO", "MP": "NRO", "NO": "NRO",
        "Nro.  :": "NRO:", "N9": "NRO", "Ne": "NRO", "nE": "NRO", "Nro": "NRO",
        "Nro ": "NRO", "Nro  ": "NRO", "Folio N?": "NRO", "FOLION?": "NRO ",
        "FOLIO N°": "NRO ", "FCLIO": "NRO ", "NUMERO": "NRO ",
        "NUMERO .": "NRO ", "No": "NRO ", "Nra": "NRO ",
        "Folio /": "NRO ", "RUMERO :": "NRO ", "Né": "NRO ", "N '": "NRO ",
        "NT ": "NRO ", "Nt": "NRO ", "eg.6n N": "NRO ","ND": "NRO ","Folio H": "NRO ",
        "N 0": "NRO ","Nm": "NRO ","Ni": "NRO ","KUMERO": "NRO ","HUMERO": "NRO ",
        "NUHERO": "NRO ","XUMERO": "NRO ","Nro: ": "NRO","NP": "NRO","NI": "NRO",
        
    }
    for k, v in reemplazos.items():
        texto = texto.replace(k, v)

    texto = texto.replace('=', ':').replace('f.1', 'ELECTRONICA').replace('f 1', 'ELECTRONICA')
    texto = texto.replace('Nro =', 'NRO:').replace('Nro :', 'NRO:').replace('Nro::', 'NRO:')
    texto = texto.replace('N:', 'NRO:').replace('NRO:::', 'NRO:')
    texto = texto.replace("NRO :", "NRO:").replace("NRO  :", "NRO:")
    texto = re.sub(r'F[A4@][CÇ][\s\-]*U[R][A4@][\s\-]*E[L1I][E3]C[T7][R][O0][N][I1][C][A4@]', "FACTURA ELECTRONICA", texto, flags=re.IGNORECASE)
    texto = re.sub(r'F[\s\-]*A[\s\-]*C[\s\-]*U[\s\-]*R[\s\-]*A[\s\-]*[E3][L1I][E3]C[T7][R][O0][N][I1][C][A4@]', "FACTURA ELECTRONICA", texto, flags=re.IGNORECASE)
    texto = texto.replace("FACTURAELECTRONICA", "FACTURA ELECTRONICA")
    texto = texto.replace("FACTURAELECTRÓNICA", "FACTURA ELECTRONICA")
    texto = texto.replace("FACTURAELECTRONICP", "FACTURA ELECTRONICA")
    texto = texto.replace("FACTVRAELECTRoNICA", "FACTURA ELECTRONICA")
    texto = texto.replace("FACTURA ELECTRÓNICA", "FACTURA ELECTRONICA")
    texto = texto.replace("FACTURA ELECTRNICA", "FACTURA ELECTRONICA")
    texto = texto.replace("FACTURA FLECTRONICA", "FACTURA ELECTRONICA")
    texto = texto.replace("FAC URA FLECTRONICA", "FACTURA ELECTRONICA")
    texto = texto.replace("FACTURA ELECIRONICA", "FACTURA ELECTRONICA")
    texto = texto.replace("FACTURLELECTRONICA", "FACTURA ELECTRONICA")
    texto = texto.replace("FACTURA Electronica", "FACTURA ELECTRONICA")
    texto = texto.replace("FACTURA ELECTRNICA FOLIO", "FACTURA ELECTRONICA")
    texto = texto.replace("FACTURA ELECTRNICA FOLIO NRO", "FACTURA ELECTRONICA")
    texto = texto.replace("FACTURA ELECTRONICA NO1", "FACTURA ELECTRONICA")
    texto = texto.upper()
    texto = re.sub(r'[^\x00-\x7F]+', '', texto)
    texto = re.sub(r'\bN\s*R\s*O\b', 'NRO', texto)  # unifica "N R O" → "NRO"

    # Unifica "NRO: <num>" para facilitar la extracción
    texto = re.sub(
        r'(NRO|NUMERO)[\s:=\.\-]+([0-9OQBILSZDEUA]{3,20})',
        lambda m: f"NRO:{corregir_ocr_numero(m.group(2).upper())}",
        texto, flags=re.IGNORECASE
    )
    texto = re.sub(
        r'(NRO)[\s]{1,5}([0-9OQBILSZDEUA]{1,5})[\s\.]{1,2}([0-9OQBILSZDEUA]{1,5})',
        lambda m: f"NRO:{corregir_ocr_numero(m.group(2) + m.group(3))}",
        texto
    )

    print("🟢 Texto tras limpieza completa (Número Factura):\n", texto)
    lineas = texto.splitlines()
    candidatos = []

    distractores = {"CONTROL", "DOC", "INTERNO", "DOC INTERNO", "ORDEN", "KP/G.D", "GUIA", "GD", "VEND", "CLIENTE"}

    def es_posible_numero_factura(num: str) -> bool:
        num = corregir_ocr_numero(num)
        if re.match(r'^\+?56\s?\d{2}\s?\d{4,}$', num):  # teléfono
            return False
        if len(num) < 3 or len(num) > 12:
            return False
        if not num.isdigit():
            return False
        return True

    # === Prioridad fuerte 1: "FACTURA ELECTRONICA ... N <num>" (tolera ruido intermedio) ===
    m_factura_n = re.search(
        r'FACTURA\s+ELECTRONICA(?:[^\d]{0,40})\bN[\s\.:=\-]*([0-9OQBILSZDEUA]{6,12})',
        texto
    )
    if m_factura_n:
        raw = m_factura_n.group(1)          # ya sin espacios por el patrón
        cand = corregir_ocr_numero(raw)
        if es_posible_numero_factura(cand):
            candidatos.append((cand, "FacturaN"))

    # === Búsqueda por líneas (tu lógica + toques mínimos) ===
    for i, linea in enumerate(lineas):
        linea_upper = linea.upper()
        cerca_de_rut = ("RUT" in linea_upper)

        match_exacto = re.search(r'NRO[\s:=\.\-]*([0-9OQBILSZDEUA]{3,20})', linea_upper)
        if match_exacto:
            candidato = corregir_ocr_numero(match_exacto.group(1))
            if es_posible_numero_factura(candidato):
                etiqueta = "NRO: exacto (cerca RUT)" if cerca_de_rut else "NRO: exacto"
                candidatos.append((candidato, etiqueta))
            continue

        match_tolerante = re.search(r'\bN[\s°ºOo0]?R?O?[\.:=\- ]{0,10}([0-9OQBILSZDEUA]{3,20})\b', linea_upper)
        if match_tolerante:
            candidato = corregir_ocr_numero(match_tolerante.group(1))
            if es_posible_numero_factura(candidato):
                etiqueta = "Prefijo tolerante (cerca RUT)" if cerca_de_rut else "Prefijo tolerante"
                candidatos.append((candidato, etiqueta))
            continue

        match_factura = re.search(
            r'FACTURA\s+(?:F\s*1\s+)?ELECTRONICA[^0-9A-Z]{0,10}([0-9OQBILSZDEUA\.]{3,20})\b',
            linea_upper
        )
        if match_factura:
            raw = match_factura.group(1).split()[0]  # primer token
            candidato = corregir_ocr_numero(raw)
            if es_posible_numero_factura(candidato):
                etiqueta = "Factura + número en línea (cerca RUT)" if cerca_de_rut else "Factura + número en línea"
                candidatos.append((candidato, etiqueta))
            continue

        match_nopegado = re.search(r'\bNO([0-9OQBILSZDEUA]{3,20})\b', linea_upper)
        if match_nopegado:
            candidato = corregir_ocr_numero(match_nopegado.group(1))
            if es_posible_numero_factura(candidato):
                candidatos.append((candidato, "NO+Número sin espacio"))
            continue

        if "FACTURA" in linea_upper and i + 1 < len(lineas):
            siguiente = lineas[i + 1].upper()
            match_sig = re.search(r'\b([0-9OQBILSZDEUA]{3,20})\b', siguiente)
            if match_sig:
                candidato = corregir_ocr_numero(match_sig.group(1))
                if es_posible_numero_factura(candidato):
                    candidatos.append((candidato, "Número en línea inferior"))
            continue

        match_solo = re.fullmatch(r'\s*([0-9OQBILSZDEUA]{3,20})\s*', linea_upper)
        if match_solo:
            candidato = corregir_ocr_numero(match_solo.group(1))
            if es_posible_numero_factura(candidato):
                if any(w in linea_upper for w in distractores):
                    candidatos.append((candidato, "Línea numérica pura (distractor)"))
                else:
                    candidatos.append((candidato, "Línea numérica pura"))
            continue

    # Respaldo
    if not candidatos:
        for m in re.findall(r'\b([0-9OQBILSZDEUA]{3,20})\b', texto):
            candidato = corregir_ocr_numero(m)
            if es_posible_numero_factura(candidato):
                candidatos.append((candidato, "Respaldo: número general"))

    if not candidatos:
        return ""

    # Si existe algún candidato con >=6 dígitos, descarta los más cortos (evita “300” de DOC INTERNO)
    if any(len(c[0]) >= 6 for c in candidatos):
        candidatos = [c for c in candidatos if len(c[0]) >= 6]

    prioridad = {
        "FacturaN": 5,
        "NRO: exacto": 4,
        "Factura + número en línea": 4,
        "Prefijo tolerante": 3,
        "Número en línea inferior": 2,
        "Línea numérica pura": 2,
        "NO+Número sin espacio": 2,
        "Respaldo: número general": 1,
        "NRO: exacto (cerca RUT)": 2,
        "Factura + número en línea (cerca RUT)": 2,
        "Prefijo tolerante (cerca RUT)": 1,
        "Línea numérica pura (distractor)": 1,
    }

    if not any(etq in prioridad for _, etq in candidatos):
        numero_crudo, origen = max(candidatos, key=lambda x: len(x[0]))
        return numero_crudo

    # Selección final: prioridad → largo → (ya filtrado por >=6 si existía)
    numero_crudo, origen = max(candidatos, key=lambda x: (prioridad.get(x[1], 0), len(x[0])))
    return numero_crudo
