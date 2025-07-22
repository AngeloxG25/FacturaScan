import hide_subprocess
import re
import os,sys
import io
import logging
import contextlib
from datetime import datetime

from log_utils import registrar_log_proceso
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
                reader = easyocr.Reader(['es'], gpu=False)
                sys.stdout = sys.__stdout__
                sys.stderr = sys.__stderr__

inicializar_ocr()

# def limpiar_texto_ocr(texto: str) -> str:
#     """
#     Limpia texto OCR reemplazando caracteres mal reconocidos y normalizando el contenido.
#     """
#     # Reemplazos comunes OCR
#     reemplazos = {
#         'O': '0', 'o': '0', 'Q': '0', 'D': '0', 'G': '6',
#         'I': '1', 'l': '1', 'L': '1', 'Z': '2',
#         'B': '8', 'S': '5', 'E': '8', 'A': '4', 'U': '0',
#         'C': '0'
#     }

#     # Reemplazar caracteres uno a uno
#     texto = texto.upper()
#     texto = texto.translate(str.maketrans(reemplazos))

#     # Reemplazos adicionales comunes en OCR
#     texto = texto.replace(',', '.')
#     texto = texto.replace('–', '-').replace('—', '-')
#     texto = texto.replace('=', ':').replace(';', ':')
#     texto = re.sub(r'[^\x00-\x7F]+', '', texto)  # eliminar caracteres no ASCII

#     return texto


# def ocr_zona_factura_desde_png(imagen_entrada, ruta_debug=None):
#     """
#     Realiza OCR en la zona superior derecha de una factura.
#     Puede recibir una ruta a imagen o un objeto PIL.Image directamente.
#     Si se especifica `ruta_debug`, guarda el recorte como PNG.
#     """
#     from PIL import Image
#     import numpy as np

#     assert hasattr(imagen_entrada, "crop"), "imagen_entrada debe ser un objeto PIL.Image"
#     imagen = imagen_entrada

#     ancho, alto = imagen.size

#     # Recorte de zona probable de factura
#     zona = imagen.crop((
#         int(ancho * 0.58),  # x1
#         int(alto * 0.00),   # y1
#         int(ancho * 0.98),  # x2
#         int(alto * 0.25)    # y2
#     ))

#     if ruta_debug:
#         zona.save(ruta_debug)

#     # Reducción de tamaño (resample bicubic) para acelerar OCR
#     zona_reducida = zona.resize((zona.width // 2, zona.height // 2), resample=Image.BICUBIC)

#     zona_np = np.array(zona_reducida)

#     # OCR solo con texto (más rápido)
#     resultado = reader.readtext(zona_np, detail=0, batch_size=1)

#     return " ".join(resultado).strip()

def ocr_zona_factura_desde_png(imagen_entrada, ruta_debug=None):
    """
    Realiza OCR en la zona superior derecha de una factura.
    Si `ruta_debug` es None, guarda un recorte temporal solo en modo debug.
    """
    from PIL import Image
    import numpy as np
    import os
    import sys
    from datetime import datetime

    if isinstance(imagen_entrada, str):
        imagen = Image.open(imagen_entrada)
    elif hasattr(imagen_entrada, "crop"):
        imagen = imagen_entrada
    else:
        raise ValueError("imagen_entrada debe ser una ruta o un objeto PIL.Image")

    ancho, alto = imagen.size
    zona = imagen.crop((
        int(ancho * 0.65),
        int(alto * 0.01),
        int(ancho * 0.97),
        int(alto * 0.20)
    ))

    zona_reducida = zona.resize((zona.width // 2, zona.height // 2), resample=Image.BICUBIC)
    ruta_debug_creado = None

    if ruta_debug is None:
        try:
            base_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)
            debug_dir = os.path.join(base_dir, "debug")
            os.makedirs(debug_dir, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            ruta_debug_creado = os.path.join(debug_dir, f"recorte_{timestamp}.png")
            zona.save(ruta_debug_creado)
        except:
            ruta_debug_creado = None
    elif ruta_debug:
        try:
            zona.save(ruta_debug)
        except:
            pass

    zona_np = np.array(zona_reducida)
    resultado = reader.readtext(zona_np, detail=0, batch_size=1)

    if ruta_debug_creado:
        try:
            os.remove(ruta_debug_creado)
        except:
            pass

    return " ".join(resultado).strip()

# def extraer_rut(texto):
#     print("🟡 Texto OCR original (RUT):\n", texto)
#     texto_original = texto
#     # Reemplazos OCR adicionales para prefijos erróneos o confusos
#     reemplazos = {
#         "RUT.": "RUT",
#         "R.U.T.": "RUT",
#         "R-U-T": "RUT",
#         "RUT:": "RUT",
#         "RUT;": "RUT",
#         "RUT=": "RUT",
#         "RU.T": "RUT",
#         "RU:T": "RUT",
#         "R:UT": "RUT",
#         "RU.T.": "RUT",
#         "RUI": "RUT",
#         "RU1": "RUT",
#         "R.UT.": "RUT",
#         "Ru": "RUT",
#         "RU": "RUT",
#         "RU:T,;": "RUT",
#         "RuT;": "RUT",
#         "RUTTT;": "RUT",
#         "Ru:,n.": "RUT",
#         "Ru.t:": "RUT",
#         "RVT ;": "RUT",
#         "RVT ": "RUT",
#         "RVT": "RUT",

#     }

#     for k, v in reemplazos.items():
#         texto = texto.replace(k, v)

#     print("🟠 Texto tras reemplazos de prefijo (RUT):\n", texto)

#     # Reemplazos comunes de caracteres mal reconocidos
#     texto = texto.replace(',', '.').replace(' ', '')
#     texto = texto.replace('O', '0').replace('o', '0').replace('I', '1').replace('l', '1')
#     texto = texto.replace('B', '8').replace('Z', '2').replace('G', '6')
#     texto = texto.replace('–', '-').replace('—', '-')

#     print("🟢 Texto tras limpieza final (RUT):\n", texto)

#     # Buscar patrones estándar de RUT
#     posibles = re.findall(r'\d{1,2}[\.]?\d{3}[\.]?\d{3}-[\dkK]', texto)
#     if posibles:
#         rut = posibles[0].replace('.', '').upper()
#         # print(f" RUT detectado (directo): {rut}")
#         return rut

#     # Buscar patrones más flexibles si no se encontró ninguno directo
#     posibles2 = re.findall(r'(\d{1,2})[^\d]{0,2}(\d{3})[^\d]{0,2}(\d{3})[^\dkK]{0,2}([\dkK])', texto_original)
#     if posibles2:
#         rut = f"{posibles2[0][0]}{posibles2[0][1]}{posibles2[0][2]}-{posibles2[0][3].upper()}"
#         # print(f"✅ RUT detectado (flexible): {rut}")
#         return rut

#     registrar_log_proceso("⚠️ RUT no detectado.")
#     return "desconocido"

# nueva funcion:
def extraer_rut(texto):
    # print("🟡 Texto OCR original (RUT):\n", texto)

    texto_original = texto

    # Reemplazos OCR adicionales para prefijos erróneos o confusos (menos agresivos)
    reemplazos = {
        "RUT.": "RUT",
        "R.U.T.": "RUT",
        "R-U-T": "RUT",
        "RUT:": "RUT",
        "RUT;": "RUT",
        "RUT=": "RUT",
        "RU.T": "RUT",
        "RU:T": "RUT",
        "R:UT": "RUT",
        "RU.T.": "RUT",
        "RUI": "RUT",
        "RU1": "RUT",
        "R.UT.": "RUT",
        "RuT;": "RUT",
        "RUTTT;": "RUT",
        "Ru:,n.": "RUT",
        "Ru.t:": "RUT",
        "RVT ;": "RUT",
        "RVT ": "RUT",
        "RVT": "RUT"
        # 🔥 ¡Ojo! Se eliminó "RU": "RUT" para evitar 'RUTT'
    }

    for k, v in reemplazos.items():
        texto = texto.replace(k, v)

    # print("🟠 Texto tras reemplazos de prefijo (RUT):\n", texto)

    # Reemplazos comunes de caracteres mal reconocidos (sin eliminar espacios)
    texto = texto.replace(',', '.')
    texto = texto.replace('O', '0').replace('o', '0').replace('I', '1').replace('l', '1')
    texto = texto.replace('B', '8').replace('Z', '2').replace('G', '6')
    texto = texto.replace('–', '-').replace('—', '-')

    # print("🟢 Texto tras limpieza final (RUT):\n", texto)

    # Buscar patrones estándar de RUT
    posibles = re.findall(r'\d{1,2}[\.]?\d{3}[\.]?\d{3}-[\dkK]', texto)
    if posibles:
        rut = posibles[0].replace('.', '').upper()
        # print("✅ RUT detectado (directo):", rut)
        return rut

    # Buscar patrones más flexibles si no se encontró ninguno directo
    posibles2 = re.findall(r'(\d{1,2})[^\d]{0,2}(\d{3})[^\d]{0,2}(\d{3})[^\dkK]{0,2}([\dkK])', texto_original)
    if posibles2:
        rut = f"{posibles2[0][0]}{posibles2[0][1]}{posibles2[0][2]}-{posibles2[0][3].upper()}"
        # print("✅ RUT detectado (flexible):", rut)
        return rut

    registrar_log_proceso("⚠️ RUT no detectado.")
    return "desconocido"


def extraer_numero_factura(texto: str) -> str:
    # print("🟡 Texto OCR original (Número Factura):\n", texto)

    """
    Extrae el número de factura desde texto OCR, aplicando limpieza y múltiples patrones de búsqueda.
    Devuelve el número de factura más probable o una cadena vacía si no se encuentra.
    """

    reemplazos = {
        "N°": "NRO ",
        "N'": "NRO ",
        'N"': "NRO ",
        "N :": "NRO ",
        "N.": "NRO ",
        "Nº": "NRO ",
        "N:": "NRO ",
        "NE": "NRO ",
        "N?": "NRO ",
        "FNLC": "NRO ",
        "FNL": "NRO ",
        "FNLD": "NRO ",
        "FULD": "NRO ",
        "FOLIO": "NRO ",
        "NC:": "NRO:",
        "NC ": "NRO ",
        "N C": "NRO ",
        '"NC': "NRO ",
        "'NC": "NRO ",
        "NP ": "NRO ",
        "N°P": "NRO ",
        "N P": "NRO ",
        '"NP': "NRO ",
        "'NP": "NRO ",
        "NP:": "NRO:",
        "Nro.:": "NRO:",
        "Nro. :": "NRO:",
        "Nro :": "NRO:",
        "Nro.": "NRO",
        "NF":"NRO",
        "NiP" : "NRO",
        "MP":"NRO",
        "NO":"NRO",
        "Nro.  :":"NRO:",
        "N?":"NRO",
        "N\"": "NRO",
        'N"': "NRO",
        "N9":"NRO",
        "Ne":"NRO",
        "nE":"NRO",
        "Nro":"NRO",
        "Nro ":"NRO",
        "Nro  ":"NRO",
        
    }

    for k, v in reemplazos.items():
        texto = texto.replace(k, v)

    texto = texto.replace('=', ':').replace('f.1', 'ELECTRONICA').replace('f 1', 'ELECTRONICA')
    texto = texto.replace('Nro =', 'NRO:').replace('Nro :', 'NRO:').replace('Nro::', 'NRO:')
    texto = texto.replace('N:', 'NRO:').replace('NRO:::', 'NRO:')
    texto = texto.replace("NRO :", "NRO:").replace("NRO  :", "NRO:")

    # Otros reemplazo adicionales:
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

    texto = texto.upper()
    texto = re.sub(r'[^\x00-\x7F]+', '', texto)
    # print("🟢 Texto tras limpieza completa (Número Factura):\n", texto)

    lineas = texto.splitlines()
    candidatos = []

    def es_posible_numero_factura(num: str) -> bool:
        if re.match(r'^\+?56\s?\d{2}\s?\d{4,}$', num):
            return False
        if len(num) < 4 or len(num) > 12:
            return False
        return True

    for i, linea in enumerate(lineas):
        match_exacto = re.search(r'NRO[:\.\- ]{1,5}([0-9OQBILSZDE]{5,12})', linea)
        if match_exacto:
            candidato = match_exacto.group(1)
            if es_posible_numero_factura(candidato):
                candidatos.append((candidato, "NRO: exacto"))
            continue

        match_tolerante = re.search(r'\bN[\s°ºOo0]?R?O?[\.:=\- ]{0,10}([0-9OQBILSZDE]{5,12})\b', linea)
        if match_tolerante:
            candidato = match_tolerante.group(1)
            if es_posible_numero_factura(candidato):
                candidatos.append((candidato, "Prefijo tolerante"))
            continue

        match_factura = re.search(r'FACTURA\s+(?:F\s*1\s+)?ELECTRONICA[^0-9]{0,10}([0-9OQBILSZDE]{5,12})\b', linea)
        if match_factura:
            candidato = match_factura.group(1)
            if es_posible_numero_factura(candidato):
                candidatos.append((candidato, "Factura + número en línea"))
            continue

        if "FACTURA" in linea and i + 1 < len(lineas):
            siguiente = lineas[i + 1]
            match_sig = re.search(r'\b([0-9OQBILSZDE]{5,12})\b', siguiente)
            if match_sig:
                candidato = match_sig.group(1)
                if es_posible_numero_factura(candidato):
                    candidatos.append((candidato, "Número en línea inferior"))
            continue

        match_solo = re.fullmatch(r'\s*([0-9OQBILSZDE]{5,12})\s*', linea)
        if match_solo:
            candidato = match_solo.group(1)
            if es_posible_numero_factura(candidato):
                candidatos.append((candidato, "Línea numérica pura"))
            continue

    # Patrón 6: Último número de 5 a 12 dígitos como respaldo si nada anterior coincidió
    if not candidatos:
        match_respaldo = re.findall(r'\b([0-9OQBILSZDE]{5,12})\b', texto)
        for m in match_respaldo:
            if es_posible_numero_factura(m):
                candidatos.append((m, "Respaldo: número general"))

    if candidatos:
        numero_crudo, origen = max(candidatos, key=lambda x: len(x[0]))
        numero_limpio = numero_crudo.upper().translate(str.maketrans({
            'O': '0', 'Q': '0', 'B': '8', 'I': '1', 'L': '1', 'S': '5', 'Z': '2', 'D': '0', 'E': '8','A': '4'
        }))
        # print(f"✅ Número de factura detectado ({origen}): {numero_limpio}")
        return numero_limpio
    # print("❌ Número de factura no detectado.")
    return ""


