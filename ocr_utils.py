import hide_subprocess
import re
import os,sys
import io
import logging
import contextlib
from datetime import datetime
import itertools

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
                # reader = easyocr.Reader(['es'], gpu=False)
                reader = easyocr.Reader(['es'], gpu=False, verbose=False)
                sys.stdout = sys.__stdout__
                sys.stderr = sys.__stderr__

inicializar_ocr()

# def ocr_zona_factura_desde_png(imagen_entrada, ruta_debug=None):
#     """
#     Realiza OCR en la zona superior derecha de una factura.
#     Si `ruta_debug` es None, guarda un recorte temporal solo en modo debug.
#     """
#     from PIL import Image
#     import numpy as np
#     import os
#     import sys
#     from datetime import datetime

#     if isinstance(imagen_entrada, str):
#         imagen = Image.open(imagen_entrada)
#     elif hasattr(imagen_entrada, "crop"):
#         imagen = imagen_entrada
#     else:
#         raise ValueError("imagen_entrada debe ser una ruta o un objeto PIL.Image")

#     ancho, alto = imagen.size
#     zona = imagen.crop((
#         int(ancho * 0.65),
#         int(alto * 0.01),
#         int(ancho * 0.97),
#         int(alto * 0.30)
#     ))

#     zona_reducida = zona.resize((zona.width // 2, zona.height // 2), resample=Image.BICUBIC)
#     ruta_debug_creado = None

#     if ruta_debug is None:
#         try:
#             base_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)
#             debug_dir = os.path.join(base_dir, "debug")
#             os.makedirs(debug_dir, exist_ok=True)
#             timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
#             ruta_debug_creado = os.path.join(debug_dir, f"recorte_{timestamp}.png")
#             zona.save(ruta_debug_creado)
#         except:
#             ruta_debug_creado = None
#     elif ruta_debug:
#         try:
#             zona.save(ruta_debug)
#         except:
#             pass

#     zona_np = np.array(zona_reducida)
#     resultado = reader.readtext(zona_np, detail=0, batch_size=1)

#     if ruta_debug_creado:
#         try:
#             os.remove(ruta_debug_creado)
#         except:
#             pass

#     return " ".join(resultado).strip()

# def ocr_zona_factura_desde_png(imagen_entrada, ruta_debug=None):
#     from PIL import Image
#     import numpy as np
#     from datetime import datetime
#     import os

#     if isinstance(imagen_entrada, str):
#         imagen_original = Image.open(imagen_entrada)
#     elif hasattr(imagen_entrada, "crop"):
#         imagen_original = imagen_entrada
#     else:
#         raise ValueError("imagen_entrada debe ser una ruta o un objeto PIL.Image")

#     mejores_resultados = []
#     rotaciones = [0, 90, 180, 270]

#     for angulo in rotaciones:
#         imagen_rotada = imagen_original.rotate(angulo, expand=True)
#         ancho, alto = imagen_rotada.size

#         zona = imagen_rotada.crop((
#             int(ancho * 0.65),
#             int(alto * 0.01),
#             int(ancho * 0.97),
#             int(alto * 0.30)
#         ))

#         zona_reducida = zona.resize((zona.width // 2, zona.height // 2), resample=Image.BICUBIC)
#         zona_np = np.array(zona_reducida)

#         resultado = reader.readtext(zona_np, detail=0, batch_size=1)
#         texto = " ".join(resultado).strip()
#         cantidad_palabras = len(texto.split())

#         mejores_resultados.append((cantidad_palabras, angulo, texto, zona))

#     # Elegir rotación con más palabras reconocidas
#     mejores_resultados.sort(reverse=True)
#     _, mejor_angulo, mejor_texto, mejor_recorte = mejores_resultados[0]

#     registrar_log_proceso(f"🔁 Imagen rotada automáticamente {mejor_angulo}°")
#     registrar_log_proceso(f"📄 OCR mejor orientación ({mejor_angulo}°): {mejor_texto[:100]}...")

#     # Guardar recorte si se especifica ruta
#     if ruta_debug:
#         try:
#             mejor_recorte.save(ruta_debug)
#             registrar_log_proceso(f"📎 Recorte guardado en: {ruta_debug}")
#         except:
#             pass
#     else:
#         # Ruta temporal solo si es en modo debug
#         try:
#             base_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)
#             debug_dir = os.path.join(base_dir, "debug")
#             os.makedirs(debug_dir, exist_ok=True)
#             timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
#             ruta_debug_creado = os.path.join(debug_dir, f"recorte_{timestamp}.png")
#             mejor_recorte.save(ruta_debug_creado)
#         except:
#             pass

#     return mejor_texto

def ocr_zona_factura_desde_png(imagen_entrada, ruta_debug=None):
    """
    Detecta automáticamente la orientación del documento y realiza OCR en la zona superior derecha.
    """
    from PIL import Image
    import numpy as np
    import os
    import sys
    from datetime import datetime

    if isinstance(imagen_entrada, str):
        imagen_original = Image.open(imagen_entrada)
    elif hasattr(imagen_entrada, "crop"):
        imagen_original = imagen_entrada
    else:
        raise ValueError("imagen_entrada debe ser una ruta o un objeto PIL.Image")

    mejor_texto = ""
    mejor_puntaje = 0
    mejor_recorte = None
    mejor_angulo = 0

    for angulo in [0, 90, 180, 270]:
        imagen_rotada = imagen_original.rotate(angulo, expand=True)
        ancho, alto = imagen_rotada.size
        recorte = imagen_rotada.crop((
            int(ancho * 0.65),
            int(alto * 0.01),
            int(ancho * 0.99),
            int(alto * 0.30)
        ))

        zona_reducida = recorte.resize((recorte.width // 2, recorte.height // 2), resample=Image.BICUBIC)
        zona_np = np.array(zona_reducida)
        texto = reader.readtext(zona_np, detail=0, batch_size=1)

        texto_completo = " ".join(texto).strip()
        palabras_clave = sum(1 for palabra in texto_completo.upper().split()
                             if palabra in ["RUT", "FACTURA", "ELECTRONICA", "NRO", "SII"])
        
        if palabras_clave > mejor_puntaje:
            mejor_puntaje = palabras_clave
            mejor_texto = texto_completo
            mejor_recorte = recorte
            mejor_angulo = angulo

    if mejor_angulo != 0:
        registrar_log_proceso(f"🔁 Imagen rotada automáticamente {mejor_angulo}°")
        if ruta_debug:
            ruta_rotada = ruta_debug.replace("_recorte", f"_rotada{mejor_angulo}")
            imagen_original.rotate(mejor_angulo, expand=True).save(ruta_rotada)

    if ruta_debug:
        try:
            mejor_recorte.save(ruta_debug)
            registrar_log_proceso(f"📎 Recorte guardado en: {ruta_debug}")
        except Exception as e:
            registrar_log_proceso(f"⚠️ Error guardando recorte: {e}")

    return mejor_texto

def extraer_rut(texto):
    # print("🟡 Texto OCR original (RUT):\n", texto)
    texto_original = texto

    reemplazos = {
        "RUT.": "RUT", "R.U.T.": "RUT", "R-U-T": "RUT", "RUT:": "RUT", "RUT;": "RUT",
        "RUT=": "RUT", "RU.T": "RUT", "RU:T": "RUT", "R:UT": "RUT", "RU.T.": "RUT",
        "RUI": "RUT", "RU1": "RUT", "R.UT.": "RUT", "RuT;": "RUT", "RUTTT;": "RUT",
        "Ru:,n.": "RUT", "Ru.t:": "RUT", "RVT ;": "RUT", "RVT ": "RUT", "RVT": "RUT",
        "RUT.:": "RUT",
    }

    for k, v in reemplazos.items():
        texto = texto.replace(k, v)

    texto = texto.replace(',', '.')
    texto = texto.replace('O', '0').replace('o', '0').replace('I', '1').replace('l', '1')
    texto = texto.replace('B', '8').replace('Z', '2').replace('G', '6')

    texto = texto.replace('–', '-').replace('—', '-').replace('‐', '-')
    texto = texto.replace('+', '-')  # ⬅️ Corrección agregada para capturar mal lectura del guion

    # print("🟢 Texto tras limpieza final (RUT):\n", texto)

    # Expresión regular más robusta
    posibles = re.findall(r'\d{1,2}\s*[\.,]?\s*\d{3}\s*[\.,]?\s*\d{3}\s*[-‐–—]?\s*[\dkK]', texto)
    if posibles:
        rut = posibles[0]
        rut = re.sub(r'[^0-9kK]', '', rut[:-1]) + '-' + rut[-1].upper()
        return rut

    posibles2 = re.findall(r'(\d{1,2})[^\d]{0,2}(\d{3})[^\d]{0,2}(\d{3})[^\dkK]{0,2}([\dkK])', texto_original)
    if posibles2:
        rut = f"{posibles2[0][0]}{posibles2[0][1]}{posibles2[0][2]}-{posibles2[0][3].upper()}"
        return rut

    registrar_log_proceso("⚠️ RUT no detectado.")
    return "desconocido"



# def extraer_rut(texto):
#     # print('texto rut', texto)
#     # posibles = re.findall(r'\d{1,2}[\.]?\d{3}[\.]?\d{3}-[\dkK]', texto)
#     # if posibles:
#     #     rut = posibles[0].replace('.', '').upper()
#     #     return rut

#     # # Intentar capturar RUT sin dígito verificador y calcularlo
#     # rut_sin_dv = re.findall(r'\b\d{1,2}[\.]?\d{3}[\.]?\d{3}\b', texto)
#     # if rut_sin_dv:
#     #     base_rut = rut_sin_dv[0].replace('.', '')
#     #     dv = calcular_dv(base_rut)
#     #     rut_completo = f"{base_rut}-{dv}"
#     #     print(f"🧮 RUT reconstruido con DV: {rut_completo}")
#     #     return rut_completo

#     # print('el rut es ',rut_sin_dv)
#     # registrar_log_proceso("⚠️ RUT no detectado.")
#     # return "desconocido"



#     print("🟡 Texto OCR original (RUT):\n", texto)
#     texto_original = texto
#     # Reemplazos OCR adicionales para prefijos erróneos o confusos (menos agresivos)
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
#         "RuT;": "RUT",
#         "RUTTT;": "RUT",
#         "Ru:,n.": "RUT",
#         "Ru.t:": "RUT",
#         "RVT ;": "RUT",
#         "RVT ": "RUT",
#         "RVT": "RUT",
#         "RUT.:":"RUT",
#         # 🔥 ¡Ojo! Se eliminó "RU": "RUT" para evitar 'RUTT'
#     }
#     for k, v in reemplazos.items():
#         texto = texto.replace(k, v)

#     # print("🟠 Texto tras reemplazos de prefijo (RUT):\n", texto)
#     # Reemplazos comunes de caracteres mal reconocidos (sin eliminar espacios)
#     texto = texto.replace(',', '.')
#     texto = texto.replace('O', '0').replace('o', '0').replace('I', '1').replace('l', '1')
#     texto = texto.replace('B', '8').replace('Z', '2').replace('G', '6')
#     texto = texto.replace('–', '-').replace('—', '-')

#     print("🟢 Texto tras limpieza final (RUT):\n", texto)
#     # Buscar patrones estándar de RUT
#     posibles = re.findall(r'\d{1,2}\s*[\.,]?\s*\d{3}\s*[\.,]?\s*\d{3}\s*[-‐–—]?\s*[\dkK]', texto)
#     if posibles:
#         rut = posibles[0].replace('.', '').upper()
#         # print("✅ RUT detectado (directo):", rut)
#         return rut

#     # Buscar patrones más flexibles si no se encontró ninguno directo
#     posibles2 = re.findall(r'(\d{1,2})[^\d]{0,2}(\d{3})[^\d]{0,2}(\d{3})[^\dkK]{0,2}([\dkK])', texto_original)
#     if posibles2:
#         rut = f"{posibles2[0][0]}{posibles2[0][1]}{posibles2[0][2]}-{posibles2[0][3].upper()}"
#         # print("✅ RUT detectado (flexible):", rut)
#         return rut

#     registrar_log_proceso("⚠️ RUT no detectado.")
#     return "desconocido"

#############################################################
# def extraer_numero_factura(texto: str) -> str:
#     print("🟡 Texto OCR original (Número Factura):\n", texto)

#     """
#     Extrae el número de factura desde texto OCR, aplicando limpieza y múltiples patrones de búsqueda.
#     Devuelve el número de factura más probable o una cadena vacía si no se encuentra.
#     """

#     reemplazos = {
#         "N°": "NRO ",
#         "N'": "NRO ",
#         'N"': "NRO ",
#         "N :": "NRO ",
#         "N.": "NRO ",
#         "Nº": "NRO ",
#         "N:": "NRO ",
#         "NE": "NRO ",
#         "N?": "NRO ",
#         "FNLC": "NRO ",
#         "FNL": "NRO ",
#         "FNLD": "NRO ",
#         "FULD": "NRO ",
#         "FOLIO": "NRO ",
#         "NC:": "NRO:",
#         "NC ": "NRO ",
#         "N C": "NRO ",
#         '"NC': "NRO ",
#         "'NC": "NRO ",
#         "NP ": "NRO ",
#         "N°P": "NRO ",
#         "N P": "NRO ",
#         '"NP': "NRO ",
#         "'NP": "NRO ",
#         "NP:": "NRO:",
#         "Nro.:": "NRO:",
#         "Nro. :": "NRO:",
#         "Nro :": "NRO:",
#         "Nro.": "NRO",
#         "NF":"NRO",
#         "NiP" : "NRO",
#         "MP":"NRO",
#         "NO":"NRO",
#         "Nro.  :":"NRO:",
#         "N?":"NRO",
#         "N\"": "NRO",
#         'N"': "NRO",
#         "N9":"NRO",
#         "Ne":"NRO",
#         "nE":"NRO",
#         "Nro":"NRO",
#         "Nro ":"NRO",
#         "Nro  ":"NRO",
#         "Folio N?":"NRO",
#         "Folio N?": "NRO ",
#         "FOLION?": "NRO ",
#         "FOLIO N°": "NRO ",
#         "FOLIO": "NRO ",
#         "FCLIO": "NRO ",
#         "NUMERO": "NRO ",
#         "NUMERO .": "NRO ",
        
#     }

#     for k, v in reemplazos.items():
#         texto = texto.replace(k, v)

#     texto = texto.replace('=', ':').replace('f.1', 'ELECTRONICA').replace('f 1', 'ELECTRONICA')
#     texto = texto.replace('Nro =', 'NRO:').replace('Nro :', 'NRO:').replace('Nro::', 'NRO:')
#     texto = texto.replace('N:', 'NRO:').replace('NRO:::', 'NRO:')
#     texto = texto.replace("NRO :", "NRO:").replace("NRO  :", "NRO:")

#     # Otros reemplazo adicionales:
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

#     texto = texto.upper()
#     texto = re.sub(r'[^\x00-\x7F]+', '', texto)
#     print("🟢 Texto tras limpieza completa (Número Factura):\n", texto)

#     lineas = texto.splitlines()
#     candidatos = []

#     def es_posible_numero_factura(num: str) -> bool:
#         if re.match(r'^\+?56\s?\d{2}\s?\d{4,}$', num):
#             return False
#         if len(num) < 4 or len(num) > 12:
#             return False
#         return True

#     for i, linea in enumerate(lineas):
#         match_exacto = re.search(r'NRO[:\.\- ]{1,5}([0-9OQBILSZDE]{5,12})', linea)
#         if match_exacto:
#             candidato = match_exacto.group(1)
#             if es_posible_numero_factura(candidato):
#                 candidatos.append((candidato, "NRO: exacto"))
#             continue

#         match_tolerante = re.search(r'\bN[\s°ºOo0]?R?O?[\.:=\- ]{0,10}([0-9OQBILSZDE]{5,12})\b', linea)
#         if match_tolerante:
#             candidato = match_tolerante.group(1)
#             if es_posible_numero_factura(candidato):
#                 candidatos.append((candidato, "Prefijo tolerante"))
#             continue

#         match_factura = re.search(r'FACTURA\s+(?:F\s*1\s+)?ELECTRONICA[^0-9]{0,10}([0-9OQBILSZDE]{5,12})\b', linea)
#         if match_factura:
#             candidato = match_factura.group(1)
#             if es_posible_numero_factura(candidato):
#                 candidatos.append((candidato, "Factura + número en línea"))
#             continue

#         if "FACTURA" in linea and i + 1 < len(lineas):
#             siguiente = lineas[i + 1]
#             match_sig = re.search(r'\b([0-9OQBILSZDE]{5,12})\b', siguiente)
#             if match_sig:
#                 candidato = match_sig.group(1)
#                 if es_posible_numero_factura(candidato):
#                     candidatos.append((candidato, "Número en línea inferior"))
#             continue

#         match_solo = re.fullmatch(r'\s*([0-9OQBILSZDE]{5,12})\s*', linea)
#         if match_solo:
#             candidato = match_solo.group(1)
#             if es_posible_numero_factura(candidato):
#                 candidatos.append((candidato, "Línea numérica pura"))
#             continue

#     # Patrón 6: Último número de 5 a 12 dígitos como respaldo si nada anterior coincidió
#     if not candidatos:
#         match_respaldo = re.findall(r'\b([0-9OQBILSZDE]{5,12})\b', texto)
#         for m in match_respaldo:
#             if es_posible_numero_factura(m):
#                 candidatos.append((m, "Respaldo: número general"))

#     if candidatos:
#         numero_crudo, origen = max(candidatos, key=lambda x: len(x[0]))
#         numero_limpio = numero_crudo.upper().translate(str.maketrans({
#             'O': '0', 'Q': '0', 'B': '8', 'I': '1', 'L': '1', 'S': '5', 'Z': '2', 'D': '0', 'E': '8','A': '4'
#         }))
#         # print(f"✅ Número de factura detectado ({origen}): {numero_limpio}")
#         return numero_limpio
#     # print("❌ Número de factura no detectado.")
#     return ""




# import re

# def extraer_numero_factura(texto: str) -> str:
#     print("🟡 Texto OCR original (Número Factura):\n", texto)

#     def corregir_ocr_numero(numero: str) -> str:
#         traduccion = str.maketrans({
#             'O': '0', 'Q': '0', 'B': '8', 'I': '1', 'L': '1', 'S': '5',
#             'Z': '2', 'D': '0', 'E': '8', 'A': '4'
#         })
#         return numero.translate(traduccion).replace('.', '').replace(' ', '')

#     reemplazos = {
#         "N°": "NRO ",
#         "N'": "NRO ",
#         'N"': "NRO ",
#         "N :": "NRO ",
#         "N.": "NRO ",
#         "Nº": "NRO ",
#         "N:": "NRO ",
#         "NE": "NRO ",
#         "N?": "NRO ",
#         "FNLC": "NRO ",
#         "FNL": "NRO ",
#         "FNLD": "NRO ",
#         "FULD": "NRO ",
#         "FOLIO": "NRO ",
#         "NC:": "NRO:",
#         "NC ": "NRO ",
#         "N C": "NRO ",
#         '"NC': "NRO ",
#         "'NC": "NRO ",
#         "NP ": "NRO ",
#         "N°P": "NRO ",
#         "N P": "NRO ",
#         '"NP': "NRO ",
#         "'NP": "NRO ",
#         "NP:": "NRO:",
#         "Nro.:": "NRO:",
#         "Nro. :": "NRO:",
#         "Nro :": "NRO:",
#         "Nro.": "NRO",
#         "NF": "NRO",
#         "NiP": "NRO",
#         "MP": "NRO",
#         "NO": "NRO",
#         "Nro.  :": "NRO:",
#         "N?": "NRO",
#         "N\"": "NRO",
#         'N"': "NRO",
#         "N9": "NRO",
#         "Ne": "NRO",
#         "nE": "NRO",
#         "Nro": "NRO",
#         "Nro ": "NRO",
#         "Nro  ": "NRO",
#         "Folio N?": "NRO",
#         "FOLION?": "NRO ",
#         "FOLIO N°": "NRO ",
#         "FOLIO": "NRO ",
#         "FCLIO": "NRO ",
#         "NUMERO": "NRO ",
#         "NUMERO .": "NRO ",
#         "NO": "NRO ",
#         "No": "NRO ",

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
#     r'(NRO|NUMERO)[\s:=\.\-]+([0-9OQBILSZDE]{5,12})',
#     lambda m: f"NRO:{corregir_ocr_numero(m.group(2).upper())}",
#     texto,flags=re.IGNORECASE)
#     # Detectar NRO seguido de número separado por espacios/puntos (ej: 'NRO   15.380')
#     texto = re.sub(
#     r'(NRO)[\s]{1,5}([0-9OQBILSZDE]{1,5})[\s\.]{1,2}([0-9OQBILSZDE]{1,5})',
#     lambda m: f"NRO:{corregir_ocr_numero(m.group(2) + m.group(3))}",
#     texto)




#     print("🟢 Texto tras limpieza completa (Número Factura):\n", texto)

#     lineas = texto.splitlines()
#     candidatos = []

#     def es_posible_numero_factura(num: str) -> bool:
#         num = corregir_ocr_numero(num)
#         if re.match(r'^\+?56\s?\d{2}\s?\d{4,}$', num):  # teléfono
#             return False
#         if len(num) < 4 or len(num) > 12:
#             return False
#         if not num.isdigit():
#             return False
#         return True

#     for i, linea in enumerate(lineas):
#         match_exacto = re.search(r'NRO[\s:=\.\-]*([0-9OQBILSZDE]{5,12})', linea)
#         if match_exacto:
#             candidato = match_exacto.group(1)
#             if es_posible_numero_factura(candidato):
#                 candidatos.append((candidato, "NRO: exacto"))
#             continue

#         match_tolerante = re.search(r'\bN[\s°ºOo0]?R?O?[\.:=\- ]{0,10}([0-9OQBILSZDE]{5,12})\b', linea)
#         if match_tolerante:
#             candidato = match_tolerante.group(1)
#             if es_posible_numero_factura(candidato):
#                 candidatos.append((candidato, "Prefijo tolerante"))
#             continue

#         match_factura = re.search(r'FACTURA\s+(?:F\s*1\s+)?ELECTRONICA[^0-9]{0,10}([0-9OQBILSZDE]{5,12})\b', linea)
#         if match_factura:
#             candidato = match_factura.group(1)
#             if es_posible_numero_factura(candidato):
#                 candidatos.append((candidato, "Factura + número en línea"))
#             continue

#                 # <-- Aquí va este nuevo bloque
#         match_nopegado = re.search(r'\bNO([0-9OQBILSZDE]{3,12})\b', linea)
#         if match_nopegado:
#             candidato = match_nopegado.group(1)
#             if es_posible_numero_factura(candidato):
#                 candidatos.append((candidato, "NO+Número sin espacio"))
#             continue

#         if "FACTURA" in linea and i + 1 < len(lineas):
#             siguiente = lineas[i + 1]
#             match_sig = re.search(r'\b([0-9OQBILSZDE]{5,12})\b', siguiente)
#             if match_sig:
#                 candidato = match_sig.group(1)
#                 if es_posible_numero_factura(candidato):
#                     candidatos.append((candidato, "Número en línea inferior"))
#             continue

#         match_solo = re.fullmatch(r'\s*([0-9OQBILSZDE]{5,12})\s*', linea)
#         if match_solo:
#             candidato = match_solo.group(1)
#             if es_posible_numero_factura(candidato):
#                 candidatos.append((candidato, "Línea numérica pura"))
#             continue

#     if not candidatos:
#         match_respaldo = re.findall(r'\b([0-9OQBILSZDE]{5,12})\b', texto)
#         for m in match_respaldo:
#             if es_posible_numero_factura(m):
#                 candidatos.append((m, "Respaldo: número general"))

#     if candidatos:
#         numero_crudo, origen = max(candidatos, key=lambda x: len(x[0]))
#         numero_limpio = corregir_ocr_numero(numero_crudo.upper())
#         return numero_limpio

#     return ""



def extraer_numero_factura(texto: str) -> str:

    print("🟡 Texto OCR original (Número Factura):\n", texto)

    def corregir_ocr_numero(numero: str) -> str:
        traduccion = str.maketrans({
            'O': '0', 'Q': '0', 'B': '8', 'I': '1', 'L': '1', 'S': '5',
            'Z': '2', 'D': '0', 'E': '8', 'A': '4', 'U': '0'  # <- U como 0
        })
        return numero.translate(traduccion).replace('.', '').replace(' ', '')

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
        "NUMERO .": "NRO ", "NO": "NRO ", "No": "NRO ","Nra": "NRO ",
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

    texto = re.sub(
        r'(NRO|NUMERO)[\s:=\.\-]+([0-9OQBILSZDEUA]{3,20})',
        lambda m: f"NRO:{corregir_ocr_numero(m.group(2).upper())}",
        texto, flags=re.IGNORECASE)

    texto = re.sub(
        r'(NRO)[\s]{1,5}([0-9OQBILSZDEUA]{1,5})[\s\.]{1,2}([0-9OQBILSZDEUA]{1,5})',
        lambda m: f"NRO:{corregir_ocr_numero(m.group(2) + m.group(3))}",
        texto)
    print("\n")
    print("🟢 Texto tras limpieza completa (Número Factura):\n", texto)

    lineas = texto.splitlines()
    candidatos = []

    def es_posible_numero_factura(num: str) -> bool:
        num = corregir_ocr_numero(num)
        if re.match(r'^\+?56\s?\d{2}\s?\d{4,}$', num):  # teléfono
            return False
        if len(num) < 3 or len(num) > 12:
            return False
        if not num.isdigit():
            return False
        return True

    for i, linea in enumerate(lineas):
        match_exacto = re.search(r'NRO[\s:=\.\-]*([0-9OQBILSZDEUA]{3,20})', linea)
        if match_exacto:
            candidato = corregir_ocr_numero(match_exacto.group(1))
            if es_posible_numero_factura(candidato):
                candidatos.append((candidato, "NRO: exacto"))
            continue

        match_tolerante = re.search(r'\bN[\s°ºOo0]?R?O?[\.:=\- ]{0,10}([0-9OQBILSZDEUA]{3,20})\b', linea)
        if match_tolerante:
            candidato = corregir_ocr_numero(match_tolerante.group(1))
            if es_posible_numero_factura(candidato):
                candidatos.append((candidato, "Prefijo tolerante"))
            continue

        match_factura = re.search(r'FACTURA\s+(?:F\s*1\s+)?ELECTRONICA[^0-9A-Z]{0,10}([0-9OQBILSZDEUA]{3,20})\b', linea)
        if match_factura:
            candidato = corregir_ocr_numero(match_factura.group(1))
            if es_posible_numero_factura(candidato):
                candidatos.append((candidato, "Factura + número en línea"))
            continue

        match_nopegado = re.search(r'\bNO([0-9OQBILSZDEUA]{3,20})\b', linea)
        if match_nopegado:
            candidato = corregir_ocr_numero(match_nopegado.group(1))
            if es_posible_numero_factura(candidato):
                candidatos.append((candidato, "NO+Número sin espacio"))
            continue

        if "FACTURA" in linea and i + 1 < len(lineas):
            siguiente = lineas[i + 1]
            match_sig = re.search(r'\b([0-9OQBILSZDEUA]{3,20})\b', siguiente)
            if match_sig:
                candidato = corregir_ocr_numero(match_sig.group(1))
                if es_posible_numero_factura(candidato):
                    candidatos.append((candidato, "Número en línea inferior"))
            continue

        match_solo = re.fullmatch(r'\s*([0-9OQBILSZDEUA]{3,20})\s*', linea)
        if match_solo:
            candidato = corregir_ocr_numero(match_solo.group(1))
            if es_posible_numero_factura(candidato):
                candidatos.append((candidato, "Línea numérica pura"))
            continue

    if not candidatos:
        match_respaldo = re.findall(r'\b([0-9OQBILSZDEUA]{3,20})\b', texto)
        for m in match_respaldo:
            candidato = corregir_ocr_numero(m)
            if es_posible_numero_factura(candidato):
                candidatos.append((candidato, "Respaldo: número general"))

    if candidatos:
        numero_crudo, origen = max(candidatos, key=lambda x: len(x[0]))
        return numero_crudo

    return ""

