import re
import os
import io
import logging
import contextlib

# Silenciar advertencias GPU de torch/easyocr
os.environ["CUDA_VISIBLE_DEVICES"] = "-1"
logging.getLogger("torch").setLevel(logging.ERROR)

# Cargar EasyOCR sin mostrar advertencias
with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
    import easyocr

# 🔁 Precarga global del modelo EasyOCR (una sola vez)
reader = None
def inicializar_ocr():
    global reader
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        reader = easyocr.Reader(['es'])

inicializar_ocr()

# def limpiar_texto_ocr_general_mejorado(texto):
#     """
#     Limpia el texto OCR detectado corrigiendo errores comunes de reconocimiento
#     de caracteres y normalizando expresiones claves.
#     """

#     # 🔁 Corrección de caracteres visualmente similares o mal interpretados por OCR
#     sustituciones = {
#         'o': '0', 'O': '0', 'Q': '0', 'D': '0',
#         'I': '1', 'l': '1', 'i': '1', '|': '1',
#         'S': '5', 's': '5',
#         'Z': '2', 'B': '8', 'G': '6', 'A': '4',
#         '—': '-', '–': '-', '~': '-', ';': ':', ',': '.', '*': '.',
#         '..': '.', '.-': '-'
#     }
#     for k, v in sustituciones.items():
#         texto = texto.replace(k, v)

#     # 🔁 Normaliza patrones confusos como "N°", con o sin símbolos erróneos
#     texto = re.sub(r'N[\°º°OoPp][\s:\-]*', 'N° ', texto, flags=re.IGNORECASE)

#     # 🔁 Reemplazos específicos para tus documentos
#     reemplazos_especiales = {
#         'R.U.T.': 'RUT:', 'RU.T': 'RUT:', 'Ru.T': 'RUT:',
#         'RUT.': 'RUT:', 'UT': 'RUT:', 'RuT': 'RUT:',
#         'SLL': 'SII', '5LL': 'SII', '511': 'SII', 'S11': 'SII',
#         'SIl': 'SII', '5IL': 'SII', 'S1I': 'SII',
#         'FNL0': 'FOLIO', 'FNLC': 'FOLIO', 'FCLIOS': 'FOLIOS', 'FOLICS': 'FOLIOS',
#         'FNLO': 'FOLIO', 'FN1O': 'FOLIO',
#         'FACTURA ELECTR0N1CA': 'FACTURA ELECTRONICA',
#         'FACTURA ELECTRON1CA': 'FACTURA ELECTRONICA',
#         'FACTURA ELECTR0NICA': 'FACTURA ELECTRONICA'
#     }
#     for incorrecto, correcto in reemplazos_especiales.items():
#         texto = texto.replace(incorrecto, correcto)

#     # 🔁 Correcciones con expresiones regulares para RUT y FACTURA
#     texto = re.sub(r'R{2,}UT[:\.\-]*', 'RUT:', texto, flags=re.IGNORECASE)
#     texto = re.sub(r'FACTURA\s+ELECTR[0O]N[1I]CA', 'FACTURA ELECTRONICA', texto, flags=re.IGNORECASE)
#     texto = re.sub(r'FACTURA\s*ELECTRON1CA', 'FACTURA ELECTRONICA', texto, flags=re.IGNORECASE)
#     texto = re.sub(r'FACTURA\s*ELECTR0NICA', 'FACTURA ELECTRONICA', texto, flags=re.IGNORECASE)

#     # 🔁 Limpieza de repeticiones redundantes
#     texto = re.sub(r'(RUT[:\s\-\.]*){2,}', 'RUT: ', texto, flags=re.IGNORECASE)
#     texto = re.sub(r'(FACTURA[\s\-]*){2,}', 'FACTURA ', texto, flags=re.IGNORECASE)
#     texto = re.sub(r'(ELECTRONICA[\s\-]*){2,}', 'ELECTRONICA ', texto, flags=re.IGNORECASE)

#     # 🔁 Formato correcto de RUT: RUT: xx.xxx.xxx-x
#     texto = re.sub(r'(RUT:\s*\d{2,3})\.(\d{3})\.(\d{3})[\s\-]+(\d)', r'RUT: \1.\2.\3-\4', texto)

#     # 🔁 Limpieza final: elimina símbolos no válidos y espacios excesivos
#     texto = re.sub(r'[^A-Za-z0-9ñÑ\s\.\-:/]', ' ', texto)
#     texto = re.sub(r'\s+', ' ', texto)

#     return texto.strip()

def ocr_zona_factura_desde_png(ruta_png, ruta_debug=None):
    from PIL import Image
    import numpy as np

    imagen = Image.open(ruta_png)
    ancho, alto = imagen.size

    # Ajuste más profundo (baja más para capturar bien N° y evitar corte)

    zona = imagen.crop((
        int(ancho * 0.58),  # x1: comenzamos desde la mitad horizontal de la imagen (50%)
        int(alto * 0.00),   # y1: comenzamos desde el borde superior (0%)
        int(ancho * 0.98),  # x2: recortamos hasta casi el borde derecho (98%)
        int(alto * 0.25)    # y2: recortamos solo hasta el 18% de la altura para evitar capturar la línea de "FOLIOS"
    ))

    if ruta_debug:
        zona.save(ruta_debug)

    zona_np = np.array(zona)
    resultado = reader.readtext(zona_np)
    return " ".join([item[1] for item in resultado]).strip()


def extraer_rut(texto):
    texto_original = texto
    texto = texto.replace(',', '.').replace(' ', '')
    texto = texto.replace('O', '0').replace('o', '0').replace('I', '1').replace('l', '1')
    texto = texto.replace('B', '8').replace('Z', '2').replace('G', '6')
    texto = texto.replace('–', '-').replace('—', '-')

    # Busca todos los patrones que parezcan un RUT
    posibles = re.findall(r'\d{1,2}[\.]?\d{3}[\.]?\d{3}-[\dkK]', texto)

    if posibles:
        # Elimina puntos y estandariza mayúsculas, retorna el primero
        return posibles[0].replace('.', '').upper()

    # Intenta con una expresión más flexible si no encontró nada
    posibles2 = re.findall(r'(\d{1,2})[^\d]{0,2}(\d{3})[^\d]{0,2}(\d{3})[^\dkK]{0,2}([\dkK])', texto_original)
    if posibles2:
        return f"{posibles2[0][0]}{posibles2[0][1]}{posibles2[0][2]}-{posibles2[0][3].upper()}"

    return "desconocido"

def extraer_numero_factura(texto: str) -> str:
    # print(f'Texto OCR: ',texto)
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
    }

    for k, v in reemplazos.items():
        texto = texto.replace(k, v)

    texto = texto.replace('=', ':').replace('f.1', 'ELECTRONICA').replace('f 1', 'ELECTRONICA')
    texto = texto.replace('Nro =', 'NRO:').replace('Nro :', 'NRO:').replace('Nro::', 'NRO:')
    texto = texto.replace('N:', 'NRO:').replace('NRO:::', 'NRO:')
    texto = texto.replace("NRO :", "NRO:").replace("NRO  :", "NRO:")
    texto = texto.upper()
    texto = re.sub(r'[^\x00-\x7F]+', '', texto)
    lineas = texto.splitlines()
    # print("\n🧼 Texto:\n", texto)

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

    # ✅ Patrón 6: Último número de 5 a 12 dígitos como respaldo si nada anterior coincidió
    if not candidatos:
        match_respaldo = re.findall(r'\b([0-9OQBILSZDE]{5,12})\b', texto)
        for m in match_respaldo:
            if es_posible_numero_factura(m):
                candidatos.append((m, "Respaldo: número general"))

    if candidatos:
        numero_crudo, origen = max(candidatos, key=lambda x: len(x[0]))
        numero_limpio = numero_crudo.upper().translate(str.maketrans({
            'O': '0', 'Q': '0', 'B': '8', 'I': '1', 'L': '1', 'S': '5', 'Z': '2', 'D': '0', 'E': '8'
        }))
        # print(f"🔍 Candidato: {numero_limpio} ({origen})")
        return numero_limpio

    # print("⚠️ Ningún patrón coincidió en el texto filtrado:\n", texto)
    return ""


