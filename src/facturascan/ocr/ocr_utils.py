import os, sys, io, re, logging, contextlib, itertools
from datetime import datetime
from utils.log_utils import registrar_log
# ---- Popup simple (sin txt) ----
def _popup_error(msg: str, title="Error en OCR"):
    try:
        import tkinter as _tk
        from tkinter import messagebox as _mb
        r = _tk.Tk(); r.withdraw()
        _mb.showerror(title, msg)
        r.destroy()
    except Exception:
        # Último recurso: consola
        try: print(msg)
        except Exception: pass

# ---- Silenciar CUDA/GPU y warnings de torch si existiera ----
os.environ["CUDA_VISIBLE_DEVICES"] = "-1"
import warnings
warnings.filterwarnings("ignore")
logging.getLogger("torch").setLevel(logging.ERROR)

# ---- Intenta imports críticos y acumula faltantes para un único popup ----
_missing = []

# Pillow (crítico)
try:
    from PIL import Image, ImageOps
except Exception as e:
    _missing.append(f"- Pillow (PIL): {e}")

# NumPy (crítico para EasyOCR)
try:
    import numpy as np
except Exception as e:
    _missing.append(f"- numpy: {e}")

# Torch (no siempre requerido explícito, pero EasyOCR lo usa)
_torch_ok = True
try:
    import inspect, torch
    # Parche 'weights_only' si falta en esta versión
    try:
        _sig = inspect.signature(torch.load)
        if "weights_only" not in _sig.parameters:
            _orig_load = torch.load
            def _patched_load(*args, **kwargs):
                kwargs.pop("weights_only", None)
                return _orig_load(*args, **kwargs)
            torch.load = _patched_load
    except Exception:
        pass
except Exception as e:
    # No lo tratamos como crítico si EasyOCR puede cargar cpu-only, pero lo avisamos.
    _torch_ok = False

# EasyOCR (crítico)
try:
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        import easyocr
except Exception as e:
    _missing.append(f"- easyocr: {e}")

# Si falta algo crítico, mostramos popup y abortamos import de este módulo
if _missing:
    _popup_error("Faltan dependencias de OCR:\n\n" + "\n".join(_missing), title="Dependencias OCR faltantes")
    # Elevar ImportError para que el llamador lo maneje
    raise ImportError("Dependencias OCR faltantes: " + " | ".join(_missing))

if not _torch_ok:
    try:
        from utils.log_utils import registrar_log_proceso
        registrar_log_proceso("ℹ️ Torch no disponible; EasyOCR usará CPU (OK).")
    except Exception:
        pass

# ================== RESTO DE TUS IMPORTS/UTILS ==================
import threading

try:
    from utils.log_utils import registrar_log_proceso, is_debug
except Exception:
    def is_debug(): return False
    def registrar_log_proceso(*args, **kwargs): pass

# Palabras clave cabecera
_PALABRAS_CLAVE = {"RUT", "FACTURA", "ELECTRONICA", "NRO", "SII"}

_TRANSPOSE_POR_ANGULO = {
    0:   None,
    90:  Image.ROTATE_90,
    180: Image.ROTATE_180,
    270: Image.ROTATE_270,
}

DEBUG_COUNTER = itertools.count(1)

def _unique_path(candidate: str) -> str:
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
    if not p: return False
    if os.path.isdir(p): return True
    if p.endswith(os.sep): return True
    root, ext = os.path.splitext(p)
    return ext == ""

# ================== LECTOR OCR GLOBAL (lazy / perezoso) ==================
import threading

_READER = None
_READER_LOCK = threading.Lock()

def get_reader():
    """Devuelve un único easyocr.Reader inicializado (CPU por defecto)."""
    global _READER
    if _READER is None:
        with _READER_LOCK:
            if _READER is None:
                # silencia stdout/err del load de easyocr/torch
                old_out, old_err = sys.stdout, sys.stderr
                sys.stdout = open(os.devnull, 'w')
                sys.stderr = open(os.devnull, 'w')
                try:
                    # Usa el easyocr ya importado arriba
                    _READER = easyocr.Reader(['es'], gpu=False, verbose=False)
                finally:
                    try:
                        sys.stdout.close(); sys.stderr.close()
                    except Exception:
                        pass
                    sys.stdout, sys.stderr = old_out, old_err
    return _READER

def warmup_ocr():
    """Precarga modelos a RAM para evitar el ‘primer golpe’ al procesar."""
    try:
        from PIL import Image
        import numpy as np
        img = Image.new('L', (8, 8), 255)
        _ = get_reader().readtext(np.array(img))
    except Exception:
        pass


# OCR de cabecera (zona superior derecha)
def ocr_zona_factura_desde_png(imagen_entrada, ruta_debug=None, early_threshold=3):
    """
    Detecta la orientación (0/90/180/270) y realiza OCR en la cabecera superior derecha.
    Pensado para facturas chilenas donde 'RUT', 'FACTURA ELECTRONICA', 'NRO', 'SII' suelen aparecer ahí.

    Optimizaciones:
    - Usa transpose (90°) → más rápido que rotate(expand=True).
    - Recorta zona [x: 65%→100%, y: 1%→30%] y trabaja en escala de grises.
    - Auto-contraste y pequeña reducción para estabilizar OCR.
    - Heurística con "salida temprana": si el puntaje por palabras clave ≥ early_threshold, corta el loop.

    Parámetros:
    - imagen_entrada: ruta (str) o PIL.Image
    - ruta_debug: si se entrega, guarda el recorte (y la imagen rotada si aplica) cuando DEBUG está ON.
    - early_threshold: nivel de confianza mínimo (por keywords) que detona salida temprana.

    Retorna:
    - Texto OCR concatenado del mejor recorte (str).
    """
    # --- Carga imagen desde ruta o PIL.Image
    if isinstance(imagen_entrada, str):
        imagen_original = Image.open(imagen_entrada)
        nombre_base = os.path.splitext(os.path.basename(imagen_entrada))[0]
    elif hasattr(imagen_entrada, "crop"):
        imagen_original = imagen_entrada
        nombre_base = "imagen_en_memoria"
    else:
        raise ValueError("imagen_entrada debe ser una ruta o un objeto PIL.Image")

    # --- Configuración de debug (carpeta ./debug/ + nombre único)
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

    # --- Búsqueda por ángulos (0/90/180/270)
    mejor_texto, mejor_puntaje = "", -1
    mejor_recorte, mejor_angulo = None, 0

    for angulo in (0, 90, 180, 270):
        tr_op = _TRANSPOSE_POR_ANGULO[angulo]
        img = imagen_original if tr_op is None else imagen_original.transpose(tr_op)

        # Recorte superior derecho (65%→100% ancho, 1%→30% alto)
        ancho, alto = img.size
        x0, y0, x1, y1 = int(ancho * 0.60), int(alto * 0.01), int(ancho * 1.00), int(alto * 0.30)
        recorte = img.crop((x0, y0, x1, y1))

        # Preprocesado ligero: gris → reducción → autocontraste
        recorte = ImageOps.grayscale(recorte)
        if recorte.width > 2 and recorte.height > 2:
            recorte = recorte.resize((recorte.width // 2, recorte.height // 2), Image.LANCZOS)
        recorte = ImageOps.autocontrast(recorte, cutoff=1)

        # OCR (permitiendo caracteres típicos de cabeceras de facturas)
        zona_np = np.array(recorte, dtype=np.uint8)
        texto = get_reader().readtext(
            zona_np,
            detail=0,
            batch_size=1,
            allowlist="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-./#:() ",
            mag_ratio=1.0,   # evita magnificación costosa
            width_ths=0.6,   # segmentación menos agresiva
            slope_ths=0.999  # sin corrección de inclinación (ya probamos giros de 90°)
        )
        texto_completo = " ".join(texto).strip()

        # Puntaje por keywords: mientras más coincidan, mejor el ángulo/recorte
        puntaje = sum(1 for p in texto_completo.upper().split() if p in _PALABRAS_CLAVE) if texto_completo else 0

        if puntaje > mejor_puntaje:
            mejor_puntaje, mejor_texto = puntaje, texto_completo
            mejor_recorte, mejor_angulo = recorte, angulo

            # Salida temprana si ya es suficientemente "bueno"
            if mejor_puntaje >= early_threshold:
                break

    # --- Guardados de depuración (solo si DEBUG está ON)
    if debug_activo:
        try:
            if mejor_angulo != 0:
                registrar_log_proceso(f"🔁 Imagen rotada automáticamente {mejor_angulo}°")
                if ruta_debug_final:
                    root, ext = os.path.splitext(ruta_debug_final)
                    ruta_rotada_base = f"{root.rsplit('_recorte', 1)[0]}_rotada{mejor_angulo}{ext}"
                    ruta_rotada = _unique_path(ruta_rotada_base)
                    tr_op = _TRANSPOSE_POR_ANGULO[mejor_angulo]
                    img_rotada = imagen_original if tr_op is None else imagen_original.transpose(tr_op)
                    img_rotada.save(ruta_rotada)

            if ruta_debug_final and mejor_recorte is not None:
                mejor_recorte.save(ruta_debug_final)
                registrar_log_proceso(f"📎 Recorte guardado en: {ruta_debug_final}")
        except Exception as e:
            registrar_log_proceso(f"⚠️ Error guardando recortes de debug: {e}")

    return mejor_texto

# # Extraer RUT del proveedor version 1.9.2
# def extraer_rut(texto):
#     """
#     Extrae un RUT válido (proveedor o cliente) desde texto OCR.
#     - Normaliza variantes de 'RUT' y errores comunes (O→0, I/l→1, B→8, etc).
#     - Valida el dígito verificador (módulo 11).
#     - Prioriza RUT de proveedor (líneas con 'RUT' sin 'CLIENTE').
#       Si no hay, toma RUT cliente.
#     - Si no aparece DV explícito, intenta calcularlo cuando detecta cuerpo plausible.
#     Retorna: 'NNNNNNN-DV' o 'desconocido'.
#     """
#     texto_original = texto
#     # print('texto original: \n',texto)
#     # Normalización de variantes de "RUT" + confusiones de OCR
#     reemplazos = {
#         "RUT.": "RUT", "R.U.T.": "RUT", "R-U-T": "RUT", "RUT": "RUT", "RUT ;": "RUT",
#         "RUT=": "RUT", "RU.T": "RUT", "RU:T": "RUT", "R:UT": "RUT", "RU.T.": "RUT",
#         "RUI": "RUT", "RU1": "RUT", "R.UT.": "RUT", "RuT;": "RUT", "RUTTT;": "RUT",
#         "Ru:,n.": "RUT", "Ru.t:": "RUT", "RVT ;": "RUT", "RVT ": "RUT", "RVT": "RUT",
#         "RUT.:": "RUT", "R.UT.:": "RUT", "R.UI.": "RUT", "R.U.T ": "RUT:", "U.T.": "RUT:",
#         "RU. ": "RUT:","RuT :": "RUT:","R.U.T.::": "RUT:","R.UT": "RUT:","RU.T.::": "RUT:",
#         "R U.T": "RUT:","R.U.": "RUT:","RU:": "RUT:","R.Ut": "RUT:","Rut": "RUT",
#         "R U.I": "RUT","RuT:":"RUT","RUt":"RUT","R.U.1":"RUT","R  U. T ":"RUT",
#         "R U. T":"RUT","RuT.:":"RUT","KUT":"RUT","R.UT:: ":"RUT","RUT":"RUT ",
#         "Ru.T.::":"RUT ","RUT.::":"RUT ","FUT :":"RUT ","RU":"RUT ","R.UI":"RUT ",
#         "U.T:":"RUT ","J.T:":"RUT ","r.u.t":"RUT ","Nre":"RUT ",

#     }
#     for k, v in reemplazos.items():
#         texto = texto.replace(k, v)

#     # Limpiezas OCR (similitudes visuales)
#     texto = texto.replace(',', '.')
#     texto = (texto
#              .replace('O', '0').replace('o', '0')
#              .replace('I', '1').replace('l', '1')
#              .replace('B', '8').replace('Z', '2').replace('G', '6'))
#     texto = texto.replace('–', '-').replace('—', '-').replace('‐', '-')
#     texto = texto.replace('+', '-')
#     texto = texto.replace('u', '0')

#     # print("🟢 (RUT Limpio):\n", texto)

#     # Cálculo del DV por módulo 11
#     def calcular_dv(rut_sin_dv: str) -> str:
#         try:
#             rut = list(map(int, rut_sin_dv[::-1]))
#         except ValueError:
#             return ""
#         factores = [2, 3, 4, 5, 6, 7]
#         suma = 0
#         for i, d in enumerate(rut):
#             suma += d * factores[i % len(factores)]
#         resto = 11 - (suma % 11)
#         if resto == 11: return "0"
#         if resto == 10: return "K"
#         return str(resto)

#     # Patrones base (toleran puntos/espacios)
#     RUT_CUERPO = r'(\d{1,2}(?:\s*\.?\s*\d{3}){2})'
#     RUT_DV     = r'\s*[-‐–—]?\s*([\dkK])'

#     candidatos_proveedor, candidatos_cliente = [], []

#     # Primero, búsqueda línea por línea para usar el contexto semántico
#     for linea in texto.splitlines():
#         u = linea.upper()

#         # Líneas tipo "RUT CLIENTE ..."
#         if "RUT" in u and "CLIENTE" in u:
#             m = re.search(rf'RUT\b[^\dKk]{{0,10}}{RUT_CUERPO}{RUT_DV}', u)
#             if m:
#                 cuerpo, dv = m.group(1), m.group(2).upper()
#                 rut_sin = re.sub(r'\D', '', cuerpo)
#                 if len(rut_sin) in (7, 8) and dv == calcular_dv(rut_sin):
#                     candidatos_cliente.append(f"{rut_sin}-{dv}")

#             # Variante más tolerante (más ruido entre tokens)
#             m2 = re.search(rf'RUT\s*CLIENTE\b[^\dKk]{{0,15}}{RUT_CUERPO}{RUT_DV}', u)
#             if m2:
#                 cuerpo, dv = m2.group(1), m2.group(2).upper()
#                 rut_sin = re.sub(r'\D', '', cuerpo)
#                 if len(rut_sin) in (7, 8) and dv == calcular_dv(rut_sin):
#                     candidatos_cliente.append(f"{rut_sin}-{dv}")
#             continue

#         # Líneas con "RUT" que NO sean "CLIENTE" (se asumen proveedor)
#         if "RUT" in u and "CLIENTE" not in u:
#             m = re.search(rf'RUT\b[^\dKk]{{0,10}}{RUT_CUERPO}{RUT_DV}', u)
#             if m:
#                 cuerpo, dv = m.group(1), m.group(2).upper()
#                 rut_sin = re.sub(r'\D', '', cuerpo)
#                 if len(rut_sin) in (7, 8) and dv == calcular_dv(rut_sin):
#                     candidatos_proveedor.append(f"{rut_sin}-{dv}")
#             continue

#     # Si no hubo DV explícito, intenta deducirlo en líneas donde aparece "RUT"
#     if not candidatos_proveedor and not candidatos_cliente:
#         for linea in texto.splitlines():
#             u = linea.upper()
#             if "RUT" not in u:
#                 continue
#             m = re.search(rf'RUT\b[^\dKk]{{0,10}}(\d[\d\.\s]{{6,11}})(?![-0-9Kk])', u)
#             if m:
#                 rut_sin = re.sub(r'\D', '', m.group(1))
#                 if 7 <= len(rut_sin) <= 8:
#                     dv = calcular_dv(rut_sin)
#                     if dv:
#                         candidatos_proveedor.append(f"{rut_sin}-{dv}")
    
#     # Selección final (preferimos proveedor; si no hay, cliente)
#     if candidatos_proveedor:
#         rut = candidatos_proveedor[0]
#         registrar_log_proceso(f"✅ RUT validado (proveedor): {rut}")
#         registrar_log(f'🪪 Rut detectado: {rut}')
#         return rut
#     if candidatos_cliente:
#         rut = candidatos_cliente[0]
#         registrar_log_proceso(f"✅ RUT validado (cliente): {rut}")
#         registrar_log(f'🪪 Rut detectado: {rut}')
#         return rut
    
#     registrar_log_proceso("⚠️ RUT no detectado.")
#     return "desconocido"

# modificado 1.9.3 rc1
# Extraer RUT del proveedor
# def extraer_rut(texto):
#     """
#     Extrae un RUT válido (proveedor o cliente) desde texto OCR.
#     - Normaliza variantes de 'RUT' y errores comunes (O→0, I/l→1, B→8, etc).
#     - Valida el dígito verificador (módulo 11).
#     - Prioriza RUT de proveedor (líneas con 'RUT' sin 'CLIENTE').
#       Si no hay, toma RUT cliente.
#     - Si no aparece DV explícito, intenta calcularlo cuando detecta cuerpo plausible.
#     Retorna: 'NNNNNNN-DV' o 'desconocido'.
#     """
    
#     print('texto original: \n',texto)
#     # Normalización de variantes de "RUT" + confusiones de OCR
#     reemplazos = {
#         "RUT.": "RUT", "R.U.T.": "RUT", "R-U-T": "RUT", 
#         # "RUT": "RUT",
#           "RUT ;": "RUT",
#         "RUT=": "RUT", "RU.T": "RUT", "RU:T": "RUT", "R:UT": "RUT", "RU.T.": "RUT",
#         "RUI": "RUT", "RU1": "RUT", "R.UT.": "RUT", "RuT;": "RUT", "RUTTT;": "RUT",
#         "Ru:,n.": "RUT", "Ru.t:": "RUT", "RVT ;": "RUT", "RVT ": "RUT", "RVT": "RUT",
#         "RUT.:": "RUT", "R.UT.:": "RUT", "R.UI.": "RUT", "R.U.T ": "RUT", "U.T.": "RUT",
#         "RU. ": "RUT",
#         # "RuT :": "RUT:",
#         "R.U.T.::": "RUT","R.UT": "RUT","RU.T.::": "RUT",
#         "R U.T": "RUT","R.U.": "RUT","RU:": "RUT","R.Ut": "RUT",
#         # "Rut": "RUT",

#         "R U.I": "RUT","RuT:":"RUT","RUt":"RUT","R.U.1":"RUT","R  U. T ":"RUT",
#         "R U. T":"RUT","RuT.:":"RUT","KUT":"RUT","R.UT:: ":"RUT",
#         # "RUT":"RUT ",

#         "Ru.T.::":"RUT ","RUT.::":"RUT ","FUT :":"RUT ",
#         # ,"R.UI":"RUT ",
#         "U.T:":"RUT ","J.T:":"RUT ","r.u.t":"RUT ","Nre":"RUT ","Rut:":"RUT "

#     }
#     for k, v in reemplazos.items():
#         texto = texto.replace(k, v)

#     # Limpiezas OCR (similitudes visuales)
#     texto = texto.replace(',', '.')
#     texto = (texto
#              .replace('O', '0').replace('o', '0')
#              .replace('I', '1').replace('l', '1')
#              .replace('B', '8').replace('Z', '2').replace('G', '6'))
#     texto = texto.replace('–', '-').replace('—', '-').replace('‐', '-')
#     texto = texto.replace('+', '-')
#     # texto = texto.replace('u', '0')

#     print("🟢 (RUT Limpio):\n", texto)

#     # Cálculo del DV por módulo 11
#     def calcular_dv(rut_sin_dv: str) -> str:
#         try:
#             rut = list(map(int, rut_sin_dv[::-1]))
#         except ValueError:
#             return ""
#         factores = [2, 3, 4, 5, 6, 7]
#         suma = 0
#         for i, d in enumerate(rut):
#             suma += d * factores[i % len(factores)]
#         resto = 11 - (suma % 11)
#         if resto == 11: return "0"
#         if resto == 10: return "K"
#         return str(resto)

#     # Patrones base (toleran puntos/espacios)
#     RUT_CUERPO = r'(\d{1,2}(?:\s*\.?\s*\d{3}){2})'
#     RUT_DV     = r'\s*[-‐–—]?\s*([\dkK])'

#     candidatos_proveedor, candidatos_cliente = [], []

#     # Primero, búsqueda línea por línea para usar el contexto semántico
#     for linea in texto.splitlines():
#         u = linea.upper()

#         # Líneas tipo "RUT CLIENTE ..."
#         if "RUT" in u and "CLIENTE" in u:
#             m = re.search(rf'RUT\b[^\dKk]{{0,10}}{RUT_CUERPO}{RUT_DV}', u)
#             if m:
#                 cuerpo, dv = m.group(1), m.group(2).upper()
#                 rut_sin = re.sub(r'\D', '', cuerpo)
#                 if len(rut_sin) in (7, 8) and dv == calcular_dv(rut_sin):
#                     candidatos_cliente.append(f"{rut_sin}-{dv}")

#             # Variante más tolerante (más ruido entre tokens)
#             m2 = re.search(rf'RUT\s*CLIENTE\b[^\dKk]{{0,15}}{RUT_CUERPO}{RUT_DV}', u)
#             if m2:
#                 cuerpo, dv = m2.group(1), m2.group(2).upper()
#                 rut_sin = re.sub(r'\D', '', cuerpo)
#                 if len(rut_sin) in (7, 8) and dv == calcular_dv(rut_sin):
#                     candidatos_cliente.append(f"{rut_sin}-{dv}")
#             continue

#         # Líneas con "RUT" que NO sean "CLIENTE" (se asumen proveedor)
#         if "RUT" in u and "CLIENTE" not in u:
#             m = re.search(rf'RUT\b[^\dKk]{{0,10}}{RUT_CUERPO}{RUT_DV}', u)
#             if m:
#                 cuerpo, dv = m.group(1), m.group(2).upper()
#                 rut_sin = re.sub(r'\D', '', cuerpo)
#                 if len(rut_sin) in (7, 8) and dv == calcular_dv(rut_sin):
#                     candidatos_proveedor.append(f"{rut_sin}-{dv}")
#             continue

#     # Si no hubo DV explícito, intenta deducirlo en líneas donde aparece "RUT"
#     if not candidatos_proveedor and not candidatos_cliente:
#         for linea in texto.splitlines():
#             u = linea.upper()
#             if "RUT" not in u:
#                 continue
#             m = re.search(rf'RUT\b[^\dKk]{{0,10}}(\d[\d\.\s]{{6,11}})(?![-0-9Kk])', u)
#             if m:
#                 rut_sin = re.sub(r'\D', '', m.group(1))
#                 if 7 <= len(rut_sin) <= 8:
#                     dv = calcular_dv(rut_sin)
#                     if dv:
#                         candidatos_proveedor.append(f"{rut_sin}-{dv}")
    
#     # Selección final (preferimos proveedor; si no hay, cliente)
#     if candidatos_proveedor:
#         rut = candidatos_proveedor[0]
#         registrar_log_proceso(f"✅ RUT validado (proveedor): {rut}")
#         registrar_log(f'🪪 Rut detectado: {rut}')
#         return rut
#     if candidatos_cliente:
#         rut = candidatos_cliente[0]
#         registrar_log_proceso(f"✅ RUT validado (cliente): {rut}")
#         registrar_log(f'🪪 Rut detectado: {rut}')
#         return rut
    
#     registrar_log_proceso("⚠️ RUT no detectado.")
#     return "desconocido"

def extraer_rut(texto: str) -> str:
    """
    Extrae un RUT válido (proveedor o cliente) desde texto OCR.
    - Normaliza variantes de 'RUT' y errores comunes (O→0, I/l→1, B→8, Z→2, G→6).
    - Valida/corrige el dígito verificador (módulo 11).
    - Prioriza RUT de proveedor (líneas con 'RUT' sin 'CLIENTE'); si no, cliente.
    - Si no aparece DV explícito, lo calcula cuando detecta cuerpo plausible.
    Retorna: 'NNNNNNN-DV' o 'desconocido'.
    """
    # print('texto original: \n', texto)

    # ---- Normalización segura de "RUT" (evita "RUT T:") ----
    texto = re.sub(r'\bR\s*[UUV]\s*[T7]{1,3}\s*[:\.\-;]?\b', 'RUT:', texto, flags=re.IGNORECASE)
    texto = texto.replace('RUT :', 'RUT:')

    # ---- Reemplazos OCR útiles (se mantienen; quitamos los peligrosos) ----
    reemplazos = {
        "RUT.": "RUT", "R.U.T.": "RUT", "R-U-T": "RUT",
        "RUT ;": "RUT",
        "RUT=": "RUT", "RU.T": "RUT", "RU:T": "RUT", "R:UT": "RUT", "RU.T.": "RUT",
        "RUI": "RUT", "RU1": "RUT", "R.UT.": "RUT", "RuT;": "RUT", "RUTTT;": "RUT",
        "Ru:,n.": "RUT", "Ru.t:": "RUT", "RVT ;": "RUT", "RVT ": "RUT", "RVT": "RUT",
        "RUT.:": "RUT", "R.UT.:": "RUT", "R.UI.": "RUT", "R.U.T ": "RUT", "U.T.": "RUT",
        "RU. ": "RUT",
        "R.U.T.::": "RUT", "R.UT": "RUT", "RU.T.::": "RUT",
        "R U.T": "RUT", "R.U.": "RUT", 
        # "RU:": "RUT",
          "R.Ut": "RUT",
        "R U.I": "RUT", "RuT:":"RUT", "RUt":"RUT", "R.U.1":"RUT", "R  U. T ":"RUT",
        "R U. T":"RUT", "RuT.:":"RUT", "KUT":"RUT", "R.UT:: ":"RUT",
        "Ru.T.::":"RUT ", "RUT.::":"RUT ", "FUT :":"RUT ",
        # ¡NO usar "RU":"RUT " ni "RUT":"RUT "!
        "U.T:":"RUT ", "J.T:":"RUT ", "r.u.t":"RUT ", "Nre":"RUT ", "Rut:":"RUT "
    }
    for k, v in reemplazos.items():
        texto = texto.replace(k, v)

    # ---- Limpiezas OCR generales (sin 'u'->'0') ----
    texto = texto.replace(',', '.')
    texto = (texto
             .replace('O', '0').replace('o', '0')
             .replace('I', '1').replace('l', '1')
             .replace('B', '8').replace('Z', '2').replace('G', '6'))
    texto = texto.replace('–', '-').replace('—', '-').replace('‐', '-')
    texto = texto.replace('+', '-')

    # print("🟢 (RUT Limpio):\n", texto)

    # ---- Cálculo del DV ----
    def calcular_dv(rut_sin_dv: str) -> str:
        try:
            nums = list(map(int, rut_sin_dv[::-1]))
        except ValueError:
            return ""
        factores = [2, 3, 4, 5, 6, 7]
        s = 0
        for i, d in enumerate(nums):
            s += d * factores[i % len(factores)]
        resto = 11 - (s % 11)
        if resto == 11: return "0"
        if resto == 10: return "K"
        return str(resto)

    # ---- Patrones ----
    RUT_CUERPO   = r'(\d{1,2}(?:\s*\.?\s*\d{3}){2})'
    RUT_DV_OPT   = r'(?:\s*[-‐–—]\s*([\dkK]))?'           # DV opcional
    PATRON_RUT   = rf'RUT\b[^\dKk]{{0,15}}{RUT_CUERPO}{RUT_DV_OPT}'
    PATRON_GLOBAL= rf'{RUT_CUERPO}{RUT_DV_OPT}'

    candidatos_proveedor, candidatos_cliente = [], []

    def procesa_match(m, es_cliente: bool):
        cuerpo, dv = m.group(1), (m.group(2) or "").upper()
        rut_sin = re.sub(r'\D', '', cuerpo)
        if 7 <= len(rut_sin) <= 8:
            dv_calc = calcular_dv(rut_sin)
            if not dv:
                rut = f"{rut_sin}-{dv_calc}"
                registrar_log_proceso(f"✅ RUT completado ({'cliente' if es_cliente else 'proveedor'}): {rut} (sin DV OCR)")
            elif dv == dv_calc:
                rut = f"{rut_sin}-{dv}"
                registrar_log_proceso(f"✅ RUT validado ({'cliente' if es_cliente else 'proveedor'}): {rut}")
            else:
                rut = f"{rut_sin}-{dv_calc}"
                registrar_log_proceso(f"✅ RUT corregido ({'cliente' if es_cliente else 'proveedor'}): {rut} (OCR DV='{dv}'→'{dv_calc}')")
            if es_cliente:
                candidatos_cliente.append(rut)
            else:
                candidatos_proveedor.append(rut)

    # ---- Línea por línea con contexto ----
    for linea in texto.splitlines():
        u = linea.upper()

        if "RUT" in u and "CLIENTE" in u:
            m = re.search(PATRON_RUT, u)
            if m:
                procesa_match(m, es_cliente=True)
            # Variante "RUT CLIENTE ... número"
            m2 = re.search(rf'RUT\s*CLIENTE\b[^\dKk]{{0,20}}{RUT_CUERPO}{RUT_DV_OPT}', u)
            if m2:
                procesa_match(m2, es_cliente=True)
            continue

        if "RUT" in u and "CLIENTE" not in u:
            m = re.search(PATRON_RUT, u)
            if m:
                procesa_match(m, es_cliente=False)
            continue

    # ---- Rescate global si no hubo nada ----
    if not candidatos_proveedor and not candidatos_cliente:
        for m in re.finditer(PATRON_GLOBAL, texto.upper()):
            cuerpo, dv = m.group(1), (m.group(2) or "").upper()
            rut_sin = re.sub(r'\D', '', cuerpo)
            if 7 <= len(rut_sin) <= 8:
                dv_calc = calcular_dv(rut_sin)
                rut = f"{rut_sin}-{dv if dv and dv == dv_calc else dv_calc}"
                if dv and dv != dv_calc:
                    registrar_log_proceso(f"ℹ️ Rescate global: DV corregido {rut} (OCR DV='{dv}')")
                candidatos_proveedor.append(rut)  # por defecto proveedor

    # ---- Selección final ----
    if candidatos_proveedor:
        rut = candidatos_proveedor[0]
        registrar_log(f'🪪 Rut detectado: {rut}')
        return rut
    if candidatos_cliente:
        rut = candidatos_cliente[0]
        registrar_log(f'🪪 Rut detectado: {rut}')
        return rut

    registrar_log_proceso("⚠️ RUT no detectado.")
    return "desconocido"

# Extraer el Número de Factura
def extraer_numero_factura(texto: str) -> str:
    """
    Extrae el número de factura desde texto OCR.
    - Normaliza prefijos (N°, Nº, N:, Nro., FOLIO, etc.) -> unifica como 'NRO:'
    - Corrige caracteres confundidos (O/Q→0, I/L→1, S→5, Z→2, U→0, '/'→'1')
    - Aplica reglas para evitar falsos positivos (teléfonos, códigos muy cortos/largos).
    - Priorización de candidatos por contexto (aparece junto a 'FACTURA ELECTRONICA', 'NRO:', etc.).
    Retorna: número como string ('' si no se detecta).
    """
    # print("🟡 Texto OCR original (Número Factura):\n", texto)

    def corregir_ocr_numero(numero: str) -> str:
        """Normaliza dígitos con confusiones típicas de OCR y elimina separadores."""
        traduccion = str.maketrans({
            'O': '0', 'Q': '0', 'B': '8', 'I': '1', 'L': '1', 'S': '5',
            'Z': '2', 'D': '0', 'E': '8', 'A': '4', 'U': '0', '/': '1'
        })
        return numero.translate(traduccion).replace('.', '').replace(' ', '')

    # --- PRIORIDAD MÁXIMA: NP ###### (antes de normalizar a NRO) ---
    # Soportamos variantes OCR: N°P, N P, 'NP, NP:, etc.
    texto_up = texto.upper()

    # m_np = re.search(
    #     r'\bN[\s°º\'"]?P\b\s*[:=\-]?\s*([0-9OQBILSZDEUA]{6,12})',
    #     texto_up
    # )
    # if m_np:
    #     cand = corregir_ocr_numero(m_np.group(1))
    #     if cand.isdigit() and 6 <= len(cand) <= 12:
    #         return cand
    
    # --- PRIORIDAD ALTA: "N° Folio: ######" y variantes OCR ---
    # Cubre: N° Folio: 12345678 | Nc Folio: 12345678 | N Folio: 12345678
    #        FOLIO N° 12345678  | FOLIO No 12345678, etc.
    m_folio_right = re.search(
        r'\bN[\s°ºcC]?\s*FOLIO\b\s*[:=\-]?\s*([0-9OQBILSZDEUA]{6,12})',
        texto_up
    )
    if m_folio_right:
        cand = corregir_ocr_numero(m_folio_right.group(1))
        if cand.isdigit() and 6 <= len(cand) <= 12:
            return cand

    m_folio_left = re.search(
        r'\bFOLIO\b\s*(?:N[\s°ºcC]?|NO\.?)\s*[:=\-]?\s*([0-9OQBILSZDEUA]{6,12})',
        texto_up
    )
    if m_folio_left:
        cand = corregir_ocr_numero(m_folio_left.group(1))
        if cand.isdigit() and 6 <= len(cand) <= 12:
            return cand


    # ---- Normalizaciones de prefijos / textos ruidosos → 'NRO'
    reemplazos = {
        "NP Folio:": "NRO","NI Folio:": "NRO",
        "N°": "NRO ", "N'": "NRO ", 'N"': "NRO ", "N :": "NRO ", "N.": "NRO ",
        "Nº": "NRO ", "N:": "NRO ", "NE": "NRO ", "N?": "NRO ", "FNLC": "NRO ",
        "FNL": "NRO ", "FNLD": "NRO ", "FULD": "NRO ", "FOLIO": "NRO ",
        "NC:": "NRO:", "NC ": "NRO ", "N C": "NRO ", '"NC': "NRO ", "'NC": "NRO ",
        "NP ": "NRO ", "N°P": "NRO ", "N P": "NRO ", '"NP': "NRO ", "'NP": "NRO ",
        "NP:": "NRO:", "Nro.:": "NRO:", "Nro. :": "NRO:", "Nro :": "NRO :",
        "Nro.": "NRO", "NF": "NRO", "NiP": "NRO", "MP": "NRO", "NO": "NRO",
        "Nro.  :": "NRO:", "N9": "NRO", "Ne": "NRO", "nE": "NRO", "Nro": "NRO",
        "Nro ": "NRO", "Nro  ": "NRO", "Folio N?": "NRO", "FOLION?": "NRO ",
        "FOLIO N°": "NRO ", "FCLIO": "NRO ", "NUMERO": "NRO ",
        "NUMERO .": "NRO ", "No": "NRO ", "Nra": "NRO ",
        "Folio /": "NRO ", "RUMERO :": "NRO ", "Né": "NRO ", "N '": "NRO ",
        "NT ": "NRO ", "Nt": "NRO ", "eg.6n N": "NRO ", "ND": "NRO ", "Folio H": "NRO ",
        "N 0": "NRO ", "Nm": "NRO ", "Ni": "NRO ", "KUMERO": "NRO ", "HUMERO": "NRO ",
        "NUHERO": "NRO ", "XUMERO": "NRO ", "Nro: ": "NRO", "NP": "NRO", "NM": "NRO",
        "MUMEn": "NRO","Munern": "NRO","Nr0": "NRO","Ng": "NRO","Np": "NRO","2N#": "NRO",
        "Nw": "NRO","N  Folio:": "NRO","NP  Folio:": "NRO"," N  Folio:": "NRO",
        "No Folio:": "NRO","NP Folio:": "NRO","Nv Folio:": "NRO","MC": "NRO",

    }
    for k, v in reemplazos.items():
        texto = texto.replace(k, v)

    # Limpiezas varias + normalización de "FACTURA ELECTRONICA"
    texto = texto.replace('=', ':').replace('f.1', 'ELECTRONICA').replace('f 1', 'ELECTRONICA')
    texto = texto.replace('Nro =', 'NRO:').replace('Nro :', 'NRO:').replace('Nro::', 'NRO:')
    texto = texto.replace('N:', 'NRO:').replace('NRO:::', 'NRO:')
    texto = texto.replace("NRO :", "NRO:").replace("NRO  :", "NRO:")
    texto = re.sub(r'F[A4@][CÇ][\s\-]*U[R][A4@][\s\-]*E[L1I][E3]C[T7][R][O0][N][I1][C][A4@]', "FACTURA ELECTRONICA", texto, flags=re.IGNORECASE)
    texto = re.sub(r'F[\s\-]*A[\s\-]*C[\s\-]*U[\s\-]*R[\s\-]*A[\s\-]*[E3][L1I][E3]C[T7][R][O0][N][I1][C][A4@]', "FACTURA ELECTRONICA", texto, flags=re.IGNORECASE)
    texto = (texto.replace("FACTURAELECTRONICA", "FACTURA ELECTRONICA")
                   .replace("FACTURAELECTRÓNICA", "FACTURA ELECTRONICA")
                   .replace("FACTURAELECTRONICP", "FACTURA ELECTRONICA")
                   .replace("FACTVRAELECTRoNICA", "FACTURA ELECTRONICA")
                   .replace("FACTURA ELECTRÓNICA", "FACTURA ELECTRONICA")
                   .replace("FACTURA ELECTRNICA", "FACTURA ELECTRONICA")
                   .replace("FACTURA FLECTRONICA", "FACTURA ELECTRONICA")
                   .replace("FAC URA FLECTRONICA", "FACTURA ELECTRONICA")
                   .replace("FACTURA ELECIRONICA", "FACTURA ELECTRONICA")
                   .replace("FACTURLELECTRONICA", "FACTURA ELECTRONICA")
                   .replace("FACTURA Electronica", "FACTURA ELECTRONICA")
                   .replace("FACTURA ELECTRNICA FOLIO", "FACTURA ELECTRONICA")
                   .replace("FACTURA ELECTRNICA FOLIO NRO", "FACTURA ELECTRONICA")
                   .replace("FACTURA ELECTRONICA NO1", "FACTURA ELECTRONICA")
                   .replace("ACTURA ELECTRONICA N", "FACTURA ELECTRONICA")
                   .replace("ALTURA ELECTRONICA", "FACTURA ELECTRONICA"))
    texto = texto.upper()
    texto = re.sub(r'[^\x00-\x7F]+', '', texto)                # elimina acentos / caracteres no ASCII
    texto = re.sub(r'\bN\s*R\s*O\b', 'NRO', texto)             # unifica "N R O" → "NRO"

    # Unifica "NRO: <num>" para facilitar extracción posterior
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

    # print("🟢 (Número Factura Limpio):\n", texto)

    lineas = texto.splitlines()
    candidatos = []

    # Palabras que suelen aparecer junto a números que NO son folios (evitar falsos positivos)
    distractores = {"CONTROL", "DOC", "INTERNO", "DOC INTERNO", "ORDEN", "KP/G.D", "GUIA", "GD", "VEND", "CLIENTE"}

    def es_posible_numero_factura(num: str) -> bool:
        """Filtra teléfonos/códigos: 3..12 dígitos, todo numérico."""
        num = corregir_ocr_numero(num)
        if re.match(r'^\+?56\s?\d{2}\s?\d{4,}$', num):  # teléfono CL
            return False
        if len(num) < 3 or len(num) > 12:
            return False
        return num.isdigit()

    # Prioridad fuerte: "FACTURA ELECTRONICA ... N <num>"
    m_factura_n = re.search(r'FACTURA\s+ELECTRONICA(?:[^\d]{0,40})\bN[\s\.:=\-]*([0-9OQBILSZDEUA]{6,12})', texto)
    if m_factura_n:
        raw = m_factura_n.group(1)
        cand = corregir_ocr_numero(raw)
        if es_posible_numero_factura(cand):
            candidatos.append((cand, "FacturaN"))

    # === NUEVO: "FACTURA ELECTRONICA <num>" directo, sin 'N' intermedio ===
    m_factura_directa = re.search(r'FACTURA\s+ELECTRONICA[^\d]{0,10}([0-9OQBILSZDEUA]{6,12})', texto)
    if m_factura_directa:
        raw = m_factura_directa.group(1)
        cand = corregir_ocr_numero(raw)
        if es_posible_numero_factura(cand):
            candidatos.append((cand, "FacturaE_strict"))

    # Búsqueda por líneas (con distintos patrones y contexto)
    for i, linea in enumerate(lineas):
        linea_upper = linea.upper()
        cerca_de_rut = ("RUT" in linea_upper)

        # NRO:12345
        match_exacto = re.search(r'NRO[\s:=\.\-]*([0-9OQBILSZDEUA]{3,20})', linea_upper)
        if match_exacto:
            candidato = corregir_ocr_numero(match_exacto.group(1))
            if es_posible_numero_factura(candidato):
                etiqueta = "NRO: exacto (cerca RUT)" if cerca_de_rut else "NRO: exacto"
                candidatos.append((candidato, etiqueta))
            continue

        # N  R O ... 12345 (tolerante a ruido)
        match_tolerante = re.search(r'\bN[\s°ºOo0]?R?O?[\.:=\- ]{0,10}([0-9OQBILSZDEUA]{3,20})\b', linea_upper)
        if match_tolerante:
            candidato = corregir_ocr_numero(match_tolerante.group(1))
            if es_posible_numero_factura(candidato):
                etiqueta = "Prefijo tolerante (cerca RUT)" if cerca_de_rut else "Prefijo tolerante"
                candidatos.append((candidato, etiqueta))
            continue

        # "FACTURA ELECTRONICA <num>" en la misma línea
        match_factura = re.search(r'FACTURA\s+(?:F\s*1\s+)?ELECTRONICA[^0-9A-Z]{0,10}([0-9OQBILSZDEUA\.]{3,20})\b', linea_upper)
        if match_factura:
            raw = match_factura.group(1).split()[0]
            candidato = corregir_ocr_numero(raw)
            if es_posible_numero_factura(candidato):
                etiqueta = "Factura + número en línea (cerca RUT)" if cerca_de_rut else "Factura + número en línea"
                candidatos.append((candidato, etiqueta))
            continue

        # "NO12345" pegado
        match_nopegado = re.search(r'\bNO([0-9OQBILSZDEUA]{3,20})\b', linea_upper)
        if match_nopegado:
            candidato = corregir_ocr_numero(match_nopegado.group(1))
            if es_posible_numero_factura(candidato):
                candidatos.append((candidato, "NO+Número sin espacio"))
            continue

        # "FACTURA" en una línea y número en la siguiente
        if "FACTURA" in linea_upper and i + 1 < len(lineas):
            siguiente = lineas[i + 1].upper()
            match_sig = re.search(r'\b([0-9OQBILSZDEUA]{3,20})\b', siguiente)
            if match_sig:
                candidato = corregir_ocr_numero(match_sig.group(1))
                if es_posible_numero_factura(candidato):
                    candidatos.append((candidato, "Número en línea inferior"))
            continue

        # Línea compuesta solo por números (posible folio). Penaliza si hay palabras distractoras.
        match_solo = re.fullmatch(r'\s*([0-9OQBILSZDEUA]{3,20})\s*', linea_upper)
        if match_solo:
            candidato = corregir_ocr_numero(match_solo.group(1))
            if es_posible_numero_factura(candidato):
                if any(w in linea_upper for w in distractores):
                    candidatos.append((candidato, "Línea numérica pura (distractor)"))
                else:
                    candidatos.append((candidato, "Línea numérica pura"))
            continue

    # Respaldo: cualquier número plausible
    if not candidatos:
        for m in re.findall(r'\b([0-9OQBILSZDEUA]{3,20})\b', texto):
            candidato = corregir_ocr_numero(m)
            if es_posible_numero_factura(candidato):
                candidatos.append((candidato, "Respaldo: número general"))

    if not candidatos:
        return ""

    # Si hay alguno de 6+ dígitos, descarta candidatos cortos (evita IDs/códigos incidentales)
    if any(len(c[0]) >= 6 for c in candidatos):
        candidatos = [c for c in candidatos if len(c[0]) >= 6]

    # Priorización por contexto + longitud
    prioridad = {
        "NP_strict": 7,                 
        "FacturaE_strict": 6,
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

    # Si ninguna etiqueta mapea, usa el más largo como heurística
    if not any(etq in prioridad for _, etq in candidatos):
        numero_crudo, _ = max(candidatos, key=lambda x: len(x[0]))
        return numero_crudo

    # Selección final: prioridad -> largo
    numero_crudo, _ = max(candidatos, key=lambda x: (prioridad.get(x[1], 0), len(x[0])))
    registrar_log(f'#️⃣  Núm. factura detectado: {numero_crudo}')
    return numero_crudo

# --- CHEP detection -----------------------------------------------------------
import re, unicodedata

def _norm(txt: str) -> str:
    """Normaliza: mayúsculas, sin acentos, espacios compactos."""
    if not txt:
        return ""
    t = unicodedata.normalize("NFKD", txt)
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    t = t.upper()
    t = re.sub(r"\s+", " ", t)
    return t

# patrón de código CHEP: B + 10 o más dígitos
_CHEP_CODE_RE   = re.compile(r"\bB\s*[-]?\s*\d{10,}\b")
# frases que validan el documento (al menos una debe estar)
_CHEP_KEYWORDS  = [
    re.compile(r"FECHA\s*DE\s*CARGA"),
    re.compile(r"FECHA\s*DE\s*ENVIO"),   # <<< añadido
]

def looks_like_chep(text: str) -> bool:
    """True si hay código B########## y (Fecha de carga o Fecha de envío)."""
    t = _norm(text)
    if not t:
        return False
    code_ok = bool(_CHEP_CODE_RE.search(t))
    kw_ok   = any(p.search(t) for p in _CHEP_KEYWORDS)
    return code_ok and kw_ok
# -----------------------------------------------------------------------------


