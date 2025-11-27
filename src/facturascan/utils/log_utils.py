import os, sys
from datetime import datetime
import re
import urllib.parse


def _get_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(sys.argv[0]))

# Carpeta necesaria
carpeta_base = _get_base_dir()
carpeta_logs = os.path.join(carpeta_base, "logs")

def _ensure_logs_dir():
    try:
        os.makedirs(carpeta_logs, exist_ok=True)
    except Exception:
        pass

# ====== bandera de debug global ======
DEBUG_MODE = False

def set_debug(enabled: bool):
    global DEBUG_MODE
    DEBUG_MODE = bool(enabled)

def is_debug() -> bool:
    return DEBUG_MODE

def registrar_log_proceso(mensaje):
    if not DEBUG_MODE:
        return
    _ensure_logs_dir()
    ahora = datetime.now()
    nombre_log = f"log_procesos_{ahora.strftime('%Y_%m_%d')}.txt"
    ruta_log = os.path.join(carpeta_logs, nombre_log)
    timestamp = ahora.strftime("[%Y-%m-%d %H:%M:%S]")
    try:
        with open(ruta_log, "a", encoding="utf-8") as f:
            f.write(f"{timestamp} {mensaje}\n")
    except Exception:
        pass

def _encode_file_uris(text: str) -> str:
    """
    Busca URIs file://... en el texto y reemplaza espacios por %20
    (u otros caracteres que necesiten encoding), sin romper lo demás.
    Asumimos que la URI va hasta el final de la línea.
    """
    # Tomar todo desde file:// hasta el fin de línea (incluye espacios)
    pattern = r"file://[^\n]+"

    def repl(match):
        uri = match.group(0)

        # Separar prefijo y path
        prefix = "file:///"
        if uri.lower().startswith(prefix):
            path_part = uri[len(prefix):]
        else:
            prefix = "file://"
            path_part = uri[len(prefix):]

        # Codificar el path (espacios → %20, etc.), pero dejar / : \ . _ - y % tal cual
        path_encoded = urllib.parse.quote(path_part, safe="/:\\._-%")

        return prefix + path_encoded

    return re.sub(pattern, repl, text)



def registrar_log(mensaje):
    _ensure_logs_dir()
    ahora = datetime.now()
    nombre_log = f"log_{ahora.strftime('%Y_%m')}_{ahora.strftime('%d')}.txt"
    ruta_log = os.path.join(carpeta_logs, nombre_log)
    timestamp = ahora.strftime("[%Y-%m-%d %H:%M:%S]")

    # Construimos la línea
    linea = f"{timestamp} {mensaje}"

    # Normalizar cualquier file:///... que venga en el mensaje
    try:
        linea = _encode_file_uris(linea)
    except Exception:
        pass

    try:
        with open(ruta_log, "a", encoding="utf-8") as f:
            f.write(linea + "\n")
    except Exception:
        pass


# ====== Links a documentos procesados (para doble click en el log GUI) ======

# --- Mapa de documentos procesados para la UI (nombre -> ruta completa) ---
_DOC_LINKS = {}

def registrar_link_documento(nombre_archivo, ruta_completa):
    """
    Guarda la ruta completa asociada a un nombre de archivo .pdf
    para poder abrirlo desde el log de la interfaz.
    """
    try:
        _DOC_LINKS[nombre_archivo] = ruta_completa
    except Exception:
        pass

def obtener_link_documento(nombre_archivo):
    """
    Devuelve la ruta completa asociada al nombre de archivo, o None si no existe.
    """
    return _DOC_LINKS.get(nombre_archivo)

