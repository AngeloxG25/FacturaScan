import os, sys
from datetime import datetime

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

def registrar_log(mensaje):
    _ensure_logs_dir()
    ahora = datetime.now()
    nombre_log = f"log_{ahora.strftime('%Y_%m')}_{ahora.strftime('%d')}.txt"
    ruta_log = os.path.join(carpeta_logs, nombre_log)
    timestamp = ahora.strftime("[%Y-%m-%d %H:%M:%S]")
    try:
        with open(ruta_log, "a", encoding="utf-8") as f:
            f.write(f"{timestamp} {mensaje}\n")
    except Exception:
        pass
