import os,sys
from datetime import datetime

# Carpeta necesaria
carpeta_base = os.path.dirname(os.path.abspath(sys.argv[0]))
carpeta_logs = os.path.join(carpeta_base, "logs")
os.makedirs(carpeta_logs, exist_ok=True)

#Para los documentos procesados correctamente, desconocidos y documentos sin rut y sin numero de factura
def registrar_log_proceso(mensaje):
    ahora = datetime.now()
    nombre_log = f"log_procesos_{ahora.strftime('%Y_%m_%d')}.txt"
    ruta_log = os.path.join(carpeta_logs, nombre_log)

    timestamp = ahora.strftime("[%Y-%m-%d %H:%M:%S]")
    with open(ruta_log, "a", encoding="utf-8") as f:
        f.write(f"{timestamp} {mensaje}\n")

# Para los procesos del propio sistema
def registrar_log(mensaje):
    ahora = datetime.now()
    nombre_log = f"log_{ahora.strftime('%Y_%m')}_{ahora.strftime('%d')}.txt"
    ruta_log = os.path.join(carpeta_logs, nombre_log)
    timestamp = ahora.strftime("[%Y-%m-%d %H:%M:%S]")
    with open(ruta_log, "a", encoding="utf-8") as f:
        f.write(f"{timestamp} {mensaje}\n")
