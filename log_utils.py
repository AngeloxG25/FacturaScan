import os
from datetime import datetime

def registrar_log_proceso(mensaje):
    ahora = datetime.now()
    carpeta_logs = r"C:\FacturaScan\logs_procesos"
    os.makedirs(carpeta_logs, exist_ok=True)

    nombre_log = f"log_procesos_{ahora.strftime('%Y_%m_%d')}.txt"
    ruta_log = os.path.join(carpeta_logs, nombre_log)

    timestamp = ahora.strftime("[%Y-%m-%d %H:%M:%S]")
    with open(ruta_log, "a", encoding="utf-8") as f:
        f.write(f"{timestamp} {mensaje}\n")
