# Utilidades para post-procesar PDFs:
#  - compresión mediante Ghostscript (GS)
#  - generación de nombres únicos
#
# Notas:
#  - Este módulo asume entorno Windows (usa flags de subprocess propios de Windows).
#  - Para ocultar la ventana de consola de GS, se combina hide_subprocess + STARTUPINFO.
#  - El llamador suele verificar que GS esté disponible (GS_PATH), pero aquí agregamos
#    verificaciones defensivas para evitar excepciones innecesarias.

# Parchea subprocess para ocultar CMDs en Windows (no hace nada en otros SO)
import hide_subprocess
import subprocess
import os

from log_utils import registrar_log_proceso


def comprimir_pdf(gs_path, input_path, calidad="screen", dpi=100, tamano_pagina="a4"):
    """
    Comprime y normaliza un PDF usando Ghostscript.

    Parámetros:
        gs_path (str): Ruta absoluta al ejecutable de Ghostscript (gswin64c.exe/gswin32c.exe).
        input_path (str): Ruta al PDF de entrada (será reemplazado en sitio si todo ok).
        calidad (str): Perfil de GS: 'screen', 'ebook', 'printer', 'prepress', 'default'.
        dpi (int): Resolución objetivo para downsampling de imágenes.
        tamano_pagina (str): Papel destino para -sPAPERSIZE (p.ej., 'a4', 'letter').

    Comportamiento:
        - Escribe un archivo temporal `<nombre>_comprimido.pdf`.
        - Si la ejecución es exitosa, reemplaza el original por el comprimido.
        - Si el reemplazo falla, conserva el comprimido como fallback `<nombre>_compresion_fallback.pdf`.
        - Loguea cada paso mediante registrar_log_proceso.
    """
    try:
        # Validaciones defensivas (evitan CalledProcessError innecesarias)
        if not gs_path or not os.path.exists(gs_path):
            registrar_log_proceso("⚠️ Ghostscript no encontrado o ruta inválida. Se omite compresión.")
            return
        if not input_path or not os.path.exists(input_path):
            registrar_log_proceso("⚠️ PDF de entrada no existe. Se omite compresión.")
            return
        if not input_path.lower().endswith(".pdf"):
            registrar_log_proceso("⚠️ Archivo de entrada no es PDF. Se omite compresión.")
            return

        base, ext = os.path.splitext(input_path)
        output_path = base + "_comprimido.pdf"

        # Comando Ghostscript (pdfwrite) con downsampling y filtros típicos
        cmd = [
            gs_path,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            f"-dPDFSETTINGS=/{calidad}",          # perfil de calidad
            "-dDownsampleColorImages=true",
            f"-dColorImageResolution={dpi}",
            "-dAutoFilterColorImages=false",
            "-dColorImageFilter=/DCTEncode",     # JPEG
            "-dDownsampleGrayImages=true",
            f"-dGrayImageResolution={dpi}",
            "-dAutoFilterGrayImages=false",
            "-dGrayImageFilter=/DCTEncode",      # JPEG
            "-dDownsampleMonoImages=true",
            f"-dMonoImageResolution={dpi}",
            "-dMonoImageFilter=/CCITTFaxEncode", # CCITT para monocromo
            "-dFIXEDMEDIA",                      # fuerza tamaño de página
            "-dPDFFitPage",                      # ajusta contenido al papel
            f"-sPAPERSIZE={tamano_pagina}",
            "-dNOPAUSE", "-dQUIET", "-dBATCH",   # modo batch silencioso
            f"-sOutputFile={output_path}",
            input_path
        ]

        # Ocultar ventana de Ghostscript (Windows)
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = subprocess.SW_HIDE

        subprocess.run(
            cmd,
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            startupinfo=si,
            creationflags=subprocess.CREATE_NO_WINDOW  # también oculta
        )

        # Si GS generó el archivo de salida, intentamos reemplazar el original
        if os.path.exists(output_path):
            try:
                os.remove(input_path)                  # elimina original
                os.rename(output_path, input_path)     # renombra comprimido -> original
                registrar_log_proceso(f"✅ PDF comprimido exitosamente: {os.path.basename(input_path)}")
            except Exception as e:
                # Si no se pudo reemplazar (lock, antivirus, permisos, etc.), guardamos fallback
                registrar_log_proceso(f"❌ Error al reemplazar PDF original tras compresión: {e}")
                fallback_path = input_path.replace(".pdf", "_compresion_fallback.pdf")
                try:
                    # Si ya existe un fallback previo, generar uno único
                    if os.path.exists(fallback_path):
                        i = 1
                        root, ext = os.path.splitext(fallback_path)
                        while os.path.exists(f"{root}_{i}{ext}"):
                            i += 1
                        fallback_path = f"{root}_{i}{ext}"
                    os.rename(output_path, fallback_path)
                    registrar_log_proceso(f"📦 Guardado como fallback: {fallback_path}")
                except Exception as e2:
                    registrar_log_proceso(f"❌ Error al mover fallback: {e2}")
        else:
            registrar_log_proceso(f"⚠️ Compresión fallida: {os.path.basename(input_path)} no fue reemplazado")

    except subprocess.CalledProcessError as e:
        # Ghostscript devolvió código de error (p.ej., PDF corrupto)
        registrar_log_proceso(f"❌ Error al comprimir PDF con Ghostscript: {e}")
    except Exception as e:
        # Cualquier otro error inesperado
        registrar_log_proceso(f"❌ Error inesperado en comprimir_pdf: {e}")


def generar_nombre_unico(base_path, nombre_base):
    """
    Genera un nombre único dentro de base_path a partir de `nombre_base`.
    Retorna algo como: <nombre_base>.pdf, <nombre_base>_1.pdf, <nombre_base>_2.pdf, ...
    (No crea el archivo; solo calcula un nombre que no colisione.)

    Ejemplo:
        generar_nombre_unico("C:/salida", "Factura_ABC") ->
            'Factura_ABC.pdf' (si no existe) o 'Factura_ABC_1.pdf' (si ya existía), etc.
    """
    nombre_final = nombre_base
    contador = 1
    while os.path.exists(os.path.join(base_path, nombre_final + ".pdf")):
        nombre_final = f"{nombre_base}_{contador}"
        contador += 1
    return nombre_final + ".pdf"
