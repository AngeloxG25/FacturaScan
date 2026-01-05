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
import utils.hide as hide_subprocess  # Aplica monkey patch al importar
import subprocess
import os

from utils.log_utils import registrar_log_proceso

def comprimir_pdf(gs_path, input_path, calidad="screen", dpi=100, tamano_pagina="a4"):
    """
    Comprime y normaliza un PDF usando Ghostscript.
    """
    try:
        # Validaciones defensivas
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

        cmd = [
            gs_path,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            f"-dPDFSETTINGS=/{calidad}",
            "-dDownsampleColorImages=true",
            f"-dColorImageResolution={dpi}",
            "-dAutoFilterColorImages=false",
            "-dColorImageFilter=/DCTEncode",
            "-dDownsampleGrayImages=true",
            f"-dGrayImageResolution={dpi}",
            "-dAutoFilterGrayImages=false",
            "-dGrayImageFilter=/DCTEncode",
            "-dDownsampleMonoImages=true",
            f"-dMonoImageResolution={dpi}",
            "-dMonoImageFilter=/CCITTFaxEncode",
            "-dFIXEDMEDIA",
            "-dPDFFitPage",
            f"-sPAPERSIZE={tamano_pagina}",
            "-dNOPAUSE", "-dQUIET", "-dBATCH",
            f"-sOutputFile={output_path}",
            input_path,
        ]

        # Gracias al monkey patch de utils.hide, esto ya se ejecuta oculto en Windows
        subprocess.run(
            cmd,
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

        if os.path.exists(output_path):
            try:
                os.remove(input_path)
                os.rename(output_path, input_path)
                registrar_log_proceso(f"✅ PDF comprimido exitosamente: {os.path.basename(input_path)}")
            except Exception as e:
                registrar_log_proceso(f"❌ Error al reemplazar PDF original tras compresión: {e}")

                # Fallback único
                fallback_base = input_path.replace(".pdf", "_compresion_fallback.pdf")
                fallback_path = fallback_base
                if os.path.exists(fallback_path):
                    i = 1
                    root, ext = os.path.splitext(fallback_base)
                    while os.path.exists(f"{root}_{i}{ext}"):
                        i += 1
                    fallback_path = f"{root}_{i}{ext}"

                try:
                    os.rename(output_path, fallback_path)
                    registrar_log_proceso(f"📦 Guardado como fallback: {fallback_path}")
                except Exception as e2:
                    registrar_log_proceso(f"❌ Error al mover fallback: {e2}")
        else:
            registrar_log_proceso(
                f"⚠️ Compresión fallida: {os.path.basename(input_path)} no fue reemplazado (no se generó salida)."
            )

    except subprocess.CalledProcessError as e:
        registrar_log_proceso(f"❌ Error al comprimir PDF con Ghostscript: {e}")
    except Exception as e:
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