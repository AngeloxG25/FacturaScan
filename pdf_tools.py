# Parchea subprocess para ocultar CMDs en Windows
import hide_subprocess
import subprocess
import os
import ctypes
from ocr_utils import registrar_log_proceso

def comprimir_pdf(gs_path, input_path, calidad="screen", dpi=100, tamano_pagina="a4"):
    try:
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
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            f"-sOutputFile={output_path}",
            input_path
        ]

        # Configurar subprocess sin mostrar ventana
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = subprocess.SW_HIDE  # Ocultar ventana

        subprocess.run(
            cmd,
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            startupinfo=si,
            creationflags=subprocess.CREATE_NO_WINDOW
        )

        if os.path.exists(output_path):
            try:
                os.remove(input_path)
                os.rename(output_path, input_path)
                registrar_log_proceso(f"✅ PDF comprimido exitosamente: {os.path.basename(input_path)}")
            except Exception as e:
                registrar_log_proceso(f"❌ Error al reemplazar PDF original tras compresión: {e}")
                # Recupera archivo original desde output_path con nombre alternativo
                fallback_path = input_path.replace(".pdf", "_compresion_fallback.pdf")
                os.rename(output_path, fallback_path)
                registrar_log_proceso(f"📦 Guardado como fallback: {fallback_path}")
        else:
            registrar_log_proceso(f"⚠️ Compresión fallida: {os.path.basename(input_path)} no fue reemplazado")

    except subprocess.CalledProcessError as e:
        registrar_log_proceso(f"❌ Error al comprimir PDF con Ghostscript: {e}")
    except Exception as e:
        registrar_log_proceso(f"❌ Error inesperado en comprimir_pdf: {e}")

def generar_nombre_unico(base_path, nombre_base):
    """
    Genera un nombre de archivo único en base a `nombre_base`, evitando sobreescribir.
    """
    nombre_final = nombre_base
    contador = 1
    while os.path.exists(os.path.join(base_path, nombre_final + ".pdf")):
        nombre_final = f"{nombre_base}_{contador}"
        contador += 1
    return nombre_final + ".pdf"
