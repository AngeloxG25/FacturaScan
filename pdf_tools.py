import subprocess
import os
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

        subprocess.run(
            cmd,
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW  #  ESTA LÍNEA OCULTA CMD SECUNDARIA
        )

        os.remove(input_path)
        os.rename(output_path, input_path)
    
    except subprocess.CalledProcessError as e:
        registrar_log_proceso(f"⚠️ Error al comprimir PDF en pdf_tools.py: {e}")


def generar_nombre_unico(base_path, nombre_base):
    nombre_final = nombre_base
    contador = 1
    while os.path.exists(os.path.join(base_path, nombre_final + ".pdf")):
        nombre_final = f"{nombre_base}_{contador}"
        contador += 1
    return nombre_final + ".pdf"
