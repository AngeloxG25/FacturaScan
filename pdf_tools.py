import subprocess
import os

def comprimir_pdf(gs_path, input_path, calidad='prepress', dpi=600):
    temp_output = input_path.replace(".pdf", "_temp.pdf")
    comando = [
        gs_path, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
        f"-dPDFSETTINGS=/{calidad}", "-dFIXEDMEDIA", "-dPDFFitPage",
        "-sPAPERSIZE=a4", "-dNOPAUSE", "-dQUIET", "-dBATCH",
        f"-r{dpi}", f"-sOutputFile={temp_output}", input_path
    ]
    subprocess.run(comando, check=True)
    os.replace(temp_output, input_path)

def generar_nombre_unico(base_path, nombre_base):
    nombre_final = nombre_base
    contador = 1
    while os.path.exists(os.path.join(base_path, nombre_final + ".pdf")):
        nombre_final = f"{nombre_base}_{contador}"
        contador += 1
    return nombre_final + ".pdf"
