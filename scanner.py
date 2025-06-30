import os
from datetime import datetime

def escanear_y_guardar_pdf(nombre_archivo_pdf, carpeta_entrada, preprocesado):
    try:
        import pythoncom
        import win32com.client
        from PIL import Image
        import shutil

        # Iniciar WIA
        pythoncom.CoInitialize()
        wia_dialog = win32com.client.Dispatch("WIA.CommonDialog")
        device = wia_dialog.ShowSelectDevice(1, True, False)
        if not device:
            print("⚠️ Escáner cancelado por el usuario.")
            return None

        item = device.Items[0]
        for prop in item.Properties:
            if prop.Name == "6147": prop.Value = 300
            elif prop.Name == "6148": prop.Value = 300
            elif prop.Name == "6146": prop.Value = 2
            elif prop.Name == "6149": prop.Value = 3508
            elif prop.Name == "6150": prop.Value = 2480

        image = wia_dialog.ShowTransfer(item, "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}")
        if not image:
            return None

        # Guardar temporalmente la imagen
        temp_png_path = os.path.join(carpeta_entrada, "temp_scan.png")
        if os.path.exists(temp_png_path):
            os.remove(temp_png_path)
        image.SaveFile(temp_png_path)

        # Convertir a PDF
        imagen = Image.open(temp_png_path).convert("RGB")
        pdf_path = os.path.join(carpeta_entrada, nombre_archivo_pdf)
        imagen.save(pdf_path, "PDF", resolution=300.0)

        # Copiar PNG preprocesado
        shutil.copy(temp_png_path, os.path.join(preprocesado, nombre_archivo_pdf.replace('.pdf', '.png')))
        os.remove(temp_png_path)

        return pdf_path

    except Exception as e:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showwarning(
            "Escáner no detectado",
            "⚠️ No se pudo encontrar un escáner conectado.\n\n"
            "Por favor, asegúrate de que:\n"
            "- El escáner esté encendido.\n"
            "- El cable USB esté correctamente conectado o el dispositivo esté en la misma red.\n"
            "- Los drivers estén correctamente instalados.\n\n"
            "Luego, vuelve a intentar escanear."
        )
        root.destroy()
        return None

def obtener_carpeta_salida_anual(carpeta_salida_base):
    """Devuelve la ruta de la carpeta de salida del año actual, creándola si no existe."""
    año_actual = datetime.now().strftime("%Y")
    carpeta_anual = os.path.join(carpeta_salida_base, año_actual)
    os.makedirs(carpeta_anual, exist_ok=True)
    return carpeta_anual
