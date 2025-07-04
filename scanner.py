import os
from datetime import datetime
from log_utils import registrar_log_proceso
from tkinter import Tk, messagebox  # ✅ Necesario para evitar error en el except

def escanear_y_guardar_pdf(nombre_archivo_pdf, carpeta_entrada):
    try:
        import pythoncom
        import win32com.client
        from PIL import Image
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.utils import ImageReader
        from tkinter import Tk, messagebox

        pythoncom.CoInitialize()
        wia_dialog = win32com.client.Dispatch("WIA.CommonDialog")
        device = wia_dialog.ShowSelectDevice(1, True, False)

        if not device:
            registrar_log_proceso("⚠️ Escáner cancelado por el usuario.")
            return None

        item = device.Items[0]
        for prop in item.Properties:
            if prop.Name == "6147": prop.Value = 600
            elif prop.Name == "6148": prop.Value = 600
            elif prop.Name == "6146": prop.Value = 2
            elif prop.Name == "6149": prop.Value = 5100
            elif prop.Name == "6150": prop.Value = 7020

        image = wia_dialog.ShowTransfer(item, "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}")
        if not image:
            registrar_log_proceso("⚠️ Transferencia de imagen fallida.")
            return None

        temp_png_path = os.path.join(carpeta_entrada, "temp_scan.png")
        if os.path.exists(temp_png_path):
            os.remove(temp_png_path)

        image.SaveFile(temp_png_path)

        imagen = Image.open(temp_png_path).convert("RGB")
        a4_width, a4_height = A4
        img_width, img_height = imagen.size
        aspect_ratio = img_width / img_height

        new_height = a4_height
        new_width = new_height * aspect_ratio
        if new_width > a4_width:
            new_width = a4_width
            new_height = new_width / aspect_ratio

        x_offset = (a4_width - new_width) / 2
        y_offset = (a4_height - new_height) / 2

        pdf_path = os.path.join(carpeta_entrada, nombre_archivo_pdf)
        c = canvas.Canvas(pdf_path, pagesize=A4)
        c.drawImage(ImageReader(imagen), x_offset, y_offset, width=new_width, height=new_height)
        c.showPage()
        c.save()

        os.remove(temp_png_path)

        return pdf_path

    except Exception as e:
        registrar_log_proceso(f"❌ Error en escaneo: {e}")
        root = Tk()
        root.withdraw()
        messagebox.showwarning(
            "Escáner no detectado",
            "⚠️ No se pudo encontrar un escáner conectado.\n\n"
            "- Asegúrate de que el escáner esté encendido.\n"
            "- El cable USB esté conectado o esté en red.\n"
            "- Los drivers estén correctamente instalados."
        )
        root.destroy()
        return None


def obtener_carpeta_salida_anual(carpeta_salida_base):
    año_actual = datetime.now().strftime("%Y")
    carpeta_anual = os.path.join(carpeta_salida_base, año_actual)
    os.makedirs(carpeta_anual, exist_ok=True)
    return carpeta_anual
