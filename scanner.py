import os
from log_utils import registrar_log_proceso
from tkinter import Tk, messagebox

def _tk_alert(titulo: str, mensaje: str, tipo="warning"):
    """
    Pequeño helper para mostrar messageboxes de Tk de forma segura.
    - Crea un root temporal oculto (withdraw) para no abrir una ventana principal.
    - Soporta info / warning / error.
    - Tolerante a fallos (try/except amplio) para no romper hilos de trabajo.
    """
    try:
        root = Tk()
        root.withdraw()
        if tipo == "info":
            messagebox.showinfo(titulo, mensaje)
        elif tipo == "error":
            messagebox.showerror(titulo, mensaje)
        else:
            messagebox.showwarning(titulo, mensaje)
        root.destroy()
    except:
        # Evita que un error de GUI tumbe el proceso (p. ej., no hay loop Tk activo)
        pass


def escanear_y_guardar_pdf(nombre_archivo_pdf, carpeta_entrada):
    """
    Escanea usando WIA y guarda un PDF multipágina en `carpeta_entrada`.
    - Intenta usar ADF y dúplex si el driver WIA los expone.
    - Si no hay dispositivo WIA o el usuario cancela, muestra un Tk messagebox y retorna None.
    - Si el driver no publica ADF, cae automáticamente a flatbed (una página).

    Parámetros:
      nombre_archivo_pdf: nombre del PDF final (ej. 'DocEscaneado_20250101_010203.pdf')
      carpeta_entrada   : carpeta donde se guardará el PDF

    Retorna:
      Ruta absoluta del PDF generado, o None si algo impide escanear.
    """
    import os
    from PIL import Image
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.utils import ImageReader
    import pythoncom, win32com.client, pywintypes  # COM y WIA

    # -------- Helpers para setear/leer propiedades WIA de forma robusta --------
    def set_prop_by_name(props, name, value):
        """Setea una propiedad WIA buscándola por nombre (más portable entre drivers)."""
        for p in props:
            if p.Name.strip().lower() == name.strip().lower():
                p.Value = value
                return True
        return False

    def get_prop_value_by_name(props, name, default=None):
        """Lee una propiedad WIA por nombre y la castea a int si corresponde."""
        for p in props:
            if p.Name.strip().lower() == name.strip().lower():
                try:
                    return int(p.Value)
                except Exception:
                    return p.Value
        return default

    def set_prop_by_id(props, pid, value):
        """Setea una propiedad WIA por ID (fallback cuando el nombre no existe)."""
        try:
            props.Item(pid).Value = value
            return True
        except Exception:
            return False
    # ---------------------------------------------------------------------------

    # Inicializa COM en el hilo actual (requerido por WIA)
    pythoncom.CoInitialize()
    try:
        wia = win32com.client.Dispatch("WIA.CommonDialog")

        try:
            # Abre el diálogo de selección de dispositivo (1 = ScannerDeviceType)
            # Aquí lanza com_error si no hay ningún dispositivo WIA disponible.
            device = wia.ShowSelectDevice(1, True, False)
        except pywintypes.com_error as e:
            # HRESULT típico: -2145320939 → "No hay disponible ningún dispositivo WIA del tipo seleccionado."
            _tk_alert(
                "Escáner no encontrado",
                "No se detectó ningún dispositivo WIA (escáner) disponible.\n\n"
                "• Verifica la conexión USB o red del escáner.\n"
                "• Instala/actualiza el driver WIA del fabricante.\n"
                "• Si tu equipo solo expone TWAIN, usa el software del fabricante o TWAIN→PDF y vuelve a intentar.",
                tipo="warning"
            )
            registrar_log_proceso("⚠️ No hay dispositivo WIA disponible.")
            return None

        if not device:
            # El usuario abrió el selector pero canceló sin elegir escáner
            _tk_alert("Operación cancelada", "No se seleccionó ningún escáner.", tipo="info")
            registrar_log_proceso("⚠️ Escaneo cancelado por el usuario.")
            return None

        # ----------------- Flags y PropertyIDs WIA más comunes -----------------
        FEEDER  = 0x00000001  # Soporte de alimentador (ADF)
        FLATBED = 0x00000002  # Cama plana
        DUPLEX  = 0x00000004  # Soporte de dúplex

        # Estos IDs son estándar en WIA, pero algunos drivers no los exponen
        WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES = 3086
        WIA_DPS_DOCUMENT_HANDLING_STATUS       = 3087
        WIA_DPS_DOCUMENT_HANDLING_SELECT       = 3088
        # ----------------------------------------------------------------------

        # Lee capacidades del dispositivo: primero por nombre, luego por ID
        caps = get_prop_value_by_name(device.Properties, "Document Handling Capabilities", None)
        if caps is None:
            try:
                caps = int(device.Properties.Item(WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES).Value)
            except Exception:
                caps = 0  # si el driver no lo expone, asumimos "sin ADF"

        use_feeder = bool(caps & FEEDER)
        use_duplex = bool(caps & DUPLEX)

        # Intenta seleccionar ADF (y dúplex si existe); si no, fuerza flatbed
        selected = False
        if use_feeder:
            sel_val = FEEDER | (DUPLEX if use_duplex else 0)
            if set_prop_by_name(device.Properties, "Document Handling Select", sel_val):
                selected = True
            elif set_prop_by_id(device.Properties, WIA_DPS_DOCUMENT_HANDLING_SELECT, sel_val):
                selected = True

        if not selected:
            # Si no pudimos seleccionar ADF, forzamos flatbed explícitamente
            if not set_prop_by_name(device.Properties, "Document Handling Select", FLATBED):
                set_prop_by_id(device.Properties, WIA_DPS_DOCUMENT_HANDLING_SELECT, FLATBED)
            use_feeder = False
            use_duplex = False

        # Obtiene el primer item (WIA usa índices 1-based)
        item = device.Items[1]

        def try_set_item(pid_or_name, value, by_name=True):
            """
            Setea propiedades del 'item' de escaneo (resolución, color, etc.).
            - Primero intenta por nombre; si falla, por ID.
            """
            try:
                if by_name:
                    for p in item.Properties:
                        if p.Name.strip().lower() == pid_or_name.strip().lower():
                            p.Value = value
                            return True
                else:
                    item.Properties.Item(pid_or_name).Value = value
                    return True
            except Exception:
                pass
            return False

        # Ajustes básicos de captura:
        # - Resolución 300 DPI (6147/6148 suelen ser X/Y DPI)
        # - Current Intent = 2 (color). Algunos drivers usan 6146.
        if not try_set_item("Horizontal Resolution", 300):
            try_set_item(6147, 300, by_name=False)
        if not try_set_item("Vertical Resolution", 300):
            try_set_item(6148, 300, by_name=False)
        if not try_set_item("Current Intent", 2):
            try_set_item(6146, 2, by_name=False)

        print(f"🖨️ Iniciando escaneo (ADF detectado: {use_feeder}, dúplex: {use_duplex})")

        # ----------------- Transferencia de páginas desde WIA ------------------
        images = []
        idx = 1
        while True:
            try:
                # ShowTransfer con formato PNG (GUID estándar para PNG)
                img = wia.ShowTransfer(item, "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}")
                if not img:
                    break

                # Guardamos cada página temporalmente como PNG en la carpeta de entrada
                tmp = os.path.join(carpeta_entrada, f"temp_scan_{idx}.png")
                if os.path.exists(tmp):
                    os.remove(tmp)  # evita “file in use” si quedó de un intento anterior
                img.SaveFile(tmp)
                images.append(tmp)
                idx += 1

                # Si NO hay ADF (flatbed), solo habrá una página
                if not use_feeder:
                    break

            except Exception:
                # En ADF, aquí suele terminar cuando se vacía el alimentador o el driver devuelve error benigno
                break
        # ----------------------------------------------------------------------

        # Si no se obtuvo al menos una imagen, avisamos al usuario
        if not images:
            _tk_alert("Sin páginas", "No se obtuvo ninguna imagen del escaneo.", tipo="warning")
            registrar_log_proceso("⚠️ No se obtuvo ninguna imagen del escaneo.")
            return None

        # ----------------- Construcción del PDF multipágina --------------------
        pdf_path = os.path.join(carpeta_entrada, nombre_archivo_pdf)
        a4_w, a4_h = A4
        c = canvas.Canvas(pdf_path, pagesize=A4)

        for path in images:
            im = Image.open(path).convert("RGB")
            iw, ih = im.size
            aspect = iw / float(ih)

            # Escalado para encajar la imagen completa dentro de A4 sin deformarla
            sw = min(a4_w, a4_h * aspect)
            sh = sw / aspect
            xo = (a4_w - sw) / 2
            yo = (a4_h - sh) / 2

            c.drawImage(ImageReader(im), xo, yo, width=sw, height=sh)
            c.showPage()

        c.save()
        # ----------------------------------------------------------------------

        # Limpieza de temporales PNG (best-effort)
        for p in images:
            try:
                os.remove(p)
            except:
                pass

        print(f"✅ Guardado: {os.path.basename(pdf_path)}")
        return pdf_path

    finally:
        # Libera COM del hilo actual
        try:
            pythoncom.CoUninitialize()
        except:
            pass


def registrar_log(mensaje):
    """
    Helper mínimo de logging a archivo + stdout.
    (Aquí se mantiene por compatibilidad con otros módulos/tests).
    """
    with open("registro.log", "a", encoding="utf-8") as f:
        f.write(mensaje + "\n")
    print(mensaje) 


def diagnostico_wia():
    """
    Utilidad interactiva: lista todas las propiedades WIA del dispositivo seleccionado.
    Útil para:
      - Ver si el driver expone ADF/Dúplex y con qué nombres/IDs.
      - Depurar por qué no se puede activar ADF en cierto modelo.
    """
    import pythoncom, os
    import win32com.client

    pythoncom.CoInitialize()
    wia = win32com.client.Dispatch("WIA.CommonDialog")
    device = wia.ShowSelectDevice(1, True, False)  # 1=Scanner

    if not device:
        print("No se seleccionó dispositivo.")
        return

    print("=== PROPIEDADES DEL DISPOSITIVO ===")
    for p in device.Properties:
        try:
            print(f"ID={p.PropertyID} | Name={p.Name} | Value={p.Value}")
        except Exception:
            pass

    # Algunos drivers publican propiedades relevantes a nivel de item
    if device.Items and device.Items.Count >= 1:
        item = device.Items[1]
        print("\n=== PROPIEDADES DEL ITEM ===")
        for p in item.Properties:
            try:
                print(f"ID={p.PropertyID} | Name={p.Name} | Value={p.Value}")
            except Exception:
                pass

    print("\nTip: busca líneas que contengan 'Document Handling' / 'Feeder' / 'Duplex' o similares.")
