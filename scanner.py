# scanner.py (ADF 1 hoja + cama con selección de item)
# - ADF: SIEMPRE 1 hoja (simplex)
# - Si ADF vacío/ocupado antes de empezar → cae a cama plana (1 hoja)
# - Si ADF marca “ocupado/offline” después de ≥1 página → fin normal
# - 300 DPI color, PDF A4, formatos fallback (BMP → JPG → PNG)

import os
import time
from tkinter import Tk, messagebox
from log_utils import registrar_log_proceso

def _tk_alert(titulo: str, mensaje: str, tipo: str = "warning"):
    try:
        r = Tk(); r.withdraw()
        {"info": messagebox.showinfo, "error": messagebox.showerror}.get(tipo, messagebox.showwarning)(titulo, mensaje)
        r.destroy()
    except Exception:
        pass

# ---------- Palabras clave para identificar items ----------
_ADF_KEYWORDS = {
    "feeder", "document feeder", "automatic document feeder",
    "adf", "alimentador", "alimentador automatico", "alimentador automático",
    "bandeja", "simplex", "duplex", "dúplex"
}
_FLATBED_KEYWORDS = {
    "flatbed", "flat bed", "platen", "cama", "plana", "superficie"
}

# ---------- Helpers de módulo ----------
def _list_items(device):
    """Devuelve lista [(idx, nombre_lower, nombre_original)] y loguea lo encontrado."""
    items = []
    try:
        cnt = device.Items.Count
    except Exception:
        cnt = 0
    for i in range(1, cnt + 1):
        nm = ""
        try:
            it = device.Items[i]
            for p in it.Properties:
                if (p.Name or "").strip().lower() in ("item name", "name"):
                    nm = str(p.Value).strip()
                    break
        except Exception:
            pass
        items.append((i, nm.lower(), nm))
        registrar_log_proceso(f"• Item {i}: {nm or '(sin nombre)'}")
    return items

def _get_item_for_source(device, prefer_feeder: bool):
    """
    Elige el item correcto según la fuente:
      1) Busca por nombre (palabras clave)
      2) Heurística: si prefer_feeder y hay >=2 items → Items[2] (suele ser ADF)
      3) Fallback: Items[1] o Items[0]
    """
    items = _list_items(device)

    adf_candidates = [idx for idx, nm, _ in items if any(k in nm for k in _ADF_KEYWORDS)]
    flat_candidates = [idx for idx, nm, _ in items if any(k in nm for k in _FLATBED_KEYWORDS)]

    try:
        if prefer_feeder and adf_candidates:
            idx = adf_candidates[0]
            registrar_log_proceso(f"→ Item {idx} identificado como ADF por nombre.")
            return device.Items[idx]
        if (not prefer_feeder) and flat_candidates:
            idx = flat_candidates[0]
            registrar_log_proceso(f"→ Item {idx} identificado como FLATBED por nombre.")
            return device.Items[idx]
    except Exception:
        pass

    try:
        cnt = device.Items.Count
    except Exception:
        cnt = 0
    if prefer_feeder and cnt >= 2:
        try:
            registrar_log_proceso("→ Usando heurística: Items[2] como ADF.")
            return device.Items[2]
        except Exception:
            pass

    try:
        registrar_log_proceso("→ Usando Items[1] por fallback.")
        return device.Items[1]
    except Exception:
        registrar_log_proceso("→ Usando Items[0] por fallback.")
        return device.Items[0]

def _set_prop_item(item, name_or_id, value) -> bool:
    """Fija propiedad del item: primero por nombre (recorriendo), luego por ID numérico."""
    try:
        for p in item.Properties:
            if p.Name.strip().lower() == str(name_or_id).strip().lower():
                p.Value = value
                return True
    except Exception:
        pass
    try:
        item.Properties.Item(int(name_or_id)).Value = value
        return True
    except Exception:
        return False
# -----------------------------------------------------------

def escanear_y_guardar_pdf(nombre_archivo_pdf, carpeta_entrada):
    """
    Escanea con WIA y guarda un PDF en `carpeta_entrada`.
    • ADF: 1 hoja (simplex). Si falla antes de empezar → flatbed (1 hoja).
    • ADF ocupado/offline tras ≥1 página → fin normal.
    """
    try:
        import pythoncom, win32com.client, pywintypes
        from typing import Optional
        from PIL import Image
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.utils import ImageReader

        # --------- Constantes / IDs WIA ---------
        FEEDER  = 0x00000001
        FLATBED = 0x00000002
        DUPLEX  = 0x00000004
        WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES = 3086
        WIA_DPS_DOCUMENT_HANDLING_SELECT       = 3088

        # Errores frecuentes (HRESULT → signed)
        WIA_ERROR_PAPER_EMPTY = -2145320957  # 0x80210003 (ADF sin papel)
        WIA_ERROR_OFFLINE     = -2145320954  # 0x80210006 ("dispositivo ocupado")

        # Formatos aceptados por drivers (probamos en orden)
        WIA_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
        WIA_JPG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
        WIA_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
        FORMATOS = [WIA_BMP, WIA_JPG, WIA_PNG]
        # ----------------------------------------

        def _select_source(device, use_feeder: bool, soporta_duplex: bool) -> bool:
            # Forzamos SIMPLEX: si use_feeder True, NO marcamos DUPLEX
            sel = FEEDER if use_feeder else FLATBED
            try:
                for p in device.Properties:
                    if p.Name.strip().lower() == "document handling select":
                        p.Value = sel
                        registrar_log_proceso(f"→ Select {'ADF' if use_feeder else 'FLATBED'} (simplex): OK (by name)")
                        return True
                device.Properties.Item(WIA_DPS_DOCUMENT_HANDLING_SELECT).Value = sel
                registrar_log_proceso(f"→ Select {'ADF' if use_feeder else 'FLATBED'} (simplex): OK (by id)")
                return True
            except Exception as e:
                registrar_log_proceso(f"→ Select {'ADF' if use_feeder else 'FLATBED'}: NO SOPORTADO ({e})")
                return False

        def _configure_item(item) -> None:
            _set_prop_item(item, "Horizontal Resolution", 300) or _set_prop_item(item, 6147, 300)
            _set_prop_item(item, "Vertical Resolution", 300)   or _set_prop_item(item, 6148, 300)
            _set_prop_item(item, "Current Intent", 2)          or _set_prop_item(item, 6146, 2)
            # _set_prop_item(item, 6149, 5100)  # width px (opcional)
            # _set_prop_item(item, 6150, 7020)  # height px (opcional)

        def _transfer(device, carpeta: str, prefer_feeder: bool, max_pages: Optional[int]):
            """Devuelve (rutas, motivo_fin). Usa el ITEM adecuado según la fuente."""
            rutas = []
            item = _get_item_for_source(device, prefer_feeder)
            _configure_item(item)

            idx = 1
            while True:
                if max_pages is not None and idx > max_pages:
                    break
                last_err = None
                for fmt in FORMATOS:
                    try:
                        img = wia_dialog.ShowTransfer(item, fmt)
                        if not img:
                            return rutas, "NO_IMAGE"
                        tmp = os.path.join(carpeta, f"temp_scan_{'adf' if prefer_feeder else 'flat'}_{idx}.bmp")
                        try:
                            if os.path.exists(tmp):
                                os.remove(tmp)
                        except Exception:
                            pass
                        img.SaveFile(tmp)
                        rutas.append(tmp)
                        registrar_log_proceso(f"✔ Página {idx} transferida ({'ADF' if prefer_feeder else 'Flatbed'})")
                        idx += 1
                        last_err = None
                        break
                    except pywintypes.com_error as e:
                        last_err = e
                        hr = int(e.excepinfo[5]) if (e.excepinfo and len(e.excepinfo) >= 6) else 0
                        # ADF: gestionar vacío/ocupado
                        if prefer_feeder and hr in (WIA_ERROR_PAPER_EMPTY, WIA_ERROR_OFFLINE):
                            if rutas:  # ya obtuvimos ≥1 página -> fin normal
                                registrar_log_proceso("ℹ️ Fin del alimentador (sin más hojas).")
                                return rutas, "OK"
                            motivo = "ADF_EMPTY" if hr == WIA_ERROR_PAPER_EMPTY else "ADF_BUSY"
                            registrar_log_proceso(f"ℹ️ {motivo} durante transferencia.")
                            return rutas, motivo
                        continue
                if last_err is not None:
                    if not prefer_feeder:
                        try:
                            registrar_log_proceso("↪ Intentando ShowAcquireImage (UI del driver) en flatbed…")
                            acquired = wia_dialog.ShowAcquireImage(1, 0, 0, None, False, True, False)
                            if not acquired:
                                return rutas, "ACQUIRE_CANCELLED"
                            tmp2 = os.path.join(carpeta, f"temp_scan_flat_{idx}_acq.bmp")
                            try:
                                if os.path.exists(tmp2):
                                    os.remove(tmp2)
                            except Exception:
                                pass
                            acquired.SaveFile(tmp2)
                            rutas.append(tmp2)
                            registrar_log_proceso("✔ Página obtenida vía UI del driver (flatbed)")
                            idx += 1
                            continue
                        except Exception as e2:
                            registrar_log_proceso(f"❗ Falló ShowAcquireImage: {e2}")
                    return rutas, "TRANSFER_ERROR"
            return rutas, "OK"

        # ===== Flujo principal =====
        pythoncom.CoInitialize()
        wia_dialog = win32com.client.Dispatch("WIA.CommonDialog")
        device = wia_dialog.ShowSelectDevice(1, True, False)
        if not device:
            registrar_log_proceso("⚠️ Escaneo cancelado por el usuario.")
            return None

        # Capacidades (informativas)
        try:
            caps = int(device.Properties.Item(WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES).Value)
        except Exception:
            caps = 0
        soporta_feeder = bool(caps & FEEDER)
        # soporta_duplex = bool(caps & DUPLEX)  # No lo usamos: forzamos simplex

        registrar_log_proceso("🖨️ Iniciando escaneo…")

        rutas_imagenes = []

        # ADF (simplex) -> sólo 1 hoja
        _select_source(device, True, False)   # False para no marcar DUPLEX
        time.sleep(0.7)
        _list_items(device)
        rutas, motivo = _transfer(device, carpeta_entrada, True, max_pages=1)
        rutas_imagenes.extend(rutas)

        # Si ADF no entregó nada, cae a flatbed (1 hoja)
        if (motivo in ("ADF_EMPTY", "ADF_BUSY", "NO_IMAGE", "TRANSFER_ERROR")) and not rutas_imagenes:
            _select_source(device, False, False)
            rutas2, _ = _transfer(device, carpeta_entrada, False, max_pages=1)
            rutas_imagenes.extend(rutas2)

        if not rutas_imagenes:
            _tk_alert("Sin páginas", "No se obtuvo ninguna imagen del escaneo.")
            registrar_log_proceso("⚠️ No se obtuvo ninguna imagen del escaneo.")
            return None

        # === Generación de PDF A4 (escalado proporcional) ===
        a4_w, a4_h = A4
        pdf_path = os.path.join(carpeta_entrada, nombre_archivo_pdf)
        c = canvas.Canvas(pdf_path, pagesize=A4)
        for ruta in rutas_imagenes:
            im = Image.open(ruta).convert("RGB")
            iw, ih = im.size
            aspect = iw / float(ih)
            sw = min(a4_w, a4_h * aspect)
            sh = sw / aspect
            xo = (a4_w - sw) / 2
            yo = (a4_h - sh) / 2
            c.drawImage(ImageReader(im), xo, yo, width=sw, height=sh)
            c.showPage()
        c.save()

        # Limpieza temporales
        for p in rutas_imagenes:
            try:
                os.remove(p)
            except Exception:
                pass

        registrar_log_proceso(f"✅ Escaneo guardado como: {os.path.basename(pdf_path)}")
        return pdf_path

    except Exception as e:
        registrar_log_proceso(f"❌ Error en escaneo: {e}")
        _tk_alert(
            "Escáner no detectado",
            (
                "⚠️ No se pudo encontrar un escáner conectado.\n\n"
                "- Asegúrate de que el escáner esté encendido.\n"
                "- Verifica el cable USB o la red.\n"
                "- Revisa/instala los drivers WIA del fabricante."
            ),
            tipo="warning",
        )
        return None
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

def registrar_log(mensaje):
    with open("registro.log", "a", encoding="utf-8") as f:
        f.write(mensaje + "\n")
    print(mensaje)

def diagnostico_wia():
    """Muestra propiedades e items disponibles (para ver cómo nombra el driver al ADF)."""
    import pythoncom, win32com.client
    pythoncom.CoInitialize()
    try:
        wia = win32com.client.Dispatch("WIA.CommonDialog")
        device = wia.ShowSelectDevice(1, True, False)
        if not device:
            print("No se seleccionó dispositivo.")
            return
        print("=== PROPIEDADES DEL DISPOSITIVO ===")
        for p in device.Properties:
            try:
                print(f"ID={p.PropertyID} | Name={p.Name} | Value={p.Value}")
            except Exception:
                pass
        if device.Items and device.Items.Count >= 1:
            print("\n=== ITEMS DISPONIBLES ===")
            cnt = device.Items.Count
            for i in range(1, cnt + 1):
                nm = ""
                it = device.Items[i]
                for p in it.Properties:
                    if (p.Name or "").strip().lower() in ("item name", "name"):
                        nm = str(p.Value)
                        break
                print(f"Item {i}: {nm or '(sin nombre)'}")
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
