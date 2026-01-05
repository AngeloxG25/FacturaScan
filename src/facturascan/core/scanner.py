# Versión funcional (con escáner WIA predeterminado)
import os, sys
import time
import json
from tkinter import Tk, messagebox
from utils.log_utils import registrar_log_proceso

# ====== Persistencia del escáner predeterminado ======

# ====== Persistencia del escáner predeterminado ======

def _get_app_dir() -> str:
    """
    Devuelve la carpeta base de la app:
    - Compilado (Nuitka/PyInstaller): carpeta del .exe
    - Desarrollo: carpeta del archivo .py
    """
    if getattr(sys, "frozen", False) or sys.executable.lower().endswith(".exe"):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def _is_dir_writable(dirpath: str) -> bool:
    """Verifica si podemos escribir en esa carpeta creando un archivo temporal."""
    try:
        os.makedirs(dirpath, exist_ok=True)
        test_path = os.path.join(dirpath, ".__write_test.tmp")
        with open(test_path, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(test_path)
        return True
    except Exception:
        return False

def _get_default_scanner_path() -> str:
    """
    1) Intenta al lado del exe.
    2) Si no se puede escribir ahí (Program Files, permisos, etc.), usa LOCALAPPDATA\\FacturaScan.
    """
    exe_dir = _get_app_dir()
    if _is_dir_writable(exe_dir):
        return os.path.join(exe_dir, "scanner_default.json")

    base_appdata = os.environ.get("LOCALAPPDATA", r"C:\FacturaScan")
    fallback_dir = os.path.join(base_appdata, "FacturaScan")
    try:
        os.makedirs(fallback_dir, exist_ok=True)
    except Exception:
        pass
    return os.path.join(fallback_dir, "scanner_default.json")

DEFAULT_SCANNER_PATH = _get_default_scanner_path()

def _load_default_scanner():
    try:
        with open(DEFAULT_SCANNER_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict) and data.get("device_id"):
            return data
    except Exception:
        pass
    return {}

def _save_default_scanner(device_id: str, name: str = ""):
    try:
        payload = {"device_id": str(device_id), "name": str(name or "")}
        tmp = DEFAULT_SCANNER_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        os.replace(tmp, DEFAULT_SCANNER_PATH)
    except Exception as e:
        registrar_log_proceso(f"⚠️ No se pudo guardar escáner predeterminado en '{DEFAULT_SCANNER_PATH}': {e}")

def seleccionar_scanner_predeterminado():
    """
    Abre el selector de dispositivos WIA y guarda el escáner predeterminado (DeviceID) en JSON.
    Devuelve dict: {"device_id": "...", "name": "..."} o {} si se cancela/falla.
    """
    pythoncom = None
    try:
        import pythoncom, win32com.client

        pythoncom.CoInitialize()
        wia_dialog = win32com.client.Dispatch("WIA.CommonDialog")

        device = wia_dialog.ShowSelectDevice(1, True, False)  # 1 = Scanner
        if not device:
            registrar_log_proceso("ℹ️ Cambio de escáner cancelado por el usuario.")
            return {}

        dev_id = _get_device_id(device)
        dev_name = _get_device_name(device)

        if dev_id:
            _save_default_scanner(dev_id, dev_name)
            registrar_log_proceso(f"💾 Escáner predeterminado actualizado: {dev_name or dev_id}")
            return {"device_id": dev_id, "name": dev_name or ""}

        registrar_log_proceso("⚠️ No se pudo obtener DeviceID del escáner seleccionado.")
        _tk_alert("Escáner", "Se seleccionó un escáner, pero no se pudo obtener su identificador (DeviceID).")
        return {}

    except Exception as e:
        registrar_log_proceso(f"❌ Error al cambiar escáner predeterminado: {e}")
        _tk_alert("Escáner", "No se pudo cambiar el escáner predeterminado. Revisa drivers WIA del fabricante.")
        return {}

    finally:
        try:
            if pythoncom is not None:
                pythoncom.CoUninitialize()
        except Exception:
            pass


def limpiar_scanner_predeterminado():
    """Elimina el predeterminado guardado (forzará selector al próximo escaneo)."""
    try:
        _clear_default_scanner()
        registrar_log_proceso("🧹 Escáner predeterminado eliminado.")
        return True
    except Exception:
        return False


def _clear_default_scanner():
    try:
        if os.path.exists(DEFAULT_SCANNER_PATH):
            os.remove(DEFAULT_SCANNER_PATH)
    except Exception:
        pass

def _get_device_id(device):
    # 1) Propiedad directa (a veces existe)
    try:
        did = getattr(device, "DeviceID", None)
        if did:
            return str(did)
    except Exception:
        pass

    # 2) Buscar por propiedades
    try:
        for p in device.Properties:
            nm = (p.Name or "").strip().lower()
            if nm in ("device id", "unique device id", "deviceid"):
                val = p.Value
                if val:
                    return str(val)
    except Exception:
        pass

    return ""


def _get_device_name(device):
    try:
        for p in device.Properties:
            nm = (p.Name or "").strip().lower()
            if nm in ("name", "device name"):
                val = p.Value
                if val:
                    return str(val)
    except Exception:
        pass

    try:
        return str(getattr(device, "Name", "")) or ""
    except Exception:
        return ""


def _connect_device_by_id(win32com_client, device_id: str):
    """
    Conecta a un WIA Device por DeviceID usando WIA.DeviceManager.
    Devuelve device o None.
    """
    try:
        dm = win32com_client.Dispatch("WIA.DeviceManager")
        for info in dm.DeviceInfos:
            try:
                if str(info.DeviceID) == str(device_id):
                    return info.Connect()
            except Exception:
                continue
    except Exception:
        pass
    return None


# ====== Alertas Tk ======
def _tk_alert(titulo: str, mensaje: str, tipo: str = "warning"):
    try:
        r = Tk()
        r.withdraw()
        {"info": messagebox.showinfo, "error": messagebox.showerror}.get(
            tipo, messagebox.showwarning
        )(titulo, mensaje)
        r.destroy()
    except Exception:
        pass


# -----------------------------------------------------------
def escanear_y_guardar_pdf(nombre_archivo_pdf, carpeta_entrada):
    """
    Escanea con WIA y guarda un PDF en `carpeta_entrada`.

    • ADF: 1 hoja (simplex). Si falla antes de empezar → flatbed (1 hoja).
    • ADF ocupado/offline tras ≥1 página → fin normal.
    • Optimizado: transferencia en memoria (sin archivos temporales de imagen).
    • NUEVO: Escáner predeterminado (DeviceID) persistido en JSON.
    """
    try:
        import pythoncom, win32com.client, pywintypes
        from io import BytesIO
        from typing import Optional
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.utils import ImageReader

        # --------- Constantes / IDs WIA ---------
        FEEDER  = 0x00000001
        FLATBED = 0x00000002
        WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES = 3086
        WIA_DPS_DOCUMENT_HANDLING_SELECT       = 3088

        # Errores frecuentes (HRESULT → signed)
        WIA_ERROR_PAPER_EMPTY = -2145320957  # 0x80210003 (ADF sin papel)
        WIA_ERROR_OFFLINE     = -2145320954  # 0x80210006 ("dispositivo ocupado")

        # Formatos aceptados por drivers (intentamos de más liviano a más pesado)
        WIA_JPG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
        WIA_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
        WIA_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
        FORMATOS = [WIA_JPG, WIA_PNG, WIA_BMP]
        # ----------------------------------------

        def _select_source(device, use_feeder: bool) -> bool:
            # Forzamos SIMPLEX: si use_feeder True, NO marcamos DUPLEX
            sel = FEEDER if use_feeder else FLATBED
            try:
                for p in device.Properties:
                    if (p.Name or "").strip().lower() == "document handling select":
                        p.Value = sel
                        registrar_log_proceso(
                            f"→ Select {'ADF' if use_feeder else 'FLATBED'} (simplex): OK (by name)"
                        )
                        return True
                device.Properties.Item(WIA_DPS_DOCUMENT_HANDLING_SELECT).Value = sel
                registrar_log_proceso(
                    f"→ Select {'ADF' if use_feeder else 'FLATBED'} (simplex): OK (by id)"
                )
                return True
            except Exception as e:
                registrar_log_proceso(
                    f"→ Select {'ADF' if use_feeder else 'FLATBED'}: NO SOPORTADO ({e})"
                )
                return False

        def _configure_item(item) -> None:
            # 300 DPI color
            def _set_prop_item(_item, name_or_id, value) -> bool:
                try:
                    for _p in _item.Properties:
                        if (_p.Name or "").strip().lower() == str(name_or_id).strip().lower():
                            _p.Value = value
                            return True
                except Exception:
                    pass
                try:
                    _item.Properties.Item(int(name_or_id)).Value = value
                    return True
                except Exception:
                    return False

            _set_prop_item(item, "Horizontal Resolution", 300) or _set_prop_item(item, 6147, 300)
            _set_prop_item(item, "Vertical Resolution", 300)   or _set_prop_item(item, 6148, 300)
            _set_prop_item(item, "Current Intent", 2)          or _set_prop_item(item, 6146, 2)  # 2=color

        def _get_item_for_source(device, prefer_feeder: bool):
            # Reusa tu heurística de selección
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

            ADF_KEYS = {
                "feeder","document feeder","automatic document feeder","adf",
                "alimentador","alimentador automatico","alimentador automático",
                "bandeja","simplex","duplex","dúplex"
            }
            FLAT_KEYS = {"flatbed","flat bed","platen","cama","plana","superficie"}

            adf_cand  = [idx for idx, nm, _ in items if any(k in nm for k in ADF_KEYS)]
            flat_cand = [idx for idx, nm, _ in items if any(k in nm for k in FLAT_KEYS)]

            try:
                if prefer_feeder and adf_cand:
                    registrar_log_proceso(f"→ Item {adf_cand[0]} identificado como ADF por nombre.")
                    return device.Items[adf_cand[0]]
                if (not prefer_feeder) and flat_cand:
                    registrar_log_proceso(f"→ Item {flat_cand[0]} identificado como FLATBED por nombre.")
                    return device.Items[flat_cand[0]]
            except Exception:
                pass

            try:
                if prefer_feeder and device.Items.Count >= 2:
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

        def _transfer_mem(device, prefer_feeder: bool, max_pages: Optional[int]):
            """
            Devuelve (list[(bytes, ext)], motivo_fin)
            Guarda cada página en memoria como (bytes, 'jpg'/'png'/'bmp').
            """
            paginas = []
            item = _get_item_for_source(device, prefer_feeder)
            _configure_item(item)

            idx = 1
            while True:
                if max_pages is not None and idx > max_pages:
                    break
                last_err = None

                for fmt in FORMATOS:
                    try:
                        img = wia_dialog.ShowTransfer(item, fmt)  # WIA.ImageFile
                        if not img:
                            return paginas, "NO_IMAGE"

                        # Intento 1: bytes directos
                        ext = "jpg" if fmt == WIA_JPG else ("png" if fmt == WIA_PNG else "bmp")
                        try:
                            data = img.FileData.BinaryData  # bytes
                            if not data:
                                raise ValueError("BinaryData vacío")
                            paginas.append((bytes(data), ext))
                            registrar_log_proceso(
                                f"✔ Página {idx} en memoria ({'ADF' if prefer_feeder else 'Flatbed'} | {ext.upper()})"
                            )
                            idx += 1
                            last_err = None
                            break
                        except Exception:
                            # Fallback: a disco temporal y leer a memoria
                            tmp = os.path.join(
                                carpeta_entrada,
                                f"temp_scan_{'adf' if prefer_feeder else 'flat'}_{idx}.{ext}"
                            )
                            try:
                                if os.path.exists(tmp):
                                    os.remove(tmp)
                            except Exception:
                                pass

                            img.SaveFile(tmp)
                            with open(tmp, "rb") as fh:
                                paginas.append((fh.read(), ext))
                            try:
                                os.remove(tmp)
                            except Exception:
                                pass

                            registrar_log_proceso(f"✔ Página {idx} (fallback disco) ({ext.upper()})")
                            idx += 1
                            last_err = None
                            break

                    except pywintypes.com_error as e:
                        last_err = e
                        hr = int(e.excepinfo[5]) if (e.excepinfo and len(e.excepinfo) >= 6) else 0

                        # ADF: gestionar vacío/ocupado
                        if prefer_feeder and hr in (WIA_ERROR_PAPER_EMPTY, WIA_ERROR_OFFLINE):
                            if paginas:  # ya obtuvimos ≥1 página -> fin normal
                                registrar_log_proceso("ℹ️ Fin del alimentador (sin más hojas).")
                                return paginas, "OK"
                            motivo = "ADF_EMPTY" if hr == WIA_ERROR_PAPER_EMPTY else "ADF_BUSY"
                            registrar_log_proceso(f"ℹ️ {motivo} durante transferencia.")
                            return paginas, motivo
                        continue

                if last_err is not None:
                    if not prefer_feeder:
                        # último intento usando UI del driver
                        try:
                            registrar_log_proceso("↪ Intentando ShowAcquireImage (UI del driver) en flatbed…")
                            acquired = wia_dialog.ShowAcquireImage(1, 0, 0, None, False, True, False)
                            if not acquired:
                                return paginas, "ACQUIRE_CANCELLED"
                            try:
                                data = acquired.FileData.BinaryData
                                paginas.append((bytes(data), "bmp"))  # suele venir BMP
                            except Exception:
                                # si no trae BinaryData, a disco temporal
                                tmp2 = os.path.join(carpeta_entrada, f"temp_scan_flat_{idx}_acq.bmp")
                                acquired.SaveFile(tmp2)
                                with open(tmp2, "rb") as fh:
                                    paginas.append((fh.read(), "bmp"))
                                try:
                                    os.remove(tmp2)
                                except Exception:
                                    pass
                            registrar_log_proceso("✔ Página obtenida vía UI del driver (flatbed)")
                            idx += 1
                            continue
                        except Exception as e2:
                            registrar_log_proceso(f"❗ Falló ShowAcquireImage: {e2}")
                    return paginas, "TRANSFER_ERROR"

            return paginas, "OK"

        # ===== Flujo principal =====
        pythoncom.CoInitialize()
        wia_dialog = win32com.client.Dispatch("WIA.CommonDialog")

        # 1) Intentar usar predeterminado sin mostrar selección
        device = None
        default = _load_default_scanner()

        if default.get("device_id"):
            device = _connect_device_by_id(win32com.client, default["device_id"])
            if device:
                registrar_log_proceso(
                    f"🖨️ Usando escáner predeterminado: {default.get('name') or default['device_id']}"
                )
            else:
                registrar_log_proceso("⚠️ Escáner predeterminado no disponible. Se solicitará selección…")

        # 2) Si no hay predeterminado o falló -> mostrar lista
        if not device:
            device = wia_dialog.ShowSelectDevice(1, True, False)
            if not device:
                registrar_log_proceso("⚠️ Escaneo cancelado por el usuario.")
                return None

            dev_id = _get_device_id(device)
            dev_name = _get_device_name(device)
            if dev_id:
                _save_default_scanner(dev_id, dev_name)
                registrar_log_proceso(f"💾 Escáner predeterminado guardado: {dev_name or dev_id}")

        # Capacidades (informativas)
        try:
            caps = int(device.Properties.Item(WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES).Value)
        except Exception:
            caps = 0
        soporta_feeder = bool(caps & FEEDER)  # (si lo quieres usar luego)

        registrar_log_proceso("🖨️ Iniciando escaneo…")

        # ADF (simplex) -> sólo 1 hoja
        _select_source(device, True)
        time.sleep(0.5)  # breve settling; algunos drivers lo necesitan
        rutas_mem, motivo = _transfer_mem(device, True, max_pages=1)

        # Si ADF no entregó nada, cae a flatbed (1 hoja)
        if (motivo in ("ADF_EMPTY", "ADF_BUSY", "NO_IMAGE", "TRANSFER_ERROR")) and not rutas_mem:
            _select_source(device, False)
            rutas2, _ = _transfer_mem(device, False, max_pages=1)
            rutas_mem.extend(rutas2)

        # 3) Si aun así no hay páginas: reintentar 1 vez pidiendo selección nuevamente
        if not rutas_mem:
            registrar_log_proceso("⚠️ No se obtuvo ninguna página. Reintentando con selección de dispositivo…")
            _clear_default_scanner()

            device2 = wia_dialog.ShowSelectDevice(1, True, False)
            if device2:
                dev_id2 = _get_device_id(device2)
                dev_name2 = _get_device_name(device2)
                if dev_id2:
                    _save_default_scanner(dev_id2, dev_name2)
                    registrar_log_proceso(f"💾 Nuevo predeterminado: {dev_name2 or dev_id2}")

                _select_source(device2, True)
                time.sleep(0.5)
                rutas_mem, motivo = _transfer_mem(device2, True, max_pages=1)

                if (motivo in ("ADF_EMPTY", "ADF_BUSY", "NO_IMAGE", "TRANSFER_ERROR")) and not rutas_mem:
                    _select_source(device2, False)
                    rutas2, _ = _transfer_mem(device2, False, max_pages=1)
                    rutas_mem.extend(rutas2)

        if not rutas_mem:
            _tk_alert("Sin páginas", "No se obtuvo ninguna imagen del escaneo.")
            registrar_log_proceso("⚠️ No se obtuvo ninguna imagen del escaneo.")
            return None

        # === Generación de PDF A4 (escalado proporcional) ===
        a4_w, a4_h = A4
        pdf_path = os.path.join(carpeta_entrada, nombre_archivo_pdf)
        c = canvas.Canvas(pdf_path, pagesize=A4)

        for data, ext in rutas_mem:
            bio = BytesIO(data)
            img = ImageReader(bio)
            iw, ih = img.getSize()

            # Ajuste proporcional a A4
            aspect = iw / float(ih)
            sw = min(a4_w, a4_h * aspect)
            sh = sw / aspect
            xo = (a4_w - sw) / 2
            yo = (a4_h - sh) / 2

            c.drawImage(img, xo, yo, width=sw, height=sh)
            c.showPage()

        c.save()

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
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass


