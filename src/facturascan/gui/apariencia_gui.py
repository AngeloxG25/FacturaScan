# gui/apariencia_gui.py
import os
import json
import tempfile
import customtkinter as ctk
from tkinter import messagebox


# ---------------------------
# Persistencia (JSON)
# ---------------------------
def _get_ui_config_path(base_dir: str, log_fn=None) -> str:
    """
    Devuelve una ruta persistente para config de UI.
    Prioriza C:\\FacturaScan, luego APPDATA\\FacturaScan, y por último base_dir.
    """
    candidates = []

    if os.name == "nt":
        candidates.append(r"C:\FacturaScan")
        appdata = os.environ.get("APPDATA")
        if appdata:
            candidates.append(os.path.join(appdata, "FacturaScan"))

    candidates.append(base_dir)

    for d in candidates:
        try:
            os.makedirs(d, exist_ok=True)

            # test de escritura
            test_path = os.path.join(d, "._ui_write_test")
            with open(test_path, "w", encoding="utf-8") as f:
                f.write("ok")
            os.remove(test_path)

            return os.path.join(d, "ui_config.json")
        except Exception as e:
            if log_fn:
                try:
                    log_fn(f"[Apariencia] No se pudo usar '{d}': {e}")
                except Exception:
                    pass

    # último recurso
    return os.path.join(tempfile.gettempdir(), "FacturaScan_ui_config.json")


def cargar_tamano_log(base_dir: str, default: int = 12, log_fn=None) -> int:
    path = _get_ui_config_path(base_dir, log_fn=log_fn)
    try:
        if os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            size = int(data.get("log_font_size", default))
            return max(8, min(24, size))
    except Exception as e:
        if log_fn:
            try:
                log_fn(f"[Apariencia] Error al leer ui_config.json: {e}")
            except Exception:
                pass
    return max(8, min(24, int(default)))


def guardar_tamano_log(base_dir: str, size: int, log_fn=None) -> bool:
    path = _get_ui_config_path(base_dir, log_fn=log_fn)
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        data = {"log_font_size": int(size)}
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        if log_fn:
            try:
                log_fn(f"[Apariencia] No se pudo guardar ui_config.json: {e}")
            except Exception:
                pass
        return False


# ---------------------------
# Modal Apariencia
# ---------------------------
def abrir_modal_apariencia(
    parent,
    base_dir: str,
    font_log: ctk.CTkFont,
    log_font_state: dict,
    mensaje_label=None,
    aplicar_icono_fn=None,
    log_fn=None
):
    """
    Abre un modal (Toplevel) para aumentar/disminuir la letra del log.
    - Cambia al instante (modifica font_log)
    - Guardar persiste en ui_config.json
    - Cancelar revierte
    """
    original = int(log_font_state.get("size", 12))

    modal = ctk.CTkToplevel(parent)
    modal.title("Apariencia")
    # IMPORTANTÍSIMO: esperar un poquito a que exista la ventana
    modal.after(10, lambda: aplicar_icono_fn(modal))
    modal.after(200, lambda: aplicar_icono_fn(modal))
    if aplicar_icono_fn:
        try:
            aplicar_icono_fn(modal)
        except Exception:
            pass

    w, h = 360, 180
    x = (parent.winfo_screenwidth() - w) // 2
    y = (parent.winfo_screenheight() - h) // 2
    modal.geometry(f"{w}x{h}+{x}+{y}")
    modal.resizable(False, False)

    # Primer plano + modal real
    modal.transient(parent)
    modal.grab_set()
    modal.focus_force()
    modal.attributes("-topmost", True)
    modal.after(120, lambda: modal.attributes("-topmost", False))

    ctk.CTkLabel(
        modal,
        text="Tamaño de letra del log",
        font=ctk.CTkFont(size=16, weight="bold")
    ).pack(pady=(14, 6))

    size_var = ctk.StringVar(value=str(log_font_state.get("size", 12)))
    ctk.CTkLabel(
        modal,
        textvariable=size_var,
        font=ctk.CTkFont(size=22, weight="bold")
    ).pack(pady=(0, 8))

    row = ctk.CTkFrame(modal, fg_color="transparent")
    row.pack(pady=(0, 10))

    def _aplicar_size(new_size: int):
        new_size = max(8, min(24, int(new_size)))
        log_font_state["size"] = new_size
        try:
            font_log.configure(size=new_size)
        except Exception:
            pass
        size_var.set(str(new_size))
        if mensaje_label:
            try:
                mensaje_label.configure(text=f"🅰️ Tamaño de letra del log: {new_size}")
            except Exception:
                pass

    def _menos():
        _aplicar_size(int(log_font_state.get("size", 12)) - 1)

    def _mas():
        _aplicar_size(int(log_font_state.get("size", 12)) + 1)

    ctk.CTkButton(row, text="–", width=60, height=36, command=_menos).pack(side="left", padx=8)
    ctk.CTkButton(row, text="+", width=60, height=36, command=_mas).pack(side="left", padx=8)

    acciones = ctk.CTkFrame(modal, fg_color="transparent")
    acciones.pack(fill="x", padx=14, pady=(0, 12))

    def _cancelar():
        _aplicar_size(original)
        modal.destroy()

    def _guardar():
        ok = guardar_tamano_log(base_dir, log_font_state.get("size", 12), log_fn=log_fn)
        if ok:
            if mensaje_label:
                try:
                    mensaje_label.configure(text="✅ Apariencia guardada.")
                except Exception:
                    pass
        else:
            messagebox.showwarning("Apariencia", "No se pudo guardar la configuración.")
        modal.destroy()

    ctk.CTkButton(acciones, text="Cancelar", fg_color="#9ca3af", hover_color="#6b7280", command=_cancelar).pack(side="left")
    ctk.CTkButton(acciones, text="Guardar", fg_color="#2563eb", hover_color="#1d4ed8", command=_guardar).pack(side="right")

    modal.protocol("WM_DELETE_WINDOW", _cancelar)
