"""
Actualizador de FacturaScan (GitHub Releases + Inno Setup)

Qué hace:
- Consulta la última versión en GitHub Releases (/releases/latest)
- Si hay actualización: muestra modal "Comprobando..."
- Luego muestra modal "Actualizando..." con progreso
- Descarga instalador (.exe), verifica SHA-256 (si existe), ejecuta instalador
- Limpia carpeta temporal al terminar

Uso desde app.py (ideal):
    from update.updater import schedule_update_prompt
    schedule_update_prompt(ventana, current_version=VERSION, apply_icono_fn=aplicar_icono)
"""

from __future__ import annotations

import os
import re
import json
import hashlib
import tempfile
import subprocess
import urllib.request
import urllib.error
import urllib.parse
import threading
from typing import Callable, Dict, Optional, Tuple, Any

# ===================== Config =====================

GITHUB_OWNER = "AngeloxG25"
GITHUB_REPO  = "FacturaScan"

# Selección preferida de instalador:
# primer .exe que contenga "setup"
INSTALLER_PREDICATE = lambda name: name.lower().endswith(".exe") and "setup" in name.lower()

UA = f"{GITHUB_REPO}-Updater/1.0 (+https://github.com/{GITHUB_OWNER}/{GITHUB_REPO})"
GITHUB_LATEST_API = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"

# ===================== Utilidades =====================

class DownloadCancelled(Exception):
    """Señala que el usuario canceló la descarga (si usas cancel_event)."""
    pass


def cleanup_temp_dir(path: str) -> None:
    """Elimina una carpeta temporal sin romper si algo falla."""
    import shutil
    try:
        shutil.rmtree(path, ignore_errors=True)
    except Exception:
        pass


def _http_get(url: str, timeout: int = 25) -> bytes:
    req = urllib.request.Request(url, headers={
        "User-Agent": UA,
        "Accept": "application/vnd.github+json, application/json;q=0.9, */*;q=0.1",
        "Accept-Encoding": "identity",
    })
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def _json_get(url: str, timeout: int = 25) -> Dict[str, Any]:
    raw = _http_get(url, timeout=timeout)
    return json.loads(raw.decode("utf-8", errors="replace"))


def _version_tuple(v: str) -> Tuple[int, ...]:
    v = (v or "").strip()
    if v.startswith(("v", "V")):
        v = v[1:]
    m = re.match(r"^(\d+(?:\.\d+)*)", v)
    core = m.group(1) if m else "0"
    return tuple(int(x) for x in core.split("."))


def _is_newer(latest: str, current: str) -> bool:
    try:
        return _version_tuple(latest) > _version_tuple(current)
    except Exception:
        return False


def _select_installer_and_sha(assets: list) -> Tuple[Optional[dict], Optional[dict]]:
    installer = None
    sha_asset = None

    for a in assets:
        name = a.get("name", "")
        if INSTALLER_PREDICATE(name):
            installer = a
            break

    if not installer:
        for a in assets:
            if a.get("name", "").lower().endswith(".exe"):
                installer = a
                break

    if not installer:
        return None, None

    inst_name = installer.get("name", "")

    # sha256 específico del instalador
    for a in assets:
        n = a.get("name", "").lower()
        if n.endswith(".sha256") and inst_name.lower() in n:
            sha_asset = a
            break

    # fallback: archivo general de checksums
    if not sha_asset:
        for cand in ("sha256sum.txt", "sha256sums.txt", "checksums.txt", "sha256sums", "SHA256SUMS"):
            for a in assets:
                if a.get("name", "").lower() == cand.lower():
                    sha_asset = a
                    break
            if sha_asset:
                break

    return installer, sha_asset


def _parse_sha256_file(text: str, target_filename: str) -> Optional[str]:
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    # buscar hash + nombre exacto
    for ln in lines:
        m = re.match(r"^([a-fA-F0-9]{64})\s+\*?\s*(.+)$", ln)
        if not m:
            continue
        h, fname = m.group(1), m.group(2)
        if os.path.basename(fname) == os.path.basename(target_filename):
            return h.lower()

    # fallback: primer hash válido
    for ln in lines:
        m = re.match(r"^([a-fA-F0-9]{64})\b", ln)
        if m:
            return m.group(1).lower()

    return None


# ===================== API lógica (sin UI) =====================

def get_latest_release_info() -> Dict[str, Any]:
    try:
        return _json_get(GITHUB_LATEST_API, timeout=25)
    except urllib.error.HTTPError as e:
        return {"error": f"HTTP {e.code}: {e.reason}"}
    except Exception as e:
        return {"error": str(e)}


def is_update_available(current_version: str) -> Dict[str, Any]:
    rel = get_latest_release_info()
    if rel.get("error"):
        return {"update_available": False, "error": rel["error"]}

    latest = (rel.get("tag_name") or "").strip()
    body   = rel.get("body") or ""
    assets = rel.get("assets") or []

    inst, sha_asset = _select_installer_and_sha(assets)

    if not latest:
        return {"update_available": False, "error": "Respuesta sin 'tag_name'."}

    out: Dict[str, Any] = {
        "update_available": _is_newer(latest, current_version),
        "latest": latest,
        "release_notes": body,
        "installer_url": inst.get("browser_download_url") if inst else None,
        "installer_name": inst.get("name") if inst else None,
        "sha256": None,
        "error": None,
    }

    if sha_asset:
        try:
            txt = _http_get(sha_asset.get("browser_download_url")).decode("utf-8", errors="replace")
            out["sha256"] = _parse_sha256_file(txt, out["installer_name"] or "")
        except Exception:
            out["sha256"] = None

    return out


def download_with_progress(
    url: str,
    dst_dir: str,
    progress_cb: Optional[Callable[[int, Optional[int]], None]] = None,
    cancel_event: Optional["threading.Event"] = None,
) -> str:
    import threading  # solo para type/compat (cancel_event)

    os.makedirs(dst_dir, exist_ok=True)
    filename = os.path.basename(urllib.parse.urlparse(url).path) or "download.exe"
    dst_path = os.path.join(dst_dir, filename)

    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
    with urllib.request.urlopen(req, timeout=120) as r, open(dst_path, "wb") as f:
        total = None
        try:
            total = int(r.headers.get("Content-Length"))
        except Exception:
            total = None

        read = 0
        chunk = 1024 * 256

        while True:
            if cancel_event is not None and cancel_event.is_set():
                try:
                    f.close()
                except Exception:
                    pass
                try:
                    os.remove(dst_path)
                except Exception:
                    pass
                raise DownloadCancelled()

            data = r.read(chunk)
            if not data:
                break

            f.write(data)
            read += len(data)

            if progress_cb:
                try:
                    progress_cb(read, total)
                except Exception:
                    pass

    return dst_path


def verify_sha256(path: str, expected_hex: Optional[str]) -> bool:
    if not expected_hex:
        return True
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest().lower() == expected_hex.lower()


def run_installer(path: str, mode: str = "progress", cleanup_dir: Optional[str] = None) -> None:
    """
    mode:
      - "progress": /SILENT (muestra ventana de progreso)
      - "silent":   /VERYSILENT (sin ventana)
      - "full":     sin flags
    """
    if not os.path.exists(path):
        raise FileNotFoundError(path)

    if os.name == "nt":
        CREATE_NO_WINDOW = 0x08000000
        DETACHED_PROCESS = 0x00000008

        args = [path]
        if mode == "progress":
            args += ["/SP-", "/SILENT", "/SUPPRESSMSGBOXES",
                     "/CLOSEAPPLICATIONS", "/RESTARTAPPLICATIONS", "/NORESTART"]
        elif mode == "silent":
            args += ["/SP-", "/VERYSILENT", "/SUPPRESSMSGBOXES",
                     "/CLOSEAPPLICATIONS", "/RESTARTAPPLICATIONS", "/NORESTART"]
        # full -> sin flags

        proc = subprocess.Popen(args, creationflags=CREATE_NO_WINDOW)

        # limpieza cuando el instalador termina
        if cleanup_dir:
            ps = rf"""
$pid  = {proc.pid};
$dir  = '{cleanup_dir}'.Trim('"');
try {{ Wait-Process -Id $pid -ErrorAction SilentlyContinue }} catch {{ }}
Start-Sleep -Seconds 2
for ($i=0; $i -lt 15; $i++) {{
  try {{ Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction Stop; break }}
  catch {{ Start-Sleep -Seconds 2 }}
}}
"""
            try:
                subprocess.Popen(
                    ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps],
                    creationflags=DETACHED_PROCESS | CREATE_NO_WINDOW
                )
            except Exception:
                subprocess.Popen(
                    ["cmd", "/c", f'ping -n 8 127.0.0.1 >nul & rmdir /s /q "{cleanup_dir}"'],
                    creationflags=DETACHED_PROCESS | CREATE_NO_WINDOW
                )

        # cerrar FacturaScan para que Inno maneje cierre/reinicio sin prompts
        os._exit(0)

    else:
        try:
            os.chmod(path, 0o755)
        except Exception:
            pass
        subprocess.Popen([path])


# ===================== UI FLOW (CustomTkinter) =====================

def schedule_update_prompt(
    parent_window,
    current_version: str,
    apply_icono_fn: Optional[Callable] = None,
    delay_ms: int = 400,
    check_timeout_ms: int = 12000,
    installer_mode: str = "progress",
) -> None:
    """Programa la comprobación de actualización sin bloquear UI."""
    try:
        parent_window.after(
            int(delay_ms),
            lambda: _comprobar_update_async(
                parent_window,
                current_version=current_version,
                apply_icono_fn=apply_icono_fn,
                check_timeout_ms=check_timeout_ms,
                installer_mode=installer_mode,
            )
        )
    except Exception:
        pass


def _apply_icon_safe(win, apply_icono_fn: Optional[Callable]) -> None:
    """Aplica icono de forma robusta: después de deiconify() + reintentos."""
    if not apply_icono_fn:
        return

    def _try():
        try:
            if win.winfo_exists():
                apply_icono_fn(win)
        except Exception:
            pass

    _try()
    try:
        win.after(120, _try)
        win.after(350, _try)
    except Exception:
        pass


def _comprobar_update_async(
    parent_window,
    current_version: str,
    apply_icono_fn: Optional[Callable],
    check_timeout_ms: int,
    installer_mode: str,
) -> None:
    try:
        import customtkinter as ctk
        import threading

        chk = ctk.CTkToplevel(parent_window)
        chk.withdraw()
        chk.title("Comprobando actualización…")
        chk.resizable(False, False)

        def _centrar(child, parent, w, h):
            parent.update_idletasks()
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            pw, ph = parent.winfo_width(), parent.winfo_height()
            x = px + max(0, (pw - w) // 2)
            y = py + max(0, (ph - h) // 2)
            child.geometry(f"{w}x{h}+{x}+{y}")

        W, H = 420, 160
        _centrar(chk, parent_window, W, H)

        ctk.CTkLabel(
            chk,
            text="🔄 Comprobando actualización…",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(pady=(18, 8))

        pb = ctk.CTkProgressBar(chk, mode="indeterminate", width=260)
        pb.pack(pady=(0, 16))
        pb.start()

        chk.protocol("WM_DELETE_WINDOW", lambda: None)
        chk.bind("<Escape>", lambda e: "break")

        chk.transient(parent_window)
        chk.attributes("-topmost", True)
        chk.after(80, lambda: chk.attributes("-topmost", False))

        chk.deiconify()
        chk.update_idletasks()
        _apply_icon_safe(chk, apply_icono_fn)  # ✅ icono después de mostrar

        alive = {"value": True}

        def _cerrar_chk():
            if not alive["value"]:
                return
            alive["value"] = False
            try:
                pb.stop()
            except Exception:
                pass
            try:
                chk.destroy()
            except Exception:
                pass

        chk.after(int(check_timeout_ms), _cerrar_chk)

        def worker():
            try:
                info = is_update_available(current_version)
            except Exception as e:
                info = {"error": str(e), "update_available": False}

            def ui():
                if not alive["value"]:
                    return
                _cerrar_chk()

                if info and (not info.get("error")) and info.get("update_available"):
                    _mostrar_dialogo_update(
                        parent_window,
                        info=info,
                        apply_icono_fn=apply_icono_fn,
                        installer_mode=installer_mode
                    )

            try:
                parent_window.after(0, ui)
            except Exception:
                pass

        threading.Thread(target=worker, daemon=True).start()

    except Exception:
        pass


def _mostrar_dialogo_update(
    parent_window,
    info: dict,
    apply_icono_fn: Optional[Callable],
    installer_mode: str = "progress",
) -> None:
    try:
        import customtkinter as ctk
        from tkinter import messagebox as _mb
        import threading

        if not info or info.get("error") or not info.get("update_available"):
            return

        latest = info.get("latest", "")
        url    = info.get("installer_url")
        sha    = info.get("sha256")
        if not url:
            return

        def _centrar(child, parent, w, h):
            parent.update_idletasks()
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            pw, ph = parent.winfo_width(), parent.winfo_height()
            x = px + max(0, (pw - w) // 2)
            y = py + max(0, (ph - h) // 2)
            child.geometry(f"{w}x{h}+{x}+{y}")

        prog = ctk.CTkToplevel(parent_window)
        prog.withdraw()
        prog.title("Actualizando FacturaScan…")
        prog.resizable(False, False)

        W, H = 460, 190
        _centrar(prog, parent_window, W, H)

        prog.deiconify()
        prog.update_idletasks()
        _apply_icon_safe(prog, apply_icono_fn)  # ✅ icono después de mostrar

        prog.transient(parent_window)
        prog.grab_set()
        prog.focus_force()
        prog.attributes("-topmost", True)
        prog.after(60, lambda: prog.attributes("-topmost", False))

        ctk.CTkLabel(
            prog,
            text=f"Hay una nueva versión {latest}.\nSe descargará e instalará automáticamente."
        ).pack(pady=(14, 8))

        pct_var = ctk.StringVar(value="0 %")
        ctk.CTkLabel(prog, textvariable=pct_var).pack(pady=(0, 6))

        bar = ctk.CTkProgressBar(prog)
        bar.set(0.0)
        bar.pack(fill="x", padx=16, pady=8)

        status_var = ctk.StringVar(value="Descargando instalador…")
        ctk.CTkLabel(prog, textvariable=status_var).pack(pady=(2, 10))

        prog.protocol("WM_DELETE_WINDOW", lambda: None)
        prog.bind("<Escape>", lambda e: "break")

        tmpdir   = os.path.join(tempfile.gettempdir(), "FacturaScan_Update")
        ui_alive = {"value": True}

        def _safe_ui(fn):
            if not ui_alive["value"]:
                return
            try:
                if prog.winfo_exists():
                    fn()
            except Exception:
                pass

        def _set_progress(read, total):
            if not ui_alive["value"]:
                return
            if total and total > 0:
                frac = max(0.0, min(1.0, read / total))
                prog.after(0, lambda: _safe_ui(lambda: (bar.set(frac), pct_var.set(f"{int(frac*100)} %"))))
            else:
                kb = read // 1024
                prog.after(0, lambda: _safe_ui(lambda: pct_var.set(f"{kb} KB")))

        def _worker():
            try:
                exe_path = download_with_progress(url, tmpdir, progress_cb=_set_progress, cancel_event=None)

                if sha:
                    prog.after(0, lambda: _safe_ui(lambda: status_var.set("Verificando integridad…")))
                    if not verify_sha256(exe_path, sha):
                        prog.after(0, lambda: _mb.showerror("Actualización", "El archivo no pasó la verificación SHA-256."))
                        try:
                            cleanup_temp_dir(tmpdir)
                        except Exception:
                            pass
                        ui_alive["value"] = False
                        try:
                            prog.destroy()
                        except Exception:
                            pass
                        return

                prog.after(0, lambda: _safe_ui(lambda: status_var.set("Instalando actualización…")))
                ui_alive["value"] = False
                try:
                    prog.destroy()
                except Exception:
                    pass

                parent_window.after(200, lambda: run_installer(exe_path, mode=installer_mode, cleanup_dir=tmpdir))

            except Exception as e:
                try:
                    prog.after(0, lambda: _mb.showerror("Actualización", f"Error al actualizar:\n{e}"))
                finally:
                    ui_alive["value"] = False
                    try:
                        prog.destroy()
                    except Exception:
                        pass
                    try:
                        cleanup_temp_dir(tmpdir)
                    except Exception:
                        pass

        threading.Thread(target=_worker, daemon=True).start()

    except Exception:
        pass
