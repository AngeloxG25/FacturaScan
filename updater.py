# updater.py
# -*- coding: utf-8 -*-
"""
Actualizador de FacturaScan
- Busca la última versión en GitHub Releases
- Propone/descarga el instalador .exe
- Lee/verifica el .sha256 (si está publicado)
- Ejecuta el instalador
- Soporta cancelación segura durante la descarga

Uso típico desde la UI:
    from updater import (
        is_update_available, download_with_progress, verify_sha256, run_installer,
        DownloadCancelled, cleanup_temp_dir
    )
"""

from __future__ import annotations
import os
import re
import sys
import json
import hashlib
import tempfile
import subprocess
import urllib.request
import urllib.error
from typing import Callable, Dict, Optional, Tuple, Any

# ===================== Config =====================

# 🚨 Cambia esto por tu repo real si es distinto
GITHUB_OWNER = "AngeloxG25"
GITHUB_REPO  = "FacturaScan"

# Patrón esperado del instalador
# (se selecciona el primer asset .exe que contenga "Setup"; personaliza si usas otra convención)
INSTALLER_PREDICATE = lambda name: name.lower().endswith(".exe") and "setup" in name.lower()

UA = f"{GITHUB_REPO}-Updater/1.0 (+https://github.com/{GITHUB_OWNER}/{GITHUB_REPO})"
GITHUB_LATEST_API = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"

# ==================================================


# ------------- Excepciones / utilidades -------------
class DownloadCancelled(Exception):
    """Señala que el usuario canceló la descarga."""
    pass


def cleanup_temp_dir(path: str) -> None:
    """Elimina una carpeta temporal sin emitir errores si algo falla."""
    import shutil
    try:
        shutil.rmtree(path, ignore_errors=True)
    except Exception:
        pass


def _http_get(url: str, timeout: int = 25) -> bytes:
    """GET simple con cabeceras amigables para GitHub."""
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
    """Convierte 'v1.8.0' o '1.8.0-rc1' a tupla comparable."""
    v = (v or "").strip()
    if v.startswith("v") or v.startswith("V"):
        v = v[1:]
    # Solo dígitos y puntos iniciales
    m = re.match(r"^(\d+(?:\.\d+)*)", v)
    core = m.group(1) if m else "0"
    return tuple(int(x) for x in core.split("."))


def _is_newer(latest: str, current: str) -> bool:
    """True si latest > current según tuplas numéricas."""
    try:
        return _version_tuple(latest) > _version_tuple(current)
    except Exception:
        return False


def _select_installer_and_sha(assets: list) -> Tuple[Optional[dict], Optional[dict]]:
    """
    Elige el asset del instalador (.exe) y, si existe, el asset .sha256 correspondiente.
    Busca .sha256 a juego o archivos tipo 'sha256sum.txt'.
    """
    installer: Optional[dict] = None
    sha_asset: Optional[dict] = None

    # 1) Selección del instalador
    for a in assets:
        name = a.get("name", "")
        if INSTALLER_PREDICATE(name):
            installer = a
            break

    if not installer:
        # último recurso: cualquier .exe
        for a in assets:
            if a.get("name", "").lower().endswith(".exe"):
                installer = a
                break

    if not installer:
        return None, None

    inst_name = installer.get("name", "")

    # 2) Buscar .sha256 con el mismo nombre base
    for a in assets:
        n = a.get("name", "").lower()
        if n.endswith(".sha256") and inst_name.lower() in n:
            sha_asset = a
            break

    # 3) Fallback: archivo general de hashes
    if not sha_asset:
        for cand in ("sha256sum.txt", "sha256sums.txt", "checksums.txt", "SHA256SUMS"):
            for a in assets:
                if a.get("name", "").lower() == cand:
                    sha_asset = a
                    break
            if sha_asset:
                break

    return installer, sha_asset


def _parse_sha256_file(text: str, target_filename: str) -> Optional[str]:
    """
    Intenta extraer el hash sha256 del archivo 'target_filename'
    de un contenido de texto (formato 'hash  nombre' o 'hash *nombre').
    """
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    # Primero intenta coincidencia exacta por nombre
    for ln in lines:
        m = re.match(r"^([a-fA-F0-9]{64})\s+\*?\s*(.+)$", ln)
        if not m:
            continue
        h, fname = m.group(1), m.group(2)
        if os.path.basename(fname) == os.path.basename(target_filename):
            return h.lower()

    # Si no se encontró, toma el primer hash válido
    for ln in lines:
        m = re.match(r"^([a-fA-F0-9]{64})\b", ln)
        if m:
            return m.group(1).lower()

    return None


# ----------------- API principal -----------------

def get_latest_release_info() -> Dict[str, Any]:
    """Devuelve el JSON de /releases/latest o {'error': ...}."""
    try:
        return _json_get(GITHUB_LATEST_API, timeout=25)
    except urllib.error.HTTPError as e:
        return {"error": f"HTTP {e.code}: {e.reason}"}
    except Exception as e:
        return {"error": str(e)}


def is_update_available(current_version: str) -> Dict[str, Any]:
    """
    Comprueba si hay actualización disponible.
    Devuelve dict con:
      - update_available: bool
      - latest: str (tag_name)
      - installer_url: str | None
      - installer_name: str | None
      - sha256: str | None
      - release_notes: str (body) | None
      - error: str | None
    """
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
            # Si falla, seguimos sin hash
            out["sha256"] = None

    return out

import threading
import urllib.parse


def download_with_progress(
    url: str,
    dst_dir: str,
    progress_cb: Optional[Callable[[int, Optional[int]], None]] = None,
    cancel_event: Optional["threading.Event"] = None,) -> str:
    """
    Descarga 'url' en 'dst_dir'. Llama progress_cb(bytes_leidos, total_o_None).
    Si cancel_event.is_set(): aborta y lanza DownloadCancelled.
    Retorna la ruta final del archivo.
    """
    os.makedirs(dst_dir, exist_ok=True)
    filename = os.path.basename(urllib.parse.urlparse(url).path) or "download.exe"
    dst_path = os.path.join(dst_dir, filename)

    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
    with urllib.request.urlopen(req, timeout=120) as r, open(dst_path, "wb") as f:
        # GitHub suele enviar 'Content-Length'
        total = None
        try:
            total = int(r.headers.get("Content-Length"))
        except Exception:
            total = None

        read = 0
        chunk = 1024 * 256
        while True:
            if cancel_event is not None and cancel_event.is_set():
                try: f.close()
                except Exception: pass
                try: os.remove(dst_path)
                except Exception: pass
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
    """Calcula el SHA-256 del archivo y compara con expected_hex (si se entrega)."""
    if not expected_hex:
        return True  # sin hash publicado, no verificamos
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest().lower() == expected_hex.lower()


def run_installer(path: str, silent: bool = False) -> None:
    """
    Ejecuta el instalador. En Windows usa os.startfile si no es silencioso,
    o subprocess con argumentos si necesitas modo /S, etc.
    """
    if not os.path.exists(path):
        raise FileNotFoundError(path)

    if os.name == "nt":
        if silent:
            # Ajusta la bandera /S si tu instalador es Inno/NSIS
            try:
                subprocess.Popen([path, "/S"], close_fds=True)
            except Exception:
                os.startfile(path)  # fallback
        else:
            os.startfile(path)
    else:
        # *nix (por si distribuyes .AppImage/.sh)
        try:
            os.chmod(path, 0o755)
        except Exception:
            pass
        subprocess.Popen([path], close_fds=True)


# --------- Compatibilidad con código existente ---------

def check_for_updates_now(current_version: str, auto_run: bool = False, silent: bool = True) -> Dict[str, Any]:
    """
    Conserva el nombre usado en tu proyecto:
    - Si auto_run=True y hay update, intenta descargar/ejecutar sin UI (no recomendado).
    - Normalmente la UI usa is_update_available(...) y maneja el flujo.
    """
    info = is_update_available(current_version)
    if info.get("error"):
        return info

    if auto_run and info.get("update_available") and info.get("installer_url"):
        tmpdir = os.path.join(tempfile.gettempdir(), f"{GITHUB_REPO}_AutoUpdate")
        try:
            path = download_with_progress(info["installer_url"], tmpdir)
            if verify_sha256(path, info.get("sha256")):
                run_installer(path, silent=silent)
                info["auto_ran"] = True
            else:
                info["error"] = "El archivo descargado no pasó verificación SHA-256."
                cleanup_temp_dir(tmpdir)
        except Exception as e:
            info["error"] = str(e)
            cleanup_temp_dir(tmpdir)

    return info


# Alias por si en alguna parte lo llamabas así
def check_and_update(current_version: str, silent: bool = True) -> Dict[str, Any]:
    return check_for_updates_now(current_version, auto_run=False, silent=silent)


# ------------------ Modo script (debug) ------------------
if __name__ == "__main__":
    cur = sys.argv[1] if len(sys.argv) > 1 else "0.0.0"
    res = is_update_available(cur)
    print(json.dumps(res, indent=2, ensure_ascii=False))
