# updater.py
import json, os, re, sys, tempfile, urllib.request, hashlib, subprocess, shutil

REPO = "AngeloxG25/FacturaScan"  # <-- tu repo
GITHUB_LATEST_API = f"https://api.github.com/repos/{REPO}/releases/latest"
USER_AGENT = "FacturaScan-Updater"

def _parse_ver(v: str):
    # "1.8.0" -> (1,8,0); tolera "v1.8"
    nums = re.findall(r"\d+", (v or ""))
    nums = (nums + ["0","0","0"])[:3]
    return tuple(int(x) for x in nums)

def _http_get_json(url: str):
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=8) as r:
        return json.loads(r.read().decode("utf-8"))

def _pick_asset_url(release_json):
    # Toma el primer .exe de assets
    assets = release_json.get("assets") or []
    for a in assets:
        url = a.get("browser_download_url","")
        if url.lower().endswith(".exe"):
            return url
    return None

def _download(url: str, dst: str, chunk=1<<20):
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=30) as r, open(dst, "wb") as f:
        while True:
            b = r.read(chunk)
            if not b: break
            f.write(b)
    return dst

def _sha256(path: str):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1<<20), b""):
            h.update(b)
    return h.hexdigest()

def check_latest_against(current_version: str):
    # Opción A: GitHub Releases
    data = _http_get_json(GITHUB_LATEST_API)
    latest_tag = (data.get("tag_name") or "").lstrip("v")
    inst_url = _pick_asset_url(data)
    return {"latest": latest_tag, "installer_url": inst_url, "raw": data}

def download_and_run(installer_url: str, silent: bool = False):
    temp_dir = os.path.join(tempfile.gettempdir(), "FacturaScan_Update")
    os.makedirs(temp_dir, exist_ok=True)
    local_path = os.path.join(temp_dir, os.path.basename(installer_url) or "FacturaScan-Setup.exe")
    _download(installer_url, local_path)

    # Ejecuta instalador (interactivo por defecto); silent usa flags Inno
    args = [local_path]
    if silent:
        # Requiere que el .iss tenga CloseApplications/RestartApplications (ver más abajo)
        args += ["/VERYSILENT", "/NORESTART", "/CLOSEAPPLICATIONS", "/RESTARTAPPLICATIONS", "/SUPPRESSMSGBOXES"]

    # Oculta ventana si es posible (Windows)
    try:
        si = subprocess.STARTUPINFO(); si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = 0
        subprocess.Popen(args, startupinfo=si, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except Exception:
        subprocess.Popen(args)

    # Cierra la app para liberar archivos
    sys.exit(0)

def check_for_updates_now(current_version: str, auto_run=True, silent=False):
    try:
        info = check_latest_against(current_version)
        latest = info.get("latest")
        url    = info.get("installer_url")
        if latest and url and _parse_ver(latest) > _parse_ver(current_version):
            # Aquí podrías mostrar un diálogo propio "Hay una nueva versión X - Descargar/Instalar"
            if auto_run:
                download_and_run(url, silent=silent)
            return {"update_available": True, "latest": latest, "url": url}
        return {"update_available": False, "latest": latest, "url": url}
    except Exception as e:
        return {"error": str(e)}
