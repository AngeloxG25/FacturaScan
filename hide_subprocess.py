import os
import subprocess
from subprocess import CREATE_NO_WINDOW

_original_popen = subprocess.Popen
_original_run = subprocess.run
_original_call = subprocess.call

class FullyHiddenPopen(_original_popen):
    def __init__(self, *args, **kwargs):
        args_list = args[0] if args and isinstance(args[0], list) else []

        # Para procesos como pdfinfo.exe: oculta ventana pero permite capturar stdout/stderr
        if any("pdfinfo" in arg.lower() for arg in args_list):
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            kwargs['startupinfo'] = startupinfo
            kwargs['creationflags'] = CREATE_NO_WINDOW
            # No redirige stdout/stderr para que pdf2image pueda leer
            super().__init__(*args, **kwargs)
            return

        # Para otros procesos: ocultar ventana y redirigir entrada/salida
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        kwargs['startupinfo'] = startupinfo
        kwargs['creationflags'] = CREATE_NO_WINDOW
        kwargs.setdefault('stdin', subprocess.DEVNULL)
        kwargs.setdefault('stdout', subprocess.DEVNULL)
        kwargs.setdefault('stderr', subprocess.DEVNULL)

        super().__init__(*args, **kwargs)

def hidden_run(*args, **kwargs):
    if os.name == 'nt':
        if 'startupinfo' not in kwargs:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            kwargs['startupinfo'] = startupinfo
        kwargs['creationflags'] = CREATE_NO_WINDOW
    return _original_run(*args, **kwargs)

def hidden_call(*args, **kwargs):
    if os.name == 'nt':
        if 'startupinfo' not in kwargs:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            kwargs['startupinfo'] = startupinfo
        kwargs['creationflags'] = CREATE_NO_WINDOW
    return _original_call(*args, **kwargs)

# Monkey patch
subprocess.Popen = FullyHiddenPopen
subprocess.run = hidden_run
subprocess.call = hidden_call

# También define explícitamente para acceso externo si es necesario
__all__ = ["FullyHiddenPopen"]
