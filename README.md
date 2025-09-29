# 📄 FacturaScan – Sistema de Escaneo y Procesamiento de Documentos Electrónicos

FacturaScan es una aplicación de escritorio desarrollada en Python para automatizar el escaneo, reconocimiento y clasificación de documentos. Utiliza un escáner compatible con WIA, aplica OCR para extraer el RUT y número de factura, y organiza los archivos PDF comprimidos de forma estructurada por año. Además, genera logs para trazabilidad y control.

## 📦 Requisitos del sistema

**Sistema y software**
- **SO:** Windows 10/11 **64-bit**
- **Python:** 3.10 (recomendado entorno virtual)
- **Escáner:** compatible **WIA**
- **Ghostscript:** [descarga](https://www.ghostscript.com/download/gsdnld.html)
- **Poppler (Windows builds):** [descarga](https://github.com/oschwartz10612/poppler-windows/releases/tag/v24.08.0-0)  
  > Asegúrate de instalar Poppler en `C:\poppler\Library\bin` y agregar esa ruta al **PATH** si no se detectan automáticamente.

**Hardware recomendado**
- **CPU (óptimo):** 6–8 núcleos (Intel Core i5/i7 10ª gen+ o Ryzen 5/7 4000+)
- **RAM (óptimo):** **16 GB**
- **Disco:** **SSD NVMe** con al menos **10 GB** libres para temporales y PDFs
- **Conexión del escáner:** USB 3.0 o red estable

**Mínimos (funciona)**
- **CPU:** 4 núcleos (Core i3 8ª gen / Ryzen 3 3000+) con **AVX2**
- **RAM:** **8 GB**
- **Disco:** SSD

**Uso intensivo / lotes grandes**
- **CPU:** 8–12 hilos reales (Core i7/i9 modernos o Ryzen 7/9)
- **RAM:** **32 GB**
- **Disco:** SSD NVMe con **50+ GB** libres

> Notas: El OCR (EasyOCR/PyTorch) y el rasterizado PDF (Poppler) son **CPU-bound** y se benefician de más núcleos e I/O rápida. No se requiere GPU.

## 📥 Clonar repositorio

1. Clonar o descargar el repositorio:

   git clone https://github.com/AngeloxG25/FacturaScan.git
   cd FacturaScan

## 🧰 Instalación de dependencias

1. Crear entorno virtual (recomendado)

   - py -3.10 -m venv venv310
   - .\venv310\Scripts\Activate.ps1

2. Instalar dependencias:

   - py -3.10 -m pip install --upgrade pip
   - py -3.10 -m pip install torch==1.12.1+cpu -f https://download.pytorch.org/whl/cpu/torch_stable.html
   - py -3.10 -m pip install torchvision==0.13.1+cpu -f https://download.pytorch.org/whl/cpu/torch_stable.html
   - py -3.10 -m pip install customtkinter, pdf2image, easyocr, pywin32, pillow, nuitka, reportlab
   - py -3.10 -m pip install "numpy==1.26.4"
   - py -3.10 -m pip install "opencv-python-headless==4.8.1.78"

## 🚀 Ejecución

1. Ejecutar configuración inicial:

   La primera vez que se ejecuta el sistema, se abrirá una ventana donde debes:
   - Seleccionar el archivo de configuración
   - Seleccionar razón social
   - Seleccionar sucursal
   - Definir carpeta de entrada y salida

   Esto guardará la configuración en: C:\FacturaScan\config_*.txt

2. Ejecutar la aplicación:

   py -3.10 FacturaScan.py

## 🖨️ Flujo de funcionamiento

1. Escanea documento desde un escáner físico
2. Se convierte a PDF y se guarda en la carpeta de entrada
3. Se genera imagen PNG en C:\FacturaScan\debug
4. Se extrae texto mediante OCR
5. Se detecta RUT y número de factura
6. Se comprime el PDF con GhostScript
7. Se renombra automáticamente y se guarda por año
8. Se registra actividad en logs

## 🗂️ Estructura del proyecto


FacturaScan/
├─ FacturaScan.py           # Interfaz principal (GUI)
├─ monitor_core.py          # Procesamiento y OCR
├─ scanner.py               # Escaneo vía WIA
├─ config_gui.py            # Asistente de configuración
├─ ocr_utils.py             # Reglas OCR (RUT y Nº de factura)
├─ pdf_tools.py             # Compresión PDF (Ghostscript)
├─ log_utils.py             # Logs y niveles de depuración
├─ updater.py               # Actualizador de FacturaScan
├─ hide_subprocess.py       # Oculta ventanas de procesos en Windows
├─ assets/
│  └─ images/               # Íconos y recursos
└─ (Carpetas de trabajo del sistema)
   ├─ C:\FacturaScan\debug  # PNGs temporales
   └─ C:\FacturaScan\logs   # Logs diarios


## 📝 Notas adicionales

- Si Poppler o GhostScript no están correctamente instalados, el sistema intentará añadir la ruta automáticamente al PATH.
- Ejecuta FacturaScan con permisos de administrador si hay problemas con acceso a escáner o configuración.

## 🧱 Compilación (Nuitka)

> Ejecuta desde **PowerShell** (con el *venv* activo). Usa un **bloque** como este para evitar el “desorden” de opciones sueltas:

powershell
python -m nuitka .\FacturaScan.py `
  --standalone `
  --enable-plugin=tk-inter `
  --windows-icon-from-ico=assets/iconoScan.ico `
  --windows-console-mode=disable `
  --output-dir=dist `
  --assume-yes-for-downloads `
  --noinclude-default-mode=nofollow `
  --nofollow-import-to=pytest `
  --nofollow-import-to=unittest `
  --nofollow-import-to=setuptools `
  --nofollow-import-to=scipy.optimize `
  --nofollow-import-to=scipy.interpolate `
  --nofollow-import-to=scipy.stats `
  --include-module=win32com `
  --include-module=win32com.client `
  --include-module=pywintypes `
  --include-module=customtkinter `
  --include-module=PIL `
  --include-module=easyocr `
  --include-module=pdf2image `
  --include-module=pydoc `
  --include-module=reportlab `
  --include-module=reportlab.pdfgen.canvas `
  --include-module=torchvision `
  --include-module=torchvision.transforms `
  --include-module=skimage `
  --include-package=numpy `
  --include-package=cv2 `
  --include-package-data=numpy `
  --include-package-data=cv2 `
  --include-data-files=assets/icono_escanear.png=assets/icono_escanear.png `
  --include-data-files=assets/icono_carpeta.png=assets/icono_carpeta.png `
  --include-data-files=assets/iconoScan.ico=assets/iconoScan.ico `
  --include-data-files=assets/iconoScan16.ico=assets/iconoScan16.ico `
  --module-parameter=torch-disable-jit=yes `
  --lto=no `
  --jobs=8
