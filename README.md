# 📄 FacturaScan – Sistema de Escaneo y Procesamiento de Documentos Electrónicos

FacturaScan es una aplicación de escritorio desarrollada en Python para automatizar el escaneo, reconocimiento y clasificación de documentos. Utiliza un escáner compatible con WIA, aplica OCR para extraer el RUT y número de factura, y organiza los archivos PDF comprimidos de forma estructurada por año. Además, genera logs para trazabilidad y control.

## 📦 Requisitos del sistema

- Windows 10 o superior
- Python 3.10
- Escáner compatible con WIA
- GhostScript instalado (https://www.ghostscript.com/download/gsdnld.html)
- Poppler instalado (https://github.com/oschwartz10612/poppler-windows/releases/tag/v24.08.0-0)
- EasyOCR (https://github.com/JaidedAI/EasyOCR)

Asegúrate de que Poppler esté instalado en: C:\poppler\Library\bin

## 🛠️ Instalación de dependencias

1. Crear entorno virtual (opcional pero recomendado):

   - py -3.10 -m venv venv310
   - .\venv310\Scripts\Activate.ps1

2. Instalar dependencias:

   - py -3.10 -m pip install torch==1.12.1+cpu -f https://download.pytorch.org/whl/cpu/torch_stable.html
   - py -3.10 -m pip install torchvision==0.13.1+cpu -f https://download.pytorch.org/whl/cpu/torch_stable.html
   - py -3.10 -m pip install customtkinter
   - py -3.10 -m pip install pdf2image
   - py -3.10 -m pip install easyocr
   - py -3.10 -m pip install pywin32
   - py -3.10 -m pip install pillow
   - py -3.10 -m pip install nuitka
   - python.exe -m pip install --upgrade pip
   - py -3.10 -m pip install reportlab


## 🚀 Instalación y ejecución

1. Clonar o descargar el repositorio:

   git clone https://github.com/tu-usuario/facturascan.git
   cd facturascan

   (o descargar el .zip y descomprimirlo)

2. Ejecutar configuración inicial:

   La primera vez que se ejecuta el sistema, se abrirá una ventana donde podrás:
   - Seleccionar razón social
   - Seleccionar sucursal
   - Definir carpeta de entrada y salida

   Esto guardará la configuración en: C:\FacturaScan\config_*.txt

3. Ejecutar la aplicación:

   python FacturaScan.py

## 🖨️ Flujo de funcionamiento

1. Escanea documento desde un escáner físico
2. Se convierte a PDF y se guarda en la carpeta de entrada
3. Se genera imagen PNG en C:\FacturaScan\debug
4. Se extrae texto mediante OCR
5. Se detecta RUT y número de factura
6. Se comprime el PDF con GhostScript
7. Se renombra automáticamente y se guarda por año
8. Se registra actividad en logs

## 📁 Estructura del proyecto

facturascan/
├── FacturaScan.py           → Interfaz principal
├── monitor_core.py          → Procesamiento y OCR
├── scanner.py               → Escaneo por WIA
├── config_gui.py            → Configuración inicial
├── ocr_utils.py             → Extracción de RUT y número de factura
├── pdf_tools.py             → Compresión de PDF
├── C:\FacturaScan\debug     → PNGs temporales
├── C:\FacturaScan\logs      → Archivos de logs por fecha

## 📝 Notas adicionales

- Si Poppler o GhostScript no están correctamente instalados, el sistema intentará añadir la ruta automáticamente al PATH.
- Ejecuta FacturaScan con permisos de administrador si hay problemas con acceso a escáner o configuración.

## COMPILACIÓN:

En powershell:

python -m nuitka FacturaScan.py `
  --standalone `
  --enable-plugin=tk-inter `
  --enable-plugin=pylint-warnings `
  --windows-icon-from-ico=iconoScan.ico `
  --windows-console-mode=disable `
  --output-dir=dist `
  --remove-output `
  --assume-yes-for-downloads `
  --nofollow-import-to=pytest `
  --nofollow-import-to=unittest `
  --nofollow-import-to=setuptools `
  --nofollow-import-to=scipy.optimize `
  --nofollow-import-to=scipy.interpolate `
  --nofollow-import-to=scipy.stats `
  --noinclude-default-mode=nofollow `
  --include-module=win32com `
  --include-module=pywintypes `
  --include-module=customtkinter `
  --include-module=PIL `
  --include-module=easyocr `
  --include-module=pdf2image `
  --include-module=pydoc `
  --include-data-files=images/icono_escanear.png=images/icono_escanear.png `
  --include-data-files=images/icono_carpeta.png=images/icono_carpeta.png `
  --include-data-files=iconoScan.ico=iconoScan.ico `
  --lto=yes `
  --jobs=8 `
  --show-progress
