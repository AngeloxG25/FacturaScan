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

   python -m venv venv
   venv\Scripts\activate

2. Instalar dependencias:

   pip install -r requirements.txt

Contenido sugerido del archivo requirements.txt:

   pillow
   pdf2image
   easyocr
   pywin32
   numpy

tkinter ya viene incluido en Python para Windows.

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

