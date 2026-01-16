# debug/debugapp.py
from dataclasses import dataclass

@dataclass
class DebugFlags:
    mostrar_ocr_rut: bool = False
    mostrar_ocr_factura: bool = False

DEBUG = DebugFlags()

def set_debug_flags(*, mostrar_ocr_rut=None, mostrar_ocr_factura=None):
    if mostrar_ocr_rut is not None:
        DEBUG.mostrar_ocr_rut = bool(mostrar_ocr_rut)
    if mostrar_ocr_factura is not None:
        DEBUG.mostrar_ocr_factura = bool(mostrar_ocr_factura)

def debug_print_rut(raw_text: str, clean_text: str):
    if not DEBUG.mostrar_ocr_rut:
        return
    print("\n===== DEBUG OCR RUT =====")
    print(">> RAW:\n", raw_text)
    print("\n>> CLEAN:\n", clean_text)
    print("=========================\n")

def debug_print_factura(raw_text: str, clean_text: str):
    if not DEBUG.mostrar_ocr_factura:
        return
    print("\n=== DEBUG OCR FACTURA ===")
    print(">> RAW:\n", raw_text)
    print("\n>> CLEAN:\n", clean_text)
    print("=========================\n")

