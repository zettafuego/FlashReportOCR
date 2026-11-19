"""Limpia la plantilla FLASH REPORT dejando solo estructura, colores y logo."""
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "plantilla_flash_report.xlsx"
# Preferir el archivo original de referencia si existe
REF = Path(r"C:\Users\Eduardo\Downloads\FLASH REPORT-jeferson javel 11-082026 (1).xlsx")
if REF.exists():
    SRC = REF

OUT = ROOT / "plantilla_flash_report.xlsx"

# Celdas de datos a vaciar (no toca etiquetas ni encabezados)
CLEAR_CELLS = [
    "H7", "H8", "H9", "H10", "H11", "H12", "H13", "H14", "H15", "H16", "H17", "H18",
    "D20", "K20", "T20",
    "E22", "J22", "O22", "V22", "Y22",  # marcas tipo (checkboxes)
    "E24", "J24",
    "D26", "H26", "K26",  # riesgo
    "H28", "H35",
    "X40", "X41",
    "B44",
    "I62", "I63",
    "AF24",  # basura residual
]

# Imágenes de accidente (mantener solo logo del encabezado: primera imagen)
KEEP_FIRST_IMAGE = True


def main():
    wb = load_workbook(SRC)
    ws = wb.active

    # Corregir título de sección 2 si quedó mal
    if ws["B19"].value in (2, "2", None):
        ws["B19"].value = "2. DATOS DEL EVENTO:"

    for addr in CLEAR_CELLS:
        cell = ws[addr]
        cell.value = None

    # Quitar fotos del accidente; conservar logo (primera imagen)
    if hasattr(ws, "_images") and ws._images:
        if KEEP_FIRST_IMAGE and len(ws._images) > 1:
            ws._images = [ws._images[0]]
        elif not KEEP_FIRST_IMAGE:
            ws._images = []

    wb.save(OUT)
    print(f"Plantilla limpia guardada en: {OUT}")


if __name__ == "__main__":
    main()
