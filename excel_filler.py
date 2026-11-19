"""Rellena la plantilla FLASH REPORT preservando colores, bordes y formas."""
from __future__ import annotations

import re
from datetime import datetime, time, date
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor, OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU, cm_to_EMU
from copy import copy

ROOT = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = ROOT / "plantilla_flash_report.xlsx"

# Celdas de datos (valores). Las etiquetas y estilos se dejan intactos.
FIELD_CELLS = {
    "nombres": "H7",
    "edad": "H8",
    "grado_instruccion": "H9",
    "tiempo_servicio": "H10",
    "area": "H11",
    "ocupacion": "H12",
    "procedencia": "H13",
    "experiencia": "H14",
    "estado_civil": "H15",
    "dni": "H16",
    "dia_inicio_guardia": "H17",
    "equipo_operado": "H18",
    "fecha_evento": "D20",
    "hora_evento": "K20",
    "lugar": "T20",
    "descripcion_evento": "H28",
    "descripcion_lesion": "H35",
    "causas": "B44",
    "persona_reporta": "I62",
    "cargo_reporta": "I63",
}

# Marcas "X" para tipo de evento / accidente / riesgo / investigación
TIPO_CHECK = {
    "dano_propiedad": "E22",
    "incidente_ambiental": "J22",
    "incidente": "O22",
    "incidente_peligroso": "V22",
    "atencion_medica": "Y22",
    "accidente_leve": "E24",
    "accidente_incapacitante": "J24",
}

RIESGO_CHECK = {
    "ALTO": "D26",
    "MEDIO": "H26",
    "BAJO": "K26",
}

# Anclas de fotos del accidente (col 0-based, row 0-based) — como en el Excel original
PHOTO_SLOTS = [
    (1, 49),   # col B, row 50
    (6, 49),   # col G
    (13, 49),  # col N
]


def _parse_date(value: Any) -> date | datetime | str | None:
    if value is None or value == "":
        return None
    if isinstance(value, (datetime, date)):
        return value
    s = str(value).strip()
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%d-%m-%y", "%d/%m/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return s


def _parse_time(value: Any) -> time | str | None:
    if value is None or value == "":
        return None
    if isinstance(value, time):
        return value
    if isinstance(value, datetime):
        return value.time()
    s = str(value).strip().replace(".", ":")
    for fmt in ("%H:%M", "%H:%M:%S", "%H.%M"):
        try:
            return datetime.strptime(s, fmt).time()
        except ValueError:
            continue
    m = re.match(r"^(\d{1,2}):(\d{2})$", s)
    if m:
        return time(int(m.group(1)), int(m.group(2)))
    return s


def _set_value(ws, addr: str, value: Any) -> None:
    """Escribe valor sin tocar estilo (fill, font, border, alignment)."""
    if value is None or value == "":
        return
    cell = ws[addr]
    # openpyxl conserva el estilo al asignar .value
    cell.value = value


def _clear_checkboxes(ws) -> None:
    for addr in list(TIPO_CHECK.values()) + list(RIESGO_CHECK.values()) + ["X40", "X41"]:
        ws[addr].value = None


def fill_flash_report(
    data: dict[str, Any],
    output_path: str | Path,
    template_path: str | Path | None = None,
    photo_paths: list[str | Path] | None = None,
) -> Path:
    """
    Copia la plantilla y rellena solo celdas de datos.
    Conserva colores (naranja #FFC000), bordes, merges, logo y anchos/altos.
    """
    template_path = Path(template_path or DEFAULT_TEMPLATE)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = load_workbook(template_path)
    ws = wb.active

    # Título sección 2
    if ws["B19"].value in (2, "2", None, ""):
        ws["B19"].value = "2. DATOS DEL EVENTO:"

    _clear_checkboxes(ws)

    # Campos de texto / fechas
    for key, addr in FIELD_CELLS.items():
        if key not in data:
            continue
        val = data[key]
        if key == "fecha_evento":
            val = _parse_date(val)
        elif key == "hora_evento":
            val = _parse_time(val)
        elif key == "dni" and val is not None:
            val = str(val).replace(".0", "").strip()
        _set_value(ws, addr, val)

    # Tipos (lista o string)
    tipos = data.get("tipos") or data.get("tipo") or []
    if isinstance(tipos, str):
        tipos = [tipos]
    tipo_aliases = {
        "daño a la propiedad": "dano_propiedad",
        "dano a la propiedad": "dano_propiedad",
        "dano_propiedad": "dano_propiedad",
        "incidente ambiental": "incidente_ambiental",
        "incidente": "incidente",
        "incidente peligroso": "incidente_peligroso",
        "atencion medica": "atencion_medica",
        "atención médica": "atencion_medica",
        "accidente leve": "accidente_leve",
        "accidente incapacitante": "accidente_incapacitante",
    }
    for t in tipos:
        key = tipo_aliases.get(str(t).strip().lower(), str(t).strip().lower())
        if key in TIPO_CHECK:
            _set_value(ws, TIPO_CHECK[key], "X")

    # Riesgo
    riesgo = str(data.get("riesgo") or "").strip().upper()
    if riesgo in RIESGO_CHECK:
        _set_value(ws, RIESGO_CHECK[riesgo], "X")

    # Investigación 24h
    requiere = data.get("requiere_investigacion")
    if requiere is True or str(requiere).lower() in ("si", "sí", "x", "true", "1", "requiere"):
        _set_value(ws, "X40", "x")
    elif requiere is False or str(requiere).lower() in ("no", "false", "0", "no requiere"):
        _set_value(ws, "X41", "x")
    else:
        # Por defecto: requiere (como en reportes de ejemplo)
        _set_value(ws, "X40", "x")

    # Fotos del accidente (además del logo)
    photos = [Path(p) for p in (photo_paths or []) if p and Path(p).exists()]
    # Quitar imágenes de accidente previas (conservar logo = primera)
    if hasattr(ws, "_images") and ws._images:
        logo = ws._images[0]
        ws._images = [logo]
    for i, photo in enumerate(photos[:3]):
        col, row = PHOTO_SLOTS[i]
        try:
            img = XLImage(str(photo))
            # Tamaño razonable para las celdas de fotos
            max_w, max_h = 280, 210
            if img.width and img.height:
                ratio = min(max_w / img.width, max_h / img.height, 1.0)
                img.width = int(img.width * ratio)
                img.height = int(img.height * ratio)
            img.anchor = f"{chr(65 + col) if col < 26 else 'A'}{row + 1}"
            # Usar ancla por celda con índice de columna/fila
            from openpyxl.utils import get_column_letter
            img.anchor = f"{get_column_letter(col + 1)}{row + 1}"
            ws.add_image(img)
        except Exception as exc:
            print(f"Aviso: no se pudo insertar foto {photo.name}: {exc}")

    wb.save(output_path)
    return output_path


def suggest_output_name(data: dict[str, Any], out_dir: Path | None = None) -> Path:
    out_dir = out_dir or ROOT / "salidas"
    out_dir.mkdir(parents=True, exist_ok=True)
    nombre = str(data.get("nombres") or "sin_nombre").strip()
    # Limpiar nombre de archivo
    nombre = re.sub(r'[<>:"/\\|?*]', "", nombre)
    nombre = re.sub(r"\s+", " ", nombre).strip() or "sin_nombre"
    fecha = data.get("fecha_evento") or datetime.now().strftime("%d-%m-%Y")
    if isinstance(fecha, (datetime, date)):
        fecha = fecha.strftime("%d-%m-%Y")
    else:
        fecha = str(fecha).replace("/", "-")
    return out_dir / f"FLASH REPORT-{nombre} {fecha}.xlsx"
