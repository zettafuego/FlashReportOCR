#!/usr/bin/env python3
"""
Generador FLASH REPORT (CIA. Minera Casma) desde fotos de manifiestos.

Uso:
  python flash_report.py "C:\\ruta\\foto.jpg"
  python flash_report.py foto1.jpg foto2.jpg --salida "C:\\salidas\\reporte.xlsx"
  python flash_report.py foto.jpg --no-gui
  python flash_report.py --json datos.json
  python flash_report.py foto.jpg --prefer gemini   # requiere GEMINI_API_KEY

Respeta colores (#FFC000 naranja), bordes, merges y logo de la plantilla oficial.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from extractor import extract_from_images, load_json_data, EMPTY_DATA
from excel_filler import fill_flash_report, suggest_output_name, DEFAULT_TEMPLATE
from gui_review import review_data

ROOT = Path(__file__).resolve().parent


def build_narrative_if_empty(data: dict) -> dict:
    """Completa textos profesionales si la IA/OCR dejó vacío pero hay pistas."""
    if not data.get("descripcion_evento") and data.get("nombres") and data.get("lugar"):
        nombre = data["nombres"]
        lugar = data["lugar"]
        equipo = data.get("equipo_operado") or "la unidad"
        fecha = data.get("fecha_evento") or ""
        data["descripcion_evento"] = (
            f"En {lugar}, el trabajador {nombre}, mientras realizaba sus labores "
            f"a bordo de {equipo}{', el ' + str(fecha) if fecha else ''}, "
            f"participó en el evento reportado. (Completar con detalle de la maniobra.)"
        )
    if data.get("supervisor") and not data.get("persona_reporta"):
        data["persona_reporta"] = data["supervisor"]
    if data.get("persona_reporta") and not data.get("cargo_reporta"):
        data["cargo_reporta"] = "JEFE DE GUARDIA"
    return data


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Reconoce fotos de manifiestos y genera FLASH REPORT en Excel"
    )
    parser.add_argument(
        "fotos",
        nargs="*",
        help="Rutas de fotos del formato manuscrito / del evento",
    )
    parser.add_argument(
        "--json",
        dest="json_path",
        help="Cargar datos ya estructurados (salta OCR/IA)",
    )
    parser.add_argument(
        "--salida",
        "-o",
        help="Ruta del Excel de salida",
    )
    parser.add_argument(
        "--plantilla",
        default=str(DEFAULT_TEMPLATE),
        help="Plantilla .xlsx (por defecto plantilla_flash_report.xlsx)",
    )
    parser.add_argument(
        "--prefer",
        choices=["auto", "gemini", "xai", "ocr"],
        default="auto",
        help="Motor de reconocimiento",
    )
    parser.add_argument(
        "--no-gui",
        action="store_true",
        help="No abrir ventana de revisión (modo automático)",
    )
    parser.add_argument(
        "--sin-fotos",
        action="store_true",
        help="No insertar las imágenes en la sección de fotografías",
    )
    parser.add_argument(
        "--guardar-json",
        help="Guardar JSON de datos extraídos/corregidos",
    )
    args = parser.parse_args(argv)

    if not args.json_path and not args.fotos:
        parser.print_help()
        print(
            "\nEjemplo:\n  python flash_report.py "
            r'"C:\Users\Eduardo\Downloads\content.jpg"'
        )
        return 1

    # 1) Datos
    if args.json_path:
        data = load_json_data(args.json_path)
        print(f"Datos cargados desde JSON: {args.json_path}")
    else:
        print("Reconociendo foto(s)…")
        for f in args.fotos:
            print(f"  - {f}")
        data = extract_from_images(args.fotos, prefer=args.prefer)
        print(f"Fuente de extracción: {data.get('fuente', '?')}")

    data = build_narrative_if_empty(data)

    # 2) Revisión
    if not args.no_gui:
        reviewed = review_data(data)
        if reviewed is None:
            print("Cancelado por el usuario.")
            return 2
        data = reviewed
    else:
        if not data.get("nombres"):
            print(
                "ERROR: no se pudo leer el nombre. "
                "Use la GUI (sin --no-gui), un JSON, o configure GEMINI_API_KEY / XAI_API_KEY."
            )
            if data.get("texto_ocr"):
                print("--- OCR parcial ---")
                print(data["texto_ocr"][:1500])
            return 3

    if args.guardar_json:
        jp = Path(args.guardar_json)
        jp.parent.mkdir(parents=True, exist_ok=True)
        jp.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"JSON guardado: {args.guardar_json}")

    # 3) Excel
    out = Path(args.salida) if args.salida else suggest_output_name(data)
    photos = [] if args.sin_fotos else list(args.fotos or [])
    path = fill_flash_report(
        data,
        output_path=out,
        template_path=args.plantilla,
        photo_paths=photos,
    )
    print(f"\n✓ FLASH REPORT generado:\n  {path}")
    print("  (Colores, bordes y logo de la plantilla conservados)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
