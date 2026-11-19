"""Genera plantilla sin dibujos (compatible con ExcelJS) + JS base64."""
from __future__ import annotations

import base64
import io
import re
import zipfile
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
CANDIDATES = [
    Path(r"C:\Users\Eduardo\Downloads\FLASH REPORT-jeferson javel 11-082026 (1).xlsx"),
    ROOT / "plantilla_flash_report.xlsx",
]
OUT = ROOT / "plantilla_flash_report.xlsx"

CLEAR = [
    "H7", "H8", "H9", "H10", "H11", "H12", "H13", "H14", "H15", "H16", "H17", "H18",
    "D20", "K20", "T20",
    "E22", "J22", "O22", "V22", "Y22", "E24", "J24",
    "D26", "H26", "K26",
    "H28", "H35", "X40", "X41", "B44", "I62", "I63", "AF24",
]


def strip_drawings_from_xlsx(path: Path) -> None:
    with zipfile.ZipFile(path, "r") as zin:
        names = zin.namelist()
        has = any(("drawing" in n) or ("media" in n) for n in names)
        if not has:
            print("sin media/drawing")
            return
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                n = item.filename
                if "drawing" in n or "media" in n or "chart" in n:
                    continue
                data = zin.read(n)
                if n.endswith(".rels"):
                    try:
                        t = data.decode("utf-8")
                        t = re.sub(
                            r"<Relationship[^>]+(?:drawing|image|media)[^/]*/>",
                            "",
                            t,
                            flags=re.I,
                        )
                        t = re.sub(
                            r'<Relationship[^>]+Target="[^"]*(?:drawing|media)[^"]*"[^/]*/>',
                            "",
                            t,
                            flags=re.I,
                        )
                        data = t.encode("utf-8")
                    except Exception as e:
                        print("rels", n, e)
                if "worksheets/sheet" in n and n.endswith(".xml"):
                    try:
                        t = data.decode("utf-8")
                        t = re.sub(r"<drawing[^/]*/>", "", t)
                        t = re.sub(r"<drawing[^>]*>.*?</drawing>", "", t, flags=re.S)
                        data = t.encode("utf-8")
                    except Exception as e:
                        print("sheet", e)
                zout.writestr(item, data)
        path.write_bytes(buf.getvalue())
        print("stripped drawings ->", path.stat().st_size)


def main() -> None:
    src = next(p for p in CANDIDATES if p.exists())
    print("src", src, src.stat().st_size)
    wb = load_workbook(src)
    ws = wb.active
    ws._images = []
    for a in CLEAR:
        try:
            ws[a].value = None
        except Exception:
            pass
    ws["B19"].value = "2. DATOS DEL EVENTO:"
    wb.save(OUT)
    print("openpyxl save", OUT.stat().st_size)
    strip_drawings_from_xlsx(OUT)

    b64 = base64.b64encode(OUT.read_bytes()).decode("ascii")
    (ROOT / "plantilla-b64.js").write_text(
        "window.PLANTILLA_FLASH_B64 = " + repr(b64) + ";\n", encoding="utf-8"
    )
    print("b64", len(b64))

    logo = ROOT / "image1.png"
    if logo.exists():
        lb = base64.b64encode(logo.read_bytes()).decode("ascii")
        (ROOT / "logo-b64.js").write_text(
            "window.LOGO_FLASH_B64 = " + repr(lb) + ";\n", encoding="utf-8"
        )
        print("logo b64", len(lb))


if __name__ == "__main__":
    main()
