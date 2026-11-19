"""Extrae datos de fotos de manifiestos / flash reports (OCR + IA opcional)."""
from __future__ import annotations

import base64
import json
import os
import re
from pathlib import Path
from typing import Any

# ---------------------------------------------------------------------------
# Esquema de datos del FLASH REPORT
# ---------------------------------------------------------------------------
EMPTY_DATA: dict[str, Any] = {
    "nombres": "",
    "edad": "",
    "grado_instruccion": "",
    "tiempo_servicio": "",
    "area": "",
    "ocupacion": "",
    "procedencia": "",
    "experiencia": "",
    "estado_civil": "",
    "dni": "",
    "dia_inicio_guardia": "",
    "equipo_operado": "",
    "fecha_evento": "",
    "hora_evento": "",
    "lugar": "",
    "tipos": [],
    "riesgo": "BAJO",
    "descripcion_evento": "",
    "descripcion_lesion": "",
    "causas": "",
    "requiere_investigacion": True,
    "persona_reporta": "",
    "cargo_reporta": "",
    "supervisor": "",
    "texto_ocr": "",
    "fuente": "",
}

VISION_PROMPT = """Eres un asistente de SSOMA minero. Lee la(s) imagen(es) de un formato de
manifestación de incidentes o notas de accidente de Compañía Minera Casma.

Devuelve SOLO un JSON válido (sin markdown) con estas claves en español:
{
  "nombres": "",
  "edad": "",
  "grado_instruccion": "",
  "tiempo_servicio": "",
  "area": "",
  "ocupacion": "",
  "procedencia": "",
  "experiencia": "",
  "estado_civil": "",
  "dni": "",
  "dia_inicio_guardia": "",
  "equipo_operado": "",
  "fecha_evento": "",
  "hora_evento": "",
  "lugar": "",
  "tipos": ["dano_propiedad"],
  "riesgo": "BAJO",
  "descripcion_evento": "",
  "descripcion_lesion": "",
  "causas": "",
  "requiere_investigacion": true,
  "persona_reporta": "",
  "cargo_reporta": "",
  "supervisor": ""
}

Reglas:
- tipos posibles: dano_propiedad, incidente_ambiental, incidente, incidente_peligroso,
  atencion_medica, accidente_leve, accidente_incapacitante
- riesgo: ALTO, MEDIO o BAJO
- Fechas en formato DD-MM-YYYY; hora HH:MM
- Transcribe la letra manuscrita lo mejor posible
- Si un campo no se lee, deja ""
- descripcion_evento: narra en tercera persona el evento de forma clara y profesional
- descripcion_lesion: daño a persona o propiedad
- causas: causa inmediata / acto o condición subestándar
- persona_reporta: usar supervisor / jefe directo si no hay otro reportero
"""


def _merge(base: dict, extra: dict) -> dict:
    out = dict(base)
    for k, v in extra.items():
        if v is None or v == "" or v == []:
            continue
        out[k] = v
    return out


def _extract_json(text: str) -> dict:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        m = re.search(r"\{[\s\S]*\}", text)
        if m:
            return json.loads(m.group(0))
        raise


def ocr_rapid(image_path: str | Path) -> str:
    from rapidocr_onnxruntime import RapidOCR

    ocr = RapidOCR()
    result, _ = ocr(str(image_path))
    if not result:
        return ""
    return "\n".join(item[1] for item in result if item and item[1])


def ocr_windows(image_path: str | Path) -> str:
    """OCR nativo de Windows (mejor en texto impreso)."""
    try:
        import asyncio
        from winrt.windows.graphics.imaging import BitmapDecoder
        from winrt.windows.media.ocr import OcrEngine
        from winrt.windows.storage import StorageFile, FileAccessMode
        from winrt.windows.globalization import Language
    except ImportError:
        return ""

    async def _run() -> str:
        path = str(Path(image_path).resolve())
        file = await StorageFile.get_file_from_path_async(path)
        stream = await file.open_async(FileAccessMode.READ)
        decoder = await BitmapDecoder.create_async(stream)
        bitmap = await decoder.get_software_bitmap_async()
        engine = None
        for tag in ("es-ES", "es", "es-MX", "en-US"):
            try:
                lang = Language(tag)
                if OcrEngine.is_language_supported(lang):
                    engine = OcrEngine.try_create_from_language(lang)
                    if engine:
                        break
            except Exception:
                pass
        if engine is None:
            engine = OcrEngine.try_create_from_user_profile_languages()
        result = await engine.recognize_async(bitmap)
        return result.text or ""

    try:
        return asyncio.run(_run())
    except Exception:
        return ""


def _norm_date(d: str) -> str:
    d = d.replace("/", "-").replace(".", "-")
    # Corregir OCR: 0g -> 08, o8, etc.
    d = re.sub(r"(\d{1,2})-0[gGq]-(\d{2,4})", r"\1-08-\2", d)
    d = re.sub(r"(\d{1,2})-o(\d)-(\d{2,4})", r"\1-0\2-\3", d, flags=re.I)
    parts = d.split("-")
    if len(parts) == 3 and len(parts[2]) == 2:
        try:
            yy = int(parts[2])
            parts[2] = f"20{yy:02d}" if yy < 70 else f"19{yy:02d}"
        except ValueError:
            pass
    try:
        return "-".join(
            f"{int(p):02d}" if i < 2 else p for i, p in enumerate(parts)
        )
    except ValueError:
        return d


def _after_label(lines: list[str], labels: list[str], max_ahead: int = 4) -> str:
    """Devuelve texto de líneas siguientes a una etiqueta OCR."""
    for i, line in enumerate(lines):
        low = re.sub(r"\s+", "", line.lower())
        for lab in labels:
            lab_c = re.sub(r"\s+", "", lab.lower())
            if lab_c in low or low.startswith(lab_c[:8]):
                buf = []
                for j in range(1, max_ahead + 1):
                    if i + j >= len(lines):
                        break
                    nxt = lines[i + j].strip()
                    if not nxt:
                        continue
                    # Parar si parece otra etiqueta
                    if re.search(
                        r"(edad|grado|tiempo|area|área|ocupaci|proceden|experienc|"
                        r"supervisor|estado|dni|guardia|equipo|lugar|fecha|hora|"
                        r"describa|preguntado|manifest)",
                        nxt,
                        re.I,
                    ) and len(nxt) < 40:
                        break
                    buf.append(nxt)
                return " ".join(buf).strip()
    return ""


def parse_ocr_heuristics(text: str) -> dict[str, Any]:
    """Intenta mapear campos típicos del formato de manifestación."""
    data: dict[str, Any] = {}
    if not text:
        return data

    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    flat = " ".join(text.split())
    lower = flat.lower()

    # --- Por etiquetas de línea (manifiesto CASMA) ---
    nom = _after_label(
        lines,
        ["nombres y apellidos", "nombresyapellidos", "nombres y"],
        max_ahead=5,
    )
    if nom:
        # Limpiar basura OCR
        nom = re.sub(r"\b\d{1,3}\b", " ", nom)
        nom = re.sub(r"\s+", " ", nom).strip(" :.-")
        if len(nom) >= 5:
            data["nombres"] = nom.title()

    dni_line = _after_label(lines, ["dni"], max_ahead=2)
    m = re.search(r"\b(\d{8})\b", dni_line or text)
    if m:
        data["dni"] = m.group(1)

    guard = _after_label(lines, ["dia de inicio", "día de inicio", "inicio de guardia"])
    m = re.search(r"\d{1,2}[-/o0gGq.]\d{1,2}[-/o0gGq.]\d{2,4}", guard or "", re.I)
    if not m:
        m = re.search(r"\d{1,2}[-/o0gGq.]\d{1,2}[-/o0gGq.]\d{2,4}", text, re.I)
    if m:
        # Preferir guardia vs evento por contexto
        pass

    # Fechas DD-MM-YY (OCR a veces escribe 0g)
    dates_raw = re.findall(r"\b(\d{1,2}[-/o0gGq.]\d{1,2}[-/o0gGq.]\d{2,4})\b", text, re.I)
    dates = [_norm_date(d) for d in dates_raw]
    # Evento suele ser la del encabezado "evento del" o la de fecha/hora
    m_ev = re.search(
        r"evento\s*del[:\s]*(\d{1,2}[-/o0gGq.]\d{1,2}[-/o0gGq.]\d{2,4})",
        text,
        re.I,
    )
    if m_ev:
        data["fecha_evento"] = _norm_date(m_ev.group(1))
    elif dates:
        data["fecha_evento"] = dates[0]

    m_g = re.search(
        r"(?:guardia|inicio)[^\n]{0,30}(\d{1,2}[-/o0gGq.]\d{1,2}[-/o0gGq.]\d{2,4})",
        text,
        re.I,
    )
    if m_g:
        data["dia_inicio_guardia"] = _norm_date(m_g.group(1))
    elif len(dates) >= 2:
        data["dia_inicio_guardia"] = dates[1]

    # Hora
    m = re.search(r"\b([01]?\d|2[0-3])[:.]([0-5]\d)\b", text)
    if m:
        data["hora_evento"] = f"{int(m.group(1)):02d}:{m.group(2)}"

    # Placa / equipo (CKE-000, y el OCR que lee el 0 como O)
    m = re.search(r"\b(CK[E0O][- ]?\d{3})\b", text, re.I)
    if m:
        placa = m.group(1).upper().replace(" ", "")
        placa = placa.replace("CK0", "CKE").replace("CKO", "CKE")
        if "-" not in placa:
            placa = placa[:3] + "-" + placa[3:]
        data["equipo_operado"] = f"Camioneta - {placa}"
    elif re.search(r"camioneta", lower):
        m2 = re.search(r"\b([A-Z]{2,3}[- ]?\d{3,4})\b", text, re.I)
        if m2:
            placa = m2.group(1).upper().replace(" ", "-")
            data["equipo_operado"] = f"Camioneta - {placa}"
        else:
            data["equipo_operado"] = "Camioneta"

    # Lugar
    lugar = _after_label(lines, ["lugar del evento", "lugar del", "lugar"])
    if lugar and re.search(r"\d", lugar):
        data["lugar"] = lugar
    else:
        m = re.search(
            r"(cruce|labor|rampa|crucero)[^\n]{0,40}\d{2,4}", text, re.I
        )
        if m:
            data["lugar"] = m.group(0).strip()
        elif re.search(r"\b(3\d{2})\b", text):
            lab = re.search(r"\b(3\d{2})\b", text)
            data["lugar"] = f"Cruce de {lab.group(1)}"

    # Supervisor
    sup = _after_label(
        lines,
        ["supervisor encargado", "supervisor", "jefe directo"],
        max_ahead=3,
    )
    if sup and len(sup) > 3:
        data["supervisor"] = re.sub(r"\s+", " ", sup).title()
        data["persona_reporta"] = data["supervisor"]

    # Edad (26 vs basura 76)
    m2 = re.search(r"edad[:\s]*([1-5]\d)", lower)
    if m2:
        data["edad"] = m2.group(1)

    # Estado civil
    for ec in ("soltero", "casado", "conviviente", "viudo", "divorciado"):
        if ec in lower:
            data["estado_civil"] = ec.capitalize()
            break

    # Área
    for ar in ("mina", "planta", "superficie", "geologia", "mantenimiento"):
        if re.search(rf"\b{ar}\b", lower):
            data["area"] = ar.upper() if ar == "mina" else ar.capitalize()
            break

    # Procedencia
    for p in ("lima", "tingo maría", "tingo maria", "huaraz", "cajamarca", "arequipa", "cusco"):
        if p in lower or p.replace(" ", "") in lower.replace(" ", ""):
            data["procedencia"] = "Lima" if "lima" in p else p.title()
            break
    # OCR Linia / L1ma
    if not data.get("procedencia") and re.search(r"\blin[ia]a\b", lower):
        data["procedencia"] = "Lima"

    # Ocupación
    if re.search(r"conduct|camion", lower):
        data["ocupacion"] = "Conductor de camioneta"

    # Instrucción
    if "secundaria" in lower or "comple" in lower:
        data["grado_instruccion"] = "Secundaria completa"
    elif "tecnica" in lower or "técnica" in lower:
        data["grado_instruccion"] = "Técnica completa"
    elif "universitaria" in lower:
        data["grado_instruccion"] = "Universitaria"

    # Tiempo servicio
    m = re.search(r"(\d+\s*mes(?:es)?)", text, re.I)
    if m:
        data["tiempo_servicio"] = m.group(1)
    else:
        m = re.search(r"(\d+\s*año(?:s)?)", text, re.I)
        if m and "experienc" not in text[max(0, m.start() - 30) : m.start()].lower():
            data["tiempo_servicio"] = m.group(1)

    m = re.search(r"(un\s+año|\d+\s*años?)", text, re.I)
    if m and "experienc" in lower:
        data["experiencia"] = m.group(1).capitalize() if m.group(1).lower().startswith("un") else m.group(1)

    # Tipo por pistas
    if any(w in lower for w in ("espejo", "hastial", "camioneta", "daño", "impacto", "cke", "ck0")):
        data["tipos"] = ["dano_propiedad"]
        data["riesgo"] = data.get("riesgo") or "BAJO"

    return data


def extract_with_gemini(image_paths: list[Path], api_key: str | None = None) -> dict[str, Any]:
    api_key = api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        raise RuntimeError("Falta GEMINI_API_KEY o GOOGLE_API_KEY")

    import google.generativeai as genai
    from PIL import Image

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-2.0-flash")
    parts: list[Any] = [VISION_PROMPT]
    for p in image_paths:
        parts.append(Image.open(p))
    resp = model.generate_content(parts)
    data = _extract_json(resp.text)
    data["fuente"] = "gemini"
    return data


def extract_with_xai(image_paths: list[Path], api_key: str | None = None) -> dict[str, Any]:
    """Visión vía API compatible OpenAI de xAI (Grok)."""
    import requests

    api_key = api_key or os.environ.get("XAI_API_KEY")
    if not api_key:
        raise RuntimeError("Falta XAI_API_KEY")

    content: list[dict] = [{"type": "text", "text": VISION_PROMPT}]
    for p in image_paths:
        b64 = base64.b64encode(Path(p).read_bytes()).decode("ascii")
        mime = "image/jpeg"
        suf = Path(p).suffix.lower()
        if suf == ".png":
            mime = "image/png"
        elif suf == ".webp":
            mime = "image/webp"
        content.append(
            {
                "type": "image_url",
                "image_url": {"url": f"data:{mime};base64,{b64}"},
            }
        )

    r = requests.post(
        "https://api.x.ai/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        json={
            "model": "grok-2-vision-1212",
            "messages": [{"role": "user", "content": content}],
            "temperature": 0.1,
        },
        timeout=120,
    )
    r.raise_for_status()
    text = r.json()["choices"][0]["message"]["content"]
    data = _extract_json(text)
    data["fuente"] = "xai"
    return data


def extract_from_images(
    image_paths: list[str | Path],
    prefer: str = "auto",
) -> dict[str, Any]:
    """
    prefer: auto | gemini | xai | ocr
    auto intenta IA si hay API key; si no, OCR + heurísticas.
    """
    paths = [Path(p) for p in image_paths]
    for p in paths:
        if not p.exists():
            raise FileNotFoundError(p)

    data = dict(EMPTY_DATA)

    # IA
    if prefer in ("auto", "gemini") and (
        os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    ):
        try:
            ai = extract_with_gemini(paths)
            data = _merge(data, ai)
            data["fuente"] = "gemini"
            return data
        except Exception as e:
            print(f"Gemini no disponible ({e}); se usa OCR local.")

    if prefer in ("auto", "xai") and os.environ.get("XAI_API_KEY"):
        try:
            ai = extract_with_xai(paths)
            data = _merge(data, ai)
            data["fuente"] = "xai"
            return data
        except Exception as e:
            print(f"xAI no disponible ({e}); se usa OCR local.")

    # OCR local combinado
    texts = []
    for p in paths:
        t1 = ""
        t2 = ""
        try:
            t1 = ocr_rapid(p)
        except Exception as e:
            print(f"RapidOCR falló en {p.name}: {e}")
        try:
            t2 = ocr_windows(p)
        except Exception as e:
            print(f"Windows OCR falló en {p.name}: {e}")
        texts.append(f"--- {p.name} ---\n{t1}\n{t2}")

    full_text = "\n".join(texts)
    heur = parse_ocr_heuristics(full_text)
    data = _merge(data, heur)
    data["texto_ocr"] = full_text
    data["fuente"] = "ocr"
    return data


def load_json_data(path: str | Path) -> dict[str, Any]:
    with open(path, encoding="utf-8") as f:
        raw = json.load(f)
    return _merge(dict(EMPTY_DATA), raw)
