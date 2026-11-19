/* Flash Report — Minera Casma (HTML) */
(function () {
  "use strict";

  const ORANGE = "FFFFC000";
  const FIELD_MAP = {
    nombres: "H7",
    edad: "H8",
    grado_instruccion: "H9",
    tiempo_servicio: "H10",
    area: "H11",
    ocupacion: "H12",
    procedencia: "H13",
    experiencia: "H14",
    estado_civil: "H15",
    dni: "H16",
    dia_inicio_guardia: "H17",
    equipo_operado: "H18",
    fecha_evento: "D20",
    hora_evento: "K20",
    lugar: "T20",
    descripcion_evento: "H28",
    descripcion_lesion: "H35",
    causas: "B44",
    persona_reporta: "I62",
    cargo_reporta: "I63",
  };

  const TIPO_CELLS = {
    dano_propiedad: "E22",
    incidente_ambiental: "J22",
    incidente: "O22",
    incidente_peligroso: "V22",
    atencion_medica: "Y22",
    accidente_leve: "E24",
    accidente_incapacitante: "J24",
  };

  const RIESGO_CELLS = { ALTO: "D26", MEDIO: "H26", BAJO: "K26" };
  const PHOTO_SLOTS = [
    { col: 2, row: 50 },
    { col: 7, row: 50 },
    { col: 14, row: 50 },
  ];

  /** @type {File[]} */
  let photoFiles = [];
  /** @type {string[]} data URLs for previews / excel */
  let photoDataUrls = [];

  const $ = (id) => document.getElementById(id);
  const statusEl = () => $("status");

  function setStatus(msg, type) {
    const el = statusEl();
    if (!el) return;
    el.textContent = msg || "";
    el.className = "status" + (type ? " " + type : "");
  }

  function getFormData() {
    const tipos = [...document.querySelectorAll('input[name="tipo"]:checked')].map(
      (el) => el.value
    );
    return {
      nombres: $("nombres").value.trim(),
      edad: $("edad").value.trim(),
      grado_instruccion: $("grado_instruccion").value.trim(),
      tiempo_servicio: $("tiempo_servicio").value.trim(),
      area: $("area").value.trim(),
      ocupacion: $("ocupacion").value.trim(),
      procedencia: $("procedencia").value.trim(),
      experiencia: $("experiencia").value.trim(),
      estado_civil: $("estado_civil").value.trim(),
      dni: $("dni").value.trim(),
      dia_inicio_guardia: $("dia_inicio_guardia").value.trim(),
      equipo_operado: $("equipo_operado").value.trim(),
      fecha_evento: $("fecha_evento").value.trim(),
      hora_evento: $("hora_evento").value.trim(),
      lugar: $("lugar").value.trim(),
      tipos,
      riesgo: ($("riesgo").value || "BAJO").toUpperCase(),
      descripcion_evento: $("descripcion_evento").value.trim(),
      descripcion_lesion: $("descripcion_lesion").value.trim(),
      causas: $("causas").value.trim(),
      requiere_investigacion: $("requiere_investigacion").checked,
      persona_reporta: $("persona_reporta").value.trim(),
      cargo_reporta: $("cargo_reporta").value.trim(),
      supervisor: $("supervisor").value.trim(),
    };
  }

  function setFormData(data) {
    if (!data) return;
    const set = (id, v) => {
      const el = $(id);
      if (el && v != null && v !== "") el.value = v;
    };
    set("nombres", data.nombres);
    set("edad", data.edad);
    set("grado_instruccion", data.grado_instruccion);
    set("tiempo_servicio", data.tiempo_servicio);
    set("area", data.area);
    set("ocupacion", data.ocupacion);
    set("procedencia", data.procedencia);
    set("experiencia", data.experiencia);
    set("estado_civil", data.estado_civil);
    set("dni", data.dni);
    set("dia_inicio_guardia", data.dia_inicio_guardia);
    set("equipo_operado", data.equipo_operado);
    set("fecha_evento", data.fecha_evento);
    set("hora_evento", data.hora_evento);
    set("lugar", data.lugar);
    set("descripcion_evento", data.descripcion_evento);
    set("descripcion_lesion", data.descripcion_lesion);
    set("causas", data.causas);
    set("persona_reporta", data.persona_reporta || data.supervisor);
    set("cargo_reporta", data.cargo_reporta);
    set("supervisor", data.supervisor);
    if (data.riesgo) $("riesgo").value = String(data.riesgo).toUpperCase();
    if (typeof data.requiere_investigacion === "boolean") {
      $("requiere_investigacion").checked = data.requiere_investigacion;
    }
    document.querySelectorAll('input[name="tipo"]').forEach((el) => {
      el.checked = (data.tipos || []).includes(el.value);
    });
  }

  function b64ToArrayBuffer(b64) {
    const bin = atob(b64);
    const len = bin.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
    return bytes.buffer;
  }

  function parseDateExcel(str) {
    if (!str) return null;
    const s = String(str).trim().replace(/\//g, "-");
    const m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{2,4})$/);
    if (!m) return s;
    let y = parseInt(m[3], 10);
    if (y < 100) y += 2000;
    return new Date(y, parseInt(m[2], 10) - 1, parseInt(m[1], 10));
  }

  function parseTimeExcel(str) {
    if (!str) return null;
    const s = String(str).trim().replace(".", ":");
    const m = s.match(/^(\d{1,2}):(\d{2})/);
    if (!m) return s;
    // ExcelJS accepts Date for time cells
    return new Date(1899, 11, 30, parseInt(m[1], 10), parseInt(m[2], 10), 0);
  }

  function setCellValue(ws, addr, value) {
    if (value === null || value === undefined || value === "") return;
    const cell = ws.getCell(addr);
    // Keep existing style; only change value
    cell.value = value;
  }

  function clearMarks(ws) {
    Object.values(TIPO_CELLS).forEach((a) => {
      ws.getCell(a).value = null;
    });
    Object.values(RIESGO_CELLS).forEach((a) => {
      ws.getCell(a).value = null;
    });
    ws.getCell("X40").value = null;
    ws.getCell("X41").value = null;
  }

  function arrayBufferFromDataUrl(dataUrl) {
    // data:[mime];base64,XXXX
    const i = dataUrl.indexOf(",");
    if (i < 0) throw new Error("data URL inválida");
    const meta = dataUrl.slice(0, i);
    const b64 = dataUrl.slice(i + 1);
    const bytes = new Uint8Array(b64ToArrayBuffer(b64));
    let ext = "jpeg";
    if (/png/i.test(meta)) ext = "png";
    else if (/webp/i.test(meta)) ext = "jpeg"; // ExcelJS no siempre acepta webp
    else if (/gif/i.test(meta)) ext = "gif";
    return { buffer: bytes, extension: ext };
  }

  function downloadBlob(blob, filename) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }

  function fileBaseName(data) {
    const nombre = (data.nombres || "sin_nombre").replace(/[<>:"/\\|?*]/g, "").trim();
    const fecha = String(data.fecha_evento || "").replace(/\//g, "-") || "sin-fecha";
    return `FLASH REPORT-${nombre} ${fecha}`;
  }

  async function fillWorkbook(data) {
    if (typeof ExcelJS === "undefined") {
      throw new Error("ExcelJS no cargó. Abre la app con Abrir_Flash_Report.bat (servidor local).");
    }
    if (!window.PLANTILLA_FLASH_B64) {
      throw new Error("Plantilla no cargada (plantilla-b64.js).");
    }

    const wb = new ExcelJS.Workbook();
    try {
      await wb.xlsx.load(b64ToArrayBuffer(window.PLANTILLA_FLASH_B64));
    } catch (e) {
      console.error(e);
      throw new Error(
        "No se pudo leer la plantilla Excel. Regenera plantilla con _build_plantilla.py. Detalle: " +
          (e.message || e)
      );
    }

    const ws = wb.worksheets[0] || wb.getWorksheet(1);
    if (!ws) throw new Error("La plantilla no tiene hojas.");

    const b19 = ws.getCell("B19");
    if (b19.value === 2 || b19.value === "2" || !b19.value) {
      b19.value = "2. DATOS DEL EVENTO:";
    }

    clearMarks(ws);

    Object.entries(FIELD_MAP).forEach(([key, addr]) => {
      let val = data[key];
      if (key === "fecha_evento") val = parseDateExcel(val);
      if (key === "hora_evento") val = parseTimeExcel(val);
      if (key === "dni" && val != null) val = String(val).replace(/\.0$/, "");
      // Fechas/horas como texto si Date falla en algunos Excel
      if (key === "fecha_evento" && val instanceof Date && isNaN(val.getTime())) {
        val = data[key];
      }
      if (key === "hora_evento" && !(val instanceof Date) && data[key]) {
        val = String(data[key]);
      }
      setCellValue(ws, addr, val);
    });

    (data.tipos || []).forEach((t) => {
      const addr = TIPO_CELLS[t];
      if (addr) setCellValue(ws, addr, "X");
    });

    const riesgo = (data.riesgo || "BAJO").toUpperCase();
    if (RIESGO_CELLS[riesgo]) setCellValue(ws, RIESGO_CELLS[riesgo], "X");

    if (data.requiere_investigacion !== false) {
      setCellValue(ws, "X40", "x");
    } else {
      setCellValue(ws, "X41", "x");
    }

    // Logo (se quitó de la plantilla para que ExcelJS no falle)
    try {
      if (window.LOGO_FLASH_B64) {
        const logoBytes = new Uint8Array(b64ToArrayBuffer(window.LOGO_FLASH_B64));
        const logoId = wb.addImage({ buffer: logoBytes, extension: "png" });
        ws.addImage(logoId, {
          tl: { col: 1, row: 0 },
          ext: { width: 95, height: 72 },
        });
      }
    } catch (e) {
      console.warn("Logo no insertado:", e);
    }

    // Fotos del evento
    for (let i = 0; i < Math.min(photoDataUrls.length, 3); i++) {
      try {
        const { buffer, extension } = arrayBufferFromDataUrl(photoDataUrls[i]);
        const imgId = wb.addImage({ buffer, extension });
        const slot = PHOTO_SLOTS[i];
        ws.addImage(imgId, {
          tl: { col: slot.col - 1, row: slot.row - 1 },
          ext: { width: 260, height: 195 },
        });
      } catch (e) {
        console.warn("No se pudo insertar foto", i, e);
      }
    }

    return wb;
  }

  async function generateExcel(data) {
    const wb = await fillWorkbook(data);
    const buf = await wb.xlsx.writeBuffer();
    const filename = fileBaseName(data) + ".xlsx";
    downloadBlob(
      new Blob([buf], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
      filename
    );
    return filename;
  }

  function esc(s) {
    return String(s || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function tipoLabels(tipos) {
    const map = {
      dano_propiedad: "Daño a la propiedad",
      incidente_ambiental: "Incidente ambiental",
      incidente: "Incidente",
      incidente_peligroso: "Incidente peligroso",
      atencion_medica: "Atención médica",
      accidente_leve: "Accidente leve",
      accidente_incapacitante: "Accidente incapacitante",
    };
    return (tipos || []).map((t) => map[t] || t).join(", ") || "—";
  }

  function buildPdfHtml(data) {
    const photos = photoDataUrls
      .slice(0, 3)
      .map(
        (u) =>
          `<img src="${u}" style="width:32%;max-height:160px;object-fit:cover;border:1px solid #333;margin:2px" />`
      )
      .join("");
    const logo = window.LOGO_FLASH_B64
      ? `<img src="data:image/png;base64,${window.LOGO_FLASH_B64}" style="height:54px" />`
      : "";

    return `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"/>
<title>${esc(fileBaseName(data))}</title>
<style>
  @page { size: A4; margin: 12mm; }
  body { font-family: Arial, Helvetica, sans-serif; font-size: 11pt; color: #111; }
  .hdr { display:flex; justify-content:space-between; align-items:center;
    background:#ffc000; border:2px solid #000; padding:8px 10px; margin-bottom:10px; }
  .hdr h1 { font-size:16pt; margin:0; text-align:center; flex:1; }
  .meta { font-size:9pt; min-width:120px; }
  .sec { background:#ffc000; border:1px solid #000; font-weight:bold; padding:5px 8px; margin-top:10px; }
  table.grid { width:100%; border-collapse:collapse; margin-top:0; }
  table.grid td { border:1px solid #000; padding:5px 7px; vertical-align:top; }
  table.grid td.lab { width:32%; font-weight:bold; background:#f5f5f5; }
  .box { border:1px solid #000; padding:8px; min-height:48px; white-space:pre-wrap; margin-top:0; }
  .marks { margin:6px 0; }
  .marks span { display:inline-block; margin-right:12px; }
  .x { font-weight:bold; border:1px solid #000; padding:0 5px; }
  .photos { margin-top:8px; }
  .foot { margin-top:16px; display:flex; justify-content:space-between; }
  @media print { .noprint { display:none !important; } body { -webkit-print-color-adjust: exact; print-color-adjust: exact; } }
</style></head><body>
  <div class="noprint" style="background:#222;color:#ffc000;padding:10px;margin-bottom:12px;font-family:sans-serif">
    <strong>Vista PDF / impresión</strong> —
    En el diálogo elige <em>Guardar como PDF</em> o impresora.
    <button onclick="window.print()" style="margin-left:12px;padding:6px 12px;font-weight:bold;cursor:pointer">Imprimir / Guardar PDF</button>
  </div>
  <div class="hdr">
    <div>${logo}</div>
    <h1>FORMATO DE REPORTE DE INCIDENTES Y ACCIDENTES<br/>FLASH REPORT</h1>
    <div class="meta"><strong>CIA. MINERA CASMA</strong><br/>Código: FOR.SIG.03<br/>Versión: 00</div>
  </div>

  <div class="sec">1. DATOS DEL INVOLUCRADO Y SU SUPERVISOR</div>
  <table class="grid">
    <tr><td class="lab">Nombres y apellidos</td><td>${esc(data.nombres)}</td></tr>
    <tr><td class="lab">Edad</td><td>${esc(data.edad)}</td></tr>
    <tr><td class="lab">Grado de instrucción</td><td>${esc(data.grado_instruccion)}</td></tr>
    <tr><td class="lab">Tiempo de servicio</td><td>${esc(data.tiempo_servicio)}</td></tr>
    <tr><td class="lab">Área</td><td>${esc(data.area)}</td></tr>
    <tr><td class="lab">Ocupación</td><td>${esc(data.ocupacion)}</td></tr>
    <tr><td class="lab">Procedencia</td><td>${esc(data.procedencia)}</td></tr>
    <tr><td class="lab">Experiencia total</td><td>${esc(data.experiencia)}</td></tr>
    <tr><td class="lab">Estado civil</td><td>${esc(data.estado_civil)}</td></tr>
    <tr><td class="lab">DNI</td><td>${esc(data.dni)}</td></tr>
    <tr><td class="lab">Día inicio de guardia</td><td>${esc(data.dia_inicio_guardia)}</td></tr>
    <tr><td class="lab">Equipo operado</td><td>${esc(data.equipo_operado)}</td></tr>
  </table>

  <div class="sec">2. DATOS DEL EVENTO</div>
  <table class="grid">
    <tr>
      <td class="lab">Fecha</td><td>${esc(data.fecha_evento)}</td>
      <td class="lab">Hora</td><td>${esc(data.hora_evento)}</td>
    </tr>
    <tr><td class="lab">Lugar</td><td colspan="3">${esc(data.lugar)}</td></tr>
    <tr><td class="lab">Tipo</td><td colspan="3">${esc(tipoLabels(data.tipos))}
      &nbsp; Riesgo: <strong>${esc(data.riesgo || "BAJO")}</strong>
      &nbsp; Investigación 24h: <strong>${data.requiere_investigacion !== false ? "Requiere" : "No requiere"}</strong>
    </td></tr>
  </table>

  <div class="sec">DESCRIPCIÓN DEL EVENTO</div>
  <div class="box">${esc(data.descripcion_evento)}</div>

  <div class="sec">DESCRIPCIÓN DE LA LESIÓN O PÉRDIDA</div>
  <div class="box">${esc(data.descripcion_lesion)}</div>

  <div class="sec">Comentarios, causas y sugerencias (información preliminar)</div>
  <div class="box">${esc(data.causas)}</div>

  <div class="sec">FOTOGRAFÍAS DEL ACCIDENTE</div>
  <div class="photos">${photos || "<em>Sin fotos adjuntas</em>"}</div>

  <div class="foot">
    <div><strong>Persona que reporta:</strong> ${esc(data.persona_reporta || data.supervisor)}<br/>
    <strong>Cargo:</strong> ${esc(data.cargo_reporta)}</div>
  </div>
  <script>setTimeout(function(){ try{ window.print(); }catch(e){} }, 400);<\/script>
</body></html>`;
  }

  function generatePdf(data) {
    const html = buildPdfHtml(data);
    const w = window.open("", "_blank");
    if (!w) {
      throw new Error(
        "El navegador bloqueó la ventana del PDF. Permite ventanas emergentes para esta página."
      );
    }
    w.document.open();
    w.document.write(html);
    w.document.close();
    return fileBaseName(data) + ".pdf";
  }

  /* ---------- OCR local (heurísticas sobre texto) ---------- */
  function normDate(d) {
    let s = d.replace(/\//g, "-").replace(/\./g, "-");
    s = s.replace(/(\d{1,2})-0[gGq]-(\d{2,4})/, "$1-08-$2");
    const p = s.split("-");
    if (p.length === 3 && p[2].length === 2) {
      const yy = parseInt(p[2], 10);
      p[2] = String(yy < 70 ? 2000 + yy : 1900 + yy);
    }
    try {
      return [
        String(parseInt(p[0], 10)).padStart(2, "0"),
        String(parseInt(p[1], 10)).padStart(2, "0"),
        p[2],
      ].join("-");
    } catch {
      return s;
    }
  }

  function parseOcrText(text) {
    const data = {};
    if (!text) return data;
    const lower = text.toLowerCase();

    const dni = text.match(/\b(\d{8})\b/);
    if (dni) data.dni = dni[1];

    const dates = [...text.matchAll(/\b(\d{1,2}[-/o0gGq.]\d{1,2}[-/o0gGq.]\d{2,4})\b/gi)].map(
      (m) => normDate(m[1])
    );
    const ev = text.match(/evento\s*del[:\s]*(\d{1,2}[-/o0gGq.]\d{1,2}[-/o0gGq.]\d{2,4})/i);
    if (ev) data.fecha_evento = normDate(ev[1]);
    else if (dates[0]) data.fecha_evento = dates[0];

    const gu = text.match(/(?:guardia|inicio)[^\n]{0,30}(\d{1,2}[-/o0gGq.]\d{1,2}[-/o0gGq.]\d{2,4})/i);
    if (gu) data.dia_inicio_guardia = normDate(gu[1]);
    else if (dates[1]) data.dia_inicio_guardia = dates[1];

    const hora = text.match(/\b([01]?\d|2[0-3])[:.]([0-5]\d)\b/);
    if (hora) data.hora_evento = `${String(hora[1]).padStart(2, "0")}:${hora[2]}`;

    let placa = text.match(/\b(CK[E0O][- ]?\d{3})\b/i);
    if (placa) {
      let p = placa[1].toUpperCase().replace(/\s/g, "").replace("CK0", "CKE").replace("CKO", "CKE");
      if (!p.includes("-")) p = p.slice(0, 3) + "-" + p.slice(3);
      data.equipo_operado = "Camioneta - " + p;
    } else if (/camioneta/i.test(text)) {
      data.equipo_operado = "Camioneta";
    }

    if (/\b(3\d{2})\b/.test(text)) {
      const lab = text.match(/\b(3\d{2})\b/);
      data.lugar = "Cruce de " + lab[1];
    }

    if (/soltero/i.test(text)) data.estado_civil = "Soltero";
    if (/\bmina\b/i.test(text)) data.area = "MINA";
    if (/lima|linia/i.test(text)) data.procedencia = "Lima";
    if (/conductor|camion/i.test(text)) data.ocupacion = "Conductor de camioneta";
    if (/secundaria|comple/i.test(text)) data.grado_instruccion = "Secundaria completa";

    const mes = text.match(/(\d+\s*mes(?:es)?)/i);
    if (mes) data.tiempo_servicio = mes[1];

    if (/espejo|hastial|camioneta|cke|ck0|impacto/i.test(text)) {
      data.tipos = ["dano_propiedad"];
      data.riesgo = "BAJO";
    }

    // Nombre: líneas tras "Nombres"
    const lines = text.split(/\n+/).map((l) => l.trim()).filter(Boolean);
    for (let i = 0; i < lines.length; i++) {
      if (/nombres/i.test(lines[i]) && /apellido/i.test(lines[i])) {
        const parts = [];
        for (let j = 1; j <= 4 && i + j < lines.length; j++) {
          const l = lines[i + j];
          if (/edad|grado|tiempo|área|area|ocupaci/i.test(l) && l.length < 40) break;
          if (/^[A-Za-zÁÉÍÓÚáéíóúñÑ.\s-]{3,}$/.test(l) || /[A-Za-z]{3,}/.test(l)) {
            parts.push(l.replace(/\d+/g, "").trim());
          }
        }
        const nom = parts.join(" ").replace(/\s+/g, " ").trim();
        if (nom.length >= 5) data.nombres = nom;
        break;
      }
    }

    return data;
  }

  async function runTesseract(file) {
    if (typeof Tesseract === "undefined") {
      throw new Error("Tesseract.js no está disponible (necesitas internet la primera vez)");
    }
    setStatus("OCR en curso… esto puede tardar 30–60 s", "info");
    const result = await Tesseract.recognize(file, "spa+eng", {
      logger: (m) => {
        if (m.status === "recognizing text" && m.progress != null) {
          setStatus(`OCR: ${Math.round(m.progress * 100)}%`, "info");
        }
      },
    });
    return result.data.text || "";
  }

  function arrayBufferToBase64(buf) {
    const bytes = new Uint8Array(buf);
    let binary = "";
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return btoa(binary);
  }

  const VISION_PROMPT = `Eres asistente SSOMA de minera. Lee el manifiesto/formulario de incidentes manuscrito.
Devuelve SOLO JSON válido (sin markdown) con:
nombres, edad, grado_instruccion, tiempo_servicio, area, ocupacion, procedencia,
experiencia, estado_civil, dni, dia_inicio_guardia, equipo_operado, fecha_evento,
hora_evento, lugar, tipos (array: dano_propiedad|incidente_ambiental|incidente|incidente_peligroso|atencion_medica|accidente_leve|accidente_incapacitante),
riesgo (ALTO|MEDIO|BAJO), descripcion_evento, descripcion_lesion, causas,
requiere_investigacion (bool), persona_reporta, cargo_reporta, supervisor.
Fechas DD-MM-YYYY, hora HH:MM. Transcribe manuscrito lo mejor posible. Campos no legibles: "".`;

  function parseModelJson(text) {
    if (!text) throw new Error("La IA no devolvió texto");
    text = String(text)
      .replace(/^```(?:json)?\s*/i, "")
      .replace(/\s*```$/i, "")
      .trim();
    const m = text.match(/\{[\s\S]*\}/);
    if (!m) throw new Error("La IA no devolvió JSON. Respuesta: " + text.slice(0, 180));
    return JSON.parse(m[0]);
  }

  async function filesToInlineGemini(files) {
    const inline = [];
    for (const f of files) {
      const buf = await f.arrayBuffer();
      const b64 = arrayBufferToBase64(buf);
      const mime = f.type || "image/jpeg";
      // camelCase y snake_case (según versión de API)
      inline.push({
        inline_data: { mime_type: mime, data: b64 },
      });
    }
    return inline;
  }

  async function filesToOpenAIImages(files) {
    const parts = [{ type: "text", text: VISION_PROMPT }];
    for (const f of files) {
      const buf = await f.arrayBuffer();
      const b64 = arrayBufferToBase64(buf);
      const mime = f.type || "image/jpeg";
      parts.push({
        type: "image_url",
        image_url: { url: `data:${mime};base64,${b64}` },
      });
    }
    return parts;
  }

  /** Modelos Gemini actuales (2.0-flash está apagado desde 2026-06). */
  const GEMINI_MODELS = [
    "gemini-2.5-flash",
    "gemini-2.5-flash-lite",
    "gemini-2.5-pro",
    "gemini-flash-latest",
    "gemini-1.5-flash",
    "gemini-1.5-flash-8b",
    "gemini-1.5-pro",
  ];

  async function runGemini(files, apiKey) {
    const key = (apiKey || "").trim();
    if (!key) throw new Error("Falta la API key de Gemini");

    // Claves típicas de AI Studio: AIza...  |  Las AQ. suelen ser de otro producto y fallan aquí
    if (key.startsWith("AQ.")) {
      throw new Error(
        "Esa clave empieza por AQ. y NO es una API key de Gemini AI Studio. " +
          "Ve a https://aistudio.google.com/apikey y crea una clave (suele empezar por AIza…). " +
          "O usa el proveedor xAI / OpenAI / OpenRouter abajo."
      );
    }

    const inline = await filesToInlineGemini(files);
    const bodyBase = {
      contents: [{ role: "user", parts: [{ text: VISION_PROMPT }, ...inline] }],
      generationConfig: { temperature: 0.1 },
    };

    const errors = [];
    for (const model of GEMINI_MODELS) {
      const url =
        `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent` +
        `?key=${encodeURIComponent(key)}`;
      try {
        const res = await fetch(url, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(bodyBase),
        });
        const raw = await res.text();
        if (!res.ok) {
          errors.push(`${model}: HTTP ${res.status} ${raw.slice(0, 120)}`);
          // 404 modelo → probar siguiente; 400/403 key → no seguir todos si es auth
          if (res.status === 400 || res.status === 401 || res.status === 403) {
            let hint = raw.slice(0, 220);
            try {
              const j = JSON.parse(raw);
              hint = j?.error?.message || hint;
            } catch (_) {}
            throw new Error(
              `Gemini rechazó la clave o el proyecto (${res.status}): ${hint}. ` +
                `Crea una clave nueva en https://aistudio.google.com/apikey (sin restricciones de app, o permite este sitio).`
            );
          }
          continue;
        }
        const json = JSON.parse(raw);
        if (json?.error) {
          errors.push(`${model}: ${json.error.message || JSON.stringify(json.error)}`);
          continue;
        }
        const text =
          json?.candidates?.[0]?.content?.parts?.map((p) => p.text).join("") || "";
        if (!text && json?.promptFeedback) {
          errors.push(`${model}: bloqueado por safety ${JSON.stringify(json.promptFeedback)}`);
          continue;
        }
        const data = parseModelJson(text);
        data._provider = "gemini/" + model;
        return data;
      } catch (e) {
        if (String(e.message || e).includes("rechazó la clave")) throw e;
        errors.push(`${model}: ${e.message || e}`);
      }
    }
    throw new Error(
      "Gemini falló con todos los modelos. Detalle:\n" + errors.slice(0, 4).join("\n")
    );
  }

  /** xAI Grok Vision — clave en https://console.x.ai/ */
  async function runXai(files, apiKey) {
    const key = (apiKey || "").trim();
    if (!key) throw new Error("Falta XAI_API_KEY");
    const content = await filesToOpenAIImages(files);
    const models = ["grok-2-vision-1212", "grok-4", "grok-2-vision"];
    let lastErr = "";
    for (const model of models) {
      const res = await fetch("https://api.x.ai/v1/chat/completions", {
        method: "POST",
        headers: {
          Authorization: "Bearer " + key,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          model,
          messages: [{ role: "user", content }],
          temperature: 0.1,
        }),
      });
      const raw = await res.text();
      if (!res.ok) {
        lastErr = `${model}: ${res.status} ${raw.slice(0, 180)}`;
        continue;
      }
      const json = JSON.parse(raw);
      const text = json?.choices?.[0]?.message?.content || "";
      const data = parseModelJson(text);
      data._provider = "xai/" + model;
      return data;
    }
    throw new Error("xAI falló: " + lastErr);
  }

  /** OpenAI GPT-4o visión — https://platform.openai.com/api-keys */
  async function runOpenAI(files, apiKey) {
    const key = (apiKey || "").trim();
    if (!key) throw new Error("Falta OPENAI_API_KEY");
    const content = await filesToOpenAIImages(files);
    const models = ["gpt-4o", "gpt-4o-mini", "gpt-4.1-mini", "gpt-4.1"];
    let lastErr = "";
    for (const model of models) {
      const res = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: {
          Authorization: "Bearer " + key,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          model,
          messages: [{ role: "user", content }],
          temperature: 0.1,
        }),
      });
      const raw = await res.text();
      if (!res.ok) {
        lastErr = `${model}: ${res.status} ${raw.slice(0, 180)}`;
        continue;
      }
      const json = JSON.parse(raw);
      const text = json?.choices?.[0]?.message?.content || "";
      const data = parseModelJson(text);
      data._provider = "openai/" + model;
      return data;
    }
    throw new Error("OpenAI falló: " + lastErr);
  }

  /**
   * OpenRouter — una sola clave para Gemini/Claude/GPT/etc.
   * https://openrouter.ai/keys
   */
  async function runOpenRouter(files, apiKey) {
    const key = (apiKey || "").trim();
    if (!key) throw new Error("Falta OpenRouter API key");
    const content = await filesToOpenAIImages(files);
    const models = [
      "google/gemini-2.5-flash",
      "google/gemini-2.5-flash-lite",
      "openai/gpt-4o-mini",
      "openai/gpt-4o",
      "anthropic/claude-3.5-sonnet",
      "google/gemini-2.0-flash-001",
    ];
    let lastErr = "";
    for (const model of models) {
      const res = await fetch("https://openrouter.ai/api/v1/chat/completions", {
        method: "POST",
        headers: {
          Authorization: "Bearer " + key,
          "Content-Type": "application/json",
          "HTTP-Referer": location.origin || "http://127.0.0.1",
          "X-Title": "Flash Report Casma",
        },
        body: JSON.stringify({
          model,
          messages: [{ role: "user", content }],
          temperature: 0.1,
        }),
      });
      const raw = await res.text();
      if (!res.ok) {
        lastErr = `${model}: ${res.status} ${raw.slice(0, 180)}`;
        continue;
      }
      const json = JSON.parse(raw);
      const text = json?.choices?.[0]?.message?.content || "";
      const data = parseModelJson(text);
      data._provider = "openrouter/" + model;
      return data;
    }
    throw new Error("OpenRouter falló: " + lastErr);
  }

  // OpenRouter por defecto (clave solo en este navegador vía localStorage; no la publiques)
  const DEFAULT_AI_PROVIDER = "openrouter";
  const DEFAULT_OPENROUTER_KEY = "";

  async function runVisionAI(files) {
    const provider =
      ($("aiProvider") && $("aiProvider").value) || DEFAULT_AI_PROVIDER;
    let apiKey = ($("geminiKey") && $("geminiKey").value.trim()) || "";
    if (!apiKey) {
      apiKey =
        localStorage.getItem("flash_ai_key") ||
        (provider === "openrouter" ? DEFAULT_OPENROUTER_KEY : "");
    }
    if (!apiKey) {
      throw new Error(
        "Pega una API key o usa OCR local (deja la clave vacía y vuelve a intentar)."
      );
    }
    localStorage.setItem("flash_ai_provider", provider);
    localStorage.setItem("flash_ai_key", apiKey);

    if (provider === "gemini") return runGemini(files, apiKey);
    if (provider === "xai") return runXai(files, apiKey);
    if (provider === "openai") return runOpenAI(files, apiKey);
    if (provider === "openrouter") return runOpenRouter(files, apiKey);
    throw new Error("Proveedor desconocido: " + provider);
  }

  function fileToDataUrl(file) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(r.result);
      r.onerror = reject;
      r.readAsDataURL(file);
    });
  }

  async function onPhotosSelected(files) {
    const list = [...files].filter((f) => f.type.startsWith("image/"));
    photoFiles = list.slice(0, 6);
    photoDataUrls = [];
    const box = $("preview");
    box.innerHTML = "";
    for (const f of photoFiles) {
      const url = await fileToDataUrl(f);
      photoDataUrls.push(url);
      const img = document.createElement("img");
      img.src = url;
      img.alt = f.name;
      img.title = f.name;
      box.appendChild(img);
    }
    setStatus(
      photoFiles.length
        ? `${photoFiles.length} foto(s) cargada(s). Puedes reconocer o completar el formulario.`
        : "Sin fotos",
      "ok"
    );
  }

  function loadExample() {
    setFormData({
      nombres: "Trabajador de ejemplo",
      edad: "30",
      grado_instruccion: "Secundaria completa",
      tiempo_servicio: "1 mes",
      area: "MINA",
      ocupacion: "Conductor",
      procedencia: "Lima",
      experiencia: "Un año",
      estado_civil: "Soltero",
      dni: "00000000",
      dia_inicio_guardia: "01-01-2026",
      equipo_operado: "Camioneta - ABC-123",
      fecha_evento: "02-01-2026",
      hora_evento: "09:00",
      lugar: "Labor de ejemplo",
      tipos: ["dano_propiedad"],
      riesgo: "BAJO",
      descripcion_evento:
        "Ejemplo de práctica. Al retroceder, el espejo de la camioneta tocó el hastial. Sin personas reales en este texto.",
      descripcion_lesion: "Daño en el espejo. Sin lesiones.",
      causas: "Ejemplo: la maniobra de retroceso no guardó distancia al hastial.",
      requiere_investigacion: false,
      persona_reporta: "Supervisor de ejemplo",
      cargo_reporta: "JEFE DE GUARDIA",
      supervisor: "Supervisor de ejemplo",
    });
    setStatus("Ejemplo de práctica cargado. Genera el Excel o sube tu foto.", "ok");
  }

  function wire() {
    const drop = $("dropzone");
    const input = $("fileInput");

    drop.addEventListener("click", () => input.click());
    drop.addEventListener("dragover", (e) => {
      e.preventDefault();
      drop.classList.add("drag");
    });
    drop.addEventListener("dragleave", () => drop.classList.remove("drag"));
    drop.addEventListener("drop", (e) => {
      e.preventDefault();
      drop.classList.remove("drag");
      if (e.dataTransfer?.files?.length) onPhotosSelected(e.dataTransfer.files);
    });
    input.addEventListener("change", () => {
      if (input.files?.length) onPhotosSelected(input.files);
    });

    $("btnReconocer").addEventListener("click", async () => {
      if (!photoFiles.length) {
        setStatus("Primero sube una foto del manifiesto.", "err");
        return;
      }
      const apiKey =
        ($("geminiKey") && $("geminiKey").value.trim()) ||
        localStorage.getItem("flash_ai_key") ||
        DEFAULT_OPENROUTER_KEY;
      const forceOcr = $("forceOcr") && $("forceOcr").checked;

      try {
        $("btnReconocer").disabled = true;
        if (apiKey && !forceOcr) {
          const prov =
            ($("aiProvider") && $("aiProvider").value) || DEFAULT_AI_PROVIDER;
          setStatus("Reconociendo con " + prov + " (visión / manuscrito)…", "info");
          const data = await runVisionAI(photoFiles.slice(0, 3));
          setFormData(data);
          $("ocrText").value = JSON.stringify(data, null, 2);
          setStatus(
            "Datos rellenados con " +
              (data._provider || prov) +
              ". Revisa y genera Excel/PDF.",
            "ok"
          );
        } else {
          const text = await runTesseract(photoFiles[0]);
          $("ocrText").value = text;
          const data = parseOcrText(text);
          setFormData(data);
          setStatus(
            "OCR local listo (manuscrito parcial). Completa/corrige el formulario y genera.",
            "ok"
          );
        }
      } catch (e) {
        console.error(e);
        setStatus(String(e.message || e), "err");
      } finally {
        $("btnReconocer").disabled = false;
      }
    });

    $("btnEjemplo").addEventListener("click", loadExample);

    function collectReadyData() {
      const data = getFormData();
      if (!data.nombres) {
        setStatus("Indica al menos Nombres y apellidos.", "err");
        $("nombres").focus();
        return null;
      }
      if (data.supervisor && !data.persona_reporta) {
        data.persona_reporta = data.supervisor;
      }
      return data;
    }

    $("btnGenerar").addEventListener("click", async () => {
      const data = collectReadyData();
      if (!data) return;
      try {
        $("btnGenerar").disabled = true;
        setStatus("Generando Excel FLASH REPORT…", "info");
        const name = await generateExcel(data);
        setStatus("✓ Excel descargado: " + name, "ok");
      } catch (e) {
        console.error(e);
        setStatus("Error al generar Excel: " + (e.message || e), "err");
      } finally {
        $("btnGenerar").disabled = false;
      }
    });

    const btnPdf = $("btnPdf");
    if (btnPdf) {
      btnPdf.addEventListener("click", () => {
        const data = collectReadyData();
        if (!data) return;
        try {
          setStatus("Abriendo vista PDF / impresión…", "info");
          generatePdf(data);
          setStatus(
            "✓ Vista PDF abierta. En el diálogo elige «Guardar como PDF».",
            "ok"
          );
        } catch (e) {
          console.error(e);
          setStatus("Error PDF: " + (e.message || e), "err");
        }
      });
    }

    // Diagnóstico al cargar
    if (typeof ExcelJS === "undefined") {
      setStatus(
        "Error: ExcelJS no cargó. Usa Abrir_Flash_Report.bat (no abras index.html a ciegas desde otra carpeta).",
        "err"
      );
    } else if (!window.PLANTILLA_FLASH_B64) {
      setStatus("Error: falta plantilla-b64.js", "err");
    } else {
      setStatus(
        "Listo. Sube una foto o carga el ejemplo → genera Excel o PDF.",
        "info"
      );
    }

    $("btnLimpiar").addEventListener("click", () => {
      document.querySelectorAll("input[type=text], input[type=time], textarea").forEach((el) => {
        if (el.id === "geminiKey") return;
        el.value = "";
      });
      document.querySelectorAll('input[name="tipo"]').forEach((el) => (el.checked = false));
      $("riesgo").value = "BAJO";
      $("requiere_investigacion").checked = true;
      $("ocrText").value = "";
      photoFiles = [];
      photoDataUrls = [];
      $("preview").innerHTML = "";
      $("fileInput").value = "";
      setStatus("Formulario limpio.", "info");
    });

    // Restore AI settings (OpenRouter por defecto)
    if ($("aiProvider")) {
      const savedProv =
        localStorage.getItem("flash_ai_provider") || DEFAULT_AI_PROVIDER;
      $("aiProvider").value = savedProv;
    }
    if ($("geminiKey")) {
      const savedKey =
        localStorage.getItem("flash_ai_key") ||
        localStorage.getItem("gemini_key") ||
        DEFAULT_OPENROUTER_KEY;
      $("geminiKey").value = savedKey;
      localStorage.setItem("flash_ai_provider", $("aiProvider")?.value || DEFAULT_AI_PROVIDER);
      localStorage.setItem("flash_ai_key", savedKey);
    }

    if ($("aiProvider")) {
      $("aiProvider").addEventListener("change", () => {
        localStorage.setItem("flash_ai_provider", $("aiProvider").value);
        updateProviderHint();
      });
      updateProviderHint();
    }
  }

  function updateProviderHint() {
    const el = $("providerHint");
    if (!el || !$("aiProvider")) return;
    const p = $("aiProvider").value;
    const hints = {
      gemini:
        "Clave en aistudio.google.com/apikey — debe empezar por AIza… (no AQ.). Modelos 2.0-flash ya no existen; la app usa 2.5-flash.",
      xai: "Clave en console.x.ai — Grok Vision. Formato: xai-…",
      openai: "Clave en platform.openai.com/api-keys — GPT-4o con visión. Formato: sk-…",
      openrouter:
        "Clave en openrouter.ai/keys — una sola key para Gemini/GPT/Claude. Ideal si Gemini falla.",
    };
    el.textContent = hints[p] || "";
  }

  document.addEventListener("DOMContentLoaded", wire);
})();
