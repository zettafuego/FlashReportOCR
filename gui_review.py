"""Ventana de revisión / corrección de campos antes de generar el Excel."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Any, Callable


FIELDS = [
    ("nombres", "Nombres y apellidos"),
    ("edad", "Edad"),
    ("grado_instruccion", "Grado de instrucción"),
    ("tiempo_servicio", "Tiempo de servicio"),
    ("area", "Área"),
    ("ocupacion", "Ocupación"),
    ("procedencia", "Procedencia"),
    ("experiencia", "Experiencia total"),
    ("estado_civil", "Estado civil"),
    ("dni", "DNI"),
    ("dia_inicio_guardia", "Día inicio de guardia"),
    ("equipo_operado", "Equipo operado"),
    ("fecha_evento", "Fecha del evento"),
    ("hora_evento", "Hora del evento"),
    ("lugar", "Lugar"),
    ("riesgo", "Riesgo (ALTO/MEDIO/BAJO)"),
    ("persona_reporta", "Persona que reporta"),
    ("cargo_reporta", "Cargo"),
    ("supervisor", "Supervisor"),
]

TEXT_FIELDS = [
    ("descripcion_evento", "Descripción del evento"),
    ("descripcion_lesion", "Descripción de la lesión o pérdida"),
    ("causas", "Causas / comentarios"),
]

TIPO_OPTIONS = [
    ("dano_propiedad", "Daño a la propiedad"),
    ("incidente_ambiental", "Incidente ambiental"),
    ("incidente", "Incidente"),
    ("incidente_peligroso", "Incidente peligroso"),
    ("atencion_medica", "Atención médica"),
    ("accidente_leve", "Accidente leve"),
    ("accidente_incapacitante", "Accidente incapacitante"),
]


def review_data(data: dict[str, Any]) -> dict[str, Any] | None:
    """
    Abre GUI modal. Devuelve dict corregido o None si se cancela.
    """
    result: dict[str, Any] | None = None

    root = tk.Tk()
    root.title("Flash Report — Revisar datos extraídos")
    root.geometry("900x720")
    root.minsize(780, 600)

    main = ttk.Frame(root, padding=10)
    main.pack(fill=tk.BOTH, expand=True)

    ttk.Label(
        main,
        text="Revisa y corrige los datos (OCR/IA). Luego genera el Excel con el formato oficial.",
        font=("Segoe UI", 10, "bold"),
    ).pack(anchor="w", pady=(0, 8))

    canvas = tk.Canvas(main, highlightthickness=0)
    scroll = ttk.Scrollbar(main, orient="vertical", command=canvas.yview)
    form = ttk.Frame(canvas)
    form.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
    )
    canvas.create_window((0, 0), window=form, anchor="nw")
    canvas.configure(yscrollcommand=scroll.set)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scroll.pack(side=tk.RIGHT, fill=tk.Y)

    entries: dict[str, tk.Variable | tk.Text] = {}

    # Campos simples
    for i, (key, label) in enumerate(FIELDS):
        ttk.Label(form, text=label).grid(row=i, column=0, sticky="w", padx=4, pady=3)
        var = tk.StringVar(value=str(data.get(key) or ""))
        ent = ttk.Entry(form, textvariable=var, width=70)
        ent.grid(row=i, column=1, sticky="ew", padx=4, pady=3)
        entries[key] = var

    row = len(FIELDS)

    # Tipos
    ttk.Label(form, text="Tipo de evento").grid(row=row, column=0, sticky="nw", padx=4, pady=3)
    tipo_frame = ttk.Frame(form)
    tipo_frame.grid(row=row, column=1, sticky="w", padx=4, pady=3)
    tipo_vars: dict[str, tk.BooleanVar] = {}
    selected = set(data.get("tipos") or [])
    for j, (key, label) in enumerate(TIPO_OPTIONS):
        v = tk.BooleanVar(value=key in selected)
        ttk.Checkbutton(tipo_frame, text=label, variable=v).grid(
            row=j // 2, column=j % 2, sticky="w", padx=6, pady=2
        )
        tipo_vars[key] = v
    row += 1

    # Investigación
    ttk.Label(form, text="Requiere investigación 24h").grid(
        row=row, column=0, sticky="w", padx=4, pady=3
    )
    req_var = tk.BooleanVar(value=bool(data.get("requiere_investigacion", True)))
    ttk.Checkbutton(tipo_frame if False else form, text="Sí", variable=req_var).grid(
        row=row, column=1, sticky="w", padx=4
    )
    row += 1

    # Textos largos
    for key, label in TEXT_FIELDS:
        ttk.Label(form, text=label).grid(row=row, column=0, sticky="nw", padx=4, pady=3)
        txt = tk.Text(form, width=70, height=5, wrap="word", font=("Segoe UI", 10))
        txt.insert("1.0", str(data.get(key) or ""))
        txt.grid(row=row, column=1, sticky="ew", padx=4, pady=3)
        entries[key] = txt
        row += 1

    form.columnconfigure(1, weight=1)

    # OCR crudo (solo lectura)
    if data.get("texto_ocr"):
        ttk.Label(form, text="Texto OCR (referencia)").grid(
            row=row, column=0, sticky="nw", padx=4, pady=3
        )
        ocr_box = tk.Text(form, width=70, height=6, wrap="word", font=("Consolas", 8))
        ocr_box.insert("1.0", str(data.get("texto_ocr"))[:4000])
        ocr_box.configure(state="disabled")
        ocr_box.grid(row=row, column=1, sticky="ew", padx=4, pady=3)

    def collect() -> dict[str, Any]:
        out = dict(data)
        for key, _ in FIELDS:
            out[key] = entries[key].get().strip()
        for key, _ in TEXT_FIELDS:
            out[key] = entries[key].get("1.0", "end").strip()
        out["tipos"] = [k for k, v in tipo_vars.items() if v.get()]
        out["requiere_investigacion"] = req_var.get()
        if out.get("supervisor") and not out.get("persona_reporta"):
            out["persona_reporta"] = out["supervisor"]
        return out

    def on_ok() -> None:
        nonlocal result
        out = collect()
        if not out.get("nombres"):
            messagebox.showwarning("Falta dato", "Indica al menos Nombres y apellidos.")
            return
        result = out
        root.destroy()

    def on_cancel() -> None:
        nonlocal result
        result = None
        root.destroy()

    btns = ttk.Frame(root, padding=10)
    btns.pack(fill=tk.X)
    ttk.Button(btns, text="Cancelar", command=on_cancel).pack(side=tk.RIGHT, padx=4)
    ttk.Button(btns, text="Generar Excel FLASH REPORT", command=on_ok).pack(
        side=tk.RIGHT, padx=4
    )

    root.protocol("WM_DELETE_WINDOW", on_cancel)
    root.mainloop()
    return result
