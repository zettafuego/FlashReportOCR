# Flash Report OCR — Minera Casma

Lee fotos de manifiestos y genera el **FLASH REPORT** en Excel respetando **colores, bordes, logo y formas** de la plantilla oficial (`FOR.SIG.03`).

## App HTML (recomendada)

```bat
cd "C:\Users\Eduardo\Documents\1. Programación\FlashReportOCR\web"
Abrir_Flash_Report.bat
```

O abre `web\index.html` en el navegador.

1. Sube la foto del manifiesto  
2. **Reconocer foto** (OCR o Gemini si pones API key)  
3. Revisa el formulario  
4. **Descargar Excel FLASH REPORT**  

Archivos: `web\index.html`, `web\app.js`, `web\styles.css`, plantilla embebida en `plantilla-b64.js`.

## CLI Python (alternativa)

### Requisitos

```bat
cd "C:\Users\Eduardo\Documents\1. Programación\FlashReportOCR"
python -m pip install -r requirements.txt
```

### Uso rápido

1. **Arrastra** una foto del manifiesto sobre `Generar_Flash_Report.bat`, **o**:

```bat
python flash_report.py "C:\Users\Eduardo\Downloads\content.jpg"
```

2. Se abre una ventana para **revisar/corregir** los datos.
3. Se genera el Excel en la carpeta `salidas\`.

### Sin ventana (automático)

```bat
python flash_report.py foto.jpg --no-gui -o "C:\Users\Eduardo\Downloads\FLASH_OUT.xlsx"
```

### Desde JSON (datos ya revisados) + insertar la foto en el Excel

```bat
python flash_report.py "C:\Users\Eduardo\Downloads\content.jpg" --json ejemplos\content_jeferson.json --no-gui
```

## Reconocimiento de letra manuscrita

| Motor | Calidad en manuscrito | Cómo activar |
|--------|------------------------|--------------|
| **Gemini** (recomendado) | Alta | `set GEMINI_API_KEY=tu_clave` |
| **xAI Grok Vision** | Alta | `set XAI_API_KEY=tu_clave` |
| **OCR local** (RapidOCR + Windows) | Media/baja | Por defecto sin API |

Sin API key el programa usa OCR local y la **ventana de revisión** para completar lo que no se lea.

```bat
set GEMINI_API_KEY=AIza...
python flash_report.py foto.jpg --prefer gemini
```

## Plantilla

- `plantilla_flash_report.xlsx` — copia limpia del formato oficial (colores naranja `#FFC000`, logo, secciones).
- Regenerar plantilla limpia:

```bat
python crear_plantilla_limpia.py
```

## Campos que rellena

Datos del involucrado, fecha/hora/lugar, tipo de evento (X), evaluación de riesgo, descripción, lesión/pérdida, causas, fotos del evento, persona que reporta.
