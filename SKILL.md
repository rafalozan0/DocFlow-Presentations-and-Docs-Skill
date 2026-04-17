---
name: docflow-presentations-and-docs-skill
description: "DocFlow - Presentations and Docs Skill. Agent-agnostic office automation for DOCX/XLSX/PDF/PPTX with bilingual guidance (English/Español)."
version: 1.0.0
author: Tao Jin (shynloc) [original], Rafael Lozano [modifications] & contributors
license: MIT
metadata:
  hermes:
    tags: [office, docx, xlsx, pdf, pptx, automation, presentations, docs]
    runtime: python
    import: "from office_suite import OfficeSuite"
  openclaw:
    requires:
      bins: ["python3"]
    trust: medium
    permissions:
      - read: .
      - write: .
---

# DocFlow - Presentations and Docs Skill

EN: Local-first office automation skill for document and presentation workflows.
ES: Skill local-first para automatización de documentos y presentaciones.

## Scope | Alcance

EN:
- Create: DOCX, XLSX, PDF, PPTX
- Convert formats (LibreOffice-backed)
- Extract document text/data
- Batch conversions and watermark helpers
- Send SMTP email explicitly (with attachments)
- Generate presentations with preflight preferences
- Build charts using:
  - Option A: native `python-pptx`
  - Option B: `matplotlib` image embedding

ES:
- Crear: DOCX, XLSX, PDF, PPTX
- Convertir formatos (con LibreOffice)
- Extraer texto/datos de documentos
- Conversión por lotes y helpers de watermark
- Enviar SMTP explícitamente (con adjuntos)
- Generar presentaciones con preferencias preflight
- Crear gráficos con:
  - Opción A: `python-pptx` nativo
  - Opción B: `matplotlib` como imagen

## Safety model | Modelo de seguridad

EN:
- Local processing by default
- No hidden network calls
- Network only for explicit SMTP send actions
- Never hardcode credentials; use `OFFICE_EMAIL_PASSWORD`

ES:
- Procesamiento local por defecto
- Sin llamadas de red ocultas
- Red solo para envíos SMTP explícitos
- Nunca hardcodear credenciales; usar `OFFICE_EMAIL_PASSWORD`

## Quick start | Inicio rápido

```bash
# Standard
python -m pip install -r requirements.txt
python -m compileall src examples setup.py
python examples/basic_usage.py

# Isolated (uv)
uv run --with python-pptx --with python-docx --with openpyxl --with reportlab --with pypdf2 --with pandas --with pillow --with numpy --with matplotlib python examples/basic_usage.py
```

## Python usage

```python
from office_suite import OfficeSuite
suite = OfficeSuite()

suite.create(
    "word",
    title="Daily Report",
    content="# Summary\nDone.",
    output_path="./output/daily_report.docx",
)
```

## PPTX preflight preferences (strict)

Required keys when `require_preflight=True`:
- `theme`: `midnight-luxe | aurora-glow | obsidian-slate | ivory-bloom | neon-velocity`
- `chart_mode`: `native | matplotlib | auto`
- `use_emojis`: `true | false`
- `tone`: `classic-formal | boardroom | conversational | laid-back`

```python
opts = suite.get_presentation_preflight_prompts()
print(opts)
```

## Known limitations | Limitaciones actuales

- EN: PDF watermarking currently performs a safe pass-through copy + warning.
- ES: El watermark en PDF actualmente hace copia segura + advertencia.
- EN: PPT transition effects are placeholders due to `python-pptx` limits.
- ES: Los efectos de transición PPT son placeholder por limitaciones de `python-pptx`.
