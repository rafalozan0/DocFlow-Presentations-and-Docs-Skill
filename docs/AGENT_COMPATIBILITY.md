# Agent Compatibility | Compatibilidad de agentes

## English

This repository is designed to be truly agent-agnostic:
- Python library can be used directly by any runtime (`from office_suite import OfficeSuite`).
- Includes root `SKILL.md` with neutral, cross-agent instructions.
- Includes optional metadata sections for Hermes and OpenClaw.

Supported usage patterns:
1) Direct Python import (scripts, notebooks, jobs)
2) Skill-driven orchestration in Hermes
3) Skill-driven orchestration in OpenClaw

Runtime notes:
- `pip` workflow is documented for standard environments.
- `uv run --with ...` workflow is documented for isolated ephemeral environments.
- Some features require external binaries (LibreOffice) for conversion.
- The file-based template engine (`slide_template`, `slide_deck_templates`,
  `doc_template`) requires `google-chrome` or `chromium` for PNG/PDF
  rasterization. Override with `OFFICE_CHROME_BIN` if needed.

Template-driven flow for agents:
1. Call `suite.list_templates()` → inspect `slides` and `documents` arrays.
2. Call `suite.get_template_meta(id, category)` → read `slots` and `example`.
3. Construct a `data` dict matching the slot contract.
4. Call `suite.create("slide_template" | "slide_deck_templates" | "doc_template", ...)`.
5. Return the produced file paths in your agent response.

Current implementation limits:
- PDF watermarking currently performs safe pass-through copy + warning.
- PPT transition effects are placeholders due to `python-pptx` limitations.

## Español

Este repositorio está diseñado para ser realmente agnóstico al agente:
- La librería Python se usa directo desde cualquier runtime (`from office_suite import OfficeSuite`).
- Incluye `SKILL.md` raíz con instrucciones neutrales y portables.
- Incluye secciones opcionales de metadata para Hermes y OpenClaw.

Patrones soportados:
1) Importación Python directa (scripts, notebooks, jobs)
2) Orquestación por skill en Hermes
3) Orquestación por skill en OpenClaw

Notas de runtime:
- Se documenta flujo con `pip` para entornos estándar.
- Se documenta flujo con `uv run --with ...` para entornos efímeros y aislados.
- Algunas funciones requieren binarios externos (LibreOffice) para conversión.
- El motor de plantillas por archivo (`slide_template`, `slide_deck_templates`,
  `doc_template`) requiere `google-chrome` o `chromium` para rasterización a
  PNG/PDF. Se puede sobreescribir con la variable `OFFICE_CHROME_BIN`.

Flujo guiado por plantillas para agentes:
1. `suite.list_templates()` → inspecciona los arrays `slides` y `documents`.
2. `suite.get_template_meta(id, category)` → lee `slots` y `example`.
3. Construye un dict `data` que cumpla con el contrato de slots.
4. `suite.create("slide_template" | "slide_deck_templates" | "doc_template", ...)`.
5. Devuelve las rutas de los archivos generados en la respuesta del agente.

Limitaciones actuales:
- El watermark en PDF hoy hace copia segura + advertencia.
- Efectos de transición PPT son placeholder por limitaciones de `python-pptx`.
