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

Limitaciones actuales:
- El watermark en PDF hoy hace copia segura + advertencia.
- Efectos de transición PPT son placeholder por limitaciones de `python-pptx`.
