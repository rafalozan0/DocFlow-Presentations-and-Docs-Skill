<div align="center">

# DocFlow - Presentations and Docs Skill

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![Python >=3.8](https://img.shields.io/badge/python-%3E%3D3.8-blue.svg)](https://www.python.org/downloads/)
[![Runtime](https://img.shields.io/badge/runtimes-Hermes%20%7C%20OpenClaw-7c3aed)](#agent-bootstrap-hermes--openclaw)
[![Mode](https://img.shields.io/badge/mode-agent--only-111827)](#agent-bootstrap-hermes--openclaw)

**Agent-first skill for generating DOCX/XLSX/PDF/PPTX deliverables with strict preflight and dual chart modes.**

</div>

## 🎉 News

- **2026-04-17**
  - README simplified to **agent-only** flow for Hermes/OpenClaw.
  - Human-oriented sections removed to avoid ambiguity.
  - Added explicit bootstrap, execution contract, and preflight policy.
  - Showcase screenshots and test assets integrated.
  - Release line remains **v1.0.0**.

## Agent context (Hermes / OpenClaw)

This repository is intended for autonomous agent runtimes:
- Hermes Agent (Nous Research): terminal-native autonomous agent framework.
- OpenClaw (`openclaw/openclaw`): personal AI agent runtime and skill ecosystem.

This README is intentionally written for agent orchestration, not end-user tutorials.

## Agent bootstrap (Hermes / OpenClaw)

Use exactly this block:

```bash
git clone https://github.com/rafalozan0/DocFlow-Presentations-and-Docs-Skill.git
cd DocFlow-Presentations-and-Docs-Skill
python -m compileall src examples setup.py
python -m pip install -r requirements.txt || true
```

If `pip` is unavailable, use isolated execution:

```bash
uv run --with python-pptx --with python-docx --with openpyxl --with reportlab --with pypdf2 --with pandas --with pillow --with numpy --with matplotlib python examples/basic_usage.py
```

## Agent execution contract

When this skill is loaded, the agent should:

1) Load `SKILL.md` and treat it as primary contract.
2) Keep processing local-first (no hidden network operations).
3) Generate office deliverables via `OfficeSuite` APIs.
4) For PPTX jobs, enforce strict preflight keys:
   - `theme`
   - `chart_mode`
   - `use_emojis`
   - `tone`
5) Choose chart mode by use-case:
   - `native` for editable in-PowerPoint charts.
   - `matplotlib` for style-rich rendered charts.
6) Return structured output with paths and artifact summary.

Expected output shape (minimum):

```json
{
  "success": true,
  "artifacts": [
    {"type": "pptx", "path": "...", "notes": "..."},
    {"type": "docx", "path": "...", "notes": "..."}
  ],
  "preflight": {
    "theme": "midnight-luxe",
    "chart_mode": "native",
    "use_emojis": false,
    "tone": "boardroom"
  }
}
```

## Capability surface

- Create: `.docx`, `.xlsx`, `.pdf`, `.pptx`
- Convert: LibreOffice-backed format conversion
- Extract: text/data from Office documents
- Batch operations: conversion and watermark helpers
- Presentations:
  - themes: `midnight-luxe | aurora-glow | obsidian-slate | ivory-bloom | neon-velocity`
  - chart modes: `native | matplotlib | auto`
  - tones: `classic-formal | boardroom | conversational | laid-back`

## Screenshots

### Theme previews

![Theme - Midnight Luxe](docs/screenshots/theme-midnight-luxe.png)
![Theme - Aurora Glow](docs/screenshots/theme-aurora-glow.png)
![Theme - Obsidian Slate](docs/screenshots/theme-obsidian-slate.png)
![Theme - Ivory Bloom](docs/screenshots/theme-ivory-bloom.png)
![Theme - Neon Velocity](docs/screenshots/theme-neon-velocity.png)

### Slide examples

![Example - Title Slide](docs/screenshots/example-title-slide.png)
![Example - Native Chart Slide](docs/screenshots/example-native-chart-slide.png)
![Example - Matplotlib Chart Slide](docs/screenshots/example-matplotlib-chart-slide.png)

### Document example

![Example - Document Page 1](docs/screenshots/example-document-page-1.png)

## Demo artifacts in repo

- `demo_assets/presentations/` → showcase decks per theme
- `demo_assets/documents/` → test report (`.docx` + `.pdf`)
- `demo_assets/charts/` → source chart images used in report
- `demo_assets/README.md` → generated artifact index

## Security policy for agents

- Local-first by default.
- No hidden outbound data flows.
- SMTP is explicit-only action.
- Secrets must come from env vars (`OFFICE_EMAIL_PASSWORD`, provider keys).
- Do not hardcode credentials.

See:
- `docs/SECURITY.md`
- `docs/AGENT_COMPATIBILITY.md`

## Attribution and license

- Original project/base implementation: **Tao Jin** (`shynloc`)
- Adaptation and agent-focused restructuring: **Rafael Lozano**
- License: **MIT** (`LICENSE`, `NOTICE`)
