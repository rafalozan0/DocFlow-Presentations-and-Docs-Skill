<div align="center">

# DocFlow - Presentations and Docs Skill

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![Python >=3.8](https://img.shields.io/badge/python-%3E%3D3.8-blue.svg)](https://www.python.org/downloads/)
[![Runtime](https://img.shields.io/badge/runtimes-Hermes%20%7C%20OpenClaw-7c3aed)](#agent-bootstrap-hermes--openclaw)
[![Mode](https://img.shields.io/badge/mode-agent--only-111827)](#agent-bootstrap-hermes--openclaw)

**Agent-first skill for generating DOCX/XLSX/PDF/PPTX deliverables with strict preflight and dual chart modes.**

Out-of-the-box skill package for Hermes/OpenClaw with file-based templates and predictable agent contracts.

</div>

## 🎉 News

- **2026-04-17**
  - Repository upgraded to **v1.1.0** with file-based template catalog (13 slide templates + 6 document templates).
  - Agent-first README and compatibility docs aligned for Hermes/OpenClaw out-of-the-box usage.

## Agent context (Hermes / OpenClaw)

This repository is intended for autonomous agent runtimes.

- Hermes Agent (Nous Research): terminal-native autonomous agent framework.
- OpenClaw (`openclaw/openclaw`): personal AI agent runtime and skill ecosystem.

This README is intentionally focused on agent orchestration (not end-user tutorials).

## Agent bootstrap (Hermes / OpenClaw)

Use exactly this block.

```bash
git clone https://github.com/rafalozan0/DocFlow-Presentations-and-Docs-Skill.git
cd DocFlow-Presentations-and-Docs-Skill
python -m compileall src examples setup.py
python -m pip install -r requirements.txt || true
```

If `pip` is unavailable, use isolated execution.

```bash
uv run --with python-pptx --with python-docx --with openpyxl --with reportlab --with pypdf2 --with pandas --with pillow --with numpy --with matplotlib --with jinja2 python examples/basic_usage.py
```

## Screenshots

Latest templates are shown first.

### Featured slide templates (latest first)

Use these IDs directly when calling templates:

**Template ID:** `startup-pitch` — Startup Pitch · Hero Pill
![Slide template - startup-pitch](docs/screenshots/templates/slides/startup-pitch.png)

**Template ID:** `design-review` — Design Review · Framed Hero
![Slide template - design-review](docs/screenshots/templates/slides/design-review.png)

**Template ID:** `sales-pitch` — Sales Pitch · Braced Title
![Slide template - sales-pitch](docs/screenshots/templates/slides/sales-pitch.png)

**Template ID:** `cover-blue` — Cover · Solid Color
![Slide template - cover-blue](docs/screenshots/templates/slides/cover-blue.png)

**Template ID:** `question-3d` — Question Slide (3D Glow)
![Slide template - question-3d](docs/screenshots/templates/slides/question-3d.png)

**Template ID:** `revenue-streams` — Two-column Comparison
![Slide template - revenue-streams](docs/screenshots/templates/slides/revenue-streams.png)

**Template ID:** `brutalist-concrete` — Brutalist · Concrete Grain
![Slide template - brutalist-concrete](docs/screenshots/templates/slides/brutalist-concrete.png)

**Template ID:** `swiss-minimal` — Swiss Minimal
![Slide template - swiss-minimal](docs/screenshots/templates/slides/swiss-minimal.png)

**Template ID:** `mono-editorial` — Mono Editorial · Numbered List
![Slide template - mono-editorial](docs/screenshots/templates/slides/mono-editorial.png)

**Template ID:** `editorial-layout` — Editorial Layout
![Slide template - editorial-layout](docs/screenshots/templates/slides/editorial-layout.png)

**Template ID:** `manifesto-grain` — Manifesto (Purple Grain)
![Slide template - manifesto-grain](docs/screenshots/templates/slides/manifesto-grain.png)

**Template ID:** `chart-kpi-dashboard` — KPI Dashboard (4 KPIs + Bar Chart)
![Slide template - chart-kpi-dashboard](docs/screenshots/templates/slides/chart-kpi-dashboard.png)

**Template ID:** `chart-comparison` — Multi-series Line Comparison
![Slide template - chart-comparison](docs/screenshots/templates/slides/chart-comparison.png)

### Featured document templates (latest first)

Use these IDs directly when calling templates:

**Template ID:** `proposal-cover` — Proposal Cover Page
![Document template - proposal-cover](docs/screenshots/templates/documents/proposal-cover.png)

**Template ID:** `report-executive` — Executive Report
![Document template - report-executive](docs/screenshots/templates/documents/report-executive.png)

**Template ID:** `invoice-brutalist` — Invoice · Brutalist
![Document template - invoice-brutalist](docs/screenshots/templates/documents/invoice-brutalist.png)

**Template ID:** `invoice-modern` — Invoice · Modern
![Document template - invoice-modern](docs/screenshots/templates/documents/invoice-modern.png)

**Template ID:** `report-minimal` — Report · Minimal
![Document template - report-minimal](docs/screenshots/templates/documents/report-minimal.png)

**Template ID:** `letter-formal` — Formal Letter
![Document template - letter-formal](docs/screenshots/templates/documents/letter-formal.png)

### Theme previews (python-pptx)

![Theme - Midnight Luxe](docs/screenshots/theme-midnight-luxe.png)
![Theme - Aurora Glow](docs/screenshots/theme-aurora-glow.png)
![Theme - Obsidian Slate](docs/screenshots/theme-obsidian-slate.png)
![Theme - Ivory Bloom](docs/screenshots/theme-ivory-bloom.png)
![Theme - Neon Velocity](docs/screenshots/theme-neon-velocity.png)

### Theme previews (html template engine)

![HTML Theme - Neon Executive](docs/screenshots/html-theme-neon-executive.png)
![HTML Theme - Cobalt Corporate](docs/screenshots/html-theme-cobalt-corporate.png)
![HTML Theme - Bold Pitch](docs/screenshots/html-theme-bold-pitch.png)

### Legacy examples

![Example - Title Slide](docs/screenshots/example-title-slide.png)
![Example - Native Chart Slide](docs/screenshots/example-native-chart-slide.png)
![Example - Matplotlib Chart Slide](docs/screenshots/example-matplotlib-chart-slide.png)
![Example - Document Page 1](docs/screenshots/example-document-page-1.png)

## Agent execution contract

When this skill is loaded, the agent should:

1) Load `SKILL.md` and treat it as primary contract.
2) Keep processing local-first (no hidden network operations).
3) Generate office deliverables via `OfficeSuite` APIs.
4) For template-driven work, start with `suite.list_templates()` and
   `suite.get_template_meta(id, category)` to read the slot contract,
   then call `create(doc_type="slide_template" | "slide_deck_templates" |
   "doc_template", ...)`.
5) For PPTX jobs, enforce strict preflight keys:
   - `theme`
   - `chart_mode`
   - `use_emojis`
   - `tone`
6) Choose chart mode by use-case:
   - `native` for editable in-PowerPoint charts.
   - `matplotlib` for style-rich rendered charts.
7) Return structured output with paths and artifact summary.

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
- Presentations (python-pptx engine):
  - themes: `midnight-luxe | aurora-glow | obsidian-slate | ivory-bloom | neon-velocity`
  - chart modes: `native | matplotlib | auto`
  - tones: `classic-formal | boardroom | conversational | laid-back`
- Presentations (html template engine):
  - doc_type: `html_pptx` or `pptx_html`
  - html themes: `neon-executive | cobalt-corporate | bold-pitch`
  - render stack: `HTML/CSS -> headless Chrome PNG -> PPTX`
- File-based template catalog (v1.1):
  - doc_types: `slide_template`, `slide_deck_templates`, `doc_template`
  - 13 slide templates (editorial, manifesto, brutalist-concrete, swiss-minimal,
    mono-editorial, chart-kpi-dashboard, chart-comparison, cover-blue,
    startup-pitch, design-review, sales-pitch, question-3d, revenue-streams)
  - 6 document templates (invoice-modern, invoice-brutalist, report-executive,
    report-minimal, proposal-cover, letter-formal)
  - Agents discover and fill templates via `suite.list_templates()` and
    `suite.get_template_meta(id, category)` — no hardcoded IDs needed.
  - See `docs/TEMPLATE_CATALOG.md` for the full slot contract reference.

## Demo artifacts in repo

- `demo_assets/presentations/` → showcase decks per theme
- `demo_assets/documents/` → test report (`.docx` + `.pdf`)
- `demo_assets/charts/` → source chart images used in report
- `demo_assets/README.md` → generated artifact index
- `docs/screenshots/templates/slides/*.png` → one generated preview per slide template (13)
- `docs/screenshots/templates/documents/*.png` → one generated preview per document template (6)
- `output/html_theme_demo_*.pptx` → html-engine style demos
- `output/renders_*` → rendered html/png intermediates for style QA

## Security policy for agents

- Local-first by default.
- No hidden outbound data flows.
- Do not hardcode credentials.
- HTML render mode requires local `google-chrome` headless binary.
- You can override Chrome binary path with `OFFICE_CHROME_BIN`.

See:
- `docs/SECURITY.md`
- `docs/AGENT_COMPATIBILITY.md`
- `docs/HTML_THEME_TOKENS.md`
- `docs/TEMPLATE_CATALOG.md` — file-based template catalog (v1.1)

## Attribution and license

- Original project/base implementation: **Tao Jin** (`shynloc`)
- Adaptation and agent-focused restructuring: **Rafael Lozano**
- License: **MIT** (`LICENSE`, `NOTICE`)
