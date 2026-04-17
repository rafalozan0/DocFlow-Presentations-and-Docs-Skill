# Template Catalog

This document describes the **file-based template registry** introduced in v1.1.

Templates live on disk under `src/office_suite/templates/` — one folder per
template, each with:

- `template.html` — Jinja2 template
- `meta.json` — slot contract (what the agent must fill in)

Agents (Hermes / OpenClaw) should **always call `list_templates()` first** to
discover the catalog and read each template's slots, rather than hard-coding
template IDs.

## Agent workflow

```python
from office_suite import OfficeSuite
suite = OfficeSuite()

# 1) Discover.
catalog = suite.list_templates()
# {"slides": [...], "documents": [...]}

# 2) Read slot contract for a specific template.
meta = suite.get_template_meta("invoice-modern", category="documents")
# meta["slots"] -> list of required/optional slots with types and defaults
# meta["example"] -> a ready-to-render example payload

# 3a) Render a single slide to PNG.
suite.create(
    "slide_template",
    template="manifesto-grain",
    data={...},
    output_path="/tmp/slide.png",
)

# 3b) Build a full PPTX deck from multiple templates.
suite.create(
    "slide_deck_templates",
    slides=[
        {"template": "cover-blue", "data": {...}},
        {"template": "mono-editorial", "data": {...}},
    ],
    output_path="/tmp/deck.pptx",
)

# 3c) Render a document to PDF.
suite.create(
    "doc_template",
    template="invoice-modern",
    data={...},
    output_path="/tmp/invoice.pdf",
)
```

## doc_types (OfficeSuite.create)

| `doc_type`              | Input kwargs                                      | Output                      |
|-------------------------|---------------------------------------------------|-----------------------------|
| `slide_template`        | `template`, `data`, `output_path`, `output_format` ∈ {png, pdf, html} | Single slide render |
| `slide_deck_templates`  | `slides=[{template, data}]`, `output_path`, `keep_renders`, `renders_dir` | `.pptx` file |
| `doc_template`          | `template`, `data`, `output_path`, `output_format` ∈ {pdf, html, png} | Document render |

All three use **headless Chrome** for PNG/PDF rasterization, so they require a
local `google-chrome` or `chromium` binary. You can override the path with
the `OFFICE_CHROME_BIN` environment variable.

## Slide templates (16:9 · 1920×1080)

| ID                     | Style             | When to use |
|------------------------|-------------------|-------------|
| `editorial-layout`     | editorial         | Portfolio content slide with 2-column layout and image placeholders |
| `manifesto-grain`      | bold-gradient     | Mission/values section divider with oversized numeric hero |
| `question-3d`          | cinematic-dark    | Rhetorical prompt or section break with 3D glow |
| `revenue-streams`      | light-editorial   | Two-column comparison (A vs B, pricing, revenue streams) |
| `cover-blue`           | solid-cover       | Deck opener with solid brand color and geometric ornaments |
| `startup-pitch`        | hero-pill         | Pitch deck opener, massive title, metadata pills |
| `design-review`        | framed-hero       | Formal B/W design-review kickoff slide |
| `sales-pitch`          | braced-hero       | Partnership / sales deck opener with `{Company}` title |
| `brutalist-concrete`   | brutalist         | Brutalist manifesto slide with concrete grain & heavy type |
| `swiss-minimal`        | minimal           | Calm, modern headline + body + rule |
| `mono-editorial`       | minimal           | Agenda / framework / step list with numbered items |
| `chart-kpi-dashboard`  | data-dashboard    | Quarterly snapshot: 4 KPI cards + bar chart |
| `chart-comparison`     | data-dashboard    | Multi-series line chart with context takeaways |

Each template's exhaustive slot list is in its own `meta.json`. Read it with
`OfficeSuite.get_template_meta(id, "slides")`.

## Document templates (A4)

| ID                    | Style          | When to use |
|-----------------------|----------------|-------------|
| `invoice-modern`      | modern-clean   | Standard business invoice |
| `invoice-brutalist`   | brutalist      | High-contrast raw invoice for editorial/indie brands |
| `report-executive`    | executive      | Exec report with hero cover + KPIs + sections |
| `report-minimal`      | minimal        | Weekly update / internal memo |
| `proposal-cover`      | hero-cover     | First page of a proposal document |
| `letter-formal`       | formal         | Traditional business letter |

## Adding a new template (no Python code needed)

1. Create a new folder: `src/office_suite/templates/slides/<slug>/` (or
   `templates/documents/<slug>/` for documents).
2. Add `template.html` — a Jinja2 template. Extend the base partial:

    ```html
    {% extends "slide_base.html" %}   {# or "doc_base.html" #}
    {% block styles %} /* ... */ {% endblock %}
    {% block body %} /* ... */ {% endblock %}
    ```

3. Add `meta.json` with:

    ```json
    {
      "id": "your-slug",
      "name": "Human-friendly name",
      "category": "slides",
      "style": "short-label",
      "description": "What this template is best at.",
      "aspect": "16:9",
      "slots": [
        {"name": "title", "type": "string", "required": true},
        {"name": "accent", "type": "string (hex)", "required": false, "default": "#111"}
      ],
      "example": { "title": "Hello world" }
    }
    ```

4. Call `OfficeSuite.list_templates()` — your template appears immediately.

### Jinja gotchas

- Use `foo['values']` (subscript) when your data contains a key named
  `values` (or `keys`, `items`) to avoid resolving Python dict methods.
- The registry uses `ChainableUndefined`, so `{{ foo.bar.baz }}` won't raise
  even if parts are missing; wrap with `{% if foo %}` or `{{ foo|default('...') }}`
  for clean output.
- Keep the template self-contained: inline CSS in `{% block styles %}`, no
  external CSS/JS/font imports (Chrome headless runs offline).

## Safety notes for agents

- All rendering is **local-first**, no outbound network calls.
- Chrome is invoked via a fixed argv — no arbitrary shell expansion.
- Templates are Jinja2 with `autoescape` on; HTML in slot values is escaped
  unless the template explicitly uses `|safe` (e.g. `question-3d.question`
  allows `<b>` emphasis by design).
