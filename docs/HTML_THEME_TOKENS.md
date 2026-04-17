# HTML Theme Tokens (Reference Style Pack)

This file captures the visual token system extracted from `~/Descargas/Slides-Examples`.
It is used by the new `html_pptx` / `pptx_html` rendering mode.

## Themes

### 1) neon-executive
- background: `#07090F`
- surface: `#111827`
- text: `#F5F7FF`
- muted: `#A9B2C7`
- accent: `#4D6BFF`
- accent2: `#7A4DFF`
- style intent: cinematic dark, glow accents, premium AI/product narrative

### 2) cobalt-corporate
- background: `#F7F9FC`
- surface: `#FFFFFF`
- text: `#0F172A`
- muted: `#475569`
- accent: `#2F66D8`
- accent2: `#1E40AF`
- style intent: clean enterprise layout, strong readability, report-grade visual language

### 3) bold-pitch
- background: `#000000`
- surface: `#111111`
- text: `#F8FAFC`
- muted: `#9CA3AF`
- accent: `#00D1B2`
- accent2: `#FFE500`
- style intent: aggressive pitch aesthetic, high contrast, editorial hero titles

## Shared layout primitives

- 16:9 canvas (1920x1080 render)
- top gradient accent bar
- left-heavy title hierarchy
- two-column panel content layout
- large KPI/stat callout
- chips/tags for context metadata
- subtle geometric glow background objects

## Rendering pipeline

`HTML/CSS (Jinja) -> headless Chrome screenshot PNG -> PPTX assembly`

Implemented in:
- `src/office_suite/html_slides.py`
- `examples/html_theme_demo.py`
