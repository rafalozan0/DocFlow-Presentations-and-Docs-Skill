#!/usr/bin/env python3
"""Template-catalog demo.

Showcases the new file-based template registry for agents:

1. List all templates (slides + documents) with their slot contracts.
2. Render ONE slide template to PNG.
3. Build a full deck from a mix of slide templates into PPTX.
4. Render invoice + report documents to PDF.

Run:
    python examples/templates_demo.py
    # or
    uv run --with jinja2 --with python-pptx --with python-docx --with openpyxl \\
        --with reportlab --with pypdf2 --with pandas --with pillow --with numpy \\
        --with matplotlib python examples/templates_demo.py
"""

import json
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))
from office_suite import OfficeSuite


def main():
    suite = OfficeSuite()
    out_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "output", "templates_demo"))
    os.makedirs(out_dir, exist_ok=True)

    # 1) Inspect catalog (this is what agents should call first).
    catalog = suite.list_templates()
    print("== Slide templates ==")
    for t in catalog["slides"]:
        print(f"  - {t['id']:<24} {t['style']:<16} {t['name']}")
    print("\n== Document templates ==")
    for t in catalog["documents"]:
        print(f"  - {t['id']:<24} {t['style']:<16} {t['name']}")

    # 2) Render a single slide to PNG using the meta 'example' payload.
    meta = suite.get_template_meta("manifesto-grain", "slides")
    res = suite.create(
        "slide_template",
        template="manifesto-grain",
        data=meta["example"],
        output_path=os.path.join(out_dir, "single_slide_manifesto.png"),
    )
    print("\nSingle slide PNG:", res)

    # 3) Build a full PPTX deck mixing several templates.
    def ex(tpl_id):
        return suite.get_template_meta(tpl_id, "slides")["example"]

    deck_slides = [
        {"template": "cover-blue",            "data": ex("cover-blue")},
        {"template": "mono-editorial",        "data": ex("mono-editorial")},
        {"template": "manifesto-grain",       "data": ex("manifesto-grain")},
        {"template": "chart-kpi-dashboard",   "data": ex("chart-kpi-dashboard")},
        {"template": "chart-comparison",      "data": ex("chart-comparison")},
        {"template": "revenue-streams",       "data": ex("revenue-streams")},
        {"template": "brutalist-concrete",    "data": ex("brutalist-concrete")},
        {"template": "swiss-minimal",         "data": ex("swiss-minimal")},
        {"template": "sales-pitch",           "data": ex("sales-pitch")},
    ]
    res = suite.create(
        "slide_deck_templates",
        slides=deck_slides,
        output_path=os.path.join(out_dir, "mixed_deck.pptx"),
        keep_renders=True,
        renders_dir=os.path.join(out_dir, "mixed_deck_renders"),
    )
    print("\nPPTX deck:", json.dumps({k: v for k, v in res.items() if k != "kept_render_files"}, indent=2))

    # 4) Generate an invoice (modern + brutalist), a letter and an executive report PDF.
    for tpl_id in ["invoice-modern", "invoice-brutalist", "report-executive", "report-minimal", "proposal-cover", "letter-formal"]:
        data = suite.get_template_meta(tpl_id, "documents")["example"]
        res = suite.create(
            "doc_template",
            template=tpl_id,
            data=data,
            output_path=os.path.join(out_dir, f"{tpl_id}.pdf"),
        )
        print(f"doc[{tpl_id}]:", res.get("success"), res.get("output_path"))

    print("\nAll artifacts in:", out_dir)


if __name__ == "__main__":
    main()
