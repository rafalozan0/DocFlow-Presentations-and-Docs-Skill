#!/usr/bin/env python3
"""Presentation style + chart demo.

Shows:
- Modern presentation theme presets
- Option A charts (native python-pptx)
- Option B charts (matplotlib image embedding)
- Tone + emoji controls
"""

import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))
from office_suite import OfficeSuite


def build_slides():
    return [
        {
            "layout": "title",
            "title": "DocFlow - Presentations and Docs Skill",
            "content": "Modern deck generation with style controls",
        },
        {
            "layout": "content_with_chart",
            "title": "Option A: Native Charts",
            "content": "- Built with python-pptx charts\n- Fast and editable in PowerPoint\n- Great for standard dashboards",
            "chart": {
                "type": "column",
                "title": "Revenue by Month",
                "categories": ["Jan", "Feb", "Mar", "Apr"],
                "series": [
                    {"name": "2025", "values": [120, 138, 151, 170]},
                    {"name": "2026", "values": [132, 149, 166, 188]},
                ],
            },
        },
        {
            "layout": "content_with_chart",
            "title": "Option B: Matplotlib Charts",
            "content": "- Best for visual customization\n- Good for advanced chart styling\n- Inserted as image in the deck",
            "chart": {
                "mode": "matplotlib",
                "type": "line",
                "title": "Pipeline Coverage Trend",
                "categories": ["W1", "W2", "W3", "W4", "W5"],
                "series": [
                    {"name": "Coverage", "values": [2.1, 2.3, 2.5, 2.4, 2.7]},
                    {"name": "Target", "values": [2.0, 2.0, 2.0, 2.0, 2.0]},
                ],
            },
        },
    ]


def main():
    suite = OfficeSuite()
    out_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "output"))
    os.makedirs(out_dir, exist_ok=True)

    slides = build_slides()

    out_native = os.path.join(out_dir, "style_demo_native.pptx")
    r1 = suite.create(
        "pptx",
        title="Style Demo Native",
        slides=slides,
        output_path=out_native,
        require_preflight=True,
        theme="midnight-luxe",
        chart_mode="native",
        use_emojis=False,
        tone="boardroom",
    )

    out_mpl = os.path.join(out_dir, "style_demo_matplotlib.pptx")
    r2 = suite.create(
        "pptx",
        title="Style Demo Matplotlib",
        slides=slides,
        output_path=out_mpl,
        require_preflight=True,
        theme="aurora-glow",
        chart_mode="matplotlib",
        use_emojis=True,
        tone="conversational",
    )

    print("Native:", r1)
    print("Matplotlib:", r2)


if __name__ == "__main__":
    main()
