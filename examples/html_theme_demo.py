#!/usr/bin/env python3
"""Generate a style-heavy deck using HTML templates + Chrome screenshots.

This demonstrates the new html_pptx workflow for reference-style themes.
"""

import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))
from office_suite import OfficeSuite


def build_html_style_slides():
    return [
        {
            "kicker": "AI Enablement · 2026",
            "title": "Design language system for modern operational decks",
            "subtitle": "HTML-templated slide rendering for premium style consistency and reusable content workflows.",
            "left_title": "Visual pillars",
            "left_list": [
                "High-contrast typography with strong hierarchy",
                "Asymmetric composition and card-driven storytelling",
                "Controlled accent gradients and soft geometric backdrops",
                "Crisp executive readability at 16:9 1080p output",
            ],
            "right_title": "Target impact",
            "right_text": "Deploying reusable HTML slide templates reduces manual formatting overhead while preserving brand coherence across teams.",
            "stat": "-42%",
            "tags": ["Template Engine", "Brand Safe", "Agent-Ready"],
            "footer_left": "DocFlow HTML Theme Pipeline",
        },
        {
            "kicker": "Workflow",
            "title": "From structured content to production-ready deck assets",
            "subtitle": "Agent fills semantic fields, template compiles layout, renderer captures visual slides, and PPTX assembles final artifact.",
            "left_title": "Pipeline",
            "left_list": [
                "Load content schema (title, subtitle, cards, stats)",
                "Apply theme tokens (color, font, spacing, accents)",
                "Render HTML/CSS per slide with deterministic layout",
                "Capture 1920x1080 PNG and package into PPTX",
            ],
            "right_title": "Governance benefits",
            "right_text": "Template-first generation enforces visual guardrails and improves repeatability for sales, training, and executive reporting decks.",
            "stat": "+31%",
            "tags": ["Deterministic", "Scalable", "Consistent"],
            "footer_left": "DocFlow HTML Theme Pipeline",
        },
        {
            "kicker": "Rollout Recommendation",
            "title": "Adopt hybrid strategy: native charts + html-themed covers",
            "subtitle": "Use HTML pipeline for high-impact brand slides and Python-native chart slides for editable operational data pages.",
            "left_title": "Execution model",
            "left_list": [
                "Covers and section dividers: html_pptx theme mode",
                "Data-heavy internals: native or matplotlib chart modes",
                "Preflight remains mandatory for compliance and tone",
                "Publish artifact bundle with screenshots for QA",
            ],
            "right_title": "Business outcome",
            "right_text": "Teams get visually premium storytelling without sacrificing downstream editability and analyst control in recurring reports.",
            "stat": "V1.1",
            "tags": ["Hybrid", "Practical", "Enterprise"],
            "footer_left": "DocFlow HTML Theme Pipeline",
        },
    ]


def main():
    suite = OfficeSuite()
    out_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "output"))
    os.makedirs(out_dir, exist_ok=True)

    slides = build_html_style_slides()

    themes = ["neon-executive", "cobalt-corporate", "bold-pitch"]
    for theme in themes:
        output_path = os.path.join(out_dir, f"html_theme_demo_{theme}.pptx")
        result = suite.create(
            "html_pptx",
            title=f"HTML Theme Demo - {theme}",
            slides=slides,
            output_path=output_path,
            theme_name=theme,
            keep_renders=True,
            renders_dir=os.path.join(out_dir, f"renders_{theme}"),
        )
        print(theme, result)


if __name__ == "__main__":
    main()
