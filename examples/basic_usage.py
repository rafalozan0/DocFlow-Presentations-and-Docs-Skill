#!/usr/bin/env python3
"""Basic usage example: create and convert Office documents."""

import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))
from office_suite import OfficeSuite


def main():
    print("=== DocFlow - Presentations and Docs Skill | Basic Usage ===")

    suite = OfficeSuite()
    output_dir = "./output"
    os.makedirs(output_dir, exist_ok=True)

    print("\n1) Creating Word document...")
    word_path = os.path.join(output_dir, "sales_report_q1.docx")
    result = suite.create(
        "word",
        title="Q1 2025 Sales Report",
        content="""
# Executive Summary
Q1 revenue reached 12M USD, growing 25% YoY.

## Team Performance
- Team A: 4.5M (112.5% target)
- Team B: 3.8M (95% target)
- Team C: 3.7M (92.5% target)

## Next Steps
1. Set Q2 target to 15M USD
2. Expand East and South regions
3. Launch 3 new products
        """,
        output_path=word_path,
    )
    print(result)

    print("\n2) Creating Excel spreadsheet...")
    excel_path = os.path.join(output_dir, "sales_data_q1.xlsx")
    sales_data = [
        ["Team", "Jan", "Feb", "Mar", "Quarter Total", "Completion"],
        ["Team A", 150, 140, 160, 450, "112.5%"],
        ["Team B", 120, 130, 130, 380, "95%"],
        ["Team C", 110, 120, 140, 370, "92.5%"],
    ]
    result = suite.create(
        "excel",
        title="Q1 Sales",
        data=sales_data,
        output_path=excel_path,
        create_chart=True,
    )
    print(result)

    print("\n3) Creating PDF document...")
    pdf_path = os.path.join(output_dir, "sales_report_q1.pdf")
    result = suite.create(
        "pdf",
        title="Q1 2025 Sales Report",
        content="""
# Executive Summary
Q1 revenue reached 12M USD, growing 25% YoY.

## Team Performance
- Team A: 4.5M (112.5% target)
- Team B: 3.8M (95% target)
- Team C: 3.7M (92.5% target)
        """,
        output_path=pdf_path,
    )
    print(result)

    print("\n4) Creating PowerPoint deck...")
    ppt_path = os.path.join(output_dir, "sales_report_q1.pptx")
    slides = [
        {
            "title": "Q1 2025 Sales Report",
            "content": "Presenter: Sales Ops\nDate: April 2025",
            "layout": "title",
        },
        {
            "title": "Summary",
            "content": "• Revenue: 12M USD\n• YoY growth: 25%\n• Target completion: 108%",
            "layout": "content",
        },
    ]
    result = suite.create("pptx", title="Q1 deck", slides=slides, output_path=ppt_path)
    print(result)

    print("\n5) Converting Word to PDF...")
    converted_pdf = os.path.join(output_dir, "sales_report_q1_from_word.pdf")
    result = suite.convert(word_path, to="pdf", output_path=converted_pdf)
    print(result)

    print("\nDone. Output directory:", os.path.abspath(output_dir))


if __name__ == "__main__":
    main()
