#!/usr/bin/env python3
"""Batch processing example: convert DOCX files to PDF and add watermark."""

import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))
from office_suite import OfficeSuite


def main():
    print("=== DocFlow - Presentations and Docs Skill | Batch Processing ===")

    suite = OfficeSuite()

    source_dir = "./word_docs"
    target_dir = "./pdf_docs"
    watermarked_dir = "./pdf_with_watermark"

    os.makedirs(source_dir, exist_ok=True)
    os.makedirs(target_dir, exist_ok=True)
    os.makedirs(watermarked_dir, exist_ok=True)

    print(f"\nConverting DOCX -> PDF from {source_dir} ...")
    results = suite.batch_convert(
        source_dir=source_dir,
        target_dir=target_dir,
        source_format="docx",
        target_format="pdf",
    )

    print(f"Converted {len(results)} files")
    for filename, status in results.items():
        if status.get("success"):
            print(f"  OK  {filename} -> {status.get('output_path')}")
        else:
            print(f"  ERR {filename} -> {status.get('error')}")

    print("\nAdding watermark to generated PDFs ...")
    watermark_results = suite.batch_add_watermark(
        source_dir=target_dir,
        target_dir=watermarked_dir,
        watermark_text="INTERNAL USE ONLY",
    )

    print(f"Watermarked {len(watermark_results)} files")
    for filename, status in watermark_results.items():
        if status.get("success"):
            print(f"  OK  {filename} -> {status.get('output_path')}")
        else:
            print(f"  ERR {filename} -> {status.get('error')}")

    print("\nDone.")


if __name__ == "__main__":
    main()
