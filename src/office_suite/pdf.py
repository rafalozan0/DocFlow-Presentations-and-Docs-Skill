import os
from typing import Any, Dict, List, Optional

from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer


def create_pdf(title: str, content: str, output_path: str, **kwargs) -> Dict[str, Any]:
    """Create a PDF document from markdown-like text."""
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=72,
        leftMargin=72,
        topMargin=72,
        bottomMargin=72,
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "CustomTitle",
        parent=styles["Title"],
        fontSize=24,
        textColor=colors.HexColor("#003366"),
        alignment=1,
    )

    heading1_style = ParagraphStyle(
        "CustomHeading1",
        parent=styles["Heading1"],
        fontSize=18,
        textColor=colors.HexColor("#006699"),
    )

    normal_style = ParagraphStyle(
        "CustomNormal",
        parent=styles["Normal"],
        fontSize=12,
        leading=18,
    )

    story = [Paragraph(title, title_style), Spacer(1, 30)]

    lines = content.split("\n")
    for line in lines:
        line = line.strip()
        if not line:
            story.append(Spacer(1, 6))
            continue

        if line.startswith("# "):
            story.append(Paragraph(line[2:], heading1_style))
            story.append(Spacer(1, 12))
        elif line.startswith("## "):
            story.append(Paragraph(line[3:], styles["Heading2"]))
            story.append(Spacer(1, 10))
        elif line.startswith("### "):
            story.append(Paragraph(line[4:], styles["Heading3"]))
            story.append(Spacer(1, 8))
        elif line.startswith("- ") or line.startswith("* "):
            story.append(Paragraph(f"• {line[2:]}", normal_style))
        else:
            story.append(Paragraph(line, normal_style))

    doc.build(story)

    reader = PdfReader(output_path)
    return {
        "output_path": output_path,
        "file_size": os.path.getsize(output_path),
        "pages": len(reader.pages),
    }


def add_watermark(
    input_path: str,
    watermark_text: str,
    output_path: Optional[str] = None,
    **kwargs,
) -> Dict[str, Any]:
    """Currently performs safe pass-through copy for PDF watermark requests.

    Note: This is intentionally conservative; full visual watermarking is not yet implemented.
    """
    _ = watermark_text
    output_path = output_path or input_path

    reader = PdfReader(input_path)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "wb") as f:
        writer.write(f)

    return {
        "output_path": output_path,
        "warning": "PDF watermarking is not implemented yet; output is a clean copy of input.",
    }


def extract_text(input_path: str, **kwargs) -> str:
    """Extract text from all pages in a PDF."""
    reader = PdfReader(input_path)
    all_text = []
    for page in reader.pages:
        all_text.append(page.extract_text() or "")
    return "\n".join(all_text)


def merge_pdfs(input_paths: List[str], output_path: str, **kwargs) -> Dict[str, Any]:
    """Merge multiple PDF files into one output."""
    writer = PdfWriter()

    for path in input_paths:
        if not os.path.exists(path):
            continue
        reader = PdfReader(path)
        for page in reader.pages:
            writer.add_page(page)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "wb") as f:
        writer.write(f)

    return {"output_path": output_path, "merged_pages": len(writer.pages)}


def split_pdf(input_path: str, output_dir: str, split_by: str = "page", **kwargs) -> Dict[str, Any]:
    """Split a PDF file by page."""
    if split_by != "page":
        return {"success": False, "error": f"Unsupported split mode: {split_by}"}

    os.makedirs(output_dir, exist_ok=True)
    reader = PdfReader(input_path)
    total_pages = len(reader.pages)

    for page_num in range(total_pages):
        writer = PdfWriter()
        writer.add_page(reader.pages[page_num])
        output_path = os.path.join(output_dir, f"page_{page_num + 1}.pdf")
        with open(output_path, "wb") as f:
            writer.write(f)

    return {"output_dir": output_dir, "total_pages": total_pages}
