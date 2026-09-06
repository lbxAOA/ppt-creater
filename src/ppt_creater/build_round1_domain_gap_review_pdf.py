"""Export first-round network candidates and build a review contact-sheet PDF."""
from __future__ import annotations

import json
import shutil
from pathlib import Path

import win32com.client
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from PIL import Image

ROOT = Path(r"C:/ppt-creater")
INTAKE = ROOT / "network-supplements"
OUTPUT = ROOT / "output"
TEMP = OUTPUT / "network-round1-domain-gap-review-body.pdf"
PDF = OUTPUT / "network-round1-domain-gap-review.pdf"

CANDIDATES = [
    ("NET-14", "MEDICAL / PHARMA", "medical-cardiovascular.pptx"),
    ("NET-15", "EDUCATION / TRAINING", "education-teacher-dashboard.pptx"),
    ("NET-16", "FINANCIAL LITERACY", "financial-literacy.pptx"),
    ("NET-17", "LEGAL / COMPLIANCE", "legal-corporate.pptx"),
    ("NET-18", "INDUSTRIAL / ENGINEERING", "industrial-engineering.pptx"),
    ("NET-19", "SUSTAINABILITY / ESG", "sustainability-energy.pptx"),
    ("NET-20", "RETAIL / E-COMMERCE", "marketing-analysis.pptx"),
]


def export_slides() -> dict[str, int]:
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = 1
    report: dict[str, int] = {}
    try:
        for candidate_id, _, filename in CANDIDATES:
            source = INTAKE / candidate_id / filename
            previews = source.parent / "previews"
            if previews.exists():
                shutil.rmtree(previews)
            previews.mkdir()
            presentation = ppt.Presentations.Open(str(source), WithWindow=False)
            presentation.Export(str(previews), "PNG", 1600, 900)
            count = presentation.Slides.Count
            presentation.Close()
            report[candidate_id] = count
    finally:
        ppt.Quit()
    return report


def build_pdf(exported: dict[str, int]) -> int:
    width, height = landscape((842, 595))
    margin, title_height = 24, 46
    cell_width = (width - 2 * margin) / 3
    cell_height = (height - title_height - 2 * margin) / 2
    pdf = canvas.Canvas(str(TEMP), pagesize=(width, height))
    bookmarks: list[tuple[str, int]] = []
    page_number = 0
    for candidate_id, domain, _ in CANDIDATES:
        images = sorted((INTAKE / candidate_id / "previews").glob("*.PNG"))
        if not images:
            raise RuntimeError(f"No previews exported for {candidate_id}")
        for start in range(0, len(images), 6):
            bookmarks.append((f"{candidate_id} | {domain}", page_number))
            page_number += 1
            pdf.setFillColorRGB(0.05, 0.07, 0.10)
            pdf.rect(0, 0, width, height, fill=1, stroke=0)
            pdf.setFillColorRGB(1, 1, 1)
            pdf.setFont("Helvetica-Bold", 16)
            pdf.drawString(margin, height - 24, f"{candidate_id} | {domain}")
            pdf.setFillColorRGB(0.75, 0.80, 0.85)
            pdf.setFont("Helvetica", 8)
            pdf.drawRightString(width - margin, height - 24, f"candidate only | slides {start + 1}-{min(start + 6, len(images))} | external license review required")
            for local_index, image_path in enumerate(images[start:start + 6]):
                col, row = local_index % 3, local_index // 3
                x = margin + col * cell_width
                top = height - title_height - margin - row * cell_height
                with Image.open(image_path) as image:
                    image_width, image_height = image.size
                scale = min((cell_width - 10) / image_width, (cell_height - 24) / image_height)
                draw_width, draw_height = image_width * scale, image_height * scale
                draw_x = x + (cell_width - draw_width) / 2
                draw_y = top - draw_height - 13
                pdf.setFillColorRGB(0.2, 0.23, 0.27)
                pdf.rect(x + 1, top - cell_height + 1, cell_width - 3, cell_height - 3, fill=1, stroke=0)
                pdf.drawImage(ImageReader(str(image_path)), draw_x, draw_y, width=draw_width, height=draw_height)
                pdf.setFillColorRGB(0.85, 0.88, 0.90)
                pdf.setFont("Helvetica", 8)
                pdf.drawString(x + 5, top - 11, f"SOURCE SLIDE {start + local_index + 1}")
            pdf.showPage()
    pdf.save()
    reader = PdfReader(str(TEMP))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    seen = set()
    for title, page in bookmarks:
        if title not in seen:
            writer.add_outline_item(title, page)
            seen.add(title)
    with PDF.open("wb") as handle:
        writer.write(handle)
    return len(reader.pages)


if __name__ == "__main__":
    OUTPUT.mkdir(exist_ok=True)
    result = export_slides()
    pages = build_pdf(result)
    print(json.dumps({"exported_slides": result, "contact_sheet_pdf": str(PDF), "pdf_pages": pages}, ensure_ascii=False))
