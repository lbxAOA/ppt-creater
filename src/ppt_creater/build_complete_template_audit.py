"""Build one bookmarked PDF for human audit of every acquired template family."""
from __future__ import annotations

import shutil
from pathlib import Path

import win32com.client
from PIL import Image
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

ROOT = Path(r"C:/ppt-creater")
INTAKE = ROOT / "network-supplements"
OUT = ROOT / "output"
NEW_BODY = OUT / "network-round6-slidesgo-manual-download-review-body.pdf"
NEW_PDF = OUT / "network-round6-slidesgo-manual-download-review.pdf"
COMPLETE_PDF = OUT / "template-library-complete-audit.pdf"

NEW_FAMILIES = [
    ("NET-21", "GOVERNMENT / PUBLIC POLICY", "public-policy-project-proposal.pptx"),
    ("NET-22", "ARCHITECTURE / CONSTRUCTION", "architecture-construction-management-company-profile.pptx"),
    ("NET-23", "LOGISTICS / DISTRIBUTION", "logistics-and-distribution-of-goods.pptx"),
]

SECTIONS = [
    ("A. Local production families + NET-01 through NET-04", OUT / "full-production-template-library-review.pdf"),
    ("B. Network candidates: NET-06 through NET-08", OUT / "network-round2-candidates-style-review.pdf"),
    ("C. Network candidates: NET-09, NET-10, NET-12", OUT / "network-round4-executive-candidates-style-review.pdf"),
    ("D. Network candidate: NET-13", OUT / "network-round5-ai-tech-candidate-style-review.pdf"),
    ("E. Network domain candidates: NET-14 through NET-20", OUT / "network-round1-domain-gap-review.pdf"),
    ("F. Manually downloaded Slidesgo candidates: NET-21 through NET-23", NEW_PDF),
]


def export_new_slides() -> None:
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = 1
    try:
        for candidate_id, _, filename in NEW_FAMILIES:
            source = INTAKE / candidate_id / filename
            previews = source.parent / "previews"
            if previews.exists():
                shutil.rmtree(previews)
            previews.mkdir()
            presentation = app.Presentations.Open(str(source), WithWindow=False)
            presentation.Export(str(previews), "PNG", 1600, 900)
            presentation.Close()
    finally:
        app.Quit()


def build_new_contact_sheet() -> None:
    width, height = landscape((842, 595))
    margin, title_height = 24, 46
    cell_width = (width - 2 * margin) / 3
    cell_height = (height - title_height - 2 * margin) / 2
    pdf = canvas.Canvas(str(NEW_BODY), pagesize=(width, height))
    bookmarks: list[tuple[str, int]] = []
    page_number = 0
    for candidate_id, domain, _ in NEW_FAMILIES:
        images = sorted((INTAKE / candidate_id / "previews").glob("*.PNG"))
        if not images:
            raise RuntimeError(f"Missing previews for {candidate_id}")
        bookmarks.append((f"{candidate_id} | {domain}", page_number))
        for start in range(0, len(images), 6):
            page_number += 1
            pdf.setFillColorRGB(0.05, 0.07, 0.10)
            pdf.rect(0, 0, width, height, fill=1, stroke=0)
            pdf.setFillColorRGB(1, 1, 1)
            pdf.setFont("Helvetica-Bold", 16)
            pdf.drawString(margin, height - 24, f"{candidate_id} | {domain}")
            pdf.setFillColorRGB(0.75, 0.80, 0.85)
            pdf.setFont("Helvetica", 8)
            pdf.drawRightString(width - margin, height - 24, f"candidate only | slides {start + 1}-{min(start + 6, len(images))} | external license review required")
            for local_index, image_path in enumerate(images[start : start + 6]):
                col, row = local_index % 3, local_index // 3
                x = margin + col * cell_width
                top = height - title_height - margin - row * cell_height
                with Image.open(image_path) as image:
                    image_width, image_height = image.size
                scale = min((cell_width - 10) / image_width, (cell_height - 24) / image_height)
                draw_width, draw_height = image_width * scale, image_height * scale
                pdf.setFillColorRGB(0.2, 0.23, 0.27)
                pdf.rect(x + 1, top - cell_height + 1, cell_width - 3, cell_height - 3, fill=1, stroke=0)
                pdf.drawImage(ImageReader(str(image_path)), x + (cell_width - draw_width) / 2, top - draw_height - 13, width=draw_width, height=draw_height)
                pdf.setFillColorRGB(0.85, 0.88, 0.90)
                pdf.setFont("Helvetica", 8)
                pdf.drawString(x + 5, top - 11, f"SOURCE SLIDE {start + local_index + 1}")
            pdf.showPage()
    pdf.save()
    reader = PdfReader(str(NEW_BODY))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    for title, page in bookmarks:
        writer.add_outline_item(title, page)
    with NEW_PDF.open("wb") as handle:
        writer.write(handle)


def build_complete_audit() -> None:
    writer = PdfWriter()
    for section_title, source in SECTIONS:
        reader = PdfReader(str(source))
        section_page = len(writer.pages)
        writer.add_outline_item(section_title, section_page)
        for page in reader.pages:
            writer.add_page(page)
        for outline in reader.outline:
            if isinstance(outline, dict) and "/Title" in outline and "/Page" in outline:
                try:
                    writer.add_outline_item(str(outline["/Title"]), reader.get_destination_page_number(outline) + section_page)
                except Exception:
                    pass
    with COMPLETE_PDF.open("wb") as handle:
        writer.write(handle)


if __name__ == "__main__":
    OUT.mkdir(exist_ok=True)
    export_new_slides()
    build_new_contact_sheet()
    build_complete_audit()
    print({"new_review": str(NEW_PDF), "complete_audit": str(COMPLETE_PDF), "pages": len(PdfReader(str(COMPLETE_PDF)).pages)})
