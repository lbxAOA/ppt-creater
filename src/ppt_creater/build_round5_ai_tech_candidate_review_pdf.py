"""Build a review contact sheet and provenance registry for NET-13."""
from __future__ import annotations

import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZipFile

from PIL import Image
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

ROOT = Path(r"C:/ppt-creater")
SOURCE = ROOT / "network-supplements/NET-13/minimal-ai-tool-pitch-deck.pptx"
PREVIEWS = SOURCE.parent / "previews"
OUT = ROOT / "output/network-round5-ai-tech-candidate-style-review.pdf"
BODY = ROOT / "output/network-round5-ai-tech-candidate-style-review-body.pdf"
REGISTRY = ROOT / "catalog/network-round5-ai-tech-candidates.json"


def package_metrics() -> dict[str, int]:
    with ZipFile(SOURCE) as archive:
        names = archive.namelist()
        slide_xml = [name for name in names if name.startswith("ppt/slides/slide") and name.endswith(".xml")]
        contents = [archive.read(name).decode("utf-8", "ignore") for name in slide_xml]
    return {
        "slide_count": len(slide_xml),
        "slides_with_timing_xml": sum("<p:timing" in text for text in contents),
        "slides_with_transition_xml": sum("<p:transition" in text for text in contents),
    }


def build_pdf() -> int:
    width, height = landscape((842, 595))
    margin, title_h = 24, 46
    cell_w = (width - 2 * margin) / 3
    cell_h = (height - title_h - 2 * margin) / 2
    images = sorted(PREVIEWS.glob("slide-*.png"))
    doc = canvas.Canvas(str(BODY), pagesize=(width, height))
    pages = 0
    for start in range(0, len(images), 6):
        pages += 1
        doc.setFillColorRGB(0.05, 0.07, 0.10)
        doc.rect(0, 0, width, height, fill=1, stroke=0)
        doc.setFillColorRGB(1, 1, 1)
        doc.setFont("Helvetica-Bold", 15)
        doc.drawString(margin, height - 24, "NET-13 | MINIMAL AI TOOL PITCH DECK")
        doc.setFillColorRGB(0.75, 0.80, 0.85)
        doc.setFont("Helvetica", 8)
        doc.drawRightString(width - margin, height - 24, f"SlidesCarnival | candidate only | slides {start + 1}-{min(start + 6, len(images))}")
        for idx, image_path in enumerate(images[start : start + 6]):
            col, row = idx % 3, idx // 3
            x = margin + col * cell_w
            top = height - title_h - margin - row * cell_h
            with Image.open(image_path) as image:
                image_w, image_h = image.size
            scale = min((cell_w - 10) / image_w, (cell_h - 24) / image_h)
            draw_w, draw_h = image_w * scale, image_h * scale
            draw_x, draw_y = x + (cell_w - draw_w) / 2, top - draw_h - 13
            doc.setFillColorRGB(0.20, 0.23, 0.27)
            doc.rect(x + 1, top - cell_h + 1, cell_w - 3, cell_h - 3, fill=1, stroke=0)
            doc.drawImage(ImageReader(str(image_path)), draw_x, draw_y, width=draw_w, height=draw_h)
            doc.setFillColorRGB(0.85, 0.88, 0.90)
            doc.setFont("Helvetica", 8)
            doc.drawString(x + 5, top - 11, f"SOURCE SLIDE {start + idx + 1}")
        doc.showPage()
    doc.save()
    reader = PdfReader(str(BODY))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.add_outline_item("NET-13 | Minimal AI Tool Pitch Deck", 0)
    with OUT.open("wb") as fp:
        writer.write(fp)
    return len(reader.pages)


def build_registry(pages: int) -> None:
    payload = {
        "schema_version": 1,
        "status": "candidate_only",
        "generated_at_utc": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "review_artifact": str(OUT),
        "contact_sheet_pages": pages,
        "policy": {
            "original_library_write": False,
            "production_auto_promotion": False,
            "single_source_deck_required": True,
            "license_review_required": True,
            "style_review_required": True,
            "require_explicit_user_approval_before_promotion": True,
        },
        "candidates": [{
            "id": "NET-13",
            "title": "Minimal AI Tool Pitch Deck",
            "source": "SlidesCarnival",
            "page_url": "https://www.slidescarnival.com/template/ai-presentation-tool-pitch-deck/66043",
            "pptx_url": "https://www.slidescarnival.com/download/sc00982/ai-presentation-tool-pitch-deck/pptx",
            "targets": ["ai-industry", "technical-product", "investor-pitch"],
            "missing_page_types": ["ai_capabilities", "product_overview", "team", "pitch_narrative"],
            "source_pptx": str(SOURCE),
            "previews_dir": str(PREVIEWS),
            "file_bytes": SOURCE.stat().st_size,
            "sha256": hashlib.sha256(SOURCE.read_bytes()).hexdigest(),
            "package": package_metrics(),
            "license_status": "manual_external_review_required",
            "ingest_status": "downloaded_structurally_verified_pending_user_style_review",
            "review_note": "覆盖AI能力、服务、团队等页面，但赛博机器人和电路背景模板感较强；只建议作为独立候选模板族，不作为默认AI/融资模板。",
        }],
    }
    REGISTRY.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    pages = build_pdf()
    build_registry(pages)
    print(json.dumps({"review_pdf": str(OUT), "registry": str(REGISTRY), "pages": pages}, ensure_ascii=False))
