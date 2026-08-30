"""Build a human-review PDF and provenance registry for round-4 network PPTX candidates."""
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
OUT = ROOT / "output/network-round4-executive-candidates-style-review.pdf"
BODY = ROOT / "output/network-round4-executive-candidates-style-review-body.pdf"
REGISTRY = ROOT / "catalog/network-round4-executive-candidates.json"

CANDIDATES = [
    {
        "id": "NET-09",
        "title": "KPI Scorecard Infographic",
        "file": "kpi-scorecard-infographic.pptx",
        "targets": ["data-boardroom", "generic-corporate"],
        "missing_page_types": ["kpi_scorecard", "executive_summary", "trend_callout"],
        "source": "SlidesCarnival",
        "page_url": "https://www.slidescarnival.com/template/kpi-scorecard-infographic/",
        "pptx_url": "https://www.slidescarnival.com/download/sc02590/kpi-scorecard-infographic/pptx",
        "review_note": "单页信息图组件；不作为完整模板族替换候选。资源页与致谢页不进入生产页面。",
    },
    {
        "id": "NET-10",
        "title": "Simple Quarterly Business Review Slides",
        "file": "quarterly-business-review.pptx",
        "targets": ["data-boardroom", "generic-corporate", "investor-pitch"],
        "missing_page_types": ["quarterly_review", "kpi", "revenue_breakdown", "next_steps"],
        "source": "SlidesCarnival",
        "page_url": "https://www.slidescarnival.com/template/quarterly-business-review-slides/329420",
        "pptx_url": "https://www.slidescarnival.com/download/sc03261/quarterly-business-review-slides/pptx",
        "review_note": "完整经营复盘页面系统；视觉质量须由用户判断，封面和业务场景偏通用。",
    },
    {
        "id": "NET-12",
        "title": "Financial Management for Small Businesses Consulting",
        "file": "financial-management-consulting.pptx",
        "targets": ["data-boardroom", "generic-corporate", "investor-pitch"],
        "missing_page_types": ["financial_consulting", "business_case", "timeline", "action_plan"],
        "source": "SlidesCarnival",
        "page_url": "https://www.slidescarnival.com/template/financial-management-for-small-businesses-consulting/161609",
        "pptx_url": "https://www.slidescarnival.com/download/sc01551/financial-management-for-small-businesses-consulting/pptx",
        "review_note": "完整咨询型页面系统；高对比荧光绿适合年轻化场景，不宜直接作为严肃高管汇报默认模板。",
    },
]


def package_metrics(path: Path) -> dict[str, int]:
    with ZipFile(path) as archive:
        names = archive.namelist()
    slides = sum(name.startswith("ppt/slides/slide") and name.endswith(".xml") for name in names)
    timings = sum(name.startswith("ppt/slides/slide") and name.endswith(".xml") and "<p:timing" in ZipFile(path).read(name).decode("utf-8", "ignore") for name in names)
    transitions = sum(name.startswith("ppt/slides/slide") and name.endswith(".xml") and "<p:transition" in ZipFile(path).read(name).decode("utf-8", "ignore") for name in names)
    return {"slide_count": slides, "slides_with_timing_xml": timings, "slides_with_transition_xml": transitions}


def build_pdf() -> int:
    width, height = landscape((842, 595))
    margin, title_h = 24, 46
    cell_w = (width - 2 * margin) / 3
    cell_h = (height - title_h - 2 * margin) / 2
    doc = canvas.Canvas(str(BODY), pagesize=(width, height))
    outline_pages: dict[str, int] = {}
    page_no = 0

    for candidate in CANDIDATES:
        images = sorted((ROOT / "network-supplements" / candidate["id"] / "previews").glob("slide-*.png"))
        outline_pages[f'{candidate["id"]} | {candidate["title"]}'] = page_no
        for start in range(0, len(images), 6):
            page_no += 1
            doc.setFillColorRGB(0.05, 0.07, 0.10)
            doc.rect(0, 0, width, height, fill=1, stroke=0)
            doc.setFillColorRGB(1, 1, 1)
            doc.setFont("Helvetica-Bold", 15)
            doc.drawString(margin, height - 24, f'{candidate["id"]} | {candidate["title"]}')
            doc.setFillColorRGB(0.75, 0.80, 0.85)
            doc.setFont("Helvetica", 8)
            doc.drawRightString(
                width - margin,
                height - 24,
                f'{candidate["source"]} | candidate only | slides {start + 1}-{min(start + 6, len(images))}',
            )
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
    for title, page_index in outline_pages.items():
        writer.add_outline_item(title, page_index)
    with OUT.open("wb") as fp:
        writer.write(fp)
    return len(reader.pages)


def build_registry(contact_sheet_pages: int) -> None:
    payload = {
        "schema_version": 1,
        "status": "candidate_only",
        "generated_at_utc": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "review_artifact": str(OUT),
        "contact_sheet_pages": contact_sheet_pages,
        "policy": {
            "destination": "C:/ppt-creater/network-supplements",
            "original_library_write": False,
            "production_auto_promotion": False,
            "single_source_deck_required": True,
            "license_review_required": True,
            "style_review_required": True,
            "require_explicit_user_approval_before_promotion": True,
        },
        "candidates": [],
    }
    for candidate in CANDIDATES:
        source = ROOT / "network-supplements" / candidate["id"] / candidate["file"]
        payload["candidates"].append({
            **candidate,
            "source_pptx": str(source),
            "file_bytes": source.stat().st_size,
            "sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
            "package": package_metrics(source),
            "previews_dir": str(source.parent / "previews"),
            "license_status": "manual_external_review_required",
            "ingest_status": "downloaded_structurally_verified_pending_user_style_review",
        })
    REGISTRY.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    pages = build_pdf()
    build_registry(pages)
    print(json.dumps({"review_pdf": str(OUT), "registry": str(REGISTRY), "pages": pages}, ensure_ascii=False))
