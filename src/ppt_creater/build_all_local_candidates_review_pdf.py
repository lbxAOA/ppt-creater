#!/usr/bin/env python
"""Build a user-review PDF for all 40 curated local candidates.

Each candidate gets one contact-sheet page using the already exported six-slide
samples. This is intentionally a review artifact, not a production promotion.
"""
from __future__ import annotations

import json
from pathlib import Path

from PIL import Image
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

ROOT = Path(r"C:/ppt-creater")
FAMILIES = ROOT / "catalog" / "families.json"
ROUTING = ROOT / "catalog" / "generation-template-library.json"
SHEETS = ROOT / "candidate-previews" / "contact-sheets"
OUT = ROOT / "output" / "all-local-candidates-review.pdf"
TMP = ROOT / "output" / "all-local-candidates-review-body.pdf"

W, H = landscape((842, 595))
MARGIN = 24


def text_or_dash(value: object) -> str:
    return str(value) if value not in (None, "", []) else "—"


def main() -> None:
    families = json.loads(FAMILIES.read_text(encoding="utf-8-sig"))["families"]
    reviews = json.loads(ROUTING.read_text(encoding="utf-8-sig")).get("family_reviews", {})
    missing = [family["family_id"] for family in families if not (SHEETS / f"{family['family_id']}.jpg").exists()]
    if missing:
        raise RuntimeError(f"Missing candidate contact sheets: {', '.join(missing)}")

    c = canvas.Canvas(str(TMP), pagesize=(W, H))
    outline = {}
    for number, family in enumerate(families, start=1):
        family_id = family["family_id"]
        evidence = family.get("evidence", {})
        review = reviews.get(family_id, {})
        sheet = SHEETS / f"{family_id}.jpg"
        with Image.open(sheet) as image:
            iw, ih = image.size
        max_w, max_h = W - 2 * MARGIN, H - 150
        scale = min(max_w / iw, max_h / ih)
        dw, dh = iw * scale, ih * scale
        x, y = (W - dw) / 2, 58

        c.setFillColorRGB(0.05, 0.07, 0.10)
        c.rect(0, 0, W, H, fill=1, stroke=0)
        c.setFillColorRGB(1, 1, 1)
        c.setFont("Helvetica-Bold", 16)
        c.drawString(MARGIN, H - 26, f"{number:02d}/40  {family_id}  |  {family['name']}")
        c.setFillColorRGB(0.75, 0.80, 0.85)
        c.setFont("Helvetica", 8)
        source = Path(family["source_template"]).name
        c.drawString(MARGIN, H - 43, f"SOURCE: {source}  |  {family['source_relative_path']}")
        meta = (
            f"slides: {text_or_dash(evidence.get('slide_count'))}  |  "
            f"structural score: {text_or_dash(evidence.get('quality_score'))}  |  "
            f"charts/tables: {text_or_dash(evidence.get('charts'))}/{text_or_dash(evidence.get('tables'))}  |  "
            f"animation XML: {text_or_dash(evidence.get('has_animation_xml'))}"
        )
        c.drawString(MARGIN, H - 56, meta)
        review_meta = (
            f"prior review: {text_or_dash(review.get('status'))}  |  "
            f"visual score: {text_or_dash(review.get('visual_score'))}  |  "
            f"targets: {', '.join(review.get('targets', [])) or 'unassigned'}  |  "
            f"risks: {', '.join(review.get('risks', [])) or 'not yet recorded'}"
        )
        c.drawString(MARGIN, H - 69, review_meta[:180])
        c.drawImage(ImageReader(str(sheet)), x, y, width=dw, height=dh)
        c.setFillColorRGB(0.75, 0.80, 0.85)
        c.setFont("Helvetica-Oblique", 8)
        c.drawString(MARGIN, 30, "Review only: no candidate has been promoted or routed by this document.")
        outline[family_id] = number - 1
        c.showPage()
    c.save()

    reader, writer = PdfReader(str(TMP)), PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    for family_id, index in outline.items():
        writer.add_outline_item(family_id, index)
    with OUT.open("wb") as handle:
        writer.write(handle)
    print(f"{OUT} pages={len(reader.pages)} candidates={len(families)}")


if __name__ == "__main__":
    main()
