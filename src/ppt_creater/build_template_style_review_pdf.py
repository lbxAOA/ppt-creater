from pathlib import Path
import json
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from pypdf import PdfReader, PdfWriter

ROOT = Path(r"C:/ppt-creater")
IMAGES = ROOT / "output" / "style-review-images"
OUT = ROOT / "output" / "production-template-style-review.pdf"
MANIFEST = ROOT / "catalog" / "project-template-sources.json"
TMP = ROOT / "output" / "production-template-style-review-body.pdf"
FONT = Path(r"C:/Windows/Fonts/msyh.ttc")

source_manifest = {
    item["id"]: item
    for item in json.loads(MANIFEST.read_text(encoding="utf-8-sig"))["sources"]
}

families = [
    ("01-modern-report", "01 现代汇报 / AI / 投资人 / 科研", "modern-report-ai-investor-academic", ["ai-industry", "investor-pitch", "academic-clean"]),
    ("02-data-boardroom", "02 数据经营 / 管理层 / 企业通用", "data-boardroom-corporate", ["data-boardroom", "generic-corporate"]),
    ("03-brand-profile", "03 企业介绍 / 品牌提案", "brand-company-profile", ["brand-company-profile"]),
    ("04-event-culture", "04 专业活动 / 文化文旅", "marketing-event-culture", ["marketing-event", "culture-tourism"]),
    ("05-wedding", "05 婚礼 / 婚庆", "wedding-bridal", ["wedding-bridal"]),
    ("06-portfolio", "06 简历 / 个人作品集", "career-portfolio", ["career-portfolio"]),
]

PAGE_W, PAGE_H = landscape((842, 595))
MARGIN = 24
TITLE_H = 46
cols, rows = 3, 2
CELL_W = (PAGE_W - 2 * MARGIN) / cols
CELL_H = (PAGE_H - TITLE_H - 2 * MARGIN) / rows

font_title = ImageFont.truetype(str(FONT), 28)
font_small = ImageFont.truetype(str(FONT), 18)

c = canvas.Canvas(str(TMP), pagesize=(PAGE_W, PAGE_H))
page_map = {}
page_num = 0
for family_id, title, source_id, targets in families:
    source_name = Path(source_manifest[source_id]["original_relative_path"]).name
    page_map[title] = page_num
    imgs = sorted((IMAGES / family_id).glob("slide-*.png"))
    for start in range(0, len(imgs), 6):
        page_num += 1
        c.setFillColorRGB(0.05, 0.07, 0.10)
        c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
        review_header = {
            "01 现代汇报 / AI / 投资人 / 科研": "01 MODERN REPORT | AI / INVESTOR / ACADEMIC",
            "02 数据经营 / 管理层 / 企业通用": "02 DATA & BOARDROOM | CORPORATE",
            "03 企业介绍 / 品牌提案": "03 COMPANY PROFILE | BRAND PROPOSAL",
            "04 专业活动 / 文化文旅": "04 EVENT | CULTURE & TOURISM",
            "05 婚礼 / 婚庆": "05 WEDDING & BRIDAL",
            "06 简历 / 个人作品集": "06 CAREER & PORTFOLIO",
        }[title]
        c.setFillColorRGB(1, 1, 1)
        c.setFont("Helvetica-Bold", 16)
        c.drawString(MARGIN, PAGE_H - 24, review_header)
        c.setFont("Helvetica", 8)
        c.setFillColorRGB(0.75, 0.80, 0.85)
        c.drawRightString(PAGE_W - MARGIN, PAGE_H - 24, f"SOURCE: {source_name}  |  ID: {source_id}  |  ROUTES: {' / '.join(targets)}  |  SLIDES {start+1}-{min(start+6, len(imgs))}")
        for idx, png in enumerate(imgs[start:start+6]):
            col, row = idx % cols, idx // cols
            x = MARGIN + col * CELL_W
            top = PAGE_H - TITLE_H - MARGIN - row * CELL_H
            with Image.open(png) as im:
                im = im.convert("RGB")
                iw, ih = im.size
            scale = min((CELL_W - 10) / iw, (CELL_H - 24) / ih)
            dw, dh = iw * scale, ih * scale
            dx = x + (CELL_W - dw) / 2
            dy = top - dh - 13
            c.setFillColorRGB(0.2, 0.23, 0.27)
            c.rect(x + 1, top - CELL_H + 1, CELL_W - 3, CELL_H - 3, fill=1, stroke=0)
            c.drawImage(ImageReader(str(png)), dx, dy, width=dw, height=dh)
            c.setFillColorRGB(0.85, 0.88, 0.90)
            c.setFont("Helvetica", 8)
            c.drawString(x + 5, top - 11, f"SOURCE SLIDE {start + idx + 1}")
        c.showPage()
c.save()

reader = PdfReader(str(TMP)); writer = PdfWriter()
for page in reader.pages:
    writer.add_page(page)
for title, page_index in page_map.items():
    writer.add_outline_item(title, page_index)
with OUT.open("wb") as f:
    writer.write(f)
print(f"{OUT} pages={len(reader.pages)}")
