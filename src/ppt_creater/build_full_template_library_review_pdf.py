from pathlib import Path
from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from pypdf import PdfReader, PdfWriter
from PIL import Image

ROOT = Path(r"C:/ppt-creater")
OUT = ROOT / "output" / "full-production-template-library-review.pdf"
TMP = ROOT / "output" / "full-production-template-library-review-body.pdf"

families = [
    ("L01 Modern Report | AI / Investor / Academic", ROOT / "output/style-review-images/01-modern-report", "Local source: 101_16比9比例.pptx", ["ai-industry", "investor-pitch", "academic-clean"]),
    ("L02 Data & Boardroom | Corporate", ROOT / "output/style-review-images/02-data-boardroom", "Local source: 001_汇报总结模板.pptx", ["data-boardroom", "generic-corporate"]),
    ("L03 Company Profile | Brand", ROOT / "output/style-review-images/03-brand-profile", "Local source: 015_扁平化模板.pptx", ["brand-company-profile"]),
    ("L04 Event & Culture | Local", ROOT / "output/style-review-images/04-event-culture", "Local source: 011_PPT模板宣传类.pptx", ["marketing-event", "culture-tourism"]),
    ("L05 Wedding | Local", ROOT / "output/style-review-images/05-wedding", "Local source: 003_PPT模板.pptx", ["wedding-bridal"]),
    ("L06 Career & Portfolio | Local", ROOT / "output/style-review-images/06-portfolio", "Local source: 039_模板组.pptx", ["career-portfolio"]),
    ("N01 Investor Modern | NET-01", ROOT / "network-supplements/NET-01/previews", "Network candidate: SlidesMania", ["ai-industry", "investor-pitch"]),
    ("N02 Company Professional | NET-02", ROOT / "network-supplements/NET-02/previews", "Network candidate: SlidesMania", ["brand-company-profile", "data-boardroom", "generic-corporate"]),
    ("N03 Career Portfolio | NET-03", ROOT / "network-supplements/NET-03/previews", "Network candidate: SlidesMania", ["career-portfolio"]),
    ("N04 Wedding Amelia | NET-04", ROOT / "network-supplements/NET-04/previews", "Network candidate: SlidesMania", ["wedding-bridal"]),
]
W, H = landscape((842, 595)); M = 24; TITLE_H = 46; CW = (W-2*M)/3; CH = (H-TITLE_H-2*M)/2
c = canvas.Canvas(str(TMP), pagesize=(W,H)); page_map = {}; page_no = 0
for title, folder, source, routes in families:
    images = sorted(folder.glob('slide-*.png'))
    if not images:
        raise RuntimeError(f'No previews: {folder}')
    page_map[title] = page_no
    for start in range(0, len(images), 6):
        page_no += 1
        c.setFillColorRGB(.05,.07,.10); c.rect(0,0,W,H,fill=1,stroke=0)
        c.setFillColorRGB(1,1,1); c.setFont('Helvetica-Bold',15); c.drawString(M,H-24,title)
        c.setFillColorRGB(.75,.80,.85); c.setFont('Helvetica',8)
        c.drawRightString(W-M,H-24,f'{source} | routes: {" / ".join(routes)} | source slides {start+1}-{min(start+6,len(images))}')
        for j,png in enumerate(images[start:start+6]):
            col,row=j%3,j//3; x=M+col*CW; top=H-TITLE_H-M-row*CH
            with Image.open(png) as im: iw,ih=im.size
            scale=min((CW-10)/iw,(CH-24)/ih); dw,dh=iw*scale,ih*scale; dx=x+(CW-dw)/2; dy=top-dh-13
            c.setFillColorRGB(.2,.23,.27); c.rect(x+1,top-CH+1,CW-3,CH-3,fill=1,stroke=0)
            c.drawImage(ImageReader(str(png)),dx,dy,width=dw,height=dh)
            c.setFillColorRGB(.85,.88,.90);c.setFont('Helvetica',8);c.drawString(x+5,top-11,f'SOURCE SLIDE {start+j+1}')
        c.showPage()
c.save()
reader=PdfReader(str(TMP));writer=PdfWriter()
for page in reader.pages: writer.add_page(page)
for title,index in page_map.items(): writer.add_outline_item(title,index)
with OUT.open('wb') as fh: writer.write(fh)
print(f'{OUT} pages={len(reader.pages)} families={len(families)}')
