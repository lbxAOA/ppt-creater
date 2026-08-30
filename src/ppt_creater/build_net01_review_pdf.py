from pathlib import Path
from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from pypdf import PdfReader, PdfWriter

root=Path(r'C:/ppt-creater')
source=root/'network-supplements/NET-01/previews'
out=root/'output/network-supplement-net01-style-review.pdf'
tmp=root/'output/network-supplement-net01-style-review-body.pdf'
imgs=sorted(source.glob('slide-*.png'))
w,h=landscape((842,595)); margin=24; title_h=46; cols=3; rows=2; cell_w=(w-2*margin)/cols;cell_h=(h-title_h-2*margin)/rows
c=canvas.Canvas(str(tmp),pagesize=(w,h))
for start in range(0,len(imgs),6):
    c.setFillColorRGB(.05,.07,.10);c.rect(0,0,w,h,fill=1,stroke=0)
    c.setFillColorRGB(1,1,1);c.setFont('Helvetica-Bold',16);c.drawString(margin,h-24,'NET-01 | PITCH WITH CONFIDENCE | ISOLATED CANDIDATE')
    c.setFillColorRGB(.75,.80,.85);c.setFont('Helvetica',8);c.drawRightString(w-margin,h-24,f'SlidesMania | candidate only | slides {start+1}-{min(start+6,len(imgs))}')
    for idx,png in enumerate(imgs[start:start+6]):
        col,row=idx%cols,idx//cols;x=margin+col*cell_w;top=h-title_h-margin-row*cell_h
        from PIL import Image
        with Image.open(png) as im:iw,ih=im.size
        scale=min((cell_w-10)/iw,(cell_h-24)/ih);dw,dh=iw*scale,ih*scale;dx=x+(cell_w-dw)/2;dy=top-dh-13
        c.setFillColorRGB(.2,.23,.27);c.rect(x+1,top-cell_h+1,cell_w-3,cell_h-3,fill=1,stroke=0)
        c.drawImage(ImageReader(str(png)),dx,dy,width=dw,height=dh)
        c.setFillColorRGB(.85,.88,.90);c.setFont('Helvetica',8);c.drawString(x+5,top-11,f'SOURCE SLIDE {start+idx+1}')
    c.showPage()
c.save()
r=PdfReader(str(tmp));writer=PdfWriter()
for p in r.pages:writer.add_page(p)
writer.add_outline_item('NET-01 Pitch with Confidence',0)
with out.open('wb') as f:writer.write(f)
print(out)
