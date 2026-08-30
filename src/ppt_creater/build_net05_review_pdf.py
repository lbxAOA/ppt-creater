from pathlib import Path
from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from pypdf import PdfReader, PdfWriter
from PIL import Image
root=Path(r'C:/ppt-creater');out=root/'output/network-supplement-net05-style-review.pdf';tmp=root/'output/network-supplement-net05-style-review-body.pdf';imgs=sorted((root/'network-supplements/NET-05/previews').glob('slide-*.png'))
w,h=landscape((842,595));m=24;th=46;cw=(w-2*m)/3;ch=(h-th-2*m)/2;c=canvas.Canvas(str(tmp),pagesize=(w,h))
for start in range(0,len(imgs),6):
 c.setFillColorRGB(.05,.07,.10);c.rect(0,0,w,h,fill=1,stroke=0);c.setFillColorRGB(1,1,1);c.setFont('Helvetica-Bold',16);c.drawString(m,h-24,'NET-05 | TRAVEL & TOURISM | EVENT / CULTURE CANDIDATE');c.setFillColorRGB(.75,.80,.85);c.setFont('Helvetica',8);c.drawRightString(w-m,h-24,f'SlidesCarnival | candidate only | slides {start+1}-{min(start+6,len(imgs))}')
 for idx,png in enumerate(imgs[start:start+6]):
  col,row=idx%3,idx//3;x=m+col*cw;top=h-th-m-row*ch
  with Image.open(png) as im:iw,ih=im.size
  sc=min((cw-10)/iw,(ch-24)/ih);dw,dh=iw*sc,ih*sc;dx=x+(cw-dw)/2;dy=top-dh-13;c.setFillColorRGB(.2,.23,.27);c.rect(x+1,top-ch+1,cw-3,ch-3,fill=1,stroke=0);c.drawImage(ImageReader(str(png)),dx,dy,width=dw,height=dh);c.setFillColorRGB(.85,.88,.90);c.setFont('Helvetica',8);c.drawString(x+5,top-11,f'SOURCE SLIDE {start+idx+1}')
 c.showPage()
c.save();r=PdfReader(str(tmp));wri=PdfWriter()
for p in r.pages:wri.add_page(p)
wri.add_outline_item('NET-05 Travel & Tourism',0)
with out.open('wb') as f:wri.write(f)
print(out,len(r.pages))
