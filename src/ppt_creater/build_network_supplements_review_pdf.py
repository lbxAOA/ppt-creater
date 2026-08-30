from pathlib import Path
from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from pypdf import PdfReader, PdfWriter
from PIL import Image

root=Path(r'C:/ppt-creater')
out=root/'output/network-supplements-net02-net04-style-review.pdf'; tmp=root/'output/network-supplements-net02-net04-style-review-body.pdf'
families=[
 ('NET-02','NET-02 | SIMPLE & PROFESSIONAL | DATA / COMPANY CANDIDATE','SlidesMania','simple-professional.pptx'),
 ('NET-03','NET-03 | MY PORTFOLIO | CAREER CANDIDATE','SlidesMania','my-portfolio.pptx'),
 ('NET-04','NET-04 | AMELIA | WEDDING CANDIDATE','SlidesMania','amelia-wedding.pptx'),
]
w,h=landscape((842,595));m=24;th=46;cw=(w-2*m)/3;ch=(h-th-2*m)/2
c=canvas.Canvas(str(tmp),pagesize=(w,h));page_map={};page_no=0
for ident,title,source,file in families:
 imgs=sorted((root/'network-supplements'/ident/'previews').glob('slide-*.png'));page_map[title]=page_no
 for start in range(0,len(imgs),6):
  page_no+=1;c.setFillColorRGB(.05,.07,.10);c.rect(0,0,w,h,fill=1,stroke=0)
  c.setFillColorRGB(1,1,1);c.setFont('Helvetica-Bold',16);c.drawString(m,h-24,title)
  c.setFillColorRGB(.75,.80,.85);c.setFont('Helvetica',8);c.drawRightString(w-m,h-24,f'{source} | candidate only | source: {file} | slides {start+1}-{min(start+6,len(imgs))}')
  for idx,png in enumerate(imgs[start:start+6]):
   col,row=idx%3,idx//3;x=m+col*cw;top=h-th-m-row*ch
   with Image.open(png) as im:iw,ih=im.size
   sc=min((cw-10)/iw,(ch-24)/ih);dw,dh=iw*sc,ih*sc;dx=x+(cw-dw)/2;dy=top-dh-13
   c.setFillColorRGB(.2,.23,.27);c.rect(x+1,top-ch+1,cw-3,ch-3,fill=1,stroke=0);c.drawImage(ImageReader(str(png)),dx,dy,width=dw,height=dh)
   c.setFillColorRGB(.85,.88,.90);c.setFont('Helvetica',8);c.drawString(x+5,top-11,f'SOURCE SLIDE {start+idx+1}')
  c.showPage()
c.save();r=PdfReader(str(tmp));writer=PdfWriter()
for p in r.pages:writer.add_page(p)
for title,num in page_map.items():writer.add_outline_item(title,num)
with out.open('wb') as f:writer.write(f)
print(out, len(r.pages))
