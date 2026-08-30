from pathlib import Path
from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from pypdf import PdfReader, PdfWriter
from PIL import Image
root=Path(r'C:/ppt-creater');out=root/'output/network-round2-candidates-style-review.pdf';tmp=root/'output/network-round2-candidates-style-review-body.pdf'
families=[
 ('NET-06 | MEDELEY BUSINESS | INVESTOR / STRATEGY CANDIDATE','NET-06','SlidesMania','33 source slides'),
 ('NET-07 | MARKETING ANALYSIS | B2B / MARKET CANDIDATE','NET-07','SlidesMania','29 source slides'),
 ('NET-08 | INTERFACE FILE FOLDERS | TECH PRODUCT CANDIDATE','NET-08','SlidesMania','24 source slides'),
]
w,h=landscape((842,595));m=24;th=46;cw=(w-2*m)/3;ch=(h-th-2*m)/2;c=canvas.Canvas(str(tmp),pagesize=(w,h));mp={};pno=0
for title,ident,source,meta in families:
 imgs=sorted((root/'network-supplements'/ident/'previews').glob('slide-*.png'));mp[title]=pno
 for start in range(0,len(imgs),6):
  pno+=1;c.setFillColorRGB(.05,.07,.10);c.rect(0,0,w,h,fill=1,stroke=0);c.setFillColorRGB(1,1,1);c.setFont('Helvetica-Bold',15);c.drawString(m,h-24,title);c.setFillColorRGB(.75,.80,.85);c.setFont('Helvetica',8);c.drawRightString(w-m,h-24,f'{source} | candidate only | {meta} | slides {start+1}-{min(start+6,len(imgs))}')
  for j,png in enumerate(imgs[start:start+6]):
   col,row=j%3,j//3;x=m+col*cw;top=h-th-m-row*ch
   with Image.open(png) as im:iw,ih=im.size
   sc=min((cw-10)/iw,(ch-24)/ih);dw,dh=iw*sc,ih*sc;dx=x+(cw-dw)/2;dy=top-dh-13;c.setFillColorRGB(.2,.23,.27);c.rect(x+1,top-ch+1,cw-3,ch-3,fill=1,stroke=0);c.drawImage(ImageReader(str(png)),dx,dy,width=dw,height=dh);c.setFillColorRGB(.85,.88,.90);c.setFont('Helvetica',8);c.drawString(x+5,top-11,f'SOURCE SLIDE {start+j+1}')
  c.showPage()
c.save();r=PdfReader(str(tmp));wr=PdfWriter()
for p in r.pages:wr.add_page(p)
for t,i in mp.items():wr.add_outline_item(t,i)
with out.open('wb') as f:wr.write(f)
print(out,len(r.pages))
