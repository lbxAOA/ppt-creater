#!/usr/bin/env python
"""Build contact sheets from exported candidate template previews."""
from __future__ import annotations
import json
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT=Path(r'C:/ppt-creater/candidate-previews')
OUT=ROOT/'contact-sheets'
OUT.mkdir(exist_ok=True)
manifest=json.loads((ROOT/'export-manifest.json').read_text(encoding='utf-8-sig'))
font=ImageFont.truetype('C:/Windows/Fonts/msyh.ttc', 20)
small=ImageFont.truetype('C:/Windows/Fonts/msyh.ttc', 15)
for item in manifest:
    if item.get('status')!='ok': continue
    images=[]
    for p in item['exports']:
        try:
            im=Image.open(p).convert('RGB').resize((320,180))
            images.append((Path(p).name,im))
        except: pass
    canvas=Image.new('RGB',(1040,80+((len(images)+2)//3)*225),'white')
    d=ImageDraw.Draw(canvas)
    d.text((18,14),f"{item['family_id']} | {Path(item['source']).name} | {item['slide_count']} slides",font=font,fill='black')
    for k,(name,im) in enumerate(images):
        x=18+(k%3)*340; y=65+(k//3)*225
        canvas.paste(im,(x,y)); d.text((x,y+184),name,font=small,fill='black')
    canvas.save(OUT/f"{item['family_id']}.jpg",quality=90)
print(json.dumps({'sheets':len(list(OUT.glob('*.jpg'))),'dir':str(OUT)},ensure_ascii=False))
