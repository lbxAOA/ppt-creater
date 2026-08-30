#!/usr/bin/env python
from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('registry', type=Path)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--mode', choices=['hardlink', 'copy'], default='hardlink')
    args = parser.parse_args()
    data = json.loads(args.registry.read_text(encoding='utf-8'))
    imported=[]
    for family in data['families']:
        src=Path(family['source_template'])
        dest=args.output / family['family_id'] / src.name
        dest.parent.mkdir(parents=True, exist_ok=True)
        if not dest.exists():
            try:
                if args.mode == 'hardlink': dest.hardlink_to(src)
                else: shutil.copy2(src,dest)
            except OSError:
                shutil.copy2(src,dest)
        family['imported_template']=str(dest)
        imported.append(str(dest))
    args.registry.write_text(json.dumps(data,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({'imported':len(imported),'output':str(args.output),'mode':args.mode},ensure_ascii=False))

if __name__ == '__main__': main()
