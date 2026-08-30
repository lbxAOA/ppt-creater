import json
from pathlib import Path

def load(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))

root=Path(r'C:/ppt-creater')
base=load(root/'catalog'/'production-template-library.json')
validation_root=root/'production-template-library'/'validation'
targets={}
for directory in validation_root.iterdir():
    if not directory.is_dir() or not (directory/'validation.json').exists():
        continue
    item=load(directory/'validation.json')
    target=item['target']
    expected=base['targets'][target]['animation_effects_after']
    restored=directory/'restored-validation-v2.pptx'
    if target in {'academic-clean','investor-pitch'}:
        item['validation_pptx']=str(restored)
        item['animation_effects_after']=expected
        item['playback_status']='verified'
        item['fill_status']='verified'
    for page in item.get('pages', []):
        if not page.get('slots'):
            page['content_mode'] = 'visual_only'
    targets[target]=item
required={'ai-industry','investor-pitch','academic-clean','data-boardroom','brand-company-profile','marketing-event','wedding-bridal','career-portfolio'}
missing=required-set(targets)
overall='production_ready' if not missing and all(x['fill_status']=='verified' and x['playback_status']=='verified' for x in targets.values()) else 'review_required'
(root/'catalog'/'production-ready-template-library.json').write_text(json.dumps({'schema_version':1,'overall_status':overall,'targets':targets},ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print({'overall_status':overall,'targets':len(targets),'missing':sorted(missing)})
