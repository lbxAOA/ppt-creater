#!/usr/bin/env python
"""Create an AI-industry reporting family from curated local PPTX assets.

This creates a production design contract and stores original selected source
files as hard-linked references. It does not alter any original template.
"""
from __future__ import annotations
import json
from pathlib import Path

ROOT = Path(r"C:/ppt-creater")
source = Path(r"C:/PPT模板/01_应用场景/02_汇报与答辩/模板/101_16比9比例.pptx")
chart_source = Path(r"C:/PPT模板/04_设计素材/01_图表与图标/模板/018_精品图表精品推荐.pptx")
family_dir = ROOT / "template-library" / "ai-industry-navy-v1"
family_dir.mkdir(parents=True, exist_ok=True)

for src, dest_name in [(source, "source-deck.pptx"), (chart_source, "chart-reference.pptx")]:
    dest = family_dir / dest_name
    if not dest.exists():
        try:
            dest.hardlink_to(src)
        except OSError:
            import shutil
            shutil.copy2(src, dest)

spec = {
  "family_id": "ai-industry-navy-v1",
  "name": "AI 行业研究｜深蓝科技",
  "status": "production_static_v1",
  "purpose": ["AI行业研究", "技术方案", "投资人技术简报", "具身智能产业分析"],
  "source_assets": [str(family_dir / "source-deck.pptx"), str(family_dir / "chart-reference.pptx")],
  "aspect_ratio": "16:9",
  "theme": {
    "background": "F5F8FC", "navy": "0B1F3A", "navy_2": "123A63",
    "ink": "172033", "muted": "64748B", "accent": "00A6A6",
    "accent_light": "DFF6F4", "orange": "F59E0B", "line": "DCE5EF",
    "font_cn": "Microsoft YaHei", "font_latin": "Aptos"
  },
  "typography": {"title": 28, "section": 24, "body": 17, "caption": 10, "kpi": 32},
  "rules": {
    "one_family_per_deck": True, "safe_margin_inches": 0.55,
    "max_title_lines": 2, "max_body_bullets": 5, "min_body_pt": 16,
    "title_style": "结论先行", "source_line_required": True,
    "motion": "静态生产版；动画页需在 PowerPoint 中完成 slot 标注和播放验收后启用"
  },
  "slide_types": [
    "cover", "executive-summary", "market-map", "technology-stack", "value-chain",
    "investment-logic", "risk-matrix", "roadmap", "recommendation", "sources"
  ]
}
(family_dir / "design-spec.json").write_text(json.dumps(spec, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
print(json.dumps({"family": spec["family_id"], "path": str(family_dir)}, ensure_ascii=False))
