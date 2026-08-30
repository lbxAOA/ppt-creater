#!/usr/bin/env python
from __future__ import annotations

import argparse
import json
from pathlib import Path


def load(path: Path):
    return [json.loads(line) for line in path.read_text(encoding='utf-8').splitlines() if line.strip()]


def family_id(category: str, index: int) -> str:
    normalized = ''.join(ch.lower() if ch.isascii() and ch.isalnum() else '-' for ch in category).strip('-')
    return f"{normalized or 'template'}-{index:03d}"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('catalog', type=Path)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--per-category', type=int, default=5)
    args = parser.parse_args()
    records = [r for r in load(args.catalog) if r.get('status') == 'ok' and r.get('read_status') == 'complete']
    by_category = {}
    for record in records:
        by_category.setdefault(record['category'], []).append(record)
    families = []
    for category, entries in sorted(by_category.items()):
        entries.sort(key=lambda r: (r.get('has_animation_xml', False), r.get('quality_score', 0), r.get('chart_count', 0) + r.get('table_count', 0)), reverse=True)
        for index, record in enumerate(entries[:args.per_category], start=1):
            patterns = record.get('slide_patterns', {})
            families.append({
                'family_id': family_id(category, index),
                'name': f"候选：{category} #{index}",
                'status': 'needs_slot_annotation',
                'source_template': record['absolute_path'],
                'source_relative_path': record['relative_path'],
                'use_cases': [category],
                'visual_tags': [category, 'animated' if record.get('has_animation_xml') else 'static'],
                'aspect_ratio': f"{record['slide_size_inches']['width']}:{record['slide_size_inches']['height']}",
                'theme': record.get('theme', {}),
                'evidence': {
                    'quality_score': record.get('quality_score'),
                    'slide_count': record.get('slide_count'),
                    'has_animation_xml': record.get('has_animation_xml'),
                    'has_transition_xml': record.get('has_transition_xml'),
                    'slide_patterns': patterns,
                    'charts': record.get('chart_count'),
                    'tables': record.get('table_count'),
                },
                'animation_policy': {
                    'allowed_patterns': ['preserve_original'] if record.get('has_animation_xml') else ['none'],
                    'requires_powerpoint_playback_verification': True,
                },
                'slides': [],
                'next_action': 'Open selected source in PowerPoint, name reusable shapes slot_*, then register slide definitions and animation locks.'
            })
    output = {'schema_version': 1, 'families': families}
    args.output.write_text(json.dumps(output, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'families': len(families), 'output': str(args.output)}, ensure_ascii=False))

if __name__ == '__main__':
    main()
