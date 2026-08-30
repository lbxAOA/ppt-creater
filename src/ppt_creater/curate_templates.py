#!/usr/bin/env python
from __future__ import annotations

import argparse
import json
import shutil
from collections import Counter
from pathlib import Path


def load_jsonl(path: Path):
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]


def select(records):
    records = [record for record in records if record.get("status") == "ok"]
    selected = []
    selected_paths = set()

    # Preserve every animation-bearing deck as a review candidate, up to a sane
    # cap per top-level category. Animation cannot be inferred safely later.
    by_category = {}
    for record in records:
        by_category.setdefault(record["category"], []).append(record)
    for category, candidates in by_category.items():
        candidates.sort(key=lambda item: item.get("quality_score", 0), reverse=True)
        animated = [item for item in candidates if item.get("has_animation_xml")]
        for item in animated[:12]:
            if item["relative_path"] not in selected_paths:
                selected.append(item)
                selected_paths.add(item["relative_path"])
        # Add broadly useful/high-scoring static sources to support full decks.
        for item in candidates:
            if len([value for value in selected if value["category"] == category]) >= 20:
                break
            if item["relative_path"] not in selected_paths:
                selected.append(item)
                selected_paths.add(item["relative_path"])

    # Guarantee high-value global candidates even if a category's selection was tight.
    for item in sorted(records, key=lambda item: item.get("quality_score", 0), reverse=True)[:150]:
        if item["relative_path"] not in selected_paths:
            selected.append(item)
            selected_paths.add(item["relative_path"])

    return selected


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("catalog", type=Path)
    parser.add_argument("library_root", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--mode", choices=["manifest", "hardlink", "copy"], default="manifest")
    args = parser.parse_args()
    root = args.library_root.resolve()
    selected = select(load_jsonl(args.catalog))
    args.output.mkdir(parents=True, exist_ok=True)
    manifest = []
    for record in selected:
        src = Path(record["absolute_path"])
        dest = args.output / record["category"] / src.name
        manifest.append({
            "source": str(src),
            "selected_path": str(dest),
            "mode": args.mode,
            "quality_score": record.get("quality_score"),
            "has_animation_xml": record.get("has_animation_xml"),
            "slide_count": record.get("slide_count"),
        })
        if args.mode == "hardlink":
            dest.parent.mkdir(parents=True, exist_ok=True)
            if not dest.exists():
                try:
                    dest.hardlink_to(src)
                except OSError:
                    shutil.copy2(src, dest)
        elif args.mode == "copy":
            dest.parent.mkdir(parents=True, exist_ok=True)
            if not dest.exists():
                shutil.copy2(src, dest)
    (args.output / "selection-manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"selected": len(selected), "mode": args.mode, "categories": dict(Counter(item["category"] for item in selected))}, ensure_ascii=False))

if __name__ == "__main__":
    main()
