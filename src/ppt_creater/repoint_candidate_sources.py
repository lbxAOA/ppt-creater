#!/usr/bin/env python
"""Repoint all candidate registry sources to their immutable project copies."""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(r"C:/ppt-creater")
REGISTRY = ROOT / "catalog" / "families.json"
CURATED = ROOT / "curated" / "template-families"


def main() -> None:
    data = json.loads(REGISTRY.read_text(encoding="utf-8-sig"))
    changed = 0
    for family in data["families"]:
        family_id = family["family_id"]
        folder = CURATED / family_id
        choices = sorted(folder.glob("*.pptx"))
        if not choices:
            raise RuntimeError(f"No project-contained copy for {family_id}: {folder}")
        # The folder normally contains one source. When historic duplicate imports
        # left two, choose the filename matching the old registered source.
        old_name = Path(family["source_template"]).name
        source = next((path for path in choices if path.name == old_name), choices[0])
        family["source_template"] = str(source)
        family["source_relative_path"] = str(source.relative_to(ROOT)).replace("\\", "/")
        family["imported_template"] = str(source.relative_to(ROOT)).replace("\\", "/")
        changed += 1
    REGISTRY.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"updated_families": changed, "registry": str(REGISTRY)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
