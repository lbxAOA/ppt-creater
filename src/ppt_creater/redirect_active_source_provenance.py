#!/usr/bin/env python
"""Remove stale C:/PPT模板 provenance from active project registries."""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(r"C:/ppt-creater")
CURATED = ROOT / "curated" / "template-families"

TARGETS = {
    "ai-industry": CURATED / "02-005" / "004_文艺唯美模板.pptx",
    "investor-pitch": CURATED / "02-005" / "004_文艺唯美模板.pptx",
    "academic-clean": CURATED / "02-005" / "004_文艺唯美模板.pptx",
    "data-boardroom": CURATED / "04-008" / "001_汇报总结模板.pptx",
    "brand-company-profile": CURATED / "02-001" / "015_扁平化模板.pptx",
    "marketing-event": CURATED / "02-004" / "052_中国风模板.pptx",
    "culture-tourism": CURATED / "02-003" / "030_中国风模板.pptx",
    "wedding-bridal": CURATED / "03-004" / "004_婚礼相册模板.pptx",
    "career-portfolio": CURATED / "01-006" / "039_模板组.pptx",
}


def update(path: Path) -> int:
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    changed = 0
    for target, source in TARGETS.items():
        item = data.get("targets", {}).get(target)
        if item is not None and source.exists():
            item["source_pptx"] = str(source)
            changed += 1
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return changed


def main() -> None:
    result = {
        str(path): update(path)
        for path in [
            ROOT / "catalog" / "production-template-library.json",
            ROOT / "catalog" / "complete-production-template-library.json",
        ]
    }
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
