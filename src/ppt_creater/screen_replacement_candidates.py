#!/usr/bin/env python
"""Screen every indexed template for potential production-route replacement.

The score is a structural triage signal only. Promotion remains subject to visual
review, safe slot mapping, native fill validation, and PowerPoint playback.
"""
from __future__ import annotations

import json
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(r"C:/ppt-creater")
CATALOG = ROOT / "catalog" / "all-templates-enriched.jsonl"
FAMILIES = ROOT / "catalog" / "families.json"
ROUTING = ROOT / "catalog" / "generation-template-library.json"
OUT = ROOT / "catalog" / "replacement-screening.json"

TARGETS = (
    "ai-industry",
    "investor-pitch",
    "academic-clean",
    "data-boardroom",
    "brand-company-profile",
    "marketing-event",
    "culture-tourism",
    "generic-corporate",
    "wedding-bridal",
    "career-portfolio",
)


def read_jsonl(path: Path) -> list[dict]:
    return [json.loads(line) for line in path.read_text(encoding="utf-8-sig").splitlines() if line.strip()]


def target_matches(record: dict, target: str) -> bool:
    path = record.get("relative_path", "").lower()
    text = (record.get("text_excerpt") or "").lower()
    blob = f"{path}\n{text[:3000]}"
    rules = {
        "ai-industry": ("商务", "商业", "汇报", "答辩", "动态", "图表", "工作"),
        "investor-pitch": ("商业", "商务", "项目策划", "汇报", "工作"),
        "academic-clean": ("答辩", "教育", "培训", "学术", "论文", "汇报"),
        "data-boardroom": ("图表", "汇报", "商务", "工作", "商业"),
        "brand-company-profile": ("商务", "工作", "动态", "扁平", "企业"),
        "marketing-event": ("主题活动", "节日", "中国风", "宣传", "文艺"),
        "culture-tourism": ("主题活动", "节日", "中国风", "宣传", "文艺"),
        "generic-corporate": ("商务", "工作", "汇报", "商业", "动态"),
        "wedding-bridal": ("婚礼", "相册", "婚庆"),
        "career-portfolio": ("求职", "竞聘", "作品集", "个人"),
    }
    return any(token in blob for token in rules[target])


def triage_score(record: dict, target: str) -> float:
    patterns = record.get("slide_patterns") or {}
    score = float(record.get("quality_score") or 0)
    score += min(record.get("slide_count") or 0, 40) * 0.08
    score += min((record.get("chart_count") or 0) * 2 + (record.get("table_count") or 0) * 2, 12)
    score += 2 if record.get("has_animation_xml") else 0
    score += 1 if record.get("has_transition_xml") else 0
    if target in {"ai-industry", "investor-pitch", "data-boardroom", "generic-corporate"}:
        score += min(patterns.get("chart", 0) * 1.5 + patterns.get("table", 0) * 1.5, 8)
    if target in {"marketing-event", "culture-tourism", "wedding-bridal", "career-portfolio"}:
        score += min(patterns.get("minimal", 0) * 0.6 + patterns.get("content", 0) * 0.2, 5)
    return round(score, 2)


def main() -> None:
    records = [r for r in read_jsonl(CATALOG) if r.get("status") == "ok" and r.get("read_status") == "complete"]
    families = json.loads(FAMILIES.read_text(encoding="utf-8-sig"))["families"]
    reviews = json.loads(ROUTING.read_text(encoding="utf-8-sig")).get("family_reviews", {})
    family_by_source = defaultdict(list)
    for family in families:
        family_by_source[family.get("source_template")].append(family)

    assessed = []
    by_target: dict[str, list[dict]] = {target: [] for target in TARGETS}
    for record in records:
        source = record.get("absolute_path")
        linked_families = family_by_source.get(source, [])
        linked = []
        for family in linked_families:
            review = reviews.get(family["family_id"], {})
            candidate_copy = ROOT / "curated" / "template-families" / family["family_id"] / Path(source).name
            linked.append(
                {
                    "family_id": family["family_id"],
                    "candidate_copy": str(candidate_copy),
                    "candidate_copy_exists": candidate_copy.exists(),
                    "visual_status": review.get("status"),
                    "visual_score": review.get("visual_score"),
                    "review_targets": review.get("targets", []),
                    "risks": review.get("risks", []),
                }
            )
        assessed.append(
            {
                "relative_path": record.get("relative_path"),
                "absolute_path": source,
                "source_exists": Path(source).exists(),
                "sha256": record.get("sha256"),
                "slides": record.get("slide_count"),
                "charts": record.get("chart_count"),
                "tables": record.get("table_count"),
                "animation_xml": record.get("has_animation_xml"),
                "transition_xml": record.get("has_transition_xml"),
                "quality_score": record.get("quality_score"),
                "patterns": record.get("slide_patterns"),
                "linked_candidates": linked,
            }
        )
        for target in TARGETS:
            if target_matches(record, target):
                by_target[target].append(
                    {
                        "score": triage_score(record, target),
                        "relative_path": record.get("relative_path"),
                        "absolute_path": source,
                        "source_exists": Path(source).exists(),
                        "sha256": record.get("sha256"),
                        "slides": record.get("slide_count"),
                        "charts": record.get("chart_count"),
                        "tables": record.get("table_count"),
                        "animation_xml": record.get("has_animation_xml"),
                        "patterns": record.get("slide_patterns"),
                        "linked_candidates": linked,
                    }
                )
    for target in TARGETS:
        by_target[target].sort(
            key=lambda item: (
                max((candidate.get("visual_score") or 0 for candidate in item["linked_candidates"]), default=0),
                item["score"],
                item["source_exists"],
            ),
            reverse=True,
        )
        by_target[target] = by_target[target][:25]

    output = {
        "schema_version": 1,
        "screening_scope": "all indexed readable PPTX files",
        "screening_note": "Structural scores are triage only and do not equal visual approval. No production route was changed.",
        "records_read": len(records),
        "unique_hashes": len({record.get("sha256") for record in records if record.get("sha256")}),
        "by_category": dict(Counter(record.get("category") for record in records)),
        "candidate_families_with_existing_project_copies": sum(
            1 for family in families
            if (ROOT / "curated" / "template-families" / family["family_id"] / Path(family["source_template"]).name).exists()
        ),
        "ranked_by_target": by_target,
    }
    OUT.write_text(json.dumps(output, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"records_read": len(records), "output": str(OUT)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
