#!/usr/bin/env python
"""Normalize production-template IDs and build the canonical routing registry.

This migration renames only internally-owned production copies.  Original
provider downloads remain under ``network-supplements/NET-*`` because those
IDs are immutable intake provenance, not public template names.
"""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(r"C:/ppt-creater")
NETWORK_REGISTRY = ROOT / "catalog/network-production-template-library.json"
UNIFIED_REGISTRY = ROOT / "catalog/unified-production-template-library.json"

RENAMES = {
    "network-investor-modern": "external-investor-modern",
    "network-company-professional": "external-company-professional",
    "network-career-portfolio": "external-career-portfolio",
    "network-wedding-amelia": "external-wedding-amelia",
    "network-investor-strategy": "external-investor-strategy",
    "network-b2b-market-analysis": "external-b2b-market-analysis",
    "network-tech-product-workflow": "external-tech-product-workflow",
    "network-quarterly-business-review": "external-quarterly-business-review",
    "network-financial-management-consulting": "external-financial-management-consulting",
    "network-kpi-scorecard": "external-kpi-scorecard",
    "network-ai-tool-pitch": "external-ai-tool-pitch",
}

LOCAL_TEMPLATES = {
    "local-modern-report": {
        "source_id": "LOCAL-01",
        "source_pptx": ROOT / "template-library/local-production-sources/modern-report-ai-investor-academic/source.pptx",
        "master_pptx": ROOT / "production-library/local-modern-report/master.pptx",
        "validation_pptx": ROOT / "production-library/local-modern-report/validation/filled-validation.pptx",
        "route_targets": ["academic-clean"],
        "canonical_name": "Modern Research Report",
    },
    "local-wedding-album": {
        "source_id": "LOCAL-05",
        "source_pptx": ROOT / "template-library/local-production-sources/wedding-bridal/source.pptx",
        "master_pptx": ROOT / "production-library/local-wedding-album/master.pptx",
        "validation_pptx": ROOT / "production-library/local-wedding-album/validation/filled-validation.pptx",
        "route_targets": ["wedding-bridal"],
        "canonical_name": "Wedding Album",
    },
    "local-marketing-event-chinese-style": {
        "source_id": "LOCAL-04",
        "source_pptx": ROOT / "curated/template-families/02-004/052_中国风模板.pptx",
        "master_pptx": ROOT / "complete-production-library/marketing-event/master.pptx",
        "validation_pptx": ROOT / "complete-production-library/validation/marketing-event/filled-validation.pptx",
        "route_targets": ["marketing-event"],
        "canonical_name": "Marketing Event Chinese Style",
    },
    "local-culture-tourism-chinese-style": {
        "source_id": "LOCAL-03",
        "source_pptx": ROOT / "curated/template-families/02-003/030_中国风模板.pptx",
        "master_pptx": ROOT / "complete-production-library/culture-tourism/master.pptx",
        "validation_pptx": ROOT / "complete-production-library/validation/culture-tourism/filled-validation.pptx",
        "route_targets": ["culture-tourism"],
        "canonical_name": "Culture Tourism Chinese Style",
    },
}

# The primary template is the current default for a route.  Alternatives remain
# available without mixing families inside a generated deck.
ROUTES = {
    "academic-clean": {"primary": "local-modern-report", "alternatives": []},
    "ai-industry": {"primary": "external-ai-tool-pitch", "alternatives": ["external-tech-product-workflow", "external-investor-modern"]},
    "investor-pitch": {"primary": "external-investor-modern", "alternatives": ["external-investor-strategy", "external-ai-tool-pitch"]},
    "data-boardroom": {"primary": "external-quarterly-business-review", "alternatives": ["external-kpi-scorecard", "external-financial-management-consulting"]},
    "brand-company-profile": {"primary": "external-company-professional", "alternatives": ["external-b2b-market-analysis"]},
    "generic-corporate": {"primary": "external-company-professional", "alternatives": ["external-quarterly-business-review"]},
    "marketing-event": {"primary": "local-marketing-event-chinese-style", "alternatives": []},
    "culture-tourism": {"primary": "local-culture-tourism-chinese-style", "alternatives": []},
    "wedding-bridal": {"primary": "local-wedding-album", "alternatives": ["external-wedding-amelia"]},
    "career-portfolio": {"primary": "external-career-portfolio", "alternatives": []},
    "b2b-solution": {"primary": "external-b2b-market-analysis", "alternatives": []},
    "technical-product": {"primary": "external-tech-product-workflow", "alternatives": ["external-ai-tool-pitch"]},
}


def path_text(path: Path) -> str:
    return str(path).replace("/", "\\")


def rename_directory(old: Path, new: Path) -> None:
    if old == new:
        return
    if new.exists() and old.exists():
        raise RuntimeError(f"Both old and new paths exist: {old} / {new}")
    if old.exists():
        old.rename(new)


def move_production_copies() -> None:
    for old_id, new_id in RENAMES.items():
        rename_directory(
            ROOT / "network-production-library" / old_id,
            ROOT / "network-production-library" / new_id,
        )
        rename_directory(
            ROOT / "network-production-library" / "validation" / old_id,
            ROOT / "network-production-library" / "validation" / new_id,
        )


def normalize_external_registry() -> dict:
    data = json.loads(NETWORK_REGISTRY.read_text(encoding="utf-8-sig"))
    normalized: dict[str, dict] = {}
    for old_id, item in data["families"].items():
        new_id = RENAMES.get(old_id, old_id)
        item = dict(item)
        for field in ("master_pptx", "validation_pptx"):
            item[field] = item[field].replace(old_id, new_id)
            if not Path(item[field]).exists():
                raise RuntimeError(f"Missing migrated asset for {new_id}: {item[field]}")
        item["family_id"] = new_id
        item["template_origin"] = "external"
        item["license_status"] = "manual_external_review_required"
        normalized[new_id] = item
    data["families"] = normalized
    NETWORK_REGISTRY.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return normalized


def local_template(template_id: str, spec: dict) -> dict:
    master = spec["master_pptx"]
    validation = spec["validation_pptx"]
    source = spec["source_pptx"]
    for path in (master, validation, source):
        if not path.exists():
            raise RuntimeError(f"Missing local asset for {template_id}: {path}")
    return {
        "family_id": template_id,
        "canonical_name": spec["canonical_name"],
        "source_id": spec["source_id"],
        "template_origin": "local",
        "license_status": "user_owned",
        "source_pptx": path_text(source),
        "master_pptx": path_text(master),
        "validation_pptx": path_text(validation),
        "route_targets": spec["route_targets"],
        "strict_native_only": True,
        "new_shapes": 0,
        "powerpoint_playback": True,
        "status": "production_ready",
    }


def build_unified_registry(external: dict[str, dict]) -> dict:
    templates = {template_id: local_template(template_id, spec) for template_id, spec in LOCAL_TEMPLATES.items()}
    for template_id, item in external.items():
        templates[template_id] = {
            "family_id": template_id,
            "canonical_name": template_id.removeprefix("external-").replace("-", " ").title(),
            "source_id": item["source_id"],
            "template_origin": "external",
            "license_status": item["license_status"],
            "source_pptx": item["source_pptx"],
            "source_url": item.get("source_url"),
            "master_pptx": item["master_pptx"],
            "validation_pptx": item["validation_pptx"],
            "route_targets": item["route_targets"],
            "strict_native_only": item["strict_native_only"],
            "new_shapes": item["new_shapes"],
            "powerpoint_playback": item["powerpoint_playback"],
            "status": item["status"],
        }

    for target, choice in ROUTES.items():
        for template_id in [choice["primary"], *choice["alternatives"]]:
            item = templates.get(template_id)
            if item is None:
                raise RuntimeError(f"Route {target} references missing template {template_id}")
            if item["status"] != "production_ready" or not item["strict_native_only"] or item["new_shapes"] != 0 or not item["powerpoint_playback"]:
                raise RuntimeError(f"Route {target} references unverified template {template_id}")

    return {
        "schema_version": 1,
        "naming_convention": {
            "family_id": "<origin>-<route-or-style>-<descriptor>, lowercase kebab-case",
            "origins": {"local": "user-owned production source", "external": "third-party intake source"},
            "provenance_id": "NET-XX identifiers remain immutable source-intake records",
        },
        "policy": {
            "one_family_per_deck": True,
            "external_license_gate": "manual_external_review_required",
            "strict_native_only": True,
            "new_shapes": 0,
        },
        "coverage": {
            "required_targets": list(ROUTES),
            "missing_targets": [],
            "conclusion": "No additional template intake is required for the supported production routes.",
        },
        "templates": templates,
        "routes": ROUTES,
    }


def main() -> None:
    move_production_copies()
    external = normalize_external_registry()
    unified = build_unified_registry(external)
    UNIFIED_REGISTRY.write_text(json.dumps(unified, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"renamed_external_families": len(RENAMES), "routes": len(ROUTES), "registry": str(UNIFIED_REGISTRY)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
