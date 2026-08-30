import json
from pathlib import Path


def test_local_production_sources_are_project_contained_with_verified_hashes():
    root = Path(__file__).parents[1]
    manifest = json.loads(
        (root / "catalog" / "project-template-sources.json").read_text(encoding="utf-8-sig")
    )

    assert manifest["schema_version"] == 1
    assert len(manifest["sources"]) == 6
    for item in manifest["sources"]:
        copied = root / item["project_path"]
        assert copied.exists(), item["id"]
        assert copied.stat().st_size == item["bytes"], item["id"]
        assert item["verification"] == "source_hash_matches"


def test_unified_registry_routes_local_entries_to_project_contained_sources():
    root = Path(__file__).parents[1]
    registry = json.loads(
        (root / "catalog" / "unified-production-template-library.json").read_text(
            encoding="utf-8-sig"
        )
    )
    local = registry["templates"]
    assert Path(local["local-modern-report"]["source_pptx"]) == (
        root / "template-library/local-production-sources/modern-report-ai-investor-academic/source.pptx"
    )
    assert Path(local["local-wedding-album"]["source_pptx"]) == (
        root / "template-library/local-production-sources/wedding-bridal/source.pptx"
    )
