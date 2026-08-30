import json
from pathlib import Path


def test_external_investor_modern_family_is_strictly_validated():
    root = Path(__file__).parents[1]
    registry = json.loads(
        (root / "catalog" / "network-production-template-library.json").read_text(
            encoding="utf-8-sig"
        )
    )
    item = registry["families"]["external-investor-modern"]
    assert registry["overall_status"] == "production_ready"
    assert Path(item["master_pptx"]).exists()
    assert Path(item["validation_pptx"]).exists()
    assert item["source_id"] == "NET-01"
    assert item["strict_native_only"] is True
    assert item["new_shapes"] == 0
    assert item["fill_count"] > 0
    assert item["animation_effects_before"] == item["animation_effects_after"] == 0
    assert item["powerpoint_playback"] is True
    assert item["license_status"] == "manual_external_review_required"
