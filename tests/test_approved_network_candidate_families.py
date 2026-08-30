import json
from pathlib import Path


def test_external_network_families_are_promoted_as_separate_strict_native_families():
    root = Path(__file__).parents[1]
    registry = json.loads(
        (root / "catalog" / "network-production-template-library.json").read_text(
            encoding="utf-8-sig"
        )
    )
    required = {
        "external-investor-modern": "NET-01",
        "external-company-professional": "NET-02",
        "external-career-portfolio": "NET-03",
        "external-wedding-amelia": "NET-04",
    }
    assert set(required) <= set(registry["families"])
    for family_id, source_id in required.items():
        item = registry["families"][family_id]
        assert item["source_id"] == source_id
        assert item["family_id"] == family_id
        assert Path(item["master_pptx"]).exists()
        assert Path(item["validation_pptx"]).exists()
        assert item["strict_native_only"] is True
        assert item["new_shapes"] == 0
        assert item["fill_count"] > 0
        assert item["animation_effects_before"] == item["animation_effects_after"] == 0
        assert item["powerpoint_playback"] is True
        assert item["license_status"] == "manual_external_review_required"
