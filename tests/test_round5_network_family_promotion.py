import json
from pathlib import Path


def test_round5_external_families_are_strict_native_and_review_gated():
    root = Path(__file__).parents[1]
    registry = json.loads(
        (root / "catalog" / "network-production-template-library.json").read_text(
            encoding="utf-8-sig"
        )
    )
    expected = {
        "external-kpi-scorecard": ("NET-09", [1], [2, 3]),
        "external-ai-tool-pitch": ("NET-13", [1, *range(3, 21)], [2]),
    }
    for family_id, (source_id, retained, excluded) in expected.items():
        item = registry["families"][family_id]
        assert item["source_id"] == source_id
        assert Path(item["master_pptx"]).exists()
        assert Path(item["validation_pptx"]).exists()
        assert item["retained_source_slide_indices"] == retained
        assert item["excluded_source_slide_indices"] == excluded
        assert item["strict_native_only"] is True
        assert item["new_shapes"] == 0
        assert item["fill_count"] > 0
        assert item["animation_effects_before"] == item["animation_effects_after"]
        assert item["powerpoint_playback"] is True
        assert item["license_status"] == "manual_external_review_required"
        assert item["status"] == "production_ready"
