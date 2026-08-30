import json
from pathlib import Path


def test_external_event_culture_family_preserves_its_native_overlap_exclusion():
    root = Path(__file__).parents[1]
    registry = json.loads(
        (root / "catalog" / "network-production-template-library.json").read_text(
            encoding="utf-8-sig"
        )
    )
    item = registry["families"]["external-event-culture-tourism"]
    assert item["source_id"] == "NET-05"
    assert Path(item["master_pptx"]).exists()
    assert Path(item["validation_pptx"]).exists()
    assert item["strict_native_only"] is True
    assert item["new_shapes"] == 0
    assert item["fill_count"] > 0
    assert item["animation_effects_before"] == item["animation_effects_after"] == 0
    assert item["powerpoint_playback"] is True
    assert item["license_status"] == "manual_external_review_required"
    repair = item["native_overlap_repair"]
    assert repair["source_slide_index"] == 2
    assert repair["new_shapes"] == 0
    assert repair["action"] == "exclude_noneditable_visual_overlap_slide"
    assert 2 not in item["retained_source_slide_indices"]
