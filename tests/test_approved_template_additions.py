import json
from pathlib import Path


def test_legacy_approved_additions_are_retained_as_historical_validation_evidence():
    root = Path(__file__).parents[1]
    registry = json.loads(
        (root / "catalog" / "approved-template-additions.json").read_text(
            encoding="utf-8-sig"
        )
    )

    expected = {
        "modern-report-ai-investor-academic",
        "marketing-event-chinese-style",
        "culture-tourism-chinese-style",
        "wedding-album-expanded",
        "data-visual-reference",
        "data-chart-reference",
    }
    assert expected == set(registry["families"])

    for family_id, item in registry["families"].items():
        assert Path(item["source_pptx"]).exists(), family_id
        assert Path(item["master_pptx"]).exists(), family_id
        assert Path(item["validation_pptx"]).exists(), family_id
        assert item["strict_native_only"] is True, family_id
        assert item["new_shapes"] == 0, family_id

    # This old replacement family lost three animation entries during a historic
    # validation pass; it remains review evidence, not a canonical routed family.
    modern = registry["families"]["modern-report-ai-investor-academic"]
    assert modern["animation_effects_before"] == 87
    assert modern["animation_effects_after"] == 84
    assert modern["status"] == "review_required"
