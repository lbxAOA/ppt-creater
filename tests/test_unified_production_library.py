import json
from pathlib import Path


def test_unified_production_library_has_consistent_ids_complete_route_coverage_and_verified_assets():
    root = Path(__file__).parents[1]
    registry = json.loads(
        (root / "catalog" / "unified-production-template-library.json").read_text(
            encoding="utf-8-sig"
        )
    )

    assert registry["naming_convention"]["family_id"].startswith("<origin>")
    assert registry["coverage"]["missing_targets"] == []
    assert set(registry["coverage"]["required_targets"]) == set(registry["routes"])

    for family_id, item in registry["templates"].items():
        assert family_id == item["family_id"]
        assert family_id.startswith(("local-", "external-"))
        assert Path(item["source_pptx"]).exists(), family_id
        assert Path(item["master_pptx"]).exists(), family_id
        assert Path(item["validation_pptx"]).exists(), family_id
        assert item["status"] == "production_ready", family_id
        assert item["strict_native_only"] is True, family_id
        assert item["new_shapes"] == 0, family_id
        assert item["powerpoint_playback"] is True, family_id
        if item["template_origin"] == "external":
            assert item["license_status"] == "manual_external_review_required", family_id
        else:
            assert item["license_status"] == "user_owned", family_id

    for target, route in registry["routes"].items():
        assert route["primary"] in registry["templates"], target
        assert target in registry["templates"][route["primary"]]["route_targets"], target
        for alternative in route["alternatives"]:
            assert alternative in registry["templates"], target
            assert target in registry["templates"][alternative]["route_targets"], target
