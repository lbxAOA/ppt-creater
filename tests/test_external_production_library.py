import json
from pathlib import Path


def test_external_network_production_families_are_named_consistently_and_strictly_validated():
    root = Path(__file__).parents[1]
    registry = json.loads(
        (root / "catalog" / "network-production-template-library.json").read_text(
            encoding="utf-8-sig"
        )
    )
    expected = {
        "external-investor-modern": "NET-01",
        "external-company-professional": "NET-02",
        "external-career-portfolio": "NET-03",
        "external-wedding-amelia": "NET-04",
        "external-investor-strategy": "NET-06",
        "external-b2b-market-analysis": "NET-07",
        "external-tech-product-workflow": "NET-08",
        "external-kpi-scorecard": "NET-09",
        "external-quarterly-business-review": "NET-10",
        "external-financial-management-consulting": "NET-12",
        "external-ai-tool-pitch": "NET-13",
    }
    assert registry["overall_status"] == "production_ready"
    assert expected == {family_id: item["source_id"] for family_id, item in registry["families"].items()}

    for family_id, source_id in expected.items():
        item = registry["families"][family_id]
        assert item["family_id"] == family_id
        assert item["source_id"] == source_id
        assert item["template_origin"] == "external"
        assert item["license_status"] == "manual_external_review_required"
        assert Path(item["master_pptx"]).exists(), family_id
        assert Path(item["validation_pptx"]).exists(), family_id
        assert item["strict_native_only"] is True, family_id
        assert item["new_shapes"] == 0, family_id
        assert item["animation_effects_before"] == item["animation_effects_after"], family_id
        assert item["powerpoint_playback"] is True, family_id
        assert item["status"] == "production_ready", family_id


def test_external_special_page_exclusions_are_preserved_after_renaming():
    root = Path(__file__).parents[1]
    registry = json.loads(
        (root / "catalog" / "network-production-template-library.json").read_text(
            encoding="utf-8-sig"
        )
    )
    kpi = registry["families"]["external-kpi-scorecard"]
    ai = registry["families"]["external-ai-tool-pitch"]
    assert kpi["retained_source_slide_indices"] == [1]
    assert kpi["excluded_source_slide_indices"] == [2, 3]
    assert ai["retained_source_slide_indices"] == [1, *range(3, 21)]
    assert ai["excluded_source_slide_indices"] == [2]
