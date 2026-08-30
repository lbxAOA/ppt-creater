from pathlib import Path
import json


def test_network_template_candidate_registry_is_isolated_and_review_gated():
    root = Path(__file__).parents[1]
    registry = json.loads((root / "catalog" / "network-template-supplement-candidates.json").read_text(encoding="utf-8-sig"))
    policy = registry["policy"]
    assert registry["status"] == "candidate_only"
    assert policy["original_library_write"] is False
    assert policy["production_auto_promotion"] is False
    assert policy["single_source_deck_required"] is True
    assert policy["license_review_required"] is True
    assert policy["style_review_required"] is True
    assert {"ai-industry", "career-portfolio", "wedding-bridal", "culture-tourism"} <= {
        target for item in registry["candidates"] for target in item["targets"]
    }
