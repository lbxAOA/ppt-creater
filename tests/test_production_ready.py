import json
from pathlib import Path


def test_production_ready_library_has_verified_filled_decks_and_semantic_slots():
    root = Path(__file__).parents[1]
    data = json.loads((root / "catalog" / "production-ready-template-library.json").read_text(encoding="utf-8-sig"))
    assert data["overall_status"] == "production_ready"
    for target, family in data["targets"].items():
        assert Path(family["master_pptx"]).exists(), target
        assert Path(family["validation_pptx"]).exists(), target
        assert family["playback_status"] == "verified"
        assert family["fill_status"] == "verified"
        assert family["semantic_slots"]
        assert all(
            page["page_type"] and (page["slots"] or page.get("content_mode") == "visual_only")
            for page in family["pages"]
        )
