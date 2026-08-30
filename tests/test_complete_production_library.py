import json
from pathlib import Path


def test_complete_production_library_contract():
    root = Path(__file__).parents[1]
    registry = json.loads((root / "catalog" / "complete-production-template-library.json").read_text(encoding="utf-8-sig"))
    required = {
        "ai-industry", "investor-pitch", "academic-clean", "data-boardroom",
        "brand-company-profile", "marketing-event", "culture-tourism", "generic-corporate",
        "wedding-bridal", "career-portfolio",
    }
    assert registry["overall_status"] == "production_ready"
    assert required == set(registry["targets"])
    for target, item in registry["targets"].items():
        assert Path(item["master_pptx"]).exists(), target
        assert Path(item["validation_pptx"]).exists(), target
        assert item["strict_native_only"] is True, target
        assert item["new_shapes"] == 0, target
        assert item["animation_effects_before"] == item["animation_effects_after"], target
        assert item["powerpoint_playback"] is True, target
