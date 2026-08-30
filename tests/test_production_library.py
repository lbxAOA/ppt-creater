import json
from pathlib import Path


def test_production_library_covers_all_high_value_targets():
    root = Path(__file__).parents[1]
    data = json.loads((root / "catalog" / "production-template-library.json").read_text(encoding="utf-8-sig"))
    expected = {
        "ai-industry", "investor-pitch", "academic-clean", "data-boardroom",
        "brand-company-profile", "marketing-event", "wedding-bridal", "career-portfolio",
    }
    assert expected.issubset(data["targets"])
    for target in expected:
        family = data["targets"][target]
        assert Path(family["master_pptx"]).exists(), target
        assert family["animation_status"] == "verified_preserved"
        assert family["slot_count"] > 0
