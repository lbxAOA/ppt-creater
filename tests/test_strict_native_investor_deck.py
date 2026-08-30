from pathlib import Path
import json


def test_strict_native_investor_deck_contract():
    root = Path(__file__).parents[1]
    qa = json.loads((root / "catalog" / "embodied-world-model-strict-native-qa.json").read_text(encoding="utf-8-sig"))
    assert qa["slide_count"] == 12
    assert qa["new_shapes"] == 0
    assert qa["animation_effects_before"] == qa["animation_effects_after"] == 55
    assert qa["powerpoint_playback"] is True
    assert Path(qa["output_pptx"]).exists()
