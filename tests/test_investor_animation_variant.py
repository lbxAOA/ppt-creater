import json
from pathlib import Path


def test_investor_animation_variant_contract():
    root = Path(__file__).parents[1]
    qa = json.loads((root / "catalog" / "embodied-world-model-investor-animation-qa.json").read_text(encoding="utf-8-sig"))
    assert qa["source_pptx"].endswith("strict-native.pptx")
    assert qa["new_shapes"] == 0
    assert qa["animation_style"] == "native_object_fade_sequence"
    assert qa["powerpoint_playback"] is True
    assert qa["animation_effects_after"] > 0
    assert Path(qa["output_pptx"]).exists()
