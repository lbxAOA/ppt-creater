import json
from pathlib import Path


def test_every_registered_candidate_source_is_project_contained_before_source_cleanup():
    root = Path(__file__).parents[1]
    registry = json.loads((root / "catalog" / "families.json").read_text(encoding="utf-8-sig"))
    curated_root = (root / "curated" / "template-families").resolve()

    assert len(registry["families"]) == 40
    for family in registry["families"]:
        source = Path(family["source_template"]).resolve()
        assert source.exists(), family["family_id"]
        assert source.is_relative_to(curated_root), family["family_id"]
