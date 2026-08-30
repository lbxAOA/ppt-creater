import json
from pathlib import Path


def test_template_source_directory_is_removed_after_project_migration():
    root = Path(__file__).parents[1]
    audit = json.loads((root / "catalog" / "template-slimming-audit.json").read_text(encoding="utf-8-sig"))
    project_sources = root / "curated" / "template-families"

    assert audit["original_presentation_count"] == 1073
    assert project_sources.exists()
    assert any(project_sources.rglob("*.pptx"))
    assert not Path(r"C:/PPT模板").exists()
