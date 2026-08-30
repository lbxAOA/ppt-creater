import json
from pathlib import Path


def test_template_slimming_execution_records_the_completed_external_archive_removal():
    root = Path(__file__).parents[1]
    audit = json.loads((root / "catalog" / "template-slimming-audit.json").read_text(encoding="utf-8-sig"))
    result = json.loads((root / "catalog" / "template-slimming-result.json").read_text(encoding="utf-8-sig"))

    assert result["delete_executed"] is True
    assert result["deleted_count"] == len(audit["delete_candidates"])
    assert not result["failed_deletions"]
    assert all(not Path(path).exists() for path in audit["delete_candidates"])
    # The original external C:/PPT模板 source archive was intentionally removed
    # after immutable in-project source copies were verified.  Its historical
    # retain list is provenance, not a current filesystem contract.
    assert all((root / item["project_path"]).exists() for item in json.loads(
        (root / "catalog" / "project-template-sources.json").read_text(encoding="utf-8-sig")
    )["sources"])
