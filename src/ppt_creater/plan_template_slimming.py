import json
from pathlib import Path

root = Path(r"C:/PPT模板")
registry = json.loads(Path(r"C:/ppt-creater/catalog/complete-production-template-library.json").read_text(encoding="utf-8-sig"))
retained = sorted({str(Path(x["source_pptx"]).resolve()) for x in registry["targets"].values()})
presentations = sorted(
    str(p.resolve()) for p in root.rglob("*")
    if p.is_file() and p.suffix.lower() in {".ppt", ".pptx"}
)
retain_set = set(retained)
delete_candidates = [x for x in presentations if x not in retain_set]
audit = {
    "schema_version": 1,
    "scope": "original library only; non-presentation assets are excluded",
    "delete_executed": False,
    "original_presentation_count": len(presentations),
    "retained_template_count": len(retained),
    "delete_candidate_count": len(delete_candidates),
    "retain_originals": retained,
    "delete_candidates": delete_candidates,
    "reason": "Keep only source templates used by strict production-ready target routes.",
}
out = Path(r"C:/ppt-creater/catalog/template-slimming-audit.json")
out.write_text(json.dumps(audit, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
print(json.dumps({k: audit[k] for k in ["original_presentation_count", "retained_template_count", "delete_candidate_count", "delete_executed"]}))
