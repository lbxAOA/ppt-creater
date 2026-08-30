import json
from pathlib import Path

catalog = Path(r"C:/ppt-creater/catalog")
audit_path = catalog / "template-slimming-audit.json"
audit = json.loads(audit_path.read_text(encoding="utf-8-sig"))
failed = []
deleted = []
for raw in audit["delete_candidates"]:
    path = Path(raw)
    try:
        if path.exists():
            path.unlink()
        if path.exists():
            failed.append(raw)
        else:
            deleted.append(raw)
    except OSError as exc:
        failed.append(f"{raw}: {exc}")
result = {
    "delete_executed": True,
    "deleted_count": len(deleted),
    "failed_deletions": failed,
    "retained_template_count": len(audit["retain_originals"]),
    "scope": audit["scope"],
}
(catalog / "template-slimming-result.json").write_text(
    json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
)
print(json.dumps(result, ensure_ascii=False))
if failed:
    raise SystemExit(1)
