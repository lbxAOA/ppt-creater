import json
import subprocess
from pathlib import Path

root = Path(r"C:/ppt-creater")
registry_path = root / "catalog" / "complete-production-template-library.json"
registry = json.loads(registry_path.read_text(encoding="utf-8-sig"))
verifier = root / "src" / "verify-powerpoint-playback.ps1"
for name, item in registry["targets"].items():
    result = subprocess.run([
        "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(verifier),
        "-FilePath", item["validation_pptx"], "-Expected", str(item["animation_effects_before"]),
    ], capture_output=True, text=True, timeout=300)
    payload = None
    for line in result.stdout.splitlines()[::-1]:
        if line.strip().startswith("{"):
            payload = json.loads(line)
            break
    verified = bool(payload and payload.get("verified"))
    item["powerpoint_playback"] = verified
    item["status"] = "production_ready" if (
        verified and item["strict_native_only"] and item["new_shapes"] == 0
        and item["fill_count"] > 0
        and item["animation_effects_before"] == item["animation_effects_after"]
    ) else "review_required"
    print(name, item["status"])
registry["overall_status"] = "production_ready" if all(
    x["status"] == "production_ready" for x in registry["targets"].values()
) else "review_required"
registry_path.write_text(json.dumps(registry, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
print("overall", registry["overall_status"])
