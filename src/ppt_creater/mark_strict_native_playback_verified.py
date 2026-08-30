import json
from pathlib import Path

path = Path(r"C:/ppt-creater/catalog/embodied-world-model-strict-native-qa.json")
data = json.loads(path.read_text(encoding="utf-8-sig"))
data["powerpoint_playback"] = True
data["status"] = "verified"
path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
print(data["status"])
