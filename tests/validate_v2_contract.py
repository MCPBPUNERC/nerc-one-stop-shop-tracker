import json
from pathlib import Path

PROFILE = Path("docs/data/profile.json")
NERC = Path("docs/data/nerc.json")

def main():
    profile = json.loads(PROFILE.read_text(encoding="utf-8"))
    standards = profile.get("tracked_standards", [])
    assert len(standards) == 47, f"Expected 47 tracked standards, found {len(standards)}"
    assert len(set(standards)) == len(standards), "Duplicate tracked standards found"
    assert profile.get("registered_functions") == ["GO","GOP","TO","DP"]
    if NERC.exists():
        data = json.loads(NERC.read_text(encoding="utf-8"))
        assert data.get("status") in {"ok","error"}
        assert "summary" in data
    print("Version 2 data contract: OK")

if __name__ == "__main__":
    main()
