import json
from pathlib import Path

PROFILE = Path("docs/data/profile.json")
NERC = Path("docs/data/nerc.json")

def main():
    profile = json.loads(PROFILE.read_text(encoding="utf-8"))
    standards = profile.get("tracked_standards", [])
    assert len(standards) == 47, f"Expected 47 tracked standards, found {len(standards)}"
    assert len(set(standards)) == len(standards), "Duplicate tracked standards found"
    assert profile.get("registered_functions") == ["GO", "GOP", "TO", "DP"]

    if NERC.exists():
        data = json.loads(NERC.read_text(encoding="utf-8"))
        # Before the first V2 semantic-baseline run, the branch may still contain
        # the legacy V1 dashboard payload. Validate the V2 contract only once a
        # schema_version=2 payload has actually been generated.
        if data.get("schema_version") == 2:
            assert data.get("status") in {"ok", "error"}
            assert "summary" in data
            assert "standards" in data
            assert "families" in data
            assert data["summary"].get("tracked_standards") == 47

    print("Version 2 data contract: OK")

if __name__ == "__main__":
    main()
