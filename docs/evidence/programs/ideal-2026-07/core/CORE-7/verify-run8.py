"""Verify retained development receipts and their exact source association."""
from pathlib import Path
import hashlib
import json
import zipfile

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[5]


def sha(data):
    return hashlib.sha256(data).hexdigest()


manifest = json.loads((HERE / "run8-manifest.json").read_text(encoding="utf-8"))
assert manifest["schema"] == "oxvba.run8.development-evidence.v1"
archive = HERE / manifest["archive"]
assert sha(archive.read_bytes()) == manifest["archive_sha256"]
with zipfile.ZipFile(archive) as bundle:
    names = bundle.namelist()
    assert len(names) == len(set(names))
    assert set(names) == set(manifest["archive_members_sha256"])
    for name, expected in manifest["archive_members_sha256"].items():
        assert sha(bundle.read(name)) == expected, name
    for name, expected in manifest["source_sha256"].items():
        assert sha((ROOT / name).read_bytes()) == expected, name
        assert sha(bundle.read("source/" + name)) == expected, name
    check = manifest["repository_check"]
    assert check["result"] == "passed" and check["counts"]["failed"] == 0
    assert b"[oxvba] meta check complete" in bundle.read(check["log"])
    live = manifest["live_excel"]
    assert live["exit_code"] == 0 and live["excel_pid"] != 3144
    result = json.loads(bundle.read(live["results"]))
    assert result["passed"] and all(c["passed"] for c in result["cases"])
    assert all(c["cleanup_status"] == "owned-process-zero" for c in result["cases"])
    assert manifest["formal"]["strict_globals"]["result"] == "unsupported-operation"
    assert manifest["formal"]["strict_globals"]["passing_credit"] is False
    assert manifest["formal"]["jit_machine_code_proved"] is False
print(f"run8 evidence verified: {len(names)} archive members, "
      f"{len(manifest['source_sha256'])} source files")
