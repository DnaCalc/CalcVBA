"""Check retained run-7 bytes and their association with current source files."""
from pathlib import Path
import hashlib
import json
import zipfile

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[5]


def sha(data):
    return hashlib.sha256(data).hexdigest()


manifest = json.loads((HERE / "run7-manifest.json").read_text(encoding="utf-8"))
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
assert manifest["repository_check"]["result"] == "passed"
assert manifest["repository_check"]["counts"]["failed"] == 0
assert manifest["live_excel"]["exit_code"] == 0
assert manifest["live_excel"]["excel_pid"] != 3144
print(f"run7 evidence verified: {len(names)} archive members, "
      f"{len(manifest['source_sha256'])} source files")
