"""Verify a run-9 development archive; optionally check its source tree/ref."""
from pathlib import Path
import argparse
import hashlib
import json
import subprocess
import zipfile

here = Path(__file__).resolve().parent
root = here.parents[5]
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("--manifest", default="run9-continuations.json")
source = parser.add_mutually_exclusive_group()
source.add_argument("--current", action="store_true")
source.add_argument("--source-ref")
args = parser.parse_args()
manifest = json.loads((here / args.manifest).read_text(encoding="utf-8"))
assert manifest["schema"].startswith("oxvba.run9.")
sha = lambda data: hashlib.sha256(data).hexdigest()
archive = here / manifest["archive"]
assert sha(archive.read_bytes()) == manifest["archive_sha256"]
with zipfile.ZipFile(archive) as bundle:
    names = bundle.namelist()
    assert len(names) == len(set(names))
    assert set(names) == set(manifest["archive_members_sha256"])
    for name, expected in manifest["archive_members_sha256"].items():
        assert sha(bundle.read(name)) == expected, name
    for name, expected in manifest["source_sha256"].items():
        assert sha(bundle.read("source/" + name)) == expected, name
        if args.current:
            assert sha((root / name).read_bytes()) == expected, name
        elif args.source_ref:
            data = subprocess.check_output(["git", "show", f"{args.source_ref}:{name}"], cwd=root)
            assert sha(data) == expected, name
    check = manifest["repository_check"]
    assert check["result"] == "passed" and check["counts"]["failed"] == 0
    assert b"[oxvba] meta check complete" in bundle.read(check["log"])
    live = manifest.get("live_excel")
    if live:
        result = json.loads(bundle.read(live["results"]))
        assert result["passed"] and all(case["passed"] for case in result["cases"])
        assert all(case["owned_excel_pid"] != 3144 and case["cleanup_status"] == "owned-process-zero" for case in result["cases"])
    assert manifest["formal"]["jit_machine_code_proved"] is False
print(f"run9 evidence verified: {len(names)} members, {len(manifest['source_sha256'])} source files")
