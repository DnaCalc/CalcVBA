"""Compares the VM3 and JIT sections with the VBA section of each case in a
run-inproc-probe.ps1 results.json (CalcVBA loaded inside Excel), and lists the differing lines.

python -I compare_inproc.py <results.json>
"""

import json
import sys


def sections(text):
    out, cur = {}, None
    for line in text.replace("\r\n", "\n").split("\n"):
        if line.startswith("== "):
            cur = line[3:].strip()
            out[cur] = []
        elif cur and line:
            out[cur].append(line)
    return out


def main():
    with open(sys.argv[1], encoding="utf-8-sig") as f:
        results = json.load(f)
    results = results if isinstance(results, list) else [results]
    total = differ = 0
    for r in results:
        if r.get("compile_status") != "ok":
            print(f"{r['id']}: VBE compile error (expected for refusal probes)")
            continue
        s = sections(r.get("run_value") or "")
        vba = s.get("VBA", [])
        for engine in ("VM3", "JIT"):
            got = s.get(engine, [])
            if got and got[0].startswith("run failed"):
                print(f"{r['id']} {engine}: {got[0][:160]}")
                differ += len(vba)
                continue
            for i, want in enumerate(vba):
                total += 1
                have = got[i] if i < len(got) else "<missing>"
                if have != want:
                    differ += 1
                    print(f"{r['id']} {engine}: VBA {want!r} CalcVBA {have!r}")
    print(f"{differ} differing of {total} compared lines")


if __name__ == "__main__":
    main()
