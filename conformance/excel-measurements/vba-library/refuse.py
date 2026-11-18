"""Records a probe line that Excel's VBE refuses to compile and neutralises it in the probe:
the case becomes `v = "compile-refused"`, and the original line plus VBE's message is appended
to compile-refusals.tsv. CalcVBA must refuse those lines too (checked separately).

python -I refuse.py <results.json> [...]   handles every compile-error case in those results
"""

import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PROBES = os.path.join(HERE, "probes")
LOG = os.path.join(HERE, "compile-refusals.tsv")


def refuse(module, dialog, line):
    path = os.path.join(PROBES, module + ".bas")
    with open(path, encoding="utf-8", newline="") as f:
        text = f.read()
    target = line.strip()
    out, done = [], False
    for raw in text.splitlines(keepends=True):
        body = raw.rstrip("\r\n")
        if not done and body.strip() == target:
            m = re.match(r"(\s*Case\s+\d+:\s*)", body)
            if not m:
                raise SystemExit(f"{module}: refused line is not a case line: {target}")
            raw = m.group(1) + 'v = "compile-refused"' + raw[len(body):]
            done = True
        out.append(raw)
    if not done:
        raise SystemExit(f"{module}: line not found: {target}")
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write("".join(out))
    new = not os.path.exists(LOG)
    with open(LOG, "a", encoding="utf-8", newline="\n") as f:
        if new:
            f.write("module\tvbe_message\tsource\n")
        f.write(f"{module}\t{dialog}\t{target}\n")
    print(module, "|", dialog, "|", target)


def main():
    for path in sys.argv[1:]:
        with open(path, encoding="utf-8-sig") as f:
            results = json.load(f)
        results = results if isinstance(results, list) else [results]
        spec_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(path))), "spec.json")
        spec = {}
        if os.path.exists(spec_path):
            with open(spec_path, encoding="utf-8-sig") as f:
                spec = {c["id"]: c for c in json.load(f)["cases"]}
        for r in results:
            if r.get("id") in spec:
                r["modules"] = spec[r["id"]]["modules"]
            if r.get("compile_status") != "compile-error":
                continue
            d = (r.get("compile_dialogs") or [{}])[0]
            text = d.get("dialog_text") or ""
            text = " ".join(text) if isinstance(text, list) else text
            text = " ".join(text.replace("Compile error:", "").split())
            members = [m["name"] for m in r.get("modules", []) if m["name"].startswith("P_")] or r.get("members") or []
            line = d.get("expanded_line") or ""
            owner = next((m for m in members if line.strip() and line.strip() in open(
                os.path.join(PROBES, m + ".bas"), encoding="utf-8", errors="replace").read()), None)
            if owner is None:
                print(r.get("id"), "could not place:", text, "@", line)
                continue
            refuse(owner, text, line)


if __name__ == "__main__":
    main()
