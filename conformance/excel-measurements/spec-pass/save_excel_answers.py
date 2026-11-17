"""Writes <Module>.excel.txt beside each probe module from results.json files of
run-inproc-probe.ps1 (the "== VBA" section of each case) or of run-excel-probe-session.ps1 with
entry <Module>.Run (VBA only), or "compile error: <dialog> @ <line>" where VBE refused it.

python -I save_excel_answers.py <results.json> [<results.json> ...]
"""

import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))


def vba_section(run_value):
    lines = run_value.replace("\r\n", "\n").split("\n")
    if not any(line.startswith("== ") for line in lines):  # a VBA-only run
        return [line for line in lines if line]
    out, inside = [], False
    for line in lines:
        if line.startswith("== "):
            inside = line == "== VBA"
            continue
        if inside and line:
            out.append(line)
    return out


def main():
    for path in sys.argv[1:]:
        with open(path, encoding="utf-8-sig") as f:
            results = json.load(f)
        if isinstance(results, dict):  # a one-case run
            results = [results]
        for r in results:
            name = r["id"]
            if name == "inproc":
                name = "SpecExpr"
            if r.get("compile_status") == "compile-error":
                d = r["compile_dialogs"][0]
                raw = d["dialog_text"]
                raw = "\n".join(raw) if isinstance(raw, list) else raw
                text = " ".join(raw.replace("Compile error:", "").split())
                body = f"compile error: {text} @ {d.get('expanded_line', '').strip()}\n"
            elif r.get("run_error") or not r.get("run_value"):
                print(f"{name}: no answer ({r.get('run_error')}), not written")
                continue
            else:
                body = "\n".join(vba_section(r["run_value"])) + "\n"
            with open(os.path.join(HERE, name + ".excel.txt"), "w", encoding="utf-8", newline="\n") as out:
                out.write(body)
            print(name, body.count("\n"), "lines")


if __name__ == "__main__":
    main()
