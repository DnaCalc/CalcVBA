"""Writes one bead description per library function with differences (from differences.tsv),
leaving out lines a cross-cutting bead covers (named arguments, RegExp, missing names, Err as a
value, locale text). Output: <out-dir>/<Module>.md and <out-dir>/index.tsv (module, priority,
title).

python -I make_bead_texts.py <out-dir>
"""

import csv
import os
import sys
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
ORDER = ["missing", "error", "type", "value", "lastbit"]
SYSTEMIC = ("(named argument)", "New VBA.RegExp", "PredeclaredObject(Err)", "Variable not defined",
            "Sub or Function not defined", "(member)")


def main():
    out_dir = sys.argv[1]
    os.makedirs(out_dir, exist_ok=True)
    rows = list(csv.DictReader(open(os.path.join(HERE, "differences.tsv"), encoding="utf-8"), delimiter="\t"))
    by = defaultdict(dict)
    for r in rows:
        if r["locale"] or any(s in r["calcvba"] for s in SYSTEMIC):
            continue
        by[r["module"]].setdefault(r["excel"], {})[r["executor"]] = (r["class"], r["calcvba"])
    index = []
    for module, lines in sorted(by.items()):
        name = module[2:]
        worst = min(ORDER.index(c) for d in lines.values() for c, _ in d.values())
        priority = 3 if ORDER[worst] == "lastbit" else 2
        counts = defaultdict(int)
        for d in lines.values():
            counts[min((c for c, _ in d.values()), key=ORDER.index)] += 1
        summary = ", ".join(f"{n} {c}" for c, n in sorted(counts.items(), key=lambda kv: ORDER.index(kv[0])))
        title = f"CalcVBA library: {name} differs from Excel VBA ({summary})"
        body = [f"Probe `conformance/excel-measurements/vba-library/probes/{module}.bas`, Excel's answers in "
                f"`{module}.excel.txt`; checked by `crates/calcvba-differential/tests/vba_library_excel.rs` "
                f"(listed in known-differences.txt until fixed). Bit-exact is the only passing state; the class "
                f"says how far off (missing > error > type > value > lastbit = within 4 ulps). Read the probe "
                f"for each case's exact expression.", "",
                "| Case | Excel | CalcVBA VM3 | CalcVBA JIT | Class |", "|---|---|---|---|---|"]
        for excel, d in lines.items():
            key, want = excel.split("=", 1)
            vm3 = d.get("vm3", (None, "= (matches)"))[1]
            jit = d.get("jit", (None, "= (matches)"))[1]
            vm3 = vm3.split("=", 1)[1] if "=" in vm3 else vm3
            jit = jit.split("=", 1)[1] if "=" in jit else jit
            cls = min((c for c, _ in d.values()), key=ORDER.index)
            cell = lambda s: s.replace("|", "\\|")
            body.append(f"| {key} | `{cell(want)}` | `{cell(vm3)}` | `{cell(jit)}` | {cls} |")
        body += ["", "Done: every line matches on VM3 and the JIT and leaves known-differences.txt."]
        with open(os.path.join(out_dir, module + ".md"), "w", encoding="utf-8", newline="\n") as f:
            f.write("\n".join(body) + "\n")
        index.append((module, priority, title))
    with open(os.path.join(out_dir, "index.tsv"), "w", encoding="utf-8", newline="\n") as f:
        for m, p, t in index:
            f.write(f"{m}\t{p}\t{t}\n")
    print(len(index), "function beads")


if __name__ == "__main__":
    main()
