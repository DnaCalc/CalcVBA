"""Excel side of the VBA library probes.

  python -I run_probes.py spec <run-dir>      writes <run-dir>/ProbeAll.bas and spec.json: one
                                              workbook with ProbeFmt, every probes/P_*.bas and a
                                              runner whose RunAll returns each module's Run()
                                              under a "== <Module>" header
  python -I run_probes.py save <results.json> splits Excel's answer into probes/<Module>.excel.txt

Run the spec with scripts/run-excel-probe-session.ps1 under on-private-desktop.ps1.
"""

import glob
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PROBES = os.path.join(HERE, "probes")


def modules():
    return sorted(os.path.splitext(os.path.basename(p))[0] for p in glob.glob(os.path.join(PROBES, "P_*.bas")))


def spec(run_dir, group=0, only=""):
    """One case (workbook) per group of `group` modules (0: all in one); `only` limits the run
    to a comma-separated list of modules."""
    os.makedirs(run_dir, exist_ok=True)
    names = modules()
    if only:
        wanted = set(only.split(","))
        names = [n for n in names if n in wanted]
    group = int(group) or len(names)
    cases = []
    for g in range(0, len(names), group):
        part = names[g:g + group]
        cases.append(case(run_dir, f"g{g // group:02d}", part))
    with open(os.path.join(run_dir, "spec.json"), "w", encoding="utf-8") as f:
        json.dump({"cases": cases}, f, indent=1)
    print(len(names), "probe modules in", len(cases), "workbooks")


def case(run_dir, cid, names):
    # A direct call per module, so an error escaping a module's Run is trapped here (an error
    # inside Application.Run would surface as a dialog instead).
    lines = ['Attribute VB_Name = "ProbeAll"', "Option Explicit", ""]
    for n in names:
        lines += [f"Private Function R_{n}() As String",
                  "    On Error GoTo fail",
                  f"    R_{n} = {n}.Run()",
                  "    Exit Function",
                  "fail:",
                  f"    R_{n} = \"RUNNER-ERR\" & Err.Number & \" \" & Err.Description & vbLf",
                  "End Function", ""]
    lines += ["Public Function RunAll() As String", "    Dim s As String"]
    for n in names:
        lines.append(f'    s = s & "== {n}" & vbLf & R_{n}()')
    lines += ["    RunAll = s", "End Function", ""]
    runner = os.path.join(run_dir, f"ProbeAll-{cid}.bas")
    with open(runner, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines))
    mods = [{"name": "ProbeFmt", "kind": 1, "path": os.path.join(PROBES, "ProbeFmt.bas")}]
    mods += [{"name": n, "kind": 1, "path": os.path.join(PROBES, n + ".bas")} for n in names]
    mods.append({"name": "ProbeAll", "kind": 1, "path": runner})
    return {"id": cid, "modules": mods, "entry": "ProbeAll.RunAll", "members": names}


def save(results_path):
    with open(results_path, encoding="utf-8-sig") as f:
        results = json.load(f)
    results = results if isinstance(results, list) else [results]
    total = 0
    for r in results:
        if r.get("compile_status") != "ok" or not r.get("run_value"):
            print(r.get("id"), "no answer:", r.get("compile_status"), r.get("run_error"))
            for d in r.get("compile_dialogs") or []:
                text = d.get("dialog_text")
                print("  ", " ".join(text) if isinstance(text, list) else text, "@", d.get("expanded_line"))
            continue
        current, out = None, {}
        for line in r["run_value"].replace("\r\n", "\n").split("\n"):
            if line.startswith("== "):
                current = line[3:].strip()
                out[current] = []
            elif current and line:
                out[current].append(line)
        for name, lines in out.items():
            with open(os.path.join(PROBES, name + ".excel.txt"), "w", encoding="utf-8", newline="\n") as f:
                f.write("\n".join(lines) + "\n")
        total += len(out)
    print(total, "answers written")


if __name__ == "__main__":
    {"spec": spec, "save": save}[sys.argv[1]](*sys.argv[2:])
