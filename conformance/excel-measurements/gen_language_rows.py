#!/usr/bin/env python3
"""Prepare CalcVBA's basic-language conformance fixtures for an Excel measurement.

Each fixture (conformance/tests/<name>.bas, listed with suite "basic-language" in
tests_manifest.csv) is a module whose `Sub Main` leaves its answer in module-level variables and
in Main's own locals; the engine's golden row (golden/values.csv) lists the module variables, then
the locals of Main, in declaration order. This script copies each fixture unchanged except that
`DnaRec "<name>", <name>` calls are inserted before the end of Main (before every `Exit Sub` that
stands on its own line), so Excel's VBA records the same variables after the same code ran.

  gen_language_rows.py prepare <conformance dir> <out dir> [batch size]
      writes <out>/helper.bas (language-probe.bas), <out>/modules/<name>.bas and
      <out>/manifest-<n>.json (the items of one Excel job each).
  gen_language_rows.py collect <out dir> <language-rows.excel.txt>
      turns the language.json files of the jobs (<out>/out-<n>/) into the committed answers: one
      tab-separated line per fixture, `OK` with the `name=<VarType>:<payload>` records of the
      variables (UDT locals are not recorded), `ERR` with Err.Number and Err.Description of an
      error Main raised, `COMPILE` with the text of the compile-error dialog, or `NOCAPTURE`.
"""
import csv
import json
import os
import re
import sys

PROC_START = re.compile(
    r"^\s*(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)", re.I
)
PROC_END = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.I)
TYPE_START = re.compile(r"^\s*(?:(?:Public|Private)\s+)?(Type|Enum)\s+(\w+)", re.I)
TYPE_END = re.compile(r"^\s*End\s+(Type|Enum)\b", re.I)
DECL = re.compile(
    r"^\s*(Dim|Private|Public|Global)\s+(?!Const\b|Declare\b|Sub\b|Function\b|Property\b|Type\b|Enum\b|Static\b)(.*)$",
    re.I,
)


def declarators(rest):
    """Names (and whether the type is a user-defined type name) of one Dim line's declarators."""
    rest = re.sub(r"'.*$", "", rest)
    rest = re.sub(r"^\s*WithEvents\s+", "", rest, flags=re.I)
    parts, depth, cur = [], 0, ""
    for ch in rest:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            parts.append(cur)
            cur = ""
        else:
            cur += ch
    parts.append(cur)
    out = []
    for part in parts:
        m = re.match(r"\s*(\w+)\s*(\([^)]*\))?\s*(?:As\s+(?:New\s+)?([\w.]+))?", part, re.I)
        if m:
            out.append((m.group(1), m.group(3) or "", m.group(2) is not None))
    return out


def prepare_module(text):
    lines = text.splitlines()
    udts = set()
    for line in lines:
        m = TYPE_START.match(line)
        if m and m.group(1).lower() == "type":
            udts.add(m.group(2).lower())
    names = []  # (name, kind) globals first, then Main's locals
    in_proc = None
    in_type = False
    main_range = None
    main_locals = []
    exits = []
    for i, line in enumerate(lines):
        if in_type:
            if TYPE_END.match(line):
                in_type = False
            continue
        if in_proc is None:
            if TYPE_START.match(line):
                in_type = True
                continue
            m = PROC_START.match(line)
            if m:
                in_proc = (m.group(2).lower(), i)
                continue
            m = DECL.match(line)
            if m:
                for name, typ, _ in declarators(m.group(2)):
                    names.append((name, "udt" if typ.lower() in udts else "var"))
            continue
        # inside a procedure
        if in_proc[0] == "main" and re.match(r"^\s*Dim\b", line, re.I):
            m = DECL.match(line)
            for name, typ, _ in declarators(m.group(2)):
                main_locals.append((name, "udt" if typ.lower() in udts else "var"))
        if in_proc[0] == "main" and re.match(r"^\s*Exit\s+Sub\s*$", line, re.I):
            exits.append(i)
        if PROC_END.match(line):
            if in_proc[0] == "main":
                main_range = (in_proc[1], i)
            in_proc = None
    if main_range is None:
        return None
    allnames = names + main_locals
    capture = ["DnaRec " + '"%s", %s' % (n, n) for n, kind in allnames if kind == "var"]
    capture.append("DnaCaptured = True")
    out = []
    for i, line in enumerate(lines):
        if i in exits or i == main_range[1]:
            out.extend(capture)
        out.append(line)
    return "\r\n".join(out) + "\r\n", allnames


def collect(outdir, target):
    items = {i["id"]: i for i in json.load(open(os.path.join(outdir, "all-items.json")))}
    answers = {}
    for n in range(1, 100):
        path = os.path.join(outdir, "out-%d" % n, "language.json")
        if not os.path.exists(path):
            break
        for row in json.load(open(path, encoding="utf-8-sig"))["rows"]:
            answers[row["id"]] = row
    lines = []
    for name in sorted(items):
        row = answers.get(name)
        if row is None:
            lines.append("%s\tMISSING" % name)
            continue
        dialogs = row.get("dialogs") or []
        text = " ".join(" ".join(d.get("texts") or []) for d in dialogs)
        text = text.replace("\t", " ").replace("\n", " ").strip()
        answer = row.get("answer") or ""
        if not answer:
            fallback = (row.get("run_error") or "").replace("\n", " ")
            lines.append("%s\tCOMPILE\t%s" % (name, text or fallback))
        elif answer.startswith("ERR|"):
            _, number, desc, source = answer.split("|", 3)
            lines.append("%s\tERR\t%s\t%s\t%s" % (name, number, desc, source))
        elif answer.startswith("OK\n"):
            records = [r for r in answer.split("\n")[1:] if r]
            merged = []
            for variable in items[name]["variables"]:
                if variable["kind"] == "var":
                    merged.append(records.pop(0))
                else:
                    merged.append("%s=udt:" % variable["name"])
            lines.append("%s\tOK\t%s" % (name, "\t".join(merged)))
        else:
            lines.append("%s\tNOCAPTURE\t%s" % (name, answer.split("\n")[0]))
        if dialogs and answer:
            lines[-1] += "\t#dialog: " + text
    with open(target, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines) + "\n")
    print(len(lines), "rows")


def main():
    if len(sys.argv) >= 4 and sys.argv[1] == "collect":
        collect(sys.argv[2], sys.argv[3])
        return
    if len(sys.argv) < 4 or sys.argv[1] != "prepare":
        sys.exit(__doc__)
    conf, outdir = sys.argv[2], sys.argv[3]
    batch = int(sys.argv[4]) if len(sys.argv) > 4 else 48
    os.makedirs(os.path.join(outdir, "modules"), exist_ok=True)
    here = os.path.dirname(os.path.abspath(__file__))
    with open(os.path.join(here, "language-probe.bas"), encoding="latin-1") as f:
        helper = f.read()
    with open(os.path.join(outdir, "helper.bas"), "w", encoding="latin-1", newline="") as f:
        f.write(helper)
    rows = list(csv.DictReader(open(os.path.join(conf, "tests_manifest.csv"), encoding="utf-8")))
    items = []
    for row in rows:
        if row["suite"] != "basic-language":
            continue
        name = row["file"]
        with open(os.path.join(conf, "tests", name), encoding="latin-1") as f:
            result = prepare_module(f.read())
        if result is None:
            items.append({"id": name, "error": "no Sub Main"})
            continue
        text, allnames = result
        with open(os.path.join(outdir, "modules", name), "w", encoding="latin-1", newline="") as f:
            f.write(text)
        items.append(
            {
                "id": name,
                "module": "modules/" + name,
                "variables": [{"name": n, "kind": k} for n, k in allnames],
            }
        )
    runnable = [i for i in items if "module" in i]
    for n in range(0, len(runnable), batch):
        with open(os.path.join(outdir, "manifest-%d.json" % (n // batch + 1)), "w") as f:
            json.dump({"helper": "helper.bas", "items": runnable[n : n + batch]}, f, indent=1)
    with open(os.path.join(outdir, "all-items.json"), "w") as f:
        json.dump(items, f, indent=1)
    print(len(items), "fixtures,", len(runnable), "prepared")


if __name__ == "__main__":
    main()
