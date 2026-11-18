"""Rebuilds known-differences.txt from the output of
cargo test -p calcvba-differential --test vba_library_excel   (run with an empty or header-only
known-differences.txt, so every difference is printed): each line VM3 or the JIT answers
differently from Excel (unprefixed when both differ), each module that does not run, and each
line Excel refuses to compile that CalcVBA accepts.

python -I make_known_differences.py <test-output.log>
"""

import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
HEADER = (
    "# VBA library lines CalcVBA still answers differently from Excel (see vba_library_excel.rs).\n"
    "# 'vm3 '/'jit ' prefix: only that executor differs. Rebuilt by make_known_differences.py.\n"
)


def main():
    found = {"Vm3": set(), "Jit": set()}
    refusals = set()
    section = None
    with open(sys.argv[1], encoding="utf-8", errors="replace") as f:
        for line in f:
            line = line.rstrip("\n")
            m = re.search(r"\b(Vm3|Jit) against Excel", line)
            if m:
                section = m.group(1)
                continue
            if "CalcVBA accepts" in line:
                section = "refusal"
                continue
            if line.startswith("fixed:"):
                section = None
                continue
            if section == "refusal" and line.startswith("refusal: "):
                refusals.add(line.split(" -- ", 1)[0])
            elif section in found and re.match(r"^P_\w+: ", line):
                module, rest = line.split(": ", 1)
                if rest.startswith(("failed:", "declined:", "no result")):
                    found[section].add(f"{module}: run fails")
                else:
                    found[section].add(line.split(" -- ", 1)[0])
    both = found["Vm3"] & found["Jit"]
    lines = sorted(both) + sorted("vm3 " + e for e in found["Vm3"] - both)
    lines += sorted("jit " + e for e in found["Jit"] - both) + sorted(refusals)
    with open(os.path.join(HERE, "known-differences.txt"), "w", encoding="utf-8", newline="\n") as out:
        out.write(HEADER + "".join(l + "\n" for l in lines))
    print(len(both), "both,", len(found["Vm3"] - both), "vm3 only,", len(found["Jit"] - both),
          "jit only,", len(refusals), "accepted refusals")


if __name__ == "__main__":
    main()
