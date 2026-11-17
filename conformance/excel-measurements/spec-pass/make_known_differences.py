"""Rebuilds known-differences.txt from the output of
cargo test -p calcvba-differential --test spec_pass_excel -- --nocapture  (or a saved log):
every line VM3 or the JIT answers differently from Excel, unprefixed when both differ.

python -I make_known_differences.py <test-output.log>
"""

import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
HEADER = (
    "# Lines Excel answers that CalcVBA still answers differently, as 'Module: line'"
    " (see spec_pass_excel.rs).\n"
    "# Prefixed 'vm3 ' or 'jit ' when only that executor differs. Rebuilt by make_known_differences.py.\n"
)


def entries(log):
    found = {"Vm3": set(), "Jit": set()}
    current = None
    for line in log.splitlines():
        m = re.search(r"\b(Vm3|Jit) against Excel", line)
        if m:
            current = m.group(1)
            continue
        if current is None or line.startswith("fixed:"):
            if line.startswith("fixed:"):
                current = None
            continue
        if not re.match(r"^(Spec|Cmp)\w+: ", line):
            continue
        module, rest = line.split(": ", 1)
        if rest.startswith(("failed:", "declined:", "no result")):
            found[current].add(f"{module}: run fails")
        elif " -- " in line:
            found[current].add(line.split(" -- ", 1)[0])
        else:
            found[current].add(line)
    return found


def main():
    with open(sys.argv[1], encoding="utf-8", errors="replace") as f:
        found = entries(f.read())
    both = found["Vm3"] & found["Jit"]
    lines = sorted(both)
    lines += sorted("vm3 " + e for e in found["Vm3"] - both)
    lines += sorted("jit " + e for e in found["Jit"] - both)
    with open(os.path.join(HERE, "known-differences.txt"), "w", encoding="utf-8", newline="\n") as out:
        out.write(HEADER + "".join(l + "\n" for l in lines))
    print(len(both), "both,", len(found["Vm3"] - both), "vm3 only,", len(found["Jit"] - both), "jit only")


if __name__ == "__main__":
    main()
