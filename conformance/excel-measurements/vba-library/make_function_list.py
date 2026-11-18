"""Builds functions.tsv, the complete list of VBA library procedures and properties, from
vbe7-library.tsv (VBE7.DLL's type library, dump-vbe7-typelib.ps1) and vbal-names.txt (MS-VBAL
section 6 headings): one row per name with its VBE7 module, $ form, signature, hidden flag, and
whether MS-VBAL lists it.

python -I make_function_list.py
"""

import csv
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
MODULES = {
    "Conversion", "DateTime", "FileSystem", "Financial", "Information", "Interaction", "Math",
    "Strings", "_HiddenModule", "VBEGlobal",
}
CLASSES = {"_Collection": "Collection", "_ErrObject": "ErrObject", "IRegExp": "RegExp",
           "IMatch": "Match", "IMatchCollection": "MatchCollection", "ISubMatches": "SubMatches"}
IUNKNOWN = {"QueryInterface", "AddRef", "Release", "GetTypeInfoCount", "GetTypeInfo",
            "GetIDsOfNames", "Invoke"}


def main():
    with open(os.path.join(HERE, "vbe7-library.tsv"), encoding="utf-8") as f:
        rows = list(csv.DictReader(f, delimiter="\t"))
    spec = {l.strip().rstrip("$").lower() for l in open(os.path.join(HERE, "vbal-names.txt"), encoding="utf-8") if l.strip()}
    spec_raw = {l.strip().lower() for l in open(os.path.join(HERE, "vbal-names.txt"), encoding="utf-8") if l.strip()}
    out = {}
    for r in rows:
        ti, member = r["typeinfo"], r["member"]
        if not member or member in IUNKNOWN:
            continue
        if ti in MODULES:
            owner = ti
        elif ti in CLASSES:
            owner = CLASSES[ti]
        else:
            continue
        flags = int(r["member_flags"] or "0", 16)
        hidden = bool(flags & 0x41)  # FUNCFLAG_FRESTRICTED | FUNCFLAG_FHIDDEN
        m = re.match(r"_B_(str|var)_(.+)", member)
        name, dollar = (m.group(2), m.group(1) == "str") if m else (member, False)
        key = (owner, name)
        e = out.setdefault(key, {"module": owner, "name": name, "dollar": False, "kinds": set(),
                                 "returns": "", "params": "", "hidden": hidden})
        e["dollar"] |= dollar
        e["kinds"].add(r["member_kind"])
        e["hidden"] &= hidden
        if not dollar and r["member_kind"] != "PROPERTYPUT":
            e["returns"], e["params"] = r["returns"], r["params"]
        elif not e["params"]:
            e["returns"], e["params"] = r["returns"], r["params"]
    with open(os.path.join(HERE, "functions.tsv"), "w", encoding="utf-8", newline="\n") as f:
        f.write("module\tname\tdollar_form\tkind\thidden\tin_ms_vbal\treturns\tparams\n")
        for (owner, name), e in sorted(out.items(), key=lambda kv: (kv[0][0].lower(), kv[0][1].lower())):
            in_spec = name.lower() in spec or name.lower() in spec_raw
            kind = "property" if any(k.startswith("PROPERTY") for k in e["kinds"]) else "function"
            f.write(f"{owner}\t{name}\t{'yes' if e['dollar'] else ''}\t{kind}\t{'yes' if e['hidden'] else ''}"
                    f"\t{'yes' if in_spec else 'NO'}\t{e['returns']}\t{e['params']}\n")
    vbe7 = {name.lower() for (_, name) in out}
    only_spec = sorted(n for n in spec if n not in vbe7)
    print(len(out), "VBE7 procedures/properties;", sum(1 for (k, e) in out.items() if not (k[1].lower() in spec)), "not in MS-VBAL")
    print("in MS-VBAL but not VBE7:", only_spec)


if __name__ == "__main__":
    main()
