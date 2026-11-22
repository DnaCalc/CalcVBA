"""Run one warmed release test process per workload and write its timing as TSV."""

import argparse
import csv
import os
from pathlib import Path
import subprocess
import time


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("executable", type=Path, help="release vba_bench test executable")
    parser.add_argument("directory", type=Path, help="new scratch directory")
    parser.add_argument("--only", action="append", help="select a workload, repeatable")
    parser.add_argument("--backend", choices=("vm3", "jit"), action="append")
    parser.add_argument("--size", type=int, help="override size for one selected workload")
    parser.add_argument("--timeout", type=int, default=90, help="seconds per case, up to 900")
    args = parser.parse_args()

    with Path(__file__).with_name("cases.tsv").open(encoding="utf-8") as source:
        cases = {row["workload"]: row for row in csv.DictReader(source, delimiter="\t")}
    if args.only and any(name not in cases or cases[name]["argument"] == "sheet" for name in args.only):
        parser.error("--only requires a CalcVBA workload from cases.tsv")
    if args.size is not None and (args.size < 1 or len(args.only or ()) != 1):
        parser.error("--size requires one workload and a positive size")
    if not 1 <= args.timeout <= 900:
        parser.error("--timeout must be between 1 and 900 seconds")

    executable = args.executable.resolve(strict=True)
    args.directory.mkdir(parents=True)
    failed = False
    deadline = time.monotonic() + 900
    with (args.directory / "calcvba.tsv").open("w", encoding="utf-8", newline="") as output:
        rows = csv.writer(output, delimiter="\t", lineterminator="\n")
        rows.writerow(("backend", "workload", "size", "checksum_bits", "median_ms", "setup_ms",
                       "run1_ms", "run2_ms", "run3_ms", "run4_ms", "run5_ms", "status"))
        for backend in args.backend or ("vm3", "jit"):
            for name in args.only or cases:
                case = cases[name]
                if case["argument"] == "sheet":
                    continue
                size = str(args.size or case["size"])
                scratch = args.directory / f"{backend}-{name}"
                scratch.mkdir()
                os.environ.update(VBA_BENCH_BACKEND=backend, VBA_BENCH_ONLY=name, VBA_BENCH_SIZE=size)
                with (scratch / "run.log").open("wb") as log:
                    remaining = deadline - time.monotonic()
                    status = "suite timeout"
                    if remaining > 0:
                        try:
                            result = subprocess.run(
                                [executable, "--ignored", "--nocapture", "--exact", "vba_bench"],
                                cwd=scratch, stdout=log, stderr=subprocess.STDOUT,
                                timeout=min(args.timeout, remaining),
                            )
                            status = "ok" if result.returncode == 0 else f"exit {result.returncode}"
                        except subprocess.TimeoutExpired:
                            status = "timeout"
                lines = (scratch / "run.log").read_text(encoding="utf-8", errors="replace").splitlines()
                marker = "BENCH\t"
                samples = [line.split(marker, 1)[1].split("\t") for line in lines if marker in line]
                if status == "ok" and (len(samples) != 1 or len(samples[0]) != 10 or samples[0][:2] != [name, size]):
                    status = "missing or invalid timing row"
                if status == "ok":
                    rows.writerow((backend, *samples[0], status))
                else:
                    rows.writerow((backend, name, size, *("",) * 8, status))
                    failed = True
                output.flush()
                print(f"{backend}\t{name}\t{status}", flush=True)
    return int(failed)


if __name__ == "__main__":
    raise SystemExit(main())
