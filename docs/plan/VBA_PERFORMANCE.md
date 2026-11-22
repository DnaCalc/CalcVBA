# VBA performance against Excel

The 2026-11-22 run compares 64-bit Excel 16.0.20527.20196 and release
CalcVBA on GOVERT-OMEN (AMD Ryzen 7 4800H). The 19 workloads and their sizes
are in [cases.tsv](../../conformance/perf/vba-bench/cases.tsv). Excel compiled
the authored modules via VBE Debug → Compile in a private desktop session.
It timed each workload inside VBA using `QueryPerformanceCounter`. CalcVBA
prepared one project session per workload and timed bound procedure calls on
the owning thread. Each result is the median of five calls after a warm-up.
The timers exclude project preparation and the checksum comparison.

Every workload returns a `Double` checksum. The runners retain its IEEE 754
bits, verify that the warm-up and all five samples agree, and record the five
individual timings. The last three workloads need a spreadsheet host and cover
block and cell-by-cell range writes and reads and Excel recalculation of a VBA
UDF over 10,000 cells. `WorksheetSum` calls the real `WorksheetFunction.Sum`
in Excel and dispatches to a minimal two-argument `Sum` host in standalone
CalcVBA. Its ratios include different host services.

## Reproduce

Run these commands from the repository root. Pass a new scratch directory to
each runner.

```powershell
./conformance/perf/vba-bench/run-excel.ps1 -RunDirectory temp/bench-excel `
    -PrivateDesktopScript ../DnaCalc-win/tools/excel-session/on-private-desktop.ps1
cargo test --release -p calcvba-differential --test vba_bench --no-run
python conformance/perf/vba-bench/run-calc.py <release-test-executable> temp/bench-calcvba
```

The CalcVBA runner works on Windows and Linux. Use `--only ArraySort --size
200000` or `--only Dict --size 100000` to replay the original sizes from the
[November 20 run](../../conformance/perf/vba-bench/calcvba-20261120.txt).
`--backend vm3` or `--backend jit` selects one engine. The Excel equivalents are
`-Only ArraySort -Size 200000` and `-Only Dict -Size 100000`. Run the timing jobs
in sequence and keep other work off the measurement machine where possible.

## Windows

| Workload | Size | Excel ms | VM3 ms (ratio) | JIT ms (ratio) | Checksum vs. Excel |
| --- | ---: | ---: | ---: | ---: | --- |
| LoopLong | 500,000 | 14.806 | 2,707.310 (182.86×) | 345.621 (23.34×) | exact |
| LoopDouble | 300,000 | 20.172 | 1,680.505 (83.31×) | 555.535 (27.54×) | 9 ULP |
| StrBuild | 20,000 | 17.712 | 389.213 (21.97×) | 205.680 (11.61×) | exact |
| StrFuncs | 20,000 | 44.475 | 551.166 (12.39×) | 321.560 (7.23×) | exact |
| ArrayScan | 20,000 | 0.985 | 180.531 (183.35×) | 36.603 (37.18×) | exact |
| ArraySort | 4,096 | 3.683 | 443.580 (120.44×) | 146.538 (39.79×) | exact |
| Coll | 5,000 | 40.274 | 192.843 (4.79×) | 117.513 (2.92×) | exact |
| Dict | 2,000 | 12.379 | 53.188 (4.30×) | 26.266 (2.12×) | exact |
| UdtArray | 50,000 | 6.402 | 747.014 (116.68×) | 313.633 (48.99×) | exact |
| ClassCalls | 50,000 | 8.325 | 543.642 (65.30×) | 273.330 (32.83×) | exact |
| LateClassCalls | 50,000 | 137.489 | 598.507 (4.35×) | 267.415 (1.94×) | exact |
| WorksheetSum | 5,000 | 27.170 | 22.956 (0.84×) | 8.616 (0.32×) | exact |
| FileLines | 1,000 | 4.804 | 48.853 (10.17×) | 38.267 (7.96×) | exact |
| Recursion | 21 | 3.066 | 171.392 (55.89×) | 131.419 (42.86×) | exact |
| VariantMix | 100,000 | 36.162 | 1,031.419 (28.52×) | 683.688 (18.91×) | exact |
| Amort | 2,000 | 56.851 | 231.603 (4.07×) | 207.073 (3.64×) | exact |
| RangeBlock | 2,000 | 7.385 | — | — | — |
| RangeCells | 2,000 | 183.588 | — | — | — |
| UdfRecalculate | 10,000 | 106.002 | — | — | — |

Windows CalcVBA used base 74b7c5b15 and the SAFEARRAY fix included with
this report. The Win64 release binaries were cross-built on dna-firehorse
with cargo-xwin and verified by SHA-256 after transfer. Both engines
finished all 16 standalone workloads. Fifteen gave Excel's exact checksum.
LoopDouble gave 40981B23AA821252 on both engines and 40981B23AA82125B in Excel.

## Linux

Release CalcVBA ran on dna-koderbot (Intel Core Processor, Haswell, no TSX).
The Linux rows are wall-clock measurements on a different CPU. Windows
Dictionary uses the native Scripting COM server. The Linux Dictionary runs
the portable VBA implementation, whose FindKey scans every existing key.

| Workload | VM3 ms | JIT ms | Checksum vs. Excel |
| --- | ---: | ---: | --- |
| LoopLong | 1,260.883 | 398.898 | exact |
| LoopDouble | 839.044 | 469.167 | 9 ULP |
| StrBuild | 295.294 | 240.999 | exact |
| StrFuncs | 273.923 | 260.821 | exact |
| ArrayScan | 78.527 | 28.368 | exact |
| ArraySort | 205.676 | 136.415 | exact |
| Coll | 97.845 | 69.855 | exact |
| Dict | timeout | timeout | — |
| UdtArray | 362.643 | 180.351 | exact |
| ClassCalls | 240.580 | 172.077 | exact |
| LateClassCalls | 287.958 | 148.382 | exact |
| WorksheetSum | 42.115 | 5.882 | exact |
| FileLines | 12.911 | 9.069 | exact |
| Recursion | 89.262 | 56.636 | exact |
| VariantMix | 633.741 | 380.795 | exact |
| Amort | 182.736 | 89.564 | exact |

The [Excel TSV](../../conformance/perf/vba-bench/excel-20261122.tsv),
[Windows TSV](../../conformance/perf/vba-bench/windows-20261122.tsv) and
[Linux TSV](../../conformance/perf/vba-bench/linux-20261122.tsv) retain
the individual runs and checksum bits. CalcVBA also records preparation time.

## Original high-volume cases

| Workload | Size | Excel ms | VM3 ms (ratio) | JIT ms (ratio) | Checksum vs. Excel |
| --- | ---: | ---: | ---: | ---: | --- |
| ArraySort | 200,000 | 235.564 | 23,976.860 (101.78×) | 15,747.911 (66.85×) | exact |
| Dict | 100,000 | 2,271.267 | 3,268.400 (1.44×) | 2,849.222 (1.25×) | exact |

The earlier benchmark ran through the deterministic differential host, which
selects the portable Dictionary even on Windows. It searches existing keys
linearly, so filling and probing 100,000 keys requires quadratic work. The
current Windows runner enables native COM and uses the same Scripting server
as Excel. On Linux, the portable implementation returned 2,600 on both
engines for 100 entries and reached the process limit at 2,000 entries.

Before the fix, Win64 JIT ran past the 35-second process limit for the
4,096-element sort and the 50,000-update UDT case. The generic JIT helpers
cloned the full SAFEARRAY on each read or write. The fix borrows the array
until it has identified the element to read or change. The object
default-member route still retains its receiver across the call.

## VM3 performance

The same-machine Windows ratios justify VM3 speed work. LoopLong,
ArrayScan, ArraySort, and UdtArray each took more than 100 times Excel's
time. LoopLong also had the largest absolute VM3 median in the set.
The JIT accelerated these workloads, and each still took substantially
longer than Excel.

On Linux, 20 GDB stack samples for each of LoopLong and UdtArray
placed both workloads in Vm3::run_loop in all 20 samples. Vm3::exec
appeared in 17/20 and 20/20 respectively. write_loc appeared in 5/20
and 6/20, and store_arith in 4/20 each. The UDT samples also reached
Variant and record cloning. These overlapping stack counts point to
VM3 instruction and value handling for follow-up. Linux perf was
unavailable under the host's existing performance-counter policy.

The three spreadsheet workloads still need VM3 and JIT runs through a
spreadsheet host. The standalone worksheet service measurement exercises
dispatch and the checksum, with the host difference described above.
