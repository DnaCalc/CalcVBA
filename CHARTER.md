# CalcVBA

CalcVBA (the DnaCalc/CalcVBA repository) is a portable, optimisable, integrable Rust
implementation of VBA as used in Office, especially Excel. We control its
implementation so it can serve two purposes:

- Open and evaluate Excel files containing VBA in DNA Calc.
- Compile and reshape VBA with agentic tools into useful forms, including
  standalone libraries, native exports, COM server DLLs and executables.

Match real VBA compile-time and runtime behaviour. Use public specifications,
published research and reproducible black-box Excel/VBA observations. Do not
use proprietary implementation sources or reverse engineer Office internals.

Keep VM3 as the readable interpreter and the Cranelift JIT as the optimising
backend. Keep useful regression tests and VM3/JIT/Excel comparisons. The
runtime is single threaded; the compiler may use threads for speed.

Support other platforms where practical. VBA UserForms are in scope through
CalcForms/OxForms. VB6 forms and support for old VB6 projects are out of scope.
Spreadsheet calculation and application object models belong to the host.

Prefer working code and clear tests over process apparatus. Formal proofs,
certification programmes, evidence archives and workset paperwork are not
required for development.
