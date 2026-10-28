# Isolated engine optimization experiments

Each optimization branch starts from the same engine commit,
`1619cd75526903045b8c5cb8c6048f93bc4160be`, plus this common harness. Changes are
independent; their speedups must not be added together.

The harness uses stdVBA's `Tools/stdRegex-build/build.py` and wrapper template to
generate temporary baseline and candidate classes. The class generator and
wrapper remain owned by stdVBA; no second maintained copy is introduced here.
Both classes use the optimized List/ListArr wrapper, so these experiments isolate
engine changes, not the earlier wrapper improvement.

## Run

Requires Windows Office Word, trusted access to its VBA project model, Python,
Git, and a sibling stdVBA checkout containing the build tools. Use VBA's **Break
on Unhandled Errors** setting so intentional error tests can reach their handlers.

```powershell
./benchmarks/engine-optimizations/run.ps1 -Candidate HEAD -OutputDir <results-directory>
```

`-StdVbaRoot`, `-Baseline`, `-Samples` (default 7), `-TargetSeconds` (default 0.025),
and `-TestsOnly` are available. The runner does not change Office security settings
or save the scratch document.

Correctness runs the maintained native tests through a small assertion adapter,
the stdVBA public API regression suite, and differential tests covering flags,
captures, assertions, backreferences, zero-width matches, processing budgets,
in-place input mutation, and changing capture layouts. The historical
TestRedBlackTree module references removed modules and is excluded; it is also
absent from the native test project.

## Measurement

The original six cases are retained. Additional cases exercise literal-prefix
failures and sparse matches, absolute anchors, long literal runs, UTF-16 text,
case-insensitive backreferences, backtracking, and Boolean tests with captures.
Each case measures List, ListArr, MatchAll, MatchAllArr, the engine's Boolean Test,
capture iteration without objects, Replace, and Compile.

Baseline and candidate execute in the same Word VBA process. Outputs are compared
before timing. Each operation receives two warm-up calls, calibrated batches, and
seven samples, with operation/pair order rotated. QueryPerformanceCounter includes
result allocation and destruction; matching benchmarks reuse compiled patterns.
The separate Compile operation includes compilation. Core Test is measured through
the engine's Test API; stdVBA's public Test uses a separate wrapper path.

Normalized elapsed time is candidate/baseline × 100: lower is better; 100% means
the same time. Aggregates use an equal-weight geometric mean of per-case/operation
ratios, not a sum dominated by dictionary allocation. Both the original 24-cell
suite and the expanded 112-cell suite are reported. Raw sample distributions and
per-cell times remain available; small changes should be considered alongside the
unchanged-engine control and observed variability.

Measurements describe this host's interpreted VBA runtime, not compiled VB6 or all
Office platforms. The UTF-16 byte-buffer experiment includes conversion, content
validation, and cache refresh costs inside timing.
