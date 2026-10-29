# Comparing each matcher with RE2

Build with VB6, then run the existing 20-suite ASCII subset separately through
DFS, NFA, and DFA:

```powershell
cd test2
.\make.ps1
.\run-tests.ps1
.\check-test-results.ps1
# Select engines or use a corpus outside the default sibling directory:
.\run-tests.ps1 -Engines nfa,dfa -TestCasesRoot C:\path\to\regex-test-cases
.\check-test-results.ps1 -TestCasesRoot C:\path\to\regex-test-cases
```

The corpus is [regex-test-cases](https://github.com/sihlfall/regex-test-cases).
Its reference results use RE2's unanchored, leftmost-first mode. The original
`\C` to `.` translation is retained for this ASCII subset. The omitted UTF-8 and
line-ending suites have additional semantic and offset differences.

Each matcher is called directly, bypassing `StaticRegex` dispatch. NFA and DFA
skip only patterns whose compiled bytecode has
`BYTECODE_PROPERTY_REQUIRES_DFS_MATCHING`; `PREFER_DFS` does not prevent testing.
DFS and NFA compare all numbered captures. DFA compares only the complete match
span; no NFA or DFS fallback supplies its results.
Each pattern gets a fresh matcher; its state is reused across that pattern's
strings.

Results are saved in `test-results/<timestamp>/{dfs,nfa,dfa}/`. A `null` entry
means the entire pattern is unsupported by that engine. `[-1, ...]` means no
match; `[-2, ...]` reports a runtime error or DFS step-limit failure, which fails
comparison. Compile errors fail the run and produce `<output>.error.txt`.
`run.json` records the requested suites and engines; missing result files fail
validation. The checker writes `summary.csv`, counts unsupported patterns and
string cases separately, and fails on any mismatch.
It prints the run directory, each suite's loading/comparison phase, and a result
when each suite finishes. Comparisons also report progress about every five
seconds. Large JSON files can take time to load; the last phase message identifies
which file is being read before comparison begins.

The executable also accepts one input file, one output file, and an optional
engine name (`dfs` by default):

```powershell
.\build\RegexTest2Project.exe .\test-cases\test.json .\results.json dfa
```

## Baseline run, 2026-10-28

All 20 existing suites were run with native VB6 (speed optimization) under
Wine 6.0.3, using corpus revision
`0ab3381ad388eaa44576fe33fa66aa3458c61c43`. There are 92,856 pattern entries and
4,651,912 pattern/string combinations per engine. The PowerShell checker and an
independent Python comparison agreed on all counts.

| Engine | Passed cases | Failed cases | Skipped patterns | Skipped cases |
|---|---:|---:|---:|---:|
| DFS | 4,651,912 | 0 | 0 | 0 |
| NFA | 4,392,540 | 13,132 | 2,080 | 246,240 |
| DFA | 4,392,540 | 13,132 | 2,080 | 246,240 |

The run fails validation: NFA/DFA differ from RE2 on nullable repetitions.
For example, on `"a"`, both `(?:|a)*` and `(?:a*?)*` return `[0, 0]` in DFS/RE2
but `[0, 1]` in NFA/DFA. These remain failures, rather than unsupported skips.
The [small fixture](test-cases/nullable-loops.json) reproduces both differences
in native VB6 and p-code. Matcher implementations were unchanged by this work.

[Per-suite counts](results-vb6-native-wine.csv) and
[run metadata](results-vb6-native-wine.json) retain the results without the large
raw matrices. Engine selection, unsupported skips, legacy CLI behavior, JSON
array loading and newline decoding also passed smoke tests in native VB6 and
p-code. The PowerShell runner was exercised through a Wine launcher on two
representative suites; the checker was exercised on all 20 suites and rejects
both an injected mismatch and a missing result file.

Two existing JSON-reader issues were corrected: object-valued Dictionary
assignment fails under Wine, and `\n` must decode to LF rather than CRLF.
Reusing a DFA state across different compiled patterns also exposed VB6 error
10 in `InitializeTables`; the per-pattern isolation used here avoids that path.

## Nullable-priority correction, 2026-10-28

After the [nullable-repetition correction](../doc/nullable_priority.md), the
same 20 suites pass in native VB6 under Wine:

| Engine | Passed cases | Failed cases | Skipped cases |
|---|---:|---:|---:|
| DFS | 4,651,912 | 0 | 0 |
| NFA | 4,405,672 | 0 | 246,240 |
| DFA | 4,405,672 | 0 | 246,240 |

All 13,132 previously failing cases now pass for each of NFA and DFA. The
unsupported-pattern policy is unchanged. An additional generated comparison
against the official RE2 Python wrapper passed 403,512 supported cases per
matcher in both native VB6 and p-code, including nested repetitions,
assertions and numbered captures. The 312 unit tests pass in VB6 native
code/p-code.

[Validation metadata](nullable-priority-validation.json) records these runs.
The original files above retain the baseline results. DFS still differs from
RE2 on some cases outside that original corpus; examples and the distinction
between those semantics are documented in the investigation.
