# Nullable repetition priority

NFA and DFA now use RE2's priority for nullable repetition. Compilation makes
this decision **after AST construction, in the existing sizing traversal**.
There is no additional AST pass, no guessed exit at runtime, and no change to
the ordinary NFA rule for discarding an already represented state.

## The demonstrated failure

Originally, the forward program for `(?:|a)*` contained:

```text
16  SPLIT1  -> prefer 18, defer 26
18  SPLIT1  -> prefer 20, defer 22
20  JUMP    -> 24
22  CHAR a  -> 24
24  JUMP    -> 16
26  SAVE match end
28  MATCH
```

A VB6 trace at input offset zero visits `16, 18, 20, 24, 16`. The second visit
to 16 is discarded. The pending order is then 22 (consume `a`), followed by
26 (continue after the repetition). That explains the incorrect preference
for consuming. After consuming a character, DFA can instead first repeat
instruction 24, so a correction limited to duplicate split instructions is
insufficient.

The diagnosis was right for this execution, but does not imply that every
discard should reorder pending paths. A prototype which claimed a visited
split's still-pending second branch passed the original 308 tests, yet
changed `^(?:[ab]*?)*b` on `abb` from `ab` to `abb`. A uniform two-split
construction for every star also fails nested lazy repetition. Those
prototypes were discarded.

## Compilation and entry identity

The sizing traversal records `NODE_NULLABLE` when returning from a child,
alongside its existing consumption and expansion properties. Assertions
are conservatively nullable; a failing expression is not nullable. Optional
bounded repetition is nullable too, including when it is later expanded.
Disabling expansion retains these facts through a single `expandRepeats`
flag; it does not require another traversal or a scan to clear the plan.

* A nonnullable star keeps its existing single-split loop.
* A nullable star uses an initial split and a separate post-body split,
  equivalent to `(body+)?` with a shared body. This replaces its two-word
  tail jump with a two-word `REOP_LOOP_GREEDY` or `REOP_LOOP_HUMBLE`.
* During that same sizing traversal, the parser's shared-body representation
  `A A*` is normalized to an internal plus node. `A{n,}` becomes `n-1`
  mandatory copies followed by a shared-body plus. This preserves capture
  and priority behavior that duplicating the first body would lose.
* `REOP_LOOP_INIT` is an explicit entry marker for plus. NFA/DFA treat it as
  an ordinary epsilon instruction. DFS needs its distinct position so that
  nested plus expressions have distinct progress landmarks.

This follows the nullable-star and shared-body-plus constructions in
[RE2's compiler](https://github.com/google/re2/blob/main/re2/compile.cc).
The original restriction on nullable-aware compilation was explicitly
relaxed for this implementation, with the condition that it add no AST pass.

For the failing example, instruction 24 is now a split which first saves
26 as a fallback and then tries the body at 18. Discarding that repeated
body resumes the newly saved 26 **before** the previously pending 22. The
continuation carries the completed iteration's captures. If it fails,
22 remains available with its own capture frame. Greedy/lazy ordering comes
from the split order, rather than a runtime rule about backward jumps.

The NFA also records the first thread after each input advance in the active
visited set, and releases a capture-frame reference when an inactive-state
merge does not enqueue its thread. The latter preserves copy-on-write
ownership when paths converge.

## Deferred assertions in DFA

The general discard rule is **dominance**: discard only the continuations
already covered by an earlier, higher-priority path.

DFA formerly keyed visits by an exact conjunction of pending assertion
flags. A state visited without a condition and that same state visited
under a word-boundary condition could therefore be explored independently,
which loses priority around cycles.

Conditions now denote sets of four disjoint next-character categories:
word, ordinary nonword, line terminator, and end of input. Assertions
intersect the current set. Each visited state records the union of categories
already explored; a later visit continues only for the remaining categories:

```text
remaining = incoming AND NOT seen
seen      = seen OR remaining
```

Fix any one category. Restricting the closure to that category gives the
usual ordered epsilon traversal: an instruction is explored only at its
first visit, and assertions either admit or reject that category. Thus
subtracting covered categories preserves ordinary first-path dominance
without dropping contexts for which the earlier assertion fails. Both
active closure visits and queued output states use this rule.

Residual masks are serialized in the ordered DFA state. Existing state
hashing, comparison, cache retention and transition-action replay therefore
include the complete condition information. The implementation uses small
signed `Long` masks with `And`, `Or`, and `Not`; no shifts or overflow-based
arithmetic are introduced.

## RE2 expectations, including differences from DFS

Direct checks used the official `google-re2` Python package,
version `1.1.20251105`. Spans below are zero-based, half-open.

| Pattern | Input | RE2 result |
| --- | --- | --- |
| `(?:\|a)*` | `aa` | `[0,0)` |
| `(?:a*?)*` | `aa` | `[0,0)` |
| `(?:\|a)*a` | `aaa` | `[0,1)` |
| `^(?:\|a)*b$` | `aab` | `[0,3)` |
| `^(?:[ab]*?)*b` | `abb` | `[0,2)` |
| `(?:\b\|.)*` | `!aaa` | `[0,4)` |
| `(?:\b\|.)*` | `aaa` | `[0,0)` |
| `(a*)*` | `a` | match `[0,1)`, group 1 `[0,1)` |

The originally proposed `!` expectation for the boundary example was a DFS
result, not a RE2 result; the regression now expects `!aaa`. DFS also captures
the final empty iteration of `(a*)*`, whereas RE2/NFA capture `a`. DFS's
existing behavior is retained: its results before and after this change
were identical on the original 248,712-case generated comparison. This
change fixes NFA/DFA; it does not establish universal DFS/RE2 equivalence.

## Regression coverage and validation

`Ast_NullableStarsSeparateEntryFromIteration` documents the exact forward
and reverse bytecode, greedy/lazy variants, and matching with expansion
both enabled and disabled. `Compiler_LoopEntryPreservesNestedPlusProgress`
protects the entry marker's DFS role. Matcher tests cover preferred empty
branches, consuming fallbacks, assertions, nested lazy repetition, bounded
optionals, numbered/named captures, and repeated DFA calls. The existing
DFA matrix also exercises offsets, modifiers, reverse scans and cache reuse.

Validation on 2026-10-28:

* Unit tests: all 312 tests pass in VB6 native code and p-code.
* Generated RE2 comparison: 403,512 supported cases pass per matcher, for
  NFA and DFA in both VB6 native code and p-code. Another 100,104 cases are
  skipped because the matcher or RE2 does not support the pattern.
* Original 20-suite RE2 corpus, VB6 native: DFS passes all 4,651,912 cases;
  NFA and DFA each pass 4,405,672 cases, with 246,240 unsupported cases
  skipped. The previous 13,132 NFA/DFA mismatches are eliminated.

These comparisons are evidence over the tested cases, not a claim of
exhaustive equivalence for every regex or Unicode convention. The existing
corpus's UTF-8/line-ending exclusions and unsupported-pattern policy remain
as documented in `test2/README.md`.
