# Excel add-in

The add-in exposes the VBA regex engine as worksheet functions. The Rx prefix
distinguishes these functions from Excel's native regex functions.

## Build and load

Requires Windows desktop Excel, PowerShell, and **Trust access to the VBA project
object model** enabled in Excel's Trust Center for building and running tests.

From the repository root:

    .\excel_addin\make_addin.ps1

The build regenerates the engine from current sources and writes
excel_addin/build/VbaRegexAddIn.xlam. In Excel, use **File > Options > Add-ins >
Manage: Excel Add-ins > Go > Browse** to select it, then enable it. Permit macros
according to your Excel security settings. VBA project access is needed for
building/testing, not for calling the installed functions.

Functions appear in the **VbaRegex** category of Insert Function, with argument
descriptions. The build also supports:

    .\excel_addin\make_addin.ps1 -createExampleFile $true

This produces build/WorksheetFunctionsExample.xlsm with the functions embedded;
it does not populate example worksheets.

## Worksheet API

Optional arguments are shown in brackets:

    RxTest(text, pattern, [flags], [startAt])
    RxMatch(text, pattern, [format], [occurrence], [flags], [startAt])
    RxFind(text, pattern, [occurrence], [flags], [startAt])
    RxLength(text, pattern, [occurrence], [flags], [startAt])
    RxCount(text, pattern, [flags], [startAt])
    RxMatches(text, pattern, [format], [flags], [startAt])
    RxReplace(text, pattern, replacement, [occurrence], [flags], [startAt])
    RxSplit(text, pattern, [ignoreEmpty], [flags])

| Function | Default behavior | Result |
| --- | --- | --- |
| RxTest | Does any match exist? | Boolean |
| RxMatch | Extract the first match | Text, or #N/A |
| RxFind | Find the first match's position | 1-based position, or #N/A |
| RxLength | Measure the first match | Length, or #N/A |
| RxCount | Count all nonoverlapping matches | Number, zero for no matches |
| RxMatches | Extract all matches | One match per row, or #N/A |
| RxReplace | Replace every match | Text |
| RxSplit | Split each source row at every delimiter match | Fields across columns; shorter rows padded with #N/A |

These functions use this repository's engine and supported regex syntax, not
Excel's native PCRE2 engine. Extracted text, replacements, and split fields remain
text even when they contain only digits. RxFind, RxLength, and RxCount return
numbers.

### Flags

Flags default to "". Combine lowercase letters in any order:

- i: case insensitive.
- m: ^ and $ recognize line boundaries.
- s: dot-all; a dot also matches line breaks.

For example, "im" combines case-insensitive and multiline matching. Repeated
flags are accepted. Unknown letters, uppercase letters, and spaces return #VALUE!.
Global matching is determined by the function, not a g flag.

### Positions and occurrences

startAt defaults to 1 and is an integer character position, measured in UTF-16
code units as in VBA and the engine. Supplementary Unicode characters occupy
two positions; this may differ from modern Excel text-function counting.

Valid positions range from 1 through the text length plus one. The terminal
position allows an empty match at the end; it also permits startAt=1 on empty
text. Searches retain the complete original input for anchors, boundaries, and
capture formatting. Replacement preserves the prefix before startAt.

occurrence counts nonoverlapping matches from startAt in the original input:

- RxMatch, RxFind, and RxLength: defaults to 1; must be a positive integer.
- RxReplace: defaults to 0 for all matches; a positive integer replaces just
  that match. An absent occurrence leaves the text unchanged.

RxFind returns the absolute position in the original text, even when startAt is
greater than 1. RxLength measures the complete match, not a capture or formatted
result. Both use UTF-16 code units. For an empty match, RxFind returns the match
boundary and RxLength returns zero; a missing occurrence returns #N/A.

RxCount counts all nonoverlapping matches from startAt, including empty matches.
It returns zero when there are none and does not construct an array of matches.
If the execution limit is exceeded while counting, it returns #NUM! rather than
a partial count.

Negative occurrences, fractional numbers, Boolean positions/occurrences, and
numbers supplied as text are invalid. Empty position/occurrence cells use their
defaults. Empty matches count as occurrences and advance to ensure termination.

### Capture formatting

format defaults to "$&" (the complete match). Formats and replacements use the
engine's existing syntax, including:

- $1, $2, ... for numbered captures.
- $<name> for named captures.
- $$ for a literal dollar.
- $& for the full match.
- A dollar followed by a backtick for the original prefix; $' for the suffix.

An empty format returns empty text for a successful match. A malformed format
returns #VALUE!, even when there is no match.

RxMatches also accepts one row of formats, from a range or array constant.
Each match becomes a row and each format becomes a column. Store formats such
as $1 in cells formatted as **Text** to prevent Excel interpreting them as
currency.

### Ranges, arrays, and spills

Every argument of RxTest, RxMatch, RxFind, RxLength, RxCount, and RxReplace
can be a scalar, contiguous
range, or worksheet array. On each axis, sizes must agree or be one. Singleton
rows/columns broadcast; incompatible sizes return #VALUE!. Per-cell validation,
pattern errors, execution limits, and input errors affect only the corresponding
output cells.

RxMatches accepts a single text, pattern, and scalar options (single-cell ranges
or one-cell arrays are allowed), plus the format row described above. It does not
combine multiple source texts into one result.

RxSplit accepts scalars or single columns for every argument. Scalar text,
patterns, ignoreEmpty, and flags broadcast down the common input height. All
non-singleton heights must agree. Each source row becomes one output row, and its
fields expand across columns. Shorter results are padded with #N/A to match the
widest result. Multiple-column source arguments are rejected with #VALUE!.

In Excel with dynamic arrays, results spill into neighboring cells. Use
TRANSPOSE to change orientation. Blocked output cells cause Excel's #SPILL!.
Older Excel versions require a selected output range and a legacy array formula;
they cannot resize the output automatically.

Text arguments, patterns, formats, and flags must be text or blank; numbers are
not implicitly converted. Use TEXT or an explicit text conversion when needed.
Blank text cells are treated as "". ignoreEmpty accepts a Boolean or a
numeric Boolean (zero is false); blank is false.

### Splitting

ignoreEmpty defaults to FALSE. Consecutive, leading, and trailing delimiters
produce empty fields; TRUE removes them. Delimiter captures do not add fields.
A missing delimiter returns the original text. Zero-length delimiters split at
the matched boundaries, including empty edge fields unless ignored.

For a column of inputs, the output preserves every source row, including rows
with no remaining fields after ignoreEmpty. Empty fields that are retained are
""; missing fields used to pad the output are #N/A. If a row produces an input,
pattern, option, execution-limit, or output-limit error, that error fills its
output row while the other rows continue normally.

For example, splitting a column containing "a,b,c", "d", and "e,f" at commas
returns three rows:

| Field 1 | Field 2 | Field 3 |
| --- | --- | --- |
| a | b | c |
| d | #N/A | #N/A |
| e | f | #N/A |

### Errors and limits

| Situation | Result |
| --- | --- |
| No match for RxTest | FALSE |
| Requested match absent for RxMatch, RxFind, or RxLength | #N/A |
| No matches for RxCount | 0 |
| No matches for RxMatches | #N/A |
| All split fields removed for a source row | #N/A across that output row |
| Successful empty match | "" for extraction; length 0; counted by RxCount |
| Invalid pattern, format, option, position, argument type, or dimensions | #VALUE! |
| Engine execution limit exceeded | #NUM! |
| Result exceeds 32,767 characters per cell or worksheet row/column limits | #NUM! |
| Input cell contains an error | That error propagates |

The engine's existing default execution limit applies to each attempted match at
a starting position. New Rx functions report exhaustion rather than treating it
as a missing match.

## Examples

Formulas below use English function names and comma argument separators. Adapt
separators to your Excel locale.

    =RxTest(A2,"^invoice","im")
    =RxMatch("a12 b34","\d+",,2)
    =RxReplace("a12 b34","\d+","X",1)
    =RxMatch("1a2b3","\d+",,1,,3)
    =IFNA(RxMatch(A2,"\d+"),"")
    =RxMatches("AB-12 CD-34","([A-Z]+)-(\d+)",{"$1","$2"})
    =RxSplit("red, green;blue","[,;]\s*")
    =RxSplit(A2:A10,"[,;]\s*")
    =RxMatch(A2:A10,"\d+")
    =RxFind("ab12cd345","\d+",2)
    =RxLength("ab12cd345","\d+",2)
    =RxCount("ab12cd345","\d+")
    =MID(A2,RxFind(A2,"\d+"),RxLength(A2,"\d+"))

The second formula returns "34"; the third returns "aX b34". The capture-table
formula spills two rows (AB, 12 and CD, 34). The single-text split formula spills
red, green, and blue across three columns.

The RxFind/RxLength/RxCount examples return 7, 3, and 2 respectively. The final
formula uses the first match's position and length to extract it with MID.

## Compatibility

RegexTest, RegexMatch, and RegexReplace retain their original signatures,
Boolean options, and result conventions, including "" for a missing RegexMatch.
Existing formulas do not acquire the new argument ordering. Excel's native regex
function names may collide with these legacy names; prefer Rx names for new
formulas.

## Tests

From the repository root:

    .\excel_addin\run-tests.ps1

The runner regenerates the engine, imports the wrapper and test modules into a
temporary hidden workbook, runs the suites, and closes it without saving.
Requires Excel with dynamic-array support for the spill-formula tests.

To also build, reopen, and test the actual add-in from a separate workbook:

    .\excel_addin\run-tests.ps1 -TestPackage

This updates build/VbaRegexAddIn.xlam but does not install it.

If the default add-in is loaded in Excel, build and validate a separately named
copy rather than overwriting the file in use:

    .\excel_addin\run-tests.ps1 -TestPackage -AddinFileName VbaRegexAddIn-next.xlam

The build script also accepts -outFileName with a file name (not a path) to select
an alternate file within the build directory. Use the same extension as the
chosen build mode: .xlam for an add-in or .xlsm for the example workbook. Unload
the previous add-in before loading a replacement so only one version supplies
the Rx function names.

- TestWorksheetDimensions.bas: original API dimension checks.
- TestWorksheetFunctions.bas: original API behavior and worksheet formulas.
- TestRxDimensions.bas: every pair of broadcast arguments, both axes and size
  orders, using arrays and actual Range objects; RxSplit additionally rejects
  multiple-column arguments.
- TestRxFunctions.bas: new API results, flags, positions, occurrences,
  formatting, capture tables, numeric positions/lengths/counts, split-column
  broadcasting and padding, zero-length matches,
  validation, local errors,
  execution/output limits, engine formatting helpers, and worksheet spills.
- test-built-addin.ps1: built add-in loading, function registration, and formulas
  for all eight functions in a consumer workbook.

Assertions check Variant types, array shapes, and values; failures identify the
case and output cell.
