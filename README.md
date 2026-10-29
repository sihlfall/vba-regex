# VbaRegex—A regular expression engine written entirely in VBA

## Overview

VbaRegex is a regular expression engine written entirely in VBA/VB 6. It is intended to support [JavaScript regular expressions](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Guide/Regular_expressions). The project started as a VBA translation of [Duktape](https://github.com/svaarala/duktape)'s regex engine, but has since deviated considerably.

## Current status

The engine supports most of the JavaScript regular expression syntax.

Currently _not_ supported are, in particular,

* named backreferences like `\k<name>` (but named capturing groups _are_ supported);
* unicode categories like `\p{L}`.

In line with the JavaScript regex specification, [bounded modifiers](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Regular_expressions/Modifier) of the form `(?flags1-flags2:...)` are supported. As an extension to the JavaScript specification, we also support unbounded modifiers (such as `(?i)`).

Your experience with case-insensitive matching may vary—that probably depends on which characters are involved. Please do not expect great results for non-latin characters (but give it a try).

As the project is still work in progress, please do expect that the API may change over time.

## Usage

The engine source code is located in the `src\` directory. You need to import all files in that directory into your project. As an alternative, you can first build a single-file version of the regex engine (see [below](#singlefile)) and import that.

`StaticRegex.bas` provides a relatively simple API.

### Examples

The below examples refer to the following example string:

```vbnet
Dim exampleString As String
exampleString = "On Jul-4-1776, independence was declared. " & _
   "On Apr-30-1789, George Washington became the first president."
```

#### Common step: Initializing a regex with a pattern

```vbnet
Dim regex As StaticRegex.RegexTy

StaticRegex.InitializeRegex regex, _
   "(?<month>\w{3})-(?<day>\d{1,2})-(?<year>\d{4})"
```

The regex itself is stateless—you can re-use it as often as you like.

#### Example 1: Testing whether a string matches the regex

```vbnet
Dim wasFound As Boolean

wasFound = StaticRegex.Test(regex, exampleString)

Debug.Print wasFound   ' prints: True
```

#### Example 2: Getting the first matching substring, as well as submatches

```vbnet
Dim wasFound As Boolean, matcherState As StaticRegex.MatcherStateTy

wasFound = StaticRegex.Match(matcherState, regex, exampleString)

Debug.Print wasFound ' prints: True
Debug.Print StaticRegex.GetCapture(matcherState, exampleString)
   ' prints: 'Jul-4-1776' (entire match)
Debug.Print StaticRegex.GetCapture(matcherState, exampleString, 2)
   ' prints: '4' (second parenthesis)
Debug.Print StaticRegex.GetCaptureByName(matcherState, regex, exampleString, "month")
   ' prints: 'Jul' (capture named "month")
```

#### Example 3: Getting all matching substrings, as well as submatches

```vbnet
Dim matcherState As StaticRegex.MatcherStateTy

Do While StaticRegex.MatchNext(matcherState, regex, exampleString)
   Debug.Print StaticRegex.GetCapture(matcherState, exampleString)
   Debug.Print StaticRegex.GetCaptureByName(matcherState, regex, exampleString, "year")
Loop

' prints:
' Jul-4-1776
' 1776
' Apr-30-1789
' 1789
```

#### Example 4: Joining all matching substrings

```vbnet
Debug.Print StaticRegex.MatchThenJoin(regex, exampleString, delimiter:=", ")
   ' prints: Jul-4-1776, Apr-30-1789
```

#### Example 5: Formatting and joining submatches

```vbnet
Debug.Print StaticRegex.MatchThenJoin( _
   regex, exampleString, delimiter:=", ", format:="$<day> $<month> $<year>" _
)
   ' prints: 4 Jul 1776, 30 Apr 1789
```

#### Example 6: Listing all matching substrings

For this example, we need an array of format strings. Since VBA does not provide a way of creating array constants, let us assume we have a function that creates an array of strings from its parameters:

```vbnet
Private Function MakeStringArray(ParamArray strings() As Variant) As String()
   Dim ary() As String, i As Long
   ReDim ary(0 To UBound(strings) - LBound(strings)) As String
   For i = LBound(strings) To UBound(strings)
      ary(i - LBound(strings)) = strings(i)
   Next
   MakeStringArray = ary
End Function
```

Then we can do the following:

```vbnet
Dim results() As String

StaticRegex.MatchThenList results, _
   regex, exampleString, _
   MakeStringArray("$&", "$<day>", "$<month>", "$<year>")
```

Now results will be a _number of matches_ × _number of format strings_ array of strings with the formatted match results. In our case, `results` will be

```
"Jul-4-1776", "4", "Jul", "1776";
"Apr-30-1789", "30", "Apr", "1789"
```

#### Example 7: Replacing matching substrings

Using the date regex and `exampleString` from above, we can reformat both dates:

```vbnet
Debug.Print StaticRegex.Replace( _
   regex, replacer:="$<day> $<month> $<year>", haystack:=exampleString _
)

' prints: On 4 Jul 1776, independence was declared. On 30 Apr 1789, George Washington became the first president.
```

`Replace` returns a new string, replacing all matches by default. Pass `localMatch:=True` to replace only the first match. If there are no matches, it returns the original string. The replacement uses the same format syntax as `MatchThenJoin` and `MatchThenList`: for example, `$&` inserts the entire match, `$1` inserts the first capturing group, `$<month>` inserts a named capture, and `$$` inserts a literal dollar sign.

#### Example 8: Splitting a string at regex matches

```vbnet
Dim separatorRegex As StaticRegex.RegexTy
Dim parts As Collection, part As Variant

StaticRegex.InitializeRegex separatorRegex, "[,;]\s*"
Set parts = StaticRegex.SplitByRegex(separatorRegex, "red, green; blue")

For Each part In parts
   Debug.Print part
Next

' prints:
' red
' green
' blue
```

`SplitByRegex` returns a `Collection` of strings. Pass `localMatch:=True` to split only at the first match. If there are no matches, the collection contains the original string as its only element. Capturing groups in the separator pattern are also included in the collection, immediately after the substring preceding each separator; use non-capturing groups `(?:...)` when you only want to group parts of the separator pattern. Empty substrings, including those before a separator at the start or after one at the end, are retained.

### Options

The Boolean options below default to `False`.

| Option | Where to set it | Effect when `True` |
| --- | --- | --- |
| `caseInsensitive` | `InitializeRegex` | Ignores case when matching. Set this when compiling the pattern. |
| `multiline` | `Test`, `Match`, `Replace`, `SplitByRegex`, `MatchThenJoin`, `MatchThenList`, or `InitializeMatcherState` | Allows `^` and `$` to match line boundaries within the string as well as the start and end of the string. |
| `dotAll` | Same as `multiline` | Allows `.` to match line terminators as well as other characters. |
| `localMatch` | `Replace`, `SplitByRegex`, `MatchThenJoin`, `MatchThenList`, or `InitializeMatcherState` | Processes only the first match. By default, these operations process all successive, non-overlapping matches. |

For example, to search without regard to case and allow a match to span multiple lines:

```vbnet
Dim messageRegex As StaticRegex.RegexTy
StaticRegex.InitializeRegex messageRegex, "begin.*end", caseInsensitive:=True
Debug.Print StaticRegex.Test(messageRegex, "BEGIN" & vbCrLf & "END", dotAll:=True)
   ' prints: True
```

`Match` also accepts `matchFrom`, the 1-based position at which to start searching (default: `1`). A match may start at that position or later.

When iterating with `MatchNext`, configure the options on the matcher state before starting:

```vbnet
Dim matcherState As StaticRegex.MatcherStateTy
StaticRegex.InitializeMatcherState matcherState, multiline:=True, dotAll:=True

Do While StaticRegex.MatchNext(matcherState, regex, exampleString)
   Debug.Print StaticRegex.GetCapture(matcherState, exampleString)
Loop
```

`InitializeMatcherState` starts iteration from the beginning and sets the three matching options. `ResetMatcherState matcherState` restarts iteration while preserving those options. A newly declared matcher state, as used in Example 3, starts at the beginning with all options set to `False`.

For patterns that require more matching work, you can adjust `regex.stepsLimit` after `InitializeRegex`. This branch defaults to `RegexDfsMatcher.DEFAULT_STEPS_LIMIT` (currently `100000000`) steps for a DFS search. Setting a nondefault budget selects DFS for the public matching methods so that the limit is honored. Reaching the limit returns no match by default; `MatchNext(..., raiseOnLimit:=True)` instead raises `REGEX_ERR_STEP_LIMIT` and selects DFS even with the default budget.

## <a id='singlefile'></a>Building a single-file version of the regex engine

In subdirectory `aio\` (“all-in-one”), you can find a PowerShell script `make_aio.ps1`, which creates a single-file version of the regex engine.

```powershell
cd aio
.\make_aio.ps1 -outModuleName StaticRegexSingle
```

This will create a file named `StaticRegexSingle.bas` in `aio\build\`, which you can then import into your project. For the module, you can choose whatever name you like, as long as it does not conflict with anything. The module you get will provide the same API as `StaticRegex.bas` does.

The shell script does not do any parsing, but is rather based on simple copy/paste and search/replace, so changes in the source code may require changes to the script.

## Tests

### Unit tests

All unit tests are intended to be run with [Rubberduck](https://github.com/rubberduck-vba).

### Testing against RE2

In addition, the regex engine was tested against (a subset of) the test cases for [RE2](https://github.com/google/re2). These test cases, as well as the results delivered by RE2, are available on [Github](https://github.com/sihlfall/regex-test-cases).

Building the test executable requires VB 6.

The [test2 runner](test2/README.md) tests DFS, NFA, and DFA separately. It skips
unsupported bytecode for NFA/DFA and compares DFA's whole-match span and DFS/NFA's
captures against RE2. Results and skip counts are reported per engine.

To run the tests and compare the results, three PowerShell scripts are provided in `test2\`. These scripts expect the following directory structure:

```
|- vba-regex
|   |- src
|   |- test2
|   ...
|- regex-test-cases
```

Build and execute the tests with

```powershell
cd test2
.\make.ps1
.\run-tests.ps1
.\check-test-results.ps1
```

## Resources

* Mozilla developer network (mdn) documentation on JavaScript regular expressions: [https://developer.mozilla.org/en-US/docs/Web/JavaScript/Guide/Regular_expressions](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Guide/Regular_expressions).
* Duktape JavaScript engine: [https://github.com/svaarala/duktape](https://github.com/svaarala/duktape).
* Russ Cox's papers on regex engines:
  * Part 1: “Regular expression matching can be simple and fast.” Jan 2007. – [https://swtch.com/~rsc/regexp/regexp1.html](https://swtch.com/~rsc/regexp/regexp1.html).
  * Part 2: “Regular expression matching: The virtual machine approach.” Dec 2009. – [https://swtch.com/~rsc/regexp/regexp2.html](https://swtch.com/~rsc/regexp/regexp2.html).
  * Part 3: “Regular expression matching in the wild.” Mar 2010. – [https://swtch.com/~rsc/regexp/regexp3.html](https://swtch.com/~rsc/regexp/regexp3.html).
  * Part 4: “Regular expression matching with a trigram index. Or: How Google code search worked.” Jan 2012. – [https://swtch.com/~rsc/regexp/regexp4.html](https://swtch.com/~rsc/regexp/regexp4.html).
* [RegularExpressions.info](https://www.regular-expressions.info/), containing a [reference](https://www.regular-expressions.info/refflavors.html) of the different flavours of regular expressions, listing which features are supported by which engine.
