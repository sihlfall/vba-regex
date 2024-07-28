# VbaRegex—A regular expression engine written entirely in VBA

## Overview

VbaRegex is a regular expression engine written entirely in VBA/VB 6. It is intended to support [JavaScript regular expressions](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Guide/Regular_expressions). The project started as a VBA translation of [Duktape](https://github.com/svaarala/duktape)'s regex engine, but has since deviated considerably.

## Current status

The engine supports most of the JavaScript regular expression syntax.

Currently _not_ supported are, in particular,
* named backreferences like `\k<name>` (but named capturing groups _are_ supported);
* unicode categories like `\p{L}`;
* mode modifiers like `(?i)`.

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
Dim matcherState As MatcherStateTy

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
   ReDim ary(0 To UBound(strings) - LBound(strings) + 1) As String
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
