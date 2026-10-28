"""Prepare scratch Office projects using the stdVBA-owned class builder."""
import argparse
import importlib.util
import json
from pathlib import Path
import re
import subprocess

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]


def write_vba(path, text):
    path.write_bytes(text.replace('\r\n', '\n').replace('\n', '\r\n').encode('cp1252'))


def read_vba(path):
    try:
        return path.read_text(encoding='utf-8-sig')
    except UnicodeDecodeError:
        return path.read_text(encoding='cp1252')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--stdvba-root', type=Path, default=ROOT.parent / 'stdVBA')
    parser.add_argument('--baseline', default='1619cd75526903045b8c5cb8c6048f93bc4160be')
    parser.add_argument('--candidate', default='HEAD')
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    tools = args.stdvba_root / 'Tools/stdRegex-build'
    spec = importlib.util.spec_from_file_location('stdvba_builder', tools / 'build.py')
    builder = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(builder)
    args.output.mkdir(parents=True, exist_ok=True)
    extra = '''
Public Function BenchEngineTest(ByVal text As String) As Boolean
  BenchEngineTest = stdRegex.protStaticRegexTest(This.regex, text, multiline:=This.isMultiline, dotAll:=This.dotAll)
End Function
Public Function BenchIterate(ByVal text As String) As Long
  Dim state As MatcherStateTy
  state.multiline = This.isMultiline: state.dotAll = This.dotAll
  Do While stdRegex.protStaticRegexMatchNext(state, This.regex, text)
    BenchIterate = BenchIterate + 1
  Loop
End Function
Public Sub BenchCompile()
  Dim compiled As RegexTy
  stdRegex.protStaticRegexInitializeRegex compiled, This.regex.pattern, This.regex.isCaseInsensitive
End Sub
'''
    manifest = json.loads((tools / 'engine.json').read_text())
    for name, ref in [('RegexBaseline', args.baseline), ('RegexCandidate', args.candidate)]:
        revision = builder.resolve_revision(ROOT, ref)
        source = builder.generate(ROOT, dict(manifest, revision=revision))
        source += extra
        source = re.sub(r'\bstdRegex\b', name, source)
        write_vba(args.output / (name + '.cls'), source)
    modules = []
    for path in sorted((ROOT / 'test').glob('Test*.bas')):
        # Historical test targets removed LongArrayBuffer/RegexRedBlackTree modules
        # and is no longer part of TestVbaRegex.vbp.
        if path.name == 'TestRedBlackTree.bas':
            continue
        text = read_vba(path)
        tests = re.findall(r"'@TestMethod[^\n]*\n(?:[^\n]*\n)*?Private Sub (\w+)\(", text)
        if not tests:
            continue
        text = text.replace('Set Assert = CreateObject("Rubberduck.AssertClass")', 'Set Assert = New PortableAssert')
        text = text.replace('Set Fakes = CreateObject("Rubberduck.FakesProvider")', 'Set Fakes = Nothing')
        text = re.sub(r'^Option Private Module\s*\n', '', text, flags=re.M)
        # Existing tests remain private and execute through one scratch entry point.
        calls = '\n'.join(('  TestInitialize\n' if 'Private Sub TestInitialize(' in text else '') +
                          '  ' + name + ('\n  TestCleanup' if 'Private Sub TestCleanup(' in text else '') for name in tests)
        text += '''
Public Function RunPortableTests() As String
  ModuleInitialize
''' + calls + '''
  RunPortableTests = "checks=" & Assert.checks & vbLf & Assert.failures
  ModuleCleanup
End Function
'''
        write_vba(args.output / path.name, text)
        modules.append(re.search(r'^Attribute VB_Name = "([^"]+)"', text, re.M)[1])
    for name in ('PortableAssert.cls', 'EngineBench.bas'):
        write_vba(args.output / name, (HERE / name).read_text(encoding='utf-8-sig'))
    for name in ('TestReporter.cls', 'TestRunner.bas'):
        write_vba(args.output / name, (tools / name).read_text(encoding='utf-8-sig'))
    tests = (args.stdvba_root / 'tests/stdRegexTests.bas').read_text(encoding='utf-8-sig')
    write_vba(args.output / 'stdRegexTests.bas', re.sub(r'\bstdRegex\b', 'RegexCandidate', tests))
    # The fixed stdVBA regression runner refers to the test module, not the class.
    (args.output / 'test-modules.json').write_text(json.dumps(modules))
    (args.output / 'metadata.json').write_text(json.dumps({
        'baseline': builder.resolve_revision(ROOT, args.baseline),
        'candidate': builder.resolve_revision(ROOT, args.candidate),
        'native_test_modules': modules,
    }, indent=2))


if __name__ == '__main__':
    main()
