"""Package the generated engine and worksheet functions without launching Excel."""

import argparse
from importlib.metadata import PackageNotFoundError, version
import os
from pathlib import Path
import re
import sys
from tempfile import TemporaryDirectory
from xml.etree import ElementTree
from zipfile import ZipFile


WRITER_VERSION = "6.5.0"
MODULE_NAMES = ("StaticRegexSingle", "VbaRegexWorksheetFunctions")
CONTENT_TYPES = {
    ".xlam": "application/vnd.ms-excel.addin.macroEnabled.main+xml",
    ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
}


def load_writer():
    if sys.version_info < (3, 10):
        raise RuntimeError("Python 3.10 or newer is required.")
    requirements = Path(__file__).with_name("requirements-build.txt")
    install = f'& "{sys.executable}" -m pip install -r "{requirements}"'
    try:
        installed = version("pyOpenVBA")
    except PackageNotFoundError:
        raise RuntimeError(f"Missing build dependency. Install it in PowerShell with: {install}") from None
    if installed != WRITER_VERSION:
        raise RuntimeError(
            f"Expected pyOpenVBA {WRITER_VERSION}, found {installed}. Run in PowerShell: {install}"
        )
    from pyopenvba import ExcelFile

    return ExcelFile


def crlf(source):
    return source.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")


def read_module(path, expected_name):
    data = Path(path).read_bytes()
    # PowerShell 7 writes UTF-8; Windows PowerShell and exported VBA modules
    # commonly use Windows-1252. Never replace undecodable source characters.
    try:
        source = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        source = data.decode("cp1252")
    names = re.findall(r'^Attribute\s+VB_Name\s*=\s*"([^"]+)"\s*$', source, re.MULTILINE)
    if names != [expected_name]:
        raise ValueError(f"{path}: expected one VB_Name attribute for {expected_name}.")
    # The writer's Excel templates use code page 1252. Reject lossy conversion
    # before writing anything, rather than silently changing the VBA source.
    try:
        source.encode("cp1252")
    except UnicodeEncodeError:
        raise ValueError(f"{path}: VBA source contains characters outside Windows-1252.") from None
    return crlf(source)


def validate_package(path, expected_modules, excel_file):
    with ZipFile(path) as archive:
        damaged = archive.testzip()
        if damaged:
            raise ValueError(f"Corrupt package entry: {damaged}")
        types = ElementTree.fromstring(archive.read("[Content_Types].xml"))
        actual_type = next(
            (item.get("ContentType") for item in types if item.get("PartName") == "/xl/workbook.xml"),
            None,
        )
        if actual_type != CONTENT_TYPES[path.suffix.lower()]:
            raise ValueError("Workbook content type does not match its extension.")
        archive.getinfo("xl/vbaProject.bin")
    with excel_file(path) as workbook:
        issues = workbook.validate()
        if issues:
            raise ValueError("Invalid VBA project: " + "; ".join(issues))
        actual_modules = workbook.vba_modules()
        if actual_modules != expected_modules:
            raise ValueError("Packaged VBA source differs from the build inputs.")


def build_addin(engine, wrapper, output):
    excel_file = load_writer()
    output = Path(output).resolve()
    if output.suffix.lower() not in CONTENT_TYPES:
        raise ValueError("Output must have the .xlam or .xlsm extension.")
    sources = {
        name: read_module(path, name)
        for name, path in zip(MODULE_NAMES, (engine, wrapper))
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    # Build and verify beside the destination, then replace it atomically. A
    # dependency, encoding, or packaging failure leaves any existing add-in intact.
    with TemporaryDirectory(prefix=".addin-", dir=output.parent) as directory:
        temporary = Path(directory) / output.name
        with excel_file.create_new(temporary) as workbook:
            project = workbook.vba_project()
            if project.code_page != 1252:
                raise ValueError("Unexpected VBA template code page.")
            project.delete_module("Module1")
            for name, source in sources.items():
                project.add_module(name, source)
            # Supplying the body preserves the template's document-module
            # attributes, including its binding to Excel's ThisWorkbook object.
            events = Path(__file__).with_name("WorkbookEvents.vba").read_text(encoding="utf-8-sig")
            workbook.set_module("ThisWorkbook", crlf(events))
            expected_modules = workbook.vba_modules()
            workbook.save()
        validate_package(temporary, expected_modules, excel_file)
        os.replace(temporary, output)
    return output


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="check build dependencies only")
    parser.add_argument("--engine", type=Path, help="generated StaticRegexSingle.bas")
    parser.add_argument("--output", type=Path, help="destination .xlam or .xlsm")
    args = parser.parse_args()
    if not args.check and (args.engine is None or args.output is None):
        parser.error("--engine and --output are required")
    try:
        if args.check:
            load_writer()
            return
        output = build_addin(
            args.engine, Path(__file__).with_name("VbaRegexWorksheetFunctions.bas"), args.output
        )
        print(f"Built {output} (without Excel).")
    except Exception as error:
        raise SystemExit(f"Add-in build failed: {error}") from None


if __name__ == "__main__":
    main()
