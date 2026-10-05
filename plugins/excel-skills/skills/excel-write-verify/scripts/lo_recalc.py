#!/usr/bin/env python3
"""Independent recalculation: LibreOffice recomputes every formula of a workbook, and the results are
compared with the results stored in the file. A stored result that a second engine cannot reproduce
is either a wrong cached value written by a script, or a formula LibreOffice does not support.

  lo_recalc.py BOOK.xlsx [--tolerance 1e-6] [--keep DIR]

Needs LibreOffice (soffice on PATH, or the macOS app) and openpyxl. Uses a throwaway LibreOffice
profile, so it does not touch the user's profile or a running LibreOffice. BOOK is never modified.
Exit code 1 if stored results differ; 0 otherwise; 2 for a usage error or when LibreOffice or the file
cannot be run or read. Cells that LibreOffice cannot compute are reported and not failed: a run with 0
differences and many "not checkable" cells has verified only part of the workbook (the last line says how
many). A difference in a column that depends on such cells is usually LibreOffice, not the file. Volatile
functions (TODAY, NOW, RAND) differ by nature between two runs.
"""
import argparse, collections, shutil, subprocess, sys, tempfile, warnings
from pathlib import Path

PROFILE = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
"""   # 0 = always recalculate on load. The default keeps the stored results, which would prove nothing.


def soffice():
    for c in ("soffice", "libreoffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        p = shutil.which(c)
        if p:
            return p
    print("LibreOffice not found (soffice). Install it, or check the workbook in Excel on a scratch copy.", file=sys.stderr)
    sys.exit(2)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("book", type=Path)
    ap.add_argument("--tolerance", type=float, default=1e-6, help="relative tolerance for numbers")
    ap.add_argument("--keep", type=Path, help="keep the recalculated copy in this folder")
    if len(sys.argv) == 1:
        ap.print_help(sys.stderr)
        sys.exit(2)
    a = ap.parse_args()
    from openpyxl import load_workbook
    warnings.simplefilter("ignore")

    work = Path(tempfile.mkdtemp(prefix="lo_recalc_"))
    try:
        profile = work / "profile"
        (profile / "user").mkdir(parents=True)
        (profile / "user" / "registrymodifications.xcu").write_text(PROFILE, encoding="utf-8")
        src = work / "in" / a.book.name
        src.parent.mkdir()
        shutil.copy2(a.book, src)
        out = work / "out"
        if not a.book.is_file():
            print(f"no such file: {a.book}", file=sys.stderr)
            sys.exit(2)
        try:
            r = subprocess.run([soffice(), "--headless", "--norestore", f"-env:UserInstallation={profile.as_uri()}",
                                "--convert-to", "xlsx", "--outdir", str(out), str(src)],
                               capture_output=True, text=True, timeout=900)
        except subprocess.TimeoutExpired:
            print("LibreOffice did not finish in 900 s", file=sys.stderr)
            sys.exit(2)
        redone = out / (a.book.stem + ".xlsx")
        if not redone.exists():
            print(f"LibreOffice did not produce a file:\n{r.stdout}\n{r.stderr}", file=sys.stderr)
            sys.exit(2)

        try:
            formulas = load_workbook(a.book, data_only=False)
            stored = load_workbook(a.book, data_only=True)
            fresh = load_workbook(redone, data_only=True)
        except Exception as e:
            print(f"cannot read the workbook: {e}", file=sys.stderr)
            sys.exit(2)
        compared = verified = 0
        diff, unchecked, no_result = collections.Counter(), collections.Counter(), collections.Counter()
        num = lambda v: isinstance(v, (int, float)) and not isinstance(v, bool)
        err = lambda v: isinstance(v, str) and v.startswith(("#", "Err:"))
        samples = collections.defaultdict(list)
        for ws in formulas.worksheets:
            if ws.title not in fresh.sheetnames:
                print(f"sheet {ws.title!r} is missing in the recalculated copy")
                continue
            S, F = stored[ws.title], fresh[ws.title]
            # Only cells that exist: a sheet can declare a used range of a million empty rows.
            for c in list(ws._cells.values()):
                v = c.value
                if not (isinstance(v, str) and v.startswith("=")) and not hasattr(v, "ref"):
                    continue                                        # only formula cells
                so, fo = S._cells.get((c.row, c.column)), F._cells.get((c.row, c.column))
                old, new = (so.value if so is not None else None), (fo.value if fo is not None else None)
                key = (ws.title, c.coordinate.rstrip("0123456789"))
                compared += 1
                if old == new or (old in ("", None) and new in ("", None)):
                    verified += 1
                    continue
                if num(old) and num(new):
                    if abs(old - new) > a.tolerance * max(1.0, abs(old), abs(new)):
                        diff[key] += 1
                        samples[key].append((c.coordinate, old, new))
                    else:
                        verified += 1
                elif old in ("", None) and not err(new):
                    no_result[key] += 1                             # nothing stored, LibreOffice computes a value
                elif err(old) or err(new) or new in ("", None):
                    unchecked[key] += 1                             # LibreOffice could not reproduce it
                else:
                    diff[key] += 1                                  # two real values of different value or type
                    samples[key].append((c.coordinate, old, new))
        print(f"{compared} formula cells compared")
        by_sheet = collections.Counter()
        for (sheet, _), n in unchecked.items():
            by_sheet[sheet] += n
        for sheet, n in sorted(by_sheet.items()):
            print(f"  {n} cells in {sheet!r} not checkable (LibreOffice gave an error or nothing, or an error is stored): check them in Excel on a scratch copy")
        for key, n in sorted(no_result.items()):
            print(f"  nothing stored, LibreOffice computes a value: {key} {n} cells")
        for key, n in sorted(diff.items()):
            print(f"  DIFFERENT: {key} {n} cells, for example {samples[key][:2]}")
        print(f"{sum(diff.values())} stored results differ")
        print(f"{verified} of {compared} formula cells verified, {sum(unchecked.values())} not checkable, "
              f"{sum(no_result.values())} with nothing stored")
        if a.keep:
            a.keep.mkdir(parents=True, exist_ok=True)
            shutil.copy2(redone, a.keep / redone.name)
        sys.exit(1 if diff else 0)
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    main()
