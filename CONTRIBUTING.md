# Contributing

Bug reports and fixes are welcome.

## Report a problem

Open an issue with:

- your system (macOS, Windows or Linux), the Python version, and the Excel or LibreOffice version if the problem involves them;
- the script and the exact command;
- the full output, including the exit code;
- what you expected.

**Do not attach a confidential workbook.** These tools exist for workbooks that matter, and those are rarely public. Describe the structure (a table loaded by a query, a shared formula, a note) or build a small file with made-up data that shows the problem. `tests/smoke_test.py` builds such files from nothing; its helpers are a good start.

## Send a change

1. Run `python tests/smoke_test.py`. It needs `openpyxl`; the LibreOffice check is skipped when LibreOffice is not installed.
2. Add a check to the smoke test for the defect you fix or the behavior you add.
3. Keep the scripts on Python 3.8 syntax, the standard library and `openpyxl`.
4. Keep examples neutral: names such as `Data`, `Sales`, `Book.xlsx`.

## What belongs here

Skills and scripts that let an agent read, change and verify existing Excel workbooks without damage. Creating new workbooks from nothing is out of scope; other tools do that well.
