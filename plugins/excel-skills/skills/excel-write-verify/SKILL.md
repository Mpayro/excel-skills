---
name: excel-write-verify
description: Use when you must change an existing Excel workbook (.xlsx or .xlsm) without breaking it, or prove that a change to a workbook is right. Covers surgical edits inside the file (values, formulas, table column formulas) that keep Power Query, tables, notes and formatting intact, a cell-by-cell diff of two versions, a structural check for the "Excel found a problem with some content" kind of damage, an independent recalculation, and a safe replace of a workbook that other people use in OneDrive, SharePoint or a shared drive. Use it whenever the task is to edit, patch, fix, update or verify a workbook that already exists ("editar", "modificar", "verificar un Excel"), even if the user only says "change this cell". Not needed to create a new workbook from nothing.
---

# Excel Write And Verify

A workbook that people depend on holds much more than cells: tables, queries, notes, styles, links, stored results. Most tools that "save an Excel file" write a new package and keep only what they understand. This skill changes the file from the inside, proves what changed, and puts it back without stepping on anybody.

The script paths below are relative to the folder of this skill. The scripts need Python 3.8 or later. `xlsxpatch.py`, `colinsert.py`, `xlsx_diff.py` and `lo_recalc.py` also need `openpyxl`, which they use only to read and to parse references. On Windows, use `python` or `py` in place of `python3`.

## First Decide Which Case You Have

| Case | What to do |
|---|---|
| A person has the workbook open (lock file `~$Name.xlsx`, Excel process, or a collaborator in the browser) | Do not replace the file under them. Ask them to close it, or make the change through Excel itself. A missing lock file does not prove that nobody edits it in the browser. |
| Existing workbook, closed, with structure that must survive | The method in this file. |
| Existing workbook, closed, plain data only and nobody depends on its structure | A library is acceptable for plain data only, and then run the checks of step 4. |
| New workbook | Use any tool. Still run `xlsx_check.py` on the result. |

In any workbook, do not turn a range into an Excel table and do not apply table formatting unless the user asked for it. It changes how formulas and queries read the range.

## Why Not Save With A Library

In a test, a plain load and save with openpyxl gave a copy with 26 fewer parts than the 89 of a real workbook. The Power Query definitions, the connections, the query tables and the threaded comments were gone; the legacy notes survived. The file still opened and looked normal. pandas `to_excel` and similar writers behave the same way, because they write a new package. Reading with a library is safe. Writing is the problem.

Writing cells through Excel automation is not the answer either. On macOS, cell writes through AppleScript crashed Excel on a 9 MB workbook and closed every workbook the user had open. Use Excel only to open, calculate, save and close.

## The Method

### 1. Copy and fingerprint

Copy the live workbook to a local work folder and record its SHA-256. That hash is the base: it names the exact version your change is built on. In a synced folder, read the file in blocks; see `references/shared-folders.md`.

Keep the copy from before the change in a folder that survives a restart. A backup kept in `/tmp` was lost one day later.

### 2. Patch inside the package

```bash
python3 scripts/xlsxpatch.py apply base.xlsx new.xlsx spec.json
```

The spec lists the operations; run the script with no arguments to see all of them. Example:

```json
{
  "set_value":   [{"sheet": "Data", "cell": "C9", "value": 40}],
  "set_formula": [{"sheet": "Data", "cell": "D9", "formula": "=C9*2"}],
  "rewrite":     [{"sheet": "Data", "cells": "N2:N147", "old": "A{r}*2", "new": "A{r}*3"}],
  "table_rewrite": [{"table": "Sales", "column": "Net", "old": "...", "new": "..."}]
}
```

These operations compare first and stop if the workbook is different: `rewrite`, `formula_replace`, `table_rewrite`, `table_formula_replace`, `table_add_default`. That is the point: a patch that applies "somewhere near" is how workbooks get damaged. `set_value` refuses a cell that holds a formula; `set_formula` replaces whatever the cell holds (it refuses the master of a formula group of several cells). A leading "=" in `old`, `new` and `formula` is removed. The script refuses when SRC and DST are the same file. All other parts of the file stay byte-identical. By default the patch tells Excel to recalculate at the next open and removes the stale calculation chain.

To insert columns, also inside an Excel table or a table that Power Query loads:

```bash
python3 scripts/colinsert.py base.xlsx new.xlsx columns.json
```

It moves everything that sits to the right (cells, formula references, merged cells, hyperlinks, formats, validations, notes, the table with its query fields and saved filter columns, defined names). It stops, and writes nothing, if another sheet, a conditional format or validation of another sheet, another table, a chart or a pivot cache points at the columns that would move, or if the sheet has a saved autofilter. Pictures and charts of the sheet keep their position. The formula of a new column is written unchanged into every row, so use structured references. Run it with no arguments for the spec.

For changes these two tools do not cover (new rows, new notes, new parts), edit the XML by hand with the patterns and traps in `references/xml-patching.md`.

### 3. Give new formulas a stored result

A formula written by a script has no result in the file until Excel calculates and saves it. A formula that `rewrite` or `formula_replace` changed also loses its old result, because that result is no longer true. Excel itself recalculates at the next open, so a person sees correct numbers. Every other reader sees an empty cell: Power Query, pandas, a link from another workbook, a preview. `xlsx_check.py` counts these cells.

Choose one:

- Compute the results yourself (mirror the formulas in Python) and store them: `python3 scripts/xlsxpatch.py cached new.xlsx new2.xlsx values.json`, where the JSON is `{"Sheet": {"D9": 80}}`.
- Let Excel do it on a scratch copy: open, calculate, save, close. On macOS:

```bash
open -a "Microsoft Excel" /abs/path/Book.xlsx
osascript -e 'tell application "Microsoft Excel" to calculate'
osascript -e 'tell application "Microsoft Excel" to save workbook "Book.xlsx"'
osascript -e 'tell application "Microsoft Excel" to close workbook "Book.xlsx" saving no'
```

  First make sure that the user has no other workbook open in Excel, because a crash closes all of them. A save by Excel rewrites the whole package, so run the checks of step 4 again on the saved file.

### 4. Verify, with three different questions

```bash
python3 scripts/xlsx_check.py base.xlsx new.xlsx     # is the package still sound?
python3 scripts/xlsx_diff.py  base.xlsx new.xlsx     # did only the intended cells change?
python3 scripts/lo_recalc.py  new.xlsx               # are the stored results right?
```

| Check | It proves | It does not prove |
|---|---|---|
| `xlsx_check.py` | None of the defects it knows: content types, relationships, row and cell order, cell types, style indexes, formula text, shared formulas, calculation chain, table headers and names, merged ranges, duplicate names, notes and their shapes | That the numbers are right, or that Excel opens the file without a message |
| `xlsx_diff.py` | Exactly which cells changed what they hold, and which stored results moved | That the new formulas are the right ones |
| `lo_recalc.py` | A second engine (LibreOffice) reproduces the stored results | Anything about functions LibreOffice lacks. It reports those cells as not checkable, and its last line says how many cells it verified. Check the rest in Excel on a scratch copy. Volatile functions (TODAY, NOW, RAND) differ by nature. |

Read the diff as a list of claims: every line must be a change you intended. A line you cannot explain is a defect or an edit by another person that you are about to erase.

For a workbook that feeds decisions or other systems, add a reviewer that is not the author. Give it the base, the new file, the spec and the output of the three checks, and ask it to find what is wrong.

### 5. Put it back

```bash
python3 scripts/xlsx_replace.py new.xlsx /path/live/Book.xlsx --base-sha <sha of step 1>          # dry run
python3 scripts/xlsx_replace.py new.xlsx /path/live/Book.xlsx --base-sha <sha of step 1> --go
```

The script refuses if the live file is no longer the base, if a lock file exists, or if a local process has it open. With `--go` it makes a backup, writes a temporary file in the same folder, renames it over the live file in one step, checks the content, waits for the sync client, checks again, and lists possible conflict copies.

If the live file changed, a person saved in the meantime. Do not force it. Start again from their version: copy it, apply the same spec, verify, deploy. Their edits survive and yours go on top.

### 6. After

- Verify by reading the content back. A file date proves nothing in a synced folder.
- If other workbooks read this one through Power Query, they hold old data until they refresh. Use the `power-query-export` skill to find them and to refresh in the right order.
- Tell the user what changed, what was verified and how, and what was not verified.

## Rules That Came From Incidents

- State the base hash again immediately before the replace. Time passes between the copy and the deploy, and people save.
- One writer per workbook. Two processes that write the same file do not merge; one of them loses.
- A fixed temporary file name is shared by overlapping runs and gets truncated. Use a name that is unique to the process, in the same folder as the target.
- "Nothing written" must say why: no change needed, or could not read. Otherwise nobody can tell a healthy run from a silent failure.
- A value typed over a formula in a table column, or inside a query-loaded column, disappears at the next query refresh. Report such cells when the diff shows them.
- After Excel's "repaired records" message, do not keep the repaired file. Excel has already removed what it could not read. Go back to the base and find the defect with `xlsx_check.py`.

## Bundled Scripts

| Script | Use | Writes? |
|---|---|---|
| `scripts/xlsxpatch.py` | Apply a spec of edits; store results; raw diff | Only the output file you name |
| `scripts/colinsert.py` | Insert columns and shift everything that depends on their position | Only the output file you name |
| `scripts/xlsx_check.py` | Structural check of one workbook, optionally against a base | No |
| `scripts/xlsx_diff.py` | Cell-by-cell diff of formulas, constants and stored results | No |
| `scripts/lo_recalc.py` | Recalculate in LibreOffice and compare stored results | No (works on a temporary copy) |
| `scripts/xlsx_replace.py` | Guarded, one-step replace of the live file, with backup | The live file, only with `--go` |

## References

- `references/xml-patching.md`: read it before you edit sheet XML by hand. It has the anatomy of the package, how to add rows, what else must be extended, and the traps that damage files.
- `references/shared-folders.md`: read it when the workbook lives in OneDrive, SharePoint, Google Drive or a network share.
