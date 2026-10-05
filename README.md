# excel-skills

A Codex/Claude plugin with two skills for Excel workbooks that people depend on.

Both skills read `.xlsx` / `.xlsm` files directly as OOXML packages. Most of the work needs no Excel, and nothing is saved through a spreadsheet library, so tables, Power Query, notes and formatting stay intact.

## power-query-export

- **Navigate.** For every query: where it loads, which file and sheet it reads, which other queries it uses, which columns of the loaded table belong to Excel and not to the query, whether the table was saved with a filter, and when the query last refreshed. Given several workbooks, it finds which one feeds which and the order to refresh them.
- **Export.** Writes each Power Query definition as a `.pq` file and generates a `behavior.md` explaining how queries connect to workbook sheets, Excel tables, loaded columns, and formulas.
- **Refresh** (macOS with Excel). Refreshes all queries of a workbook from the command line, waits until the work ends, saves, and proves from the saved file that every loaded query refreshed.

## excel-write-verify

- **Write.** Surgical edits inside the file: values, formulas, table column formulas, new columns. Every edit states what it expects to find and stops if the workbook is different. All other parts stay byte-identical.
- **Verify.** A cell-by-cell diff of two versions, a structural check for the defects behind "Excel found a problem with some content", and an independent recalculation in LibreOffice.
- **Replace.** Puts the new version in place of a shared file in one step, only if nobody changed or opened it in the meantime, with a backup.
- **Field notes.** How to edit workbook XML by hand without damage, and rules for OneDrive, SharePoint and other synced folders.

## Use

Ask your agent:

```text
Where does the Sales table in /full/path/Report.xlsx get its data, and why is my new row missing?
```

```text
Refresh /full/path/Report.xlsx and show me what changed.
```

```text
In /full/path/Budget.xlsx change the formula of column Net in table Sales to subtract freight, and prove that nothing else changed.
```

## Install

Claude Code:

```bash
claude plugin marketplace add Mpayro/power-query-export
claude plugin install excel-skills@excel-skills
```

Codex:

```bash
codex plugin marketplace add Mpayro/power-query-export
codex plugin add excel-skills@excel-skills
```

Without the plugin system, copy the skill folders:

```bash
git clone https://github.com/Mpayro/power-query-export
cp -R power-query-export/plugins/excel-skills/skills/* ~/.claude/skills/    # or ~/.codex/skills/
```

Claude and ChatGPT in the browser accept a skill as a zip of its folder. There the skills run in a remote sandbox: they can read and edit a workbook that you upload, but they cannot refresh queries or reach your files.

## Use the scripts directly

```bash
cd plugins/excel-skills/skills

python3 power-query-export/scripts/pq_map.py Book.xlsx                  # map of every query
python3 power-query-export/scripts/pq_map.py A.xlsx B.xlsx C.xlsx       # plus links and refresh order
power-query-export/scripts/refresh_excel_mac.sh /full/path/Book.xlsx    # macOS only

python3 excel-write-verify/scripts/xlsxpatch.py apply base.xlsx new.xlsx spec.json
python3 excel-write-verify/scripts/colinsert.py base.xlsx new.xlsx columns.json
python3 excel-write-verify/scripts/xlsx_check.py base.xlsx new.xlsx     # structure
python3 excel-write-verify/scripts/xlsx_diff.py  base.xlsx new.xlsx     # what changed
python3 excel-write-verify/scripts/lo_recalc.py  new.xlsx               # are stored results right
python3 excel-write-verify/scripts/xlsx_replace.py new.xlsx /live/Book.xlsx --base-sha <sha256>
```

Run a script with no arguments (or with `-h`) to see its full usage. A usage error exits with code 2; code 1 is kept for a real finding (`xlsx_check.py` found a problem, `lo_recalc.py` found a different result).

## Requirements

- Python 3.8 or later. `pq_map.py`, `xlsx_check.py` and `xlsx_replace.py` use only the standard library; `xlsxpatch.py`, `colinsert.py`, `xlsx_diff.py` and `lo_recalc.py` also need `openpyxl`.
- `lo_recalc.py` needs LibreOffice.
- Refresh needs macOS with Microsoft Excel. Developed with Excel for Mac 16.113.

## What writes to your files

Read this before you let an agent run these.

- `refresh_excel_mac.sh` copies the workbook to `~/.pq-refresh-backups`, opens it in Excel, refreshes, recalculates, **saves** and closes. It stops if the workbook is open, or if Excel has other workbooks open. A refresh rewrites every query-loaded column: values that a person typed inside a query table do not survive.
- `xlsx_replace.py` replaces the live file only with `--go`, only if its hash is still the one you give, and keeps a copy in `~/.xlsx-backups`.
- `xlsxpatch.py` and `colinsert.py` write only the output file that you name, and refuse when it is the input file. No other script changes its input; `xlsx_replace.py` changes the live file as described above.

## Limits

- Refresh on Windows is described but was not tested.
- Sources that are not local files (web, databases) were not tested for refresh.
- Queries loaded to the Data Model or to PivotTables are reported as "nothing on a sheet".
- The refresh proof covers queries. It does not cover formulas that link to other workbooks.
- LibreOffice lacks some Excel functions (for example `LET` in some versions). `lo_recalc.py` reports those cells as not checkable.
- `xlsx_check.py` finds the structural defects it knows. A clean result is strong evidence, not a guarantee that Excel opens the file without a message. Open a scratch copy in Excel before an important release.
