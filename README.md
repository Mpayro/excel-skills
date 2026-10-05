# Excel Skills for Claude Code and Codex

Let an AI agent edit, verify and refresh real Excel workbooks without breaking them.

`excel-skills` is a plugin with two agent skills for `.xlsx` and `.xlsm` files that people depend on: workbooks with Power Query, tables, notes, formatting and other people editing them in OneDrive or SharePoint.

## Why

- **Agents break Excel files.** Most tools that save an `.xlsx` write a new file and keep only what they understand. In a test, a plain load and save with openpyxl gave a copy with 26 fewer parts: the Power Query definitions, the connections, the query tables and the threaded comments were gone. The file still opened and looked normal.
- **You cannot see what an edit really changed.** One wrong formula in a shared workbook is expensive. These skills prove every change cell by cell.
- **Power Query hides where data comes from.** Which file feeds this table? Why is my new row missing? When did it last refresh? The answers are inside the file, and no Excel is needed to read them.

## Install

Claude Code:

```bash
claude plugin marketplace add Mpayro/excel-skills
claude plugin install excel-skills@excel-skills
```

OpenAI Codex:

```bash
codex plugin marketplace add Mpayro/excel-skills
codex plugin add excel-skills@excel-skills
```

Any other agent that reads Agent Skills (`SKILL.md`): copy the skill folders.

```bash
git clone https://github.com/Mpayro/excel-skills
cp -R excel-skills/plugins/excel-skills/skills/* ~/.claude/skills/    # or ~/.codex/skills/
```

## What you can ask

```text
Where does the Sales table in /path/Report.xlsx get its data, and why is my new row missing?
```

```text
Refresh the Power Query of /path/Report.xlsx and show me what changed.
```

```text
In /path/Budget.xlsx change the formula of column Net in table Sales to subtract freight,
and prove that nothing else changed.
```

```text
Excel says "We found a problem with some content" in this file. Find what is damaged.
```

## The two skills

### excel-power-query

- **Navigate.** For every query: where it loads, which file and sheet it reads, which other queries it uses, which columns of the loaded table belong to Excel and not to the query, whether the table was saved with a filter, and when the query last refreshed.
- **Trace a chain.** Given several workbooks, it finds which one feeds which and the order to refresh them.
- **Export.** Writes each Power Query (M) definition as a `.pq` file and a `behavior.md` that explains how queries connect to sheets, tables, loaded columns and formulas.
- **Refresh from the command line** (macOS with Excel). Refreshes all queries, waits until the work ends, saves, and proves from the saved file that every loaded query refreshed.

### excel-write-verify

- **Edit inside the file.** Values, formulas, table column formulas, new columns. Each edit states what it expects to find and stops if the workbook is different. Every other part of the file stays byte-identical, so Power Query, tables, notes and formatting survive.
- **Verify.** A cell-by-cell diff of two versions, a structural check for the defects behind "Excel found a problem with some content", and an independent recalculation in LibreOffice.
- **Replace a shared file safely.** Puts the new version in place in one step, only if nobody changed or opened the workbook in the meantime, with a backup.
- **Field notes.** How to edit workbook XML by hand without damage, and rules for OneDrive, SharePoint and other synced folders.

## Compared with saving through a library

| | Save with openpyxl or pandas | excel-skills |
|---|---|---|
| Power Query, connections, threaded comments | Dropped | Kept, byte-identical |
| Proof of what changed | None | Cell-by-cell diff of formulas and stored results |
| Damage check | You find out when Excel opens it | Structural check before anyone opens it |
| Are the numbers right | Unknown | Second engine recalculates and compares |
| File shared in OneDrive or SharePoint | Overwrites whatever is there | Refuses if someone saved or has it open |

## FAQ

**Does it need Microsoft Excel?**
No, for almost everything: navigation, export, edits, diff and structural check read and write the file directly. Excel for Mac is needed only to refresh Power Query. LibreOffice is needed only for the independent recalculation.

**Does it work on Windows and Linux?**
The Python scripts use portable code and the standard library plus `openpyxl`. They were developed and tested on macOS. The refresh script is macOS only.

**Why not just use openpyxl?**
openpyxl is fine to read a workbook and to create a new one. Saving an existing workbook with it rewrites the whole package and drops what it does not model. These skills change only the XML that the edit needs.

**How do I refresh Power Query from the command line on a Mac?**
`plugins/excel-skills/skills/excel-power-query/scripts/refresh_excel_mac.sh /full/path/Book.xlsx`. It opens Excel, refreshes, waits, saves, closes, and checks the fill time of every query in the saved file.

**Excel says "We found a problem with some content" after an edit. What now?**
Do not keep the repaired file: Excel already removed what it could not read. Go back to the previous version and run `xlsx_check.py` on the damaged one. It names the part and the cell for the usual causes.

**Can an agent edit a workbook that other people use in OneDrive or SharePoint?**
Yes, with `xlsx_replace.py`. It replaces the file only if it is still the version the change was built on, nobody has it open, and a backup was made.

## Use the scripts directly

```bash
cd plugins/excel-skills/skills

python3 excel-power-query/scripts/pq_map.py Book.xlsx                  # map of every query
python3 excel-power-query/scripts/pq_map.py A.xlsx B.xlsx C.xlsx       # plus links and refresh order
excel-power-query/scripts/refresh_excel_mac.sh /full/path/Book.xlsx    # macOS only

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
- `xlsxpatch.py` and `colinsert.py` write only the output file that you name, and refuse to write over their input. No script changes its input.

## Limits

- Refresh on Windows is described but was not tested.
- Sources that are not local files (web, databases) were not tested for refresh.
- Queries loaded to the Data Model or to PivotTables are reported as "nothing on a sheet".
- The refresh proof covers queries. It does not cover formulas that link to other workbooks.
- LibreOffice lacks some Excel functions. `lo_recalc.py` reports those cells as not checkable.
- `xlsx_check.py` finds the structural defects it knows. A clean result is strong evidence, not a guarantee that Excel opens the file without a message. Open a scratch copy in Excel before an important release.
- In a chat product that runs skills in a remote sandbox, the skills can read and edit a workbook that you upload, but they cannot refresh queries or reach files on your computer.

## En español

Skills de Excel para Claude Code y Codex. Permiten que un agente de IA edite libros de Excel reales sin romper Power Query, tablas ni fórmulas, que compruebe cada cambio celda por celda y que navegue o refresque Power Query desde la línea de comandos. La documentación de cada skill está en inglés.
