# power-query-export

A Codex/Claude skill for Power Query in Excel workbooks.

It reads `.xlsx` / `.xlsm` files directly as OOXML packages and extracts the embedded `DataMashup`, so most of it works without Excel and without the Power Query UI:

- **Navigate.** For every query: where it loads, which file and sheet it reads, which other queries it uses, which columns of the loaded table belong to Excel and not to the query, whether the table was saved with a filter, and when the query last refreshed. Given several workbooks, it finds which one feeds which and the order to refresh them.
- **Export.** Writes each Power Query definition as a `.pq` file and generates a `behavior.md` explaining how queries connect to workbook sheets, Excel tables, loaded columns, and formulas.
- **Refresh** (macOS with Excel). Refreshes all queries of a workbook from the command line, waits until the work ends, saves, and proves from the saved file that every loaded query refreshed.
- **Field notes.** What a refresh overwrites, how to refresh a chain of linked workbooks, how to change the data a query reads without breaking the workbook, shared and synced folders, and a table of symptoms and causes.

## Use

Ask your agent:

```text
Use the power-query-export skill on /full/path/to/workbook.xlsx
```

```text
Where does the Sales table in /full/path/Report.xlsx get its data, and why is my new row missing?
```

```text
Refresh /full/path/Report.xlsx and show me what changed.
```

The workbook path is required.

## Use the scripts directly

```bash
python3 scripts/pq_map.py Book.xlsx                    # map of every query
python3 scripts/pq_map.py A.xlsx B.xlsx C.xlsx         # plus links between workbooks and refresh order
python3 scripts/pq_map.py Book.xlsx --query "Name"     # M code of one query
python3 scripts/pq_map.py Book.xlsx --export queries   # every query as a .pq file
python3 scripts/pq_map.py Book.xlsx --json             # the map as JSON
scripts/refresh_excel_mac.sh /full/path/Book.xlsx      # macOS only
```

`pq_map.py` never modifies a workbook.

## Install

Codex:

```bash
mkdir -p ~/.codex/skills
git clone https://github.com/Mpayro/power-query-export ~/.codex/skills/power-query-export
```

Claude:

```bash
mkdir -p ~/.claude/skills
git clone https://github.com/Mpayro/power-query-export ~/.claude/skills/power-query-export
```

## Requirements

- Navigate and export: Python 3.8 or later, standard library only. macOS, Windows, or Linux.
- Refresh: macOS with Microsoft Excel. Developed with Excel for Mac 16.113.

## What the refresh script does to your files

Read this before you let an agent run it.

- It copies the workbook to `~/.pq-refresh-backups` first (change it with `--backup-dir`).
- It opens the workbook in Excel, refreshes, recalculates, **saves** and closes. It never writes cells.
- It stops if the workbook is open, or if Excel has other workbooks open, because an Excel crash would close them without saving.
- A refresh rewrites every query-loaded column. Values that a person typed inside a query table do not survive. See `references/navigate-and-refresh.md`, section 4.

## Outputs of an export

```text
queries/
  01_QueryName.pq
behavior.md
export-summary.md
```

## Notes

- Works without opening the Power Query UI.
- Detects macOS, Windows, or Linux before choosing commands.
- Focuses on workbook behavior, not just raw query export.

## Limits

- Refresh on Windows is described but was not tested.
- Sources that are not local files (web, databases) were not tested for refresh.
- Queries loaded to the Data Model or to PivotTables are reported as "nothing on a sheet".
- The refresh proof covers queries. It does not cover formulas that link to other workbooks.
