# Navigate, refresh and feed Power Query workbooks

Field notes from real runs on a chain of three linked workbooks in a shared folder (Excel for Mac, October 2026). Each rule says what happened, so that you can judge when it applies.

Contents:

1. How Power Query reads a workbook
2. Find out why data does not arrive
3. Refresh from the command line
4. What a refresh overwrites
5. Refresh a chain of workbooks
6. Changing the data a query reads
7. Shared and synced folders
8. Symptoms and causes
9. Not tested

## 1. How Power Query reads a workbook

A query that reads another workbook does not open it in Excel. It reads the file on disk and takes the value stored in each cell. Three consequences follow:

- **Unsaved work does not exist.** If the source is open with changes, the query sees the last saved version.
- **Formulas are not calculated.** The query takes the result that Excel stored at the last save. A formula written into the file by a script, with no stored result, arrives as empty or as an old value. Excel must open the source, calculate and save it before a downstream query reads it.
- **The file date proves nothing.** A sync client or a save without a refresh changes the date. Use the `last fill` data of `pq_map.py`.

What the query takes from the source depends on the navigation step:

| M code | It reads | New rows below the data |
|---|---|---|
| `{[Item="Sheet1", Kind="Sheet"]}[Data]` | The used range of the sheet | Arrive after the source is saved |
| `{[Item="Table1", Kind="Table"]}[Data]` | The range of the Excel table | Arrive only if the table range includes them |
| `{[Item="Name", Kind="DefinedName"]}[Data]` | The range of the defined name | Arrive only if the name includes them |
| `Excel.CurrentWorkbook(){[Name="X"]}[Content]` | A table or name of the same workbook | Same rule as table or name |

File paths in `File.Contents("...")` are usually absolute. They work only on the machine and user account that wrote them. Report this when a workbook is shared between people.

## 2. Find out why data does not arrive

Use `pq_map.py` on every workbook of the chain in one command. Then go upstream from the table where the row should be:

1. **Is the row there, but hidden?** If the map shows `TABLE FILTER saved on` for the table, the row can be loaded and not visible. Check the cell values in the file before you look for a query problem. In one test workbook, a table was saved with a filter on one column and showed 14 of its 66 rows.
2. **Is the last fill recent?** If `last fill` is older than the change in the source, nobody refreshed. Stop here.
3. **Does a filter drop the row?** Read the M code of each query on the path (`--query NAME`). Look for `Table.SelectRows`. The usual causes are a list of allowed categories, a "not null" filter on a key column, and a hard-coded list of excluded keys. A row that fails one of them is dropped without any message.
4. **Does a header step miss?** `Table.Skip(..., n)` plus `Table.PromoteHeaders` depends on the header row position. `Table.SelectColumns` and `Table.RenameColumns` fail if a header text changed.
5. **Does a join drop it?** An inner join (`JoinKind.Inner`) with another query drops rows that have no match there. Check the other query too.
6. **Is the source value itself stored?** Open the source with a read-only library and look at the cached value of the cell. If it is empty or old, Excel has not calculated and saved that file (section 1).
7. **Is the range wide enough?** For `Kind="Table"` and `Kind="DefinedName"`, compare the range with the last data row.
8. **Does the row arrive, but the columns next to it are wrong?** The `excel-side columns` are formulas. They often look up the key in a side table that a person maintains. A new key that is missing there returns the default of the formula, not an error. Add the key to the side table.

`cells with errors` above zero in `last fill` means that some source cells held an Excel error such as `#DIV/0!`. The rows still load. Find the cell in the source and report it.

## 3. Refresh from the command line

Only Excel can run a refresh. LibreOffice and the Python libraries cannot. If there is no Excel on the machine, say so and stop at navigation.

On macOS:

```bash
scripts/refresh_excel_mac.sh "/abs/path/Book.xlsx"            # add --timeout SECONDS for slow workbooks (default 900)
```

The script copies the workbook to a backup folder, opens it, sends `refresh all`, waits until the refresh processes go quiet, recalculates, saves, closes, and then reads the saved file to prove that each loaded query has a fill time later than the start. Measured times: 40 seconds for a 13-query workbook, 4 minutes for a 9 MB workbook with 16 queries.

Before you run it:

- **Know where the copy from before the refresh is.** The script writes one to `~/.pq-refresh-backups` (change it with `--backup-dir`) and prints the path. For work that matters, give a folder that belongs to the task. Do not use `/tmp`: it is erased on restart, and a backup kept there was lost one day later.
- **Tell the user that Excel will open.** The script stops if Excel has other workbooks open, because an Excel crash closes all of them without saving. Ask the user to save and close them. Use `--allow-open-books` only with their agreement.
- **Try it on a local copy first** when the workbook is important. A refresh of a copy reads the same sources and writes only the copy. Compare the copy before and after (section 4) and show the user what would change.
- The script stops if a lock file `~$Book.xlsx` exists. That file means a person has the workbook open.

Why the script does what it does:

- `open -a "Microsoft Excel" file` opens through the system. The AppleScript `open workbook` command raises a file-access dialog that blocks the run.
- `refresh all` returns at once and the refresh continues in `Microsoft.Mashup.Container` processes. A save sent too early stores old data. The `refreshing` property of a query table gives error -50 for Power Query tables, so the script watches the CPU use of those processes and of Excel.
- `display alerts` and `ask to update links` are switched off for the run so that no dialog blocks it. `ask to update links` is a lasting user preference. The script puts the earlier value back. Do the same if you drive Excel by hand.
- `name of every workbook` returns `missing value` when no workbook is open. Use `count of workbooks` to test for none.
- Do not write cells through AppleScript (`set value of range`, `copy range`). On a 9 MB workbook this crashed Excel with error -609 and closed the user's other workbooks. Put data in with the method of section 6 and use Excel only to refresh, calculate and save.

After the run, exit code 0 with "all N loaded queries refreshed" is the proof for the queries. Exit 8 means that the file was saved but a query kept old data: decide with the user whether to run again or to put the backup back. The quiet-CPU wait is only a sign that the work ended; the fill times are the proof. Then also check the business content: the rows and keys that you expected are in the loaded tables, and a few computed values are right.

## 4. What a refresh overwrites

Observed, and repeated on a copy:

- **Query columns are rewritten in full.** A value that a person typed over a query column is lost. In one run, three test values that a colleague had typed were lost this way.
- **A constant typed into a calculated column of a query table goes back to the column formula.** Two more test values went back to formulas on each refresh.
- **Excel-side columns and the rest of the workbook stay.** In the test, a full refresh changed only those two cells in the whole workbook.
- The table grows or shrinks with the query result, and the calculated columns fill down to new rows.

So, for a workbook that people edit:

1. Compare the workbook before and after the refresh, cell by cell, in formula mode. Read with a library in read-only mode. Do not save with it.
2. Expected changes are the query columns of rows whose source changed. Any other change is a manual edit that the refresh removed, most often a constant that became a formula.
3. Report each one. Restore it only if the user wants that, and tell the owner that such edits do not survive a refresh. Durable alternatives are a change in the source, or a separate input column that the query does not own.

## 5. Refresh a chain of workbooks

When workbook C reads B and B reads A, `pq_map.py A B C` prints the order. Work upstream first:

1. Change A. Excel must calculate and save A.
2. Refresh B, recalculate, save. B's own formulas often depend on the tables that its queries load, and C reads the results of those formulas.
3. Refresh C, save.

If you change data in the middle workbook after its refresh, save it again before you refresh the next one. Verify each step before you start the next. An error found at the end of the chain costs a repeat of all of it.

## 6. Changing the data a query reads

If the `excel-write-verify` skill is installed (it ships in the same plugin), use its tools for this work: `xlsxpatch.py` and `colinsert.py` to edit, `xlsx_check.py` and `xlsx_diff.py` to verify, `xlsx_replace.py` to put the file back. The rules below are the short version.

Do not save a workbook that has Power Query, tables or comments with openpyxl or a similar library. These libraries write a new package and drop the parts that they do not model. In a test, a plain load and save with openpyxl gave a copy with 26 fewer parts than the 89 of the workbook: the query definitions, the connections, the query tables and the threaded comments were gone. The file still opened, with no queries left. Reading in read-only mode is safe.

To add or change rows while the workbook is closed, edit the XML inside the package and leave every other part as it is:

- Copy the row that you use as a template, with its style ids and formulas, and change the row numbers. Write text as inline strings (`t="inlineStr"`) so that the shared strings table stays as it is.
- Extend what describes the range: the sheet `dimension`, the table `ref`, the `autoFilter`, and any defined name that covers the data.
- Make Excel calculate at the next open: set `fullCalcOnLoad="1"` in the workbook `calcPr` and remove `calcChain.xml` together with its relationship and its content type entry.
- For a cell that belongs to a shared formula, change only the cached value. If you rewrite the formula element, the group breaks.
- When you match cells with a regular expression, make the attribute part lazy (`<c r="K49"[^>]*?>`). A greedy pattern runs past a self-closing cell and joins it with the next one.
- Then open the workbook in Excel and save it (the refresh script does this), so that the new formulas have stored results before anything downstream reads them.
- Compare the result with the original, cell by cell. The only differences must be the ones that you intended.

If the workbook is open in Excel by a person, do not replace the file. Work through that person, or wait.

## 7. Shared and synced folders

- Copy the workbook to a local folder and work on the copy. Read it in blocks. A direct read of an online-only file can time out and give a damaged ZIP.
- Record the SHA-256 of the live file when you copy it. Before you put a new version back, compute it again. If it changed, a colleague saved in the meantime. Start again from their version. Do not write over it. This happened during the run: a colleague saved the workbook at night between the copy and the deployment.
- Replace the live file in one step: write a temporary file in the same folder, then rename it over the original.
- After the sync client had time to act, check the hash again and look for conflict copies in the folder.
- A missing lock file does not prove that nobody edits the workbook in the browser.

## 8. Symptoms and causes

| Symptom | Cause | Action |
|---|---|---|
| "No DataMashup found" | Byte search on a UTF-16 part | Decode as UTF-16, or use `pq_map.py` |
| Downstream table has old values after a refresh | The source was not saved by Excel after the change, or the refresh order was wrong | Sections 1 and 5 |
| New row is missing downstream | Row filter, inner join, empty key, or a range that is too short | Section 2 |
| New row arrives, computed columns show defaults | A hand-kept side table has no entry for the new key | Section 2, step 8 |
| Row is in the file but the user does not see it | The table was saved with a filter | Section 2, step 1 |
| Test values of a person disappeared | They were typed into a query table | Section 4 |
| Saved file has old data although `refresh all` ran | The save came before the background refresh ended | Use the script; it waits and then proves the fill time |
| Error -50 on `refreshing` | Not supported for Power Query tables | Watch the mashup processes |
| A file dialog blocks the automation | AppleScript `open workbook` | `open -a "Microsoft Excel" file` |
| Excel closed with error -609 | Cell writes through AppleScript on a large workbook | Section 6 |
| Excel no longer asks to update links | An earlier automation left `ask to update links` off | Set it back to true |
| ZIP or CRC error when reading | Online-only file in a synced folder, or a read during a sync | Copy to local in blocks, read the copy |

## 9. Not tested

State these limits when they matter. Do not present them as verified.

- **Windows.** The usual route is COM from PowerShell: `$wb.RefreshAll()`, `$xl.CalculateUntilAsyncQueriesDone()`, `$wb.Save()`. It was not run here. The fill-time proof with `pq_map.py` works on any system.
- **Sources that are not local files** (web, databases, SharePoint URLs). They can ask for credentials or privacy levels, and they can wait on the network with no CPU use, so the quiet-CPU wait can end early. The fill-time proof still shows a query that did not refresh.
- **Queries loaded to the Data Model or to PivotTables.** The map reports them as "nothing on a sheet", and the refresh script refuses a workbook where no query loads to a sheet.
- **Formulas that link to other workbooks.** The script switches the "update links" question off so that no dialog blocks the run. Whether Excel then updates those links by itself at open was not checked. The proof covers queries only, and the script prints a note when the workbook has such links.
- **Very slow recalculation.** The script calls `calculate` and then saves. On the workbooks tested (about 11,000 formulas) the saved results were right. On a slower workbook, check computed values after the run.
