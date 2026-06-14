---
name: power-query-export
description: Use when exporting Power Query or Power BI Mashup queries from Excel .xlsx or .xlsm files, or when analyzing how workbook queries feed sheets, Excel tables, formulas, and displayed workbook behavior.
---

# Power Query Export

## Purpose

Export every Power Query M definition from an Excel `.xlsx` or `.xlsm` workbook without using the Power Query UI, then generate a `behavior.md` report that explains how queries relate to workbook sheets, Excel tables, query-loaded columns, calculated columns, and formulas.

This is a universal workflow. Do not assume macOS or Windows. Detect the user's system first, then use OS-appropriate commands.

## Required User Input

The user must provide the full path to the Excel workbook.

If the path is missing, ask:

```text
Please provide the full path to the Excel file you want me to export Power Query queries from.
```

Examples:

```text
/Users/name/Documents/Workbook.xlsx
C:\Users\name\Documents\Workbook.xlsx
```

Do not guess the workbook path unless the user explicitly asks you to detect the active Excel workbook.

## Detect The System

Prefer Python:

```bash
python3 - <<'PY'
import platform
print(platform.system())
PY
```

Interpretation:

```text
Darwin  -> macOS
Windows -> Windows
Linux   -> Linux
```

If Python is unavailable, use:

macOS/Linux:

```bash
uname -s
```

Windows PowerShell:

```powershell
[System.Runtime.InteropServices.RuntimeInformation]::OSDescription
```

## Optional Active Workbook Detection

Use this only if the user asks to use the currently open workbook and has not provided a file path.

macOS:

```bash
osascript -e 'tell application "Microsoft Excel" to get {name, full name} of active workbook'
```

Windows PowerShell:

```powershell
$xl = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$xl.ActiveWorkbook.FullName
```

Prefer the explicit path provided by the user.

## Output Folder

Create an output folder next to the workbook unless the user specifies another location.

Default folder name:

```text
<workbook-stem> power query export
```

Recommended layout:

```text
<workbook-stem> power query export/
  queries/
    01_QueryName.pq
    02_AnotherQuery.pq
  behavior.md
  export-summary.md
```

## Fast Export Method

Do not automate Excel's Power Query UI. Do not refresh the workbook. Treat `.xlsx` and `.xlsm` files as OOXML ZIP packages.

Core steps:

1. Validate the input path exists and is `.xlsx` or `.xlsm`.
2. Open the workbook with ZIP parsing.
3. Search `customXml/item*.xml`.
4. Find XML containing a `DataMashup` payload.
5. Decode XML as `utf-16`; fallback to `utf-8`.
6. Extract the base64 text inside the `DataMashup` element.
7. Base64-decode it into binary.
8. Scan the binary for local ZIP headers: `PK\x03\x04`.
9. Parse local ZIP headers manually because DataMashup packages may lack a normal central directory.
10. Decompress entries:
    - method `8`: raw deflate with `zlib.decompress(payload, -15)`
    - method `0`: stored bytes
11. Find `Formulas/Section*.m`, usually `Formulas/Section1.m`.
12. Decode M files as `utf-8-sig`.
13. Split query definitions by top-level `shared` declarations.

Use this regex:

```regex
(?m)^shared\s+((?:#"(?:[^"]|"")*"|[^=\s]+))\s*=
```

Decode M query names:

```text
#"Inbound Total" -> Inbound Total
#"Query ""A"""  -> Query "A"
```

Write every query as a complete `.pq` file:

```powerquery
section Section1;

shared QueryName = ...
```

Sanitize filenames for Windows and macOS:

```text
\ / : * ? " < > | -> _
```

Normalize whitespace and prefix files with sequence numbers to preserve workbook order.

## Verification

After export:

1. Count `shared` definitions in `Formulas/Section*.m`.
2. Count `.pq` files written.
3. Parse `xl/connections.xml`.
4. Count connections named `Query - ...`.
5. Compare counts.
6. Confirm each `.pq` file:
   - starts with `section`
   - contains exactly one `shared` definition
   - ends with a semicolon

Record results in `export-summary.md`. If counts differ, state the mismatch clearly.

## behavior.md Requirement

Always generate `behavior.md` in the output folder.

`behavior.md` explains workbook behavior, not just M code. It must analyze every sheet and every Excel table in the workbook.

Inspect these OOXML parts:

```text
xl/workbook.xml
xl/_rels/workbook.xml.rels
xl/connections.xml
xl/worksheets/sheet*.xml
xl/worksheets/_rels/sheet*.xml.rels
xl/tables/table*.xml
xl/tables/_rels/table*.xml.rels
xl/queryTables/queryTable*.xml
```

For richer formula/cached-value inspection, `openpyxl` may be used when available. The core export must not depend on it.

## behavior.md Structure

Use this structure:

```markdown
# Workbook Behavior Analysis

## Workbook Summary

- Workbook path:
- Export date:
- Query count:
- Sheet count:
- Table count:
- Query-loaded table count:
- Formula table column count:

## Sheets

| Sheet | XML Part | State | Tables | Notes |
|---|---|---|---|---|

## Excel Tables

| Table | Sheet | Range | Query Connection | Columns | Calculated Columns |
|---|---|---|---|---|---|

## Power Query Connections

| Connection ID | Connection Name | Query Name | Loaded To |
|---|---|---|---|

## Query Files

| Query | File | Loaded To Sheet/Table | Direct Dependencies |
|---|---|---|---|

## Query Dependencies

Explain which queries reference other queries.

## Loaded Columns Vs Excel Formulas

Classify every table column as:

- Query-bound: loaded from Power Query
- Excel-calculated: table formula in workbook XML
- Manual/unbound: not query-bound and no calculated formula found

## Calculated Table Columns

| Table | Column | Formula | Depends On |
|---|---|---|---|

## Behavior Chains

Explain important chains, for example:

```text
Power Query loads raw inbound rows into table Inbound.
Excel formula column EOM derives month buckets.
Another sheet uses SUMIFS over the loaded table.
The displayed metric is therefore not purely a Power Query output.
```

## Risks / Quirks

Call out:

- hard-coded file paths
- absolute local paths
- duplicated import logic
- query-loaded values extended by Excel formulas
- negative availability calculations
- hidden sheets
- connection-only queries
- formulas referencing external workbooks
```

## Dependency Analysis Rules

For each exported `.pq` query, extract:

- `File.Contents(...)`
- `Excel.Workbook(...)`
- `Csv.Document(...)`
- `Web.Contents(...)`
- `Sql.Database(...)`
- direct references to other query names
- final selected or renamed columns where inferable

For each Excel table, extract:

- table name
- display name
- sheet relationship
- range
- query table relationship
- query connection id
- table columns
- calculated column formulas

For each query table, extract:

- `connectionId`
- bound fields
- unbound fields
- fields with `dataBound="0"`

Unbound fields often indicate Excel-added calculated or manual columns.

## Optimized Implementation Guidance

Use Python standard library first:

```text
zipfile
xml.etree.ElementTree
base64
struct
zlib
re
pathlib
datetime
```

Avoid reading entire worksheet data unless needed for `behavior.md`.

Prefer XML metadata over spreadsheet libraries for speed.

Use `openpyxl` only when useful for cached values or formula inspection.

## Pseudocode

```python
system = detect_os()
input_path = require_user_excel_path()
output_dir = create_output_folder(input_path)

with zipfile.ZipFile(input_path) as z:
    mashup_parts = find_custom_xml_datamashup_parts(z)
    mashup_binary = decode_datamashup(mashup_parts)
    package_entries = parse_local_zip_records(mashup_binary)
    formula_sections = extract_formulas_sections(package_entries)

queries = split_shared_queries(formula_sections)

for index, query in enumerate(queries, start=1):
    write_query_file(output_dir / "queries", index, query)

workbook_model = inspect_workbook_ooxml(input_path)
behavior = analyze_behavior(workbook_model, queries)

write_behavior_md(output_dir, behavior)
write_export_summary(output_dir, verification_results)
```

## Completion Criteria

Complete only when:

- every query is exported as a `.pq` file
- `behavior.md` exists
- `export-summary.md` exists
- query count verification was performed
- workbook path and output folder are reported
- mismatches or limitations are explicitly stated

## Final Response Format

Use this final response shape:

```text
Exported <N> Power Query definitions.

Output folder:
<path>

Main files:
- queries/
- behavior.md
- export-summary.md

Verification:
- <N> shared definitions found
- <N> .pq files written
- <N> workbook query connections found
```
