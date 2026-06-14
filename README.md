# power-query-export

A Codex/Claude skill for exporting Power Query from Excel workbooks fast.

It reads `.xlsx` / `.xlsm` files directly as OOXML packages, extracts the embedded `DataMashup`, writes each Power Query definition as a `.pq` file, and generates a `behavior.md` explaining how queries connect to workbook sheets, Excel tables, loaded columns, and formulas.

## Use

Ask your agent:

```text
Use the power-query-export skill on /full/path/to/workbook.xlsx
```

The workbook path is required.

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

## Outputs

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
