# excel-skills

A plugin for Claude Code and Codex with two agent skills for Excel workbooks that people depend on.

- **excel-power-query** maps every Power Query in a workbook: where it loads, what it reads, what uses it and when it last refreshed. It finds which workbook feeds which, exports the M code, and refreshes from the command line on macOS with proof that the refresh worked.
- **excel-write-verify** changes values, formulas and columns inside an `.xlsx` file while Power Query, tables, notes and formatting stay intact. It then proves the result with a cell-by-cell diff, a structural check and an independent recalculation, and replaces a shared file safely.

The skills run on your computer. They read and write only the workbooks you point them at, make no network calls and collect no data.

Full documentation, install commands and limits: https://github.com/Mpayro/excel-skills

License: MIT.
