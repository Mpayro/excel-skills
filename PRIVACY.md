# Privacy

The skills and scripts in this repo run on your own computer.

- They read and write only the workbook files that you or your agent name, and the backup folders described in the README (`~/.pq-refresh-backups`, `~/.xlsx-backups`, or the folder you choose).
- They make no network calls. They send nothing to the author or to any service.
- They collect no data, keep no logs outside those backup folders, and contain no telemetry.
- `lo_recalc.py` starts LibreOffice on your computer with a temporary profile. `refresh_excel_mac.sh` drives your local Microsoft Excel. A Power Query refresh reads whatever sources the workbook itself defines.

Your agent (Claude Code, Codex or another) has its own data handling. That is governed by the terms of that product, not by this repo.

Questions: open an issue at https://github.com/Mpayro/excel-skills/issues
