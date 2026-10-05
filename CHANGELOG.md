# Changelog

Versions follow the plugin version in `plugins/excel-skills/.claude-plugin/plugin.json`.

## 1.1.1 (2026-10-05)

- MIT license.
- Smoke tests on Linux, macOS and Windows, run on every change and once a month.
- Contributing guide and bug report template.

## 1.1.0 (2026-10-05)

- The repo is now `Mpayro/excel-skills`. The old address redirects.
- Skill `power-query-export` is now `excel-power-query`.
- README rewritten: why, install, example requests, comparison, FAQ.

## 1.0.0 (2026-10-05)

- Packaged as a plugin for Claude Code and Codex, with a marketplace file for each.
- New skill `excel-write-verify`: edits inside the workbook package (`xlsxpatch.py`, `colinsert.py`), structural check (`xlsx_check.py`), cell-by-cell diff (`xlsx_diff.py`), independent recalculation (`lo_recalc.py`) and a guarded replace of shared files (`xlsx_replace.py`).

## 0.2.0 (2026-10-05)

- Power Query navigation with `pq_map.py`: load targets, sources, references, saved table filters, last fill, and the refresh order across linked workbooks.
- Command-line refresh on macOS with `refresh_excel_mac.sh`, with proof from the saved file.
- Field notes on what a refresh overwrites, chains of workbooks and synced folders.

## 0.1.0 (2026-06-14)

- First release: export of Power Query definitions and a written analysis of workbook behavior.
