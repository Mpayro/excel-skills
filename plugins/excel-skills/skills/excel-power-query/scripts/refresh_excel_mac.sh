#!/bin/bash
# Refresh every Power Query of ONE workbook in Excel for Mac, wait until it settles, save, close,
# then prove from the saved file that each loaded query really refreshed.
#
#   refresh_excel_mac.sh /abs/path/Book.xlsx [--timeout SECONDS] [--backup-dir DIR] [--allow-open-books]
#
# Excel only opens, refreshes, recalculates, saves and closes. It never writes cells: cell writes
# through AppleScript crashed Excel on a large workbook and closed every book the user had open.
# A copy of the workbook from before the run is kept in --backup-dir (default ~/.pq-refresh-backups).
#
# Exit: 0 ok | 2 precondition or usage | 5 did not open | 6 refresh did not settle | 7 save failed
#       8 saved, but a loaded query is stale or not Complete (restore from the backup if needed)
set -u
usage(){ echo "usage: $0 /abs/path/Book.xlsx [--timeout SECONDS] [--backup-dir DIR] [--allow-open-books]"; exit 2; }
BOOK=""; TIMEOUT=900; ALLOW_OPEN=no; BACKUP_DIR="$HOME/.pq-refresh-backups"
while [ $# -gt 0 ]; do
  case "$1" in
    --timeout) [ $# -ge 2 ] || usage; TIMEOUT="$2"; shift 2 ;;
    --backup-dir) [ $# -ge 2 ] || usage; BACKUP_DIR="$2"; shift 2 ;;
    --allow-open-books) ALLOW_OPEN=yes; shift ;;
    -*) usage ;;
    *) [ -z "$BOOK" ] || usage; BOOK="$1"; shift ;;
  esac
done
case "$TIMEOUT" in ''|*[!0-9]*) usage ;; esac
[ -f "$BOOK" ] || usage
NAME="$(basename "$BOOK")"; DIR="$(dirname "$BOOK")"; HERE="$(cd "$(dirname "$0")" && pwd)"
log(){ echo "[$(date +%H:%M:%S)] $*"; }
sha(){ shasum -a 256 "$1" | cut -c1-16; }
# The workbook name reaches AppleScript as an argument (bookName), never as script text,
# so a quote or backslash in a file name cannot break or alter the command.
xl(){ osascript -e 'on run a' -e 'set bookName to item 1 of a' -e "tell application \"Microsoft Excel\" to $1" -e 'end run' -- "$NAME" 2>&1; }

# 0. Preconditions: the book has queries that load to a sheet, nobody has it open,
#    and no other open book is put at risk without consent.
LOADED="$(python3 -B "$HERE/pq_map.py" "$BOOK" --json | python3 -c 'import json,sys; print(sum(bool(q["loads_to"]) for q in json.load(sys.stdin)[0]["queries"]))' 2>/dev/null)"
case "$LOADED" in ''|0) log "ABORT: no query of $NAME loads to a sheet, so a refresh could not be proven"; exit 2 ;; esac
[ -e "$DIR/~\$$NAME" ] && { log "ABORT: lock file ~\$$NAME, someone has the workbook open"; exit 2; }
WAS_RUNNING=no
if pgrep -xq "Microsoft Excel"; then
  WAS_RUNNING=yes
  OPEN_BOOKS=""; [ "$(xl 'return (count of workbooks)')" = 0 ] || OPEN_BOOKS="$(xl 'return (name of every workbook)')"
  case ", $OPEN_BOOKS," in *", $NAME,"*) log "ABORT: $NAME is already open in Excel"; exit 2 ;; esac
  if [ -n "$OPEN_BOOKS" ] && [ "$ALLOW_OPEN" = no ]; then
    log "ABORT: Excel has other workbooks open ($OPEN_BOOKS). Ask the user to save and close them, or pass --allow-open-books."
    exit 2
  fi
fi
PREV_LINKS="$(xl 'return ask to update links')"
case "$PREV_LINKS" in true|false) ;; *) log "ABORT: Excel does not answer: $PREV_LINKS"; exit 2 ;; esac
BACKUP="$BACKUP_DIR/$(date -u +%Y%m%dT%H%M%SZ)_$NAME"
mkdir -p "$BACKUP_DIR" && cp -p "$BOOK" "$BACKUP" || { log "ABORT: could not write the backup $BACKUP"; exit 2; }
START="$(date -u +%Y-%m-%dT%H:%M:%S)"
log "sha before $(sha "$BOOK"); backup at $BACKUP"

# 1. Quiet Excel for the run, and put the user's settings back whatever happens.
restore(){
  if ! pgrep -xq "Microsoft Excel"; then
    [ "$PREV_LINKS" = true ] && log "WARNING: Excel is gone. Its setting 'ask to update links' was on and may now be off."
    return
  fi
  [ "$PREV_LINKS" = true ] && xl 'set ask to update links to true' >/dev/null
  xl 'set display alerts to true' >/dev/null
  if [ "$WAS_RUNNING" = no ] && [ "$(xl 'return (count of workbooks)')" = 0 ]; then xl 'quit' >/dev/null; fi
}
trap restore EXIT
xl 'set display alerts to false' >/dev/null
xl 'set ask to update links to false' >/dev/null

# 2. Open through LaunchServices. AppleScript `open workbook` raises a sandbox access dialog instead.
open -a "Microsoft Excel" "$BOOK"
opened=no; began=$SECONDS
while [ $((SECONDS-began)) -lt 180 ]; do
  sleep 2
  [ "$(xl 'return (exists workbook bookName)')" = true ] && { opened=yes; log "open after $((SECONDS-began))s"; break; }
done
[ "$opened" = yes ] || { log "ABORT: Excel did not open $NAME in 180s. A dialog may be waiting in Excel."; exit 5; }
sleep 8   # load-time recalculation

# 3. Refresh. The command returns at once; the work runs in Microsoft.Mashup.Container processes.
#    Quiet CPU is only a sign that the work ended. Step 5 is the proof.
xl 'refresh all workbook bookName' >/dev/null || { log "ABORT: refresh all failed"; exit 6; }
log "refresh all sent, waiting for the mashup containers to go quiet"
idle=0; settled=no; began=$SECONDS
while [ $((SECONDS-began)) -lt "$TIMEOUT" ]; do
  sleep 4
  cpu=$(ps -A -o %cpu,comm | grep -i "Mashup.Container" | awk '{s+=$1} END {printf "%d", s+0}')
  xcpu=$(ps -o %cpu= -p "$(pgrep -x 'Microsoft Excel' | head -1)" 2>/dev/null | awk '{printf "%d", $1+0}')
  if [ "$cpu" -lt 5 ] && [ "${xcpu:-0}" -lt 15 ]; then idle=$((idle+1)); else idle=0; fi
  if [ "$idle" -ge 5 ] && [ $((SECONDS-began)) -ge 24 ]; then settled=yes; log "quiet after $((SECONDS-began))s"; break; fi
done
[ "$settled" = yes ] || { log "ABORT: refresh still busy after ${TIMEOUT}s; workbook left open and unsaved"; exit 6; }

# 4. Recalculate, save, close.
xl 'calculate' >/dev/null; sleep 3
xl 'save workbook bookName' >/dev/null; sleep 3
[ "$(xl 'return (saved of workbook bookName)')" = true ] || { log "ABORT: save failed; workbook left open"; exit 7; }
xl 'close workbook bookName saving no' >/dev/null; sleep 2
log "saved and closed; sha after $(sha "$BOOK")"

# 5. Proof by content: every query that loads to a sheet must carry a fill time after START (UTC).
python3 -B - "$HERE/pq_map.py" "$BOOK" "$START" <<'PY'
import importlib.util, re, sys, zipfile
spec = importlib.util.spec_from_file_location('pq_map', sys.argv[1])
m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
bad = checked = 0
for q in m.build_map(sys.argv[2])['queries']:
    f = q['last_fill']
    checked += bool(q['loads_to'])
    stale = bool(q['loads_to']) and f.get('FillLastUpdated', '')[:19] < sys.argv[3]
    failed = f.get('FillStatus', 'Complete') != 'Complete'
    cell_errors = f.get('FillErrorCount', '0') not in ('', '0')
    if stale or failed or cell_errors:
        bad += stale or failed
        print(f"  {'PROBLEM' if stale or failed else 'note'}: {q['name']}: "
              f"filled {f.get('FillLastUpdated', '?')[:19]}Z, status {f.get('FillStatus', '?')}, "
              f"rows {f.get('FillCount', '?')}, cells with errors {f.get('FillErrorCount', '0')}")
links = sum(bool(re.fullmatch(r'xl/externalLinks/externalLink\d+\.xml', n)) for n in zipfile.ZipFile(sys.argv[2]).namelist())
if links:
    print(f'  note: {links} links to other workbooks. This proof covers queries only, not formulas that read those workbooks.')
if bad or not checked:
    print(f'{bad} of {checked} loaded queries did NOT refresh. The file was saved. The copy from before the run is the backup above.')
    sys.exit(8)
print(f'all {checked} loaded queries refreshed')
PY
