#!/usr/bin/env python3
"""Put a new version of a workbook in place of the live one, without stepping on anybody.

  xlsx_replace.py NEW.xlsx LIVE.xlsx --base-sha SHA256 [--backup-dir DIR] [--wait 20] [--go]

Dry run unless --go. Refuses when:
  - the live file is not the version the change was built on (its SHA-256 must start with --base-sha),
  - a lock file exists (someone has it open in Excel): a file in the same folder that starts with "~$" and
    whose rest is the full name or the end of the name (6 or more characters, for the shortened lock names),
  - a local process has the file open,
  - NEW is not a readable zip package.
With --go: copies the live file to --backup-dir (default ~/.xlsx-backups), writes a temporary file in
the same folder, renames it over the live file in one step, checks the content, waits, checks again,
and lists files that look like conflict copies. The live file is hashed again just before the rename; if it
changed, nothing is replaced. The new file gets the file mode of the live file.
Exit code: 0 only when every check passed | 2 a check refused (nothing replaced) or a usage error, a missing
or unreadable file | 3 replaced, but the content changed during the wait | 1 the replace failed its own check.
"""
import argparse, datetime, hashlib, os, shutil, subprocess, sys, time, zipfile
from pathlib import Path


def die(msg):
    print(msg, file=sys.stderr)
    sys.exit(2)


def sha(p):
    # A synced folder can hold the file as "online only": make it local first, with a time limit.
    if not Path(p).is_file():
        die(f"no such file: {p}")
    try:
        subprocess.run([sys.executable, "-c", "import sys; open(sys.argv[1], 'rb').read(1)", str(p)], check=True, timeout=90)
    except subprocess.TimeoutExpired:
        die(f"cannot read {p} within 90 s (online-only file in a synced folder? make it local first)")
    except subprocess.CalledProcessError:
        die(f"cannot read {p}")
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for block in iter(lambda: f.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def open_locally(p):
    if not shutil.which("lsof"):
        return None                     # no lsof on this system (Windows): cannot tell
    return bool(subprocess.run(["lsof", "--", str(p)], capture_output=True, text=True).stdout.strip())


def lock_files(live):
    """Lock files of Excel next to `live`: ~$ + the full name, or + the end of the name (6 or more characters)."""
    name = live.name
    out = []
    for f in os.listdir(live.parent):
        rest = f[2:]
        if f.startswith("~$") and (rest == name or (len(rest) >= 6 and name.endswith(rest))):
            out.append(f)
    return sorted(out)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("new", type=Path)
    ap.add_argument("live", type=Path)
    ap.add_argument("--base-sha", required=True, help="SHA-256 (12 or more hex digits) of the live version the change was built on")
    ap.add_argument("--backup-dir", type=Path, default=Path.home() / ".xlsx-backups")
    ap.add_argument("--wait", type=int, default=20, help="seconds before the second content check")
    ap.add_argument("--go", action="store_true")
    if len(sys.argv) == 1:
        ap.print_help(sys.stderr)
        sys.exit(2)
    a = ap.parse_args()
    if len(a.base_sha) < 12:
        die("--base-sha needs at least 12 hex digits")

    new_sha, cur = sha(a.new), sha(a.live)
    locks = lock_files(a.live)
    is_open = open_locally(a.live)
    try:
        with zipfile.ZipFile(a.new) as z:
            zip_ok = z.testzip() is None
    except zipfile.BadZipFile:
        zip_ok = False
    problems = []
    if not cur.startswith(a.base_sha.lower()):
        problems.append(f"the live file changed since the base ({cur[:12]} is not {a.base_sha[:12]}): rebuild on the new version")
    if locks:
        problems.append(f"lock file {locks[0]}: someone has the workbook open")
    if is_open:
        problems.append("a local process has the live file open")
    if not zip_ok:
        problems.append("NEW is not a readable workbook package")
    if new_sha == cur:
        problems.append("NEW is identical to the live file")
    print(f"live {cur[:12]} | new {new_sha[:12]} | lock file {bool(locks)} | open locally {is_open}")
    for p in problems:
        print("STOP:", p)
    if problems:
        sys.exit(2)
    if not a.go:
        print("dry run: all checks passed, pass --go to write")
        return

    a.backup_dir.mkdir(parents=True, exist_ok=True)
    backup = a.backup_dir / f"{datetime.datetime.now(datetime.timezone.utc):%Y%m%dT%H%M%SZ}_{a.live.name}"
    shutil.copy2(a.live, backup)
    if sha(backup) != cur:
        sys.exit(f"backup {backup} does not match the live file; nothing replaced")
    print("backup", backup)

    tmp = a.live.parent / f".{a.live.name}.{os.getpid()}.tmp"     # same folder, so the rename is one step
    try:
        with open(a.new, "rb") as fi, open(tmp, "wb") as fo:
            shutil.copyfileobj(fi, fo)
            fo.flush()
            os.fsync(fo.fileno())
        shutil.copymode(a.live, tmp)
        if sha(a.live) != cur:                       # a person saved after the first check: do not overwrite
            tmp.unlink()
            print("STOP: the live file changed while the replace was prepared; nothing replaced "
                  f"(backup {backup} is of the earlier version)")
            sys.exit(2)
        os.replace(tmp, a.live)
    finally:
        if tmp.exists():
            tmp.unlink()
    if sha(a.live) != new_sha:
        sys.exit("VERIFY FAILED right after the replace; restore from the backup")
    print(f"replaced; waiting {a.wait}s for the sync client")
    time.sleep(a.wait)
    after = sha(a.live)
    stem = a.live.stem
    others = sorted(f for f in os.listdir(a.live.parent)
                    if f.startswith(stem) and f != a.live.name and f not in locks and not f.endswith(".tmp"))
    print(f"content after {a.wait}s: {'same' if after == new_sha else 'CHANGED (' + after[:12] + ')'}")
    print(f"other files that start with {stem!r} (look for conflict copies): {others or 'none'}")
    sys.exit(0 if after == new_sha else 3)


if __name__ == "__main__":
    main()
