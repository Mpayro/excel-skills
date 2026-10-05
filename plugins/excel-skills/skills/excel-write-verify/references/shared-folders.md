# Workbooks in OneDrive, SharePoint, Google Drive or a network share

Rules from real incidents with a synced folder on macOS. Most of them apply to any sync client.

## Reading

1. **Make the file local before you read it.** A sync client can keep a file "online only". A read then hangs or fails in the middle, and a zip read gives a damaged-file error. Read one byte first, with a time limit. Then copy the file in blocks to a local folder and work on the copy. If the file does not come down, the fix is manual (mark it "always keep on this device"). Report it; do not wait forever.
2. **Put a time limit on every subprocess that touches the synced folder**, and handle the timeout. A parent process must not hang because of one file.
3. **A recursive search can miss online-only files.** In one incident a sweep with `grep -r` over a synced folder missed 12 of 30 real hits and printed no message. Go file by file, make each one local first, and list the files that you could not read. A result of "nothing found" from a recursive search over a synced folder is not evidence.
4. **Do not create virtual environments or other tool folders inside the synced folder.**

## Writing

5. **Never write directly to the final path.** Write a temporary file in the same folder, with a name that is unique to the process, and rename it over the target in one step. Delete the temporary file if anything fails. A temporary file with a fixed name is shared by overlapping runs and gets truncated.
6. **Check for Excel before you write a workbook.** A lock file `~$Name.xlsx` or an open file handle means a person has it open. Stop with a clear message. A missing lock file does not prove that nobody edits the file in the browser.
7. **One writer per file.** Two processes that write the same workbook must run one after the other, or one must become the input of the other. No lock between processes survives a sync client.
8. **Compare with the base before you replace.** Record the hash of the version you copied. If the live file has another hash when you come back, a person saved in the meantime. Start again from their version. This happened overnight between a copy and a deploy.
9. **Keep backups outside the synced folder, in a place that survives a restart.** Not in `/tmp`.
10. **Read-only permissions do not protect a file from changes that arrive from another machine.** Defend a published file with an expected hash that you can check.

## Verifying

11. **Verify by content, never by date.** File dates in a synced folder are wrong in both directions: a write can show an old date, and a file can change with no local writer. Read back what you wrote: a known cell, a row count, a hash.
12. **Check again after the sync client had time to act**, and look for conflict copies next to the file. `xlsx_replace.py` does both.
13. **"Nothing written" must say why.** A step that can end without writing must tell "no change needed" from "could not read or compare", and log it.

## macOS

14. **"Operation not permitted" on the whole synced folder while Finder shows the files normally** can be a privacy permission of the terminal application, not a sync problem. In one incident, `tccutil reset FileProviderDomain <bundle id of the terminal app>` fixed it at once; restarting the sync client did not. This was seen once and is not tested elsewhere.
15. **A background job started by launchd could not read the synced folder in one incident.** The likely reason was that it had no privacy identity that could be granted access. Keep the state of such a job on local paths, and test a scheduled job against the synced folder before you rely on it.
