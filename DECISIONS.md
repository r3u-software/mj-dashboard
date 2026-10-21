# MJ Dashboard - decisions, findings, open items

Read this before re-discussing anything. Code notes are in `CLAUDE.md`.

## Architecture (decided)
- Dashboard page hosted on GitHub Pages (`docs/index.html`, repo r3u-software/mj-dashboard, Pages serves `/docs`). Local PowerShell backend does the Outlook scanning and answers the page.
- Backend is delivered as a self-extracting `docs/MJ_Dashboard_Backend.bat` (built by `build.ps1`), downloadable from a button on the page. Company laptop: no admin rights, ever.
- **Transport: plain `http://localhost:47615`.** Browsers treat localhost as secure, so the HTTPS page may call it: no certificate, no Windows prompt. The first design used a self-signed HTTPS cert in the CurrentUser store; it worked but Windows shows an unavoidable "install root certificate" prompt, so it was dropped. `-Https` remains as an optional fallback only.
- Only the Pages origin may read (CORS + Origin check), plus a per-run token. Safari blocks http-localhost; Chrome/Edge/Firefox are fine. Chrome may show a one-time "local network access" prompt - the user must allow it.
- Page detects the backend by polling `/status`; if the backend restarts (new `runId`) the page reloads itself.

## Auto-sync (decided + bugs already found - do not re-introduce)
- Every 30 s after the first full scan; new mail is added without a rescan and a toast shows the count.
- BUG 1: filtering by "received since last check" misses mail Outlook downloads late (it keeps the old ReceivedTime). Now: when a folder's item count changes, re-read newest-first for 7 days and keep unseen EntryIDs.
- BUG 2: never use `Items.Restrict` with dates - Outlook parses the text with Windows regional settings (dd/MM vs MM/dd) and silently returns wrong counts (seen on en-PH). Compare ReceivedTime in PowerShell instead.
- BUG 3: meeting invites are Outlook class 53, not 43, so they were skipped by the scan and the sync. "Email" now = classes 43, 53, 54, 55, 56, 46.
- If Outlook is closed the backend restarts it in the background (max once a minute). Outlook may keep running hidden after its window closes (COM reference) - normal.
- Scan reads only Inbox folders and their subfolders of every mailbox in the profile (skips Deleted, Junk, Drafts, Sent, etc.).

## How it was tested
- Tested on the dev PC: HTTP/CORS/token behaviour, demo mode (`-Demo`), page rendering, the "backend not found" panel, sync logic against a live Outlook (read-only), meeting-invite fix confirmed by the user on the work laptop.
- The `-Https` path and the Windows trust prompt were exercised once by the user (works), not by automated tests.
- Test approach: `-Demo -NoBrowser -Port <other> -AllowOrigin http://localhost:<port>` plus a local `python -m http.server` copy of the page. Do not kill the user's own running backend (it holds port 47615).

## Open items / ideas not built
- "Start backend" button on the page via a per-user custom URL protocol (`mjdashboard://start`, registered in HKCU on the first `.bat` run, no admin). Browser shows its own confirmation each time unless the user ticks "always allow". Agreed in principle; the user asked for a visual explanation first; not built.
- The `.bat` opens the dashboard tab every run, so a duplicate tab appears when started from the page. Options: keep (convenient from the desktop) or drop. Undecided.
- CSV is written only for the first scan, not for synced mail.
- Mail delivered more than 7 days late is not picked up by sync.
- Not verified: HTTPS fallback on a locked-down machine whose policy blocks the trust prompt.
