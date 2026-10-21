@../../BRAND-SYSTEM.md

# MJ Dashboard (hosted dashboard + local Outlook backend)

Dashboard HTML on GitHub Pages (`docs/index.html`); a PowerShell backend runs on the
user's PC, reads Outlook (read-only) and serves JSON over `http://localhost:47615` (browsers treat localhost as secure, so the HTTPS page may call it - no certificate, no prompt).

- `src/index.html`  - dashboard (edit this)
- `src/backend.ps1` - local backend (edit this)
- `build.ps1 [-PagesUrl https://<user>.github.io/<repo>/]` - writes `docs/` (index.html + self-extracting `MJ_Dashboard_Backend.bat`, which has the Pages URL and allowed CORS origin baked in). Never edit `docs/` by hand.

## Design rules
- No admin rights, ever. Default is plain HTTP on 127.0.0.1 (TcpListener, polled from the single scan thread - Outlook COM is STA). `-Https` is an optional fallback: self-signed cert in Cert:\CurrentUser\My + Root (Windows prompts once); it uses SslStream because HttpListener+HTTPS needs admin.
- Only `$AllowOrigin` may read; all data endpoints need the per-run token; `/ping` and `/status` are what the page uses to detect the backend.
- "Email" = Outlook classes 43,53,54,55,56,46 (meeting invites are 53, not 43; the rule categories for meetings/bounces depend on this).
- Auto-sync (restarts Outlook in the background if its process is gone, max once a minute): every 30 s `Sync-NewMail` re-reads (newest-first, last 7 days) any folder whose item count changed and keeps unseen EntryIDs; bumps `seq`; the page polls `/new?after=seq`. Never use `Items.Restrict` with dates: Outlook parses the text with Windows regional settings (dd/MM vs MM/dd) and silently returns wrong results. Mail Outlook downloads late keeps its old ReceivedTime, so "since last check" misses it.
- Dev switches: `-Demo` (fake mailbox, no Outlook), `-SkipTrust` (with -Https: skip the Root-store prompt).
- Not yet verified against a real Outlook mailbox or with the real trust-store prompt.
