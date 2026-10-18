@echo off
REM ==================================================================
REM  MJ Dashboard - local backend (single file)
REM  Double-click. It reads Outlook on THIS PC (read-only: no email is
REM  ever changed or deleted) and serves the live dashboard page over
REM  http://localhost (no certificate, no prompts). Keep the black
REM  window open while you use it. No admin rights needed.
REM ==================================================================
setlocal
set "MJ_SELF=%~f0"
set "MJ_WORK=%TEMP%\MJ_Dashboard"
set "MJ_OUT=%~dp0."
set "MJ_URL=https://r3u-software.github.io/mj-dashboard/"
if not exist "%MJ_WORK%" mkdir "%MJ_WORK%"

echo Unpacking...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$t=[IO.File]::ReadAllText($env:MJ_SELF,[Text.Encoding]::UTF8);$n='backend.ps1';$a=$t.IndexOf('::BEGIN '+$n);if($a -lt 0){exit 3};$a=$t.IndexOf([char]10,$a)+1;$b=$t.IndexOf('::END '+$n,$a);[IO.File]::WriteAllText((Join-Path $env:MJ_WORK $n),$t.Substring($a,$b-$a),(New-Object Text.UTF8Encoding($true)))"
if errorlevel 1 goto unpackfail
if not exist "%MJ_WORK%\backend.ps1" goto unpackfail

echo Starting backend (32-bit PowerShell)...
C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%MJ_WORK%\backend.ps1" -CsvFolder "%MJ_OUT%" -DashboardUrl "%MJ_URL%"

if errorlevel 2 goto retry64
if errorlevel 1 goto scanfail
goto end

:retry64
echo.
echo Retrying with 64-bit PowerShell...
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy Bypass -NoProfile -File "%MJ_WORK%\backend.ps1" -CsvFolder "%MJ_OUT%" -DashboardUrl "%MJ_URL%" -FinalAttempt
if errorlevel 1 goto scanfail
goto end

:unpackfail
echo.
echo ERROR: could not unpack the files. Make sure this .bat was not edited or cut short.
pause
exit /b 1

:scanfail
echo.
echo The backend reported an error. See the message above.
pause
exit /b 1

:end
exit /b 0
::BEGIN backend.ps1
param(
    [switch]$FinalAttempt,
    [int]$Port = 47615,
    [string]$CsvFolder = "",
    [string]$DashboardUrl = "",
    [string]$AllowOrigin = "https://r3u-software.github.io",
    [int]$SyncSeconds = 30,
    [switch]$NoBrowser,
    # serve HTTPS with a self-signed cert instead of plain http://localhost (asks Windows once to trust it)
    [switch]$Https,
    # developer switches (not used by the .bat): fake mailbox / skip the trust-store prompt
    [switch]$Demo,
    [switch]$SkipTrust
)

# ==================================================================
#  MJ Dashboard - local backend (READ-ONLY: it never changes, moves
#  or deletes any email)
#
#  The dashboard page lives on GitHub Pages (HTTPS). This script runs
#  on your PC, reads Outlook, and answers the page over HTTPS on
#  http://localhost:<Port>. Browsers treat localhost as a secure origin,
#  so an HTTPS page may call it: no certificate, no Windows prompt,
#  no admin rights.
#  (-Https: serves a self-signed certificate instead, stored for the
#  current user only; Windows asks once to trust it.)
#  - Scans the Inbox (and subfolders) of every mailbox in the Outlook
#    profile, whole mailbox, no date limit
#  - Every SyncSeconds (30) after the first scan it asks Outlook for
#    mail that arrived since, and the page adds it without a rescan
#  - Only the allowed dashboard origin can read from it, and every
#    request needs the per-run token
# ==================================================================

# ------------------------ SETTINGS (edit here) ---------------------

# Shared mailboxes NOT already showing in your Outlook folder list.
# Example: $SHARED_MAILBOXES = @("cgd.support@cummins.com")
$SHARED_MAILBOXES = @()

# Folders that are never scanned
$SKIP_FOLDERS = @("Deleted Items", "Junk Email", "Junk E-mail", "Drafts", "Sent Items", "Outbox",
                  "Conflicts", "Sync Issues", "Local Failures", "Server Failures", "RSS Feeds")

$CERT_SUBJECT = "CN=MJ Dashboard Local"
$VERSION = "2.1"
$LOOKBACK_DAYS = 7

# -------------------------------------------------------------------

if ([string]::IsNullOrWhiteSpace($CsvFolder)) { $CsvFolder = $PSScriptRoot }
$CsvFolder = $CsvFolder.Trim().TrimEnd('"').TrimEnd('\')

$inv = [System.Globalization.CultureInfo]::InvariantCulture
$startedAt = Get-Date
$token = [Guid]::NewGuid().ToString("N")
$runId = [Guid]::NewGuid().ToString("N").Substring(0, 8)
$allowedOrigins = @($AllowOrigin -split ',' | ForEach-Object { $_.Trim().TrimEnd('/') } | Where-Object { $_ })

$status = [ordered]@{
    state     = "starting"
    message   = "Starting"
    current   = ""
    processed = 0
    total     = 0
    mailboxes = @()
    startedAt = $startedAt.ToString("yyyy-MM-ddTHH:mm:ss")
    updatedAt = ""
    runId     = $runId
    token     = $token
    port      = $Port
    syncSeq   = 0
    syncTotal = 0
    syncedAt  = ""
    syncEvery = $SyncSeconds
    demo      = [bool]$Demo
}

$script:allRows = New-Object System.Collections.ArrayList
$script:newRows = New-Object System.Collections.ArrayList
$script:known = New-Object 'System.Collections.Generic.HashSet[string]'
$script:seq = 0
$script:storeIds = @{}
$script:folderCounts = @{}
$script:lastRelaunch = [DateTime]::MinValue
$script:listener = $null
$script:cert = $null
$script:lastPump = [DateTime]::MinValue
$script:csvName = ""

# ---------------------------- JSON helpers -------------------------
$jsonEval = [System.Text.RegularExpressions.MatchEvaluator]{
    param($m)
    $c = [int][char]$m.Value
    if ($c -eq 34) { return '\"' }
    if ($c -eq 92) { return '\\' }
    return ('\u{0:x4}' -f $c)
}
function Json-Str($v) {
    if ($null -eq $v) { return '""' }
    return '"' + [System.Text.RegularExpressions.Regex]::Replace([string]$v, '[\\"\x00-\x1f  ]', $jsonEval) + '"'
}
function Rows-ToJson($list) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('[')
    $first = $true
    foreach ($r in $list) {
        if (-not $first) { [void]$sb.Append(',') }
        $first = $false
        [void]$sb.Append('{"Subject":' + (Json-Str $r.Subject) + ',"SenderName":' + (Json-Str $r.SenderName) + ',"SenderEmail":' + (Json-Str $r.SenderEmail) + ',"Received":' + (Json-Str $r.Received) + ',"Mailbox":' + (Json-Str $r.Mailbox) + ',"Folder":' + (Json-Str $r.Folder) + ',"EntryID":' + (Json-Str $r.EntryID) + '}')
    }
    [void]$sb.Append(']')
    return $sb.ToString()
}

# ------------------------ certificate (no admin) -------------------
function Ensure-Cert {
    $my = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
    $my.Open("ReadWrite")
    $cert = $null
    foreach ($c in $my.Certificates) {
        if ($c.Subject -eq $CERT_SUBJECT -and $c.HasPrivateKey -and $c.NotAfter -gt (Get-Date).AddDays(30)) { $cert = $c }
    }
    $my.Close()

    if (-not $cert) {
        Write-Host "First run: creating the local HTTPS certificate..."
        try {
            $cert = New-SelfSignedCertificate -Subject $CERT_SUBJECT -FriendlyName "MJ Dashboard local backend" `
                -CertStoreLocation "Cert:\CurrentUser\My" -KeyAlgorithm RSA -KeyLength 2048 `
                -KeyExportPolicy NonExportable -KeySpec KeyExchange `
                -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" `
                -KeyUsage DigitalSignature, KeyEncipherment -NotAfter (Get-Date).AddYears(5) `
                -TextExtension @("2.5.29.17={text}DNS=localhost&IPAddress=127.0.0.1")
        } catch {
            Fail ("Could not create the HTTPS certificate: " + $_.Exception.Message)
        }
    }

    if (-not $SkipTrust) {
        $root = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser")
        $root.Open("ReadWrite")
        $trusted = $root.Certificates.Find("FindByThumbprint", $cert.Thumbprint, $false).Count -gt 0
        if (-not $trusted) {
            Write-Host ""
            Write-Host "Windows will now ask you ONCE to trust the MJ Dashboard certificate."
            Write-Host "Click YES. It is stored for your user only - no administrator rights needed."
            try {
                $pub = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(, $cert.RawData)
                $root.Add($pub)
            } catch {
                $root.Close()
                Fail "The certificate was not trusted (the Windows prompt was declined or blocked by policy). Run again and click Yes."
            }
        }
        $root.Close()
    }
    return $cert
}

# ------------------------------ HTTPS server -----------------------
# HttpListener needs administrator rights for HTTPS, so this uses a plain
# TcpListener on 127.0.0.1 + SslStream instead. It is polled (Pump-Http)
# from the scan loop, so everything stays on one thread (Outlook COM is
# single-threaded).
function Send-Resp($ssl, [int]$code, [string]$reason, [string]$ctype, [string]$text, [string]$origin) {
    $body = [System.Text.Encoding]::UTF8.GetBytes($text)
    $h = "HTTP/1.1 $code $reason`r`nContent-Type: $ctype; charset=utf-8`r`nContent-Length: $($body.Length)`r`nCache-Control: no-store`r`nConnection: close`r`nX-Content-Type-Options: nosniff`r`n"
    if ($origin -and ($allowedOrigins -contains $origin)) {
        $h += "Access-Control-Allow-Origin: $origin`r`nVary: Origin`r`nAccess-Control-Allow-Methods: GET, OPTIONS`r`nAccess-Control-Allow-Headers: *`r`nAccess-Control-Allow-Private-Network: true`r`nAccess-Control-Max-Age: 600`r`n"
    }
    $hb = [System.Text.Encoding]::ASCII.GetBytes($h + "`r`n")
    $ssl.Write($hb, 0, $hb.Length)
    if ($body.Length -gt 0) {
        $off = 0
        while ($off -lt $body.Length) {
            $n = [Math]::Min(65536, $body.Length - $off)
            $ssl.Write($body, $off, $n)
            $off += $n
        }
    }
    $ssl.Flush()
}

function Get-Query([string]$qs) {
    $q = @{}
    foreach ($pair in ($qs -split '&')) {
        if (-not $pair) { continue }
        $kv = $pair -split '=', 2
        $k = [Uri]::UnescapeDataString($kv[0])
        $v = ""
        if ($kv.Count -gt 1) { $v = [Uri]::UnescapeDataString($kv[1]) }
        $q[$k] = $v
    }
    return $q
}

function Status-Json {
    $status.updatedAt = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
    return (ConvertTo-Json -InputObject $status -Depth 6 -Compress)
}

function Route-Request($ssl, [string]$method, [string]$target, $hdr) {
    $origin = [string]$hdr["origin"]
    $path = $target
    $qs = ""
    $i = $target.IndexOf('?')
    if ($i -ge 0) { $path = $target.Substring(0, $i); $qs = $target.Substring($i + 1) }
    $q = Get-Query $qs
    $json = "application/json"
    $hostOk = ([string]$hdr["host"] -in @("localhost:$Port", "127.0.0.1:$Port"))

    if (-not $hostOk) { Send-Resp $ssl 400 "Bad Request" "text/plain" "Bad host" ""; return }

    if ($method -eq "OPTIONS") {
        if ($allowedOrigins -contains $origin) { Send-Resp $ssl 204 "No Content" "text/plain" "" $origin }
        else { Send-Resp $ssl 403 "Forbidden" "text/plain" "Origin not allowed" "" }
        return
    }
    if ($method -ne "GET") { Send-Resp $ssl 405 "Method Not Allowed" "text/plain" "" ""; return }

    if ($path -eq "/" -and -not $origin) {
        Send-Resp $ssl 200 "OK" "text/html" "<!doctype html><title>MJ Dashboard backend</title><body style='font-family:Segoe UI,sans-serif;padding:32px'><h2>MJ Dashboard backend is running</h2><p>The certificate is trusted - this page loaded over HTTPS without a warning. Go back to the dashboard.</p></body>" ""
        return
    }
    if ($path -eq "/ping") {
        Send-Resp $ssl 200 "OK" $json ('{"ok":true,"app":"mj-dashboard-backend","version":' + (Json-Str $VERSION) + ',"runId":"' + $runId + '"}') $origin
        return
    }
    if ($allowedOrigins -notcontains $origin) {
        Send-Resp $ssl 403 "Forbidden" $json '{"error":"origin not allowed"}' ""
        return
    }
    if ($path -eq "/status") {
        Send-Resp $ssl 200 "OK" $json (Status-Json) $origin
        return
    }
    if ([string]$q["token"] -ne $token) {
        Send-Resp $ssl 403 "Forbidden" $json '{"error":"wrong token - reload the dashboard"}' $origin
        return
    }
    switch ($path) {
        "/rows" {
            Send-Resp $ssl 200 "OK" $json ('{"runId":"' + $runId + '","seq":' + $script:seq + ',"rows":' + (Rows-ToJson $script:allRows) + '}') $origin
        }
        "/new" {
            $after = 0
            [void][int]::TryParse([string]$q["after"], [ref]$after)
            $list = New-Object System.Collections.ArrayList
            foreach ($r in $script:newRows) { if ($r.Seq -gt $after) { [void]$list.Add($r) } }
            Send-Resp $ssl 200 "OK" $json ('{"runId":"' + $runId + '","seq":' + $script:seq + ',"syncedAt":' + (Json-Str $status.syncedAt) + ',"rows":' + (Rows-ToJson $list) + '}') $origin
        }
        "/open" {
            $res = Open-Mail ([string]$q["id"]) ([string]$q["mb"])
            Send-Resp $ssl 200 "OK" $json ('{"ok":' + $(if ($res.ok) { 'true' } else { 'false' }) + ',"msg":' + (Json-Str $res.msg) + '}') $origin
        }
        default { Send-Resp $ssl 404 "Not Found" $json '{"error":"not found"}' $origin }
    }
}

function Handle-Client($client) {
    $ssl = $null
    try {
        $client.ReceiveTimeout = 3000
        $client.SendTimeout = 15000
        if ($script:cert) {
            $ssl = New-Object System.Net.Security.SslStream($client.GetStream(), $false)
            $ssl.AuthenticateAsServer($script:cert, $false, [System.Security.Authentication.SslProtocols]::Tls12, $false)
        } else {
            $ssl = $client.GetStream()
        }

        $buf = New-Object byte[] 4096
        $text = ""
        while ($text.IndexOf("`r`n`r`n") -lt 0 -and $text.Length -lt 16384) {
            $n = $ssl.Read($buf, 0, $buf.Length)
            if ($n -le 0) { break }
            $text += [System.Text.Encoding]::ASCII.GetString($buf, 0, $n)
        }
        if ($text.IndexOf("`r`n`r`n") -lt 0) { return }
        $lines = $text -split "`r`n"
        $rl = $lines[0] -split ' '
        if ($rl.Count -lt 2) { return }
        $hdr = @{}
        for ($k = 1; $k -lt $lines.Count; $k++) {
            $c = $lines[$k].IndexOf(':')
            if ($c -gt 0) { $hdr[$lines[$k].Substring(0, $c).Trim().ToLower()] = $lines[$k].Substring($c + 1).Trim() }
        }
        Route-Request $ssl $rl[0] $rl[1] $hdr
    } catch {
        # aborted / timed-out connection (browsers open spare sockets) - ignore
        if ($Demo) { Write-Host ('[conn] ' + $_.Exception.GetType().Name + ': ' + $_.Exception.Message + ' / ' + $_.Exception.InnerException.Message) }
    } finally {
        try { if ($ssl) { $ssl.Close() } } catch { }
        try { $client.Close() } catch { }
    }
}

function Pump-Http([switch]$Force) {
    if (-not $script:listener) { return }
    if (-not $Force -and ([DateTime]::Now - $script:lastPump).TotalMilliseconds -lt 200) { return }
    $n = 0
    while ($n -lt 8 -and $script:listener.Pending()) {
        $n++
        Handle-Client ($script:listener.AcceptTcpClient())
    }
    $script:lastPump = [DateTime]::Now
}

function Push-Status { Pump-Http }

function Fail($msg) {
    Write-Host ""
    Write-Host ("ERROR: " + $msg)
    $status.state = "error"
    $status.message = $msg
    # stay up a few seconds so the dashboard can fetch the message
    $until = (Get-Date).AddSeconds(4)
    while ($script:listener -and (Get-Date) -lt $until) { Pump-Http -Force; Start-Sleep -Milliseconds 200 }
    exit 1
}

# ------------------------------ Outlook ----------------------------
$senderCache = @{}
function Get-SenderSmtp($mail) {
    try {
        $key = [string]$mail.SenderEmailAddress
        if ($key -and $senderCache.ContainsKey($key)) { return $senderCache[$key] }
        $smtp = ""
        if ($mail.SenderEmailType -eq "EX") {
            $ae = $mail.Sender
            if ($ae) {
                $exu = $ae.GetExchangeUser()
                if ($exu) { $smtp = [string]$exu.PrimarySmtpAddress }
            }
        } else {
            $smtp = $key
        }
        if ($key) { $senderCache[$key] = $smtp }
        return $smtp
    } catch { }
    return ""
}

# Returns a working Outlook namespace. If Outlook was closed after the scan,
# this starts it again (used by "Open in Outlook").
function Get-OutlookNs([switch]$Force) {
    $ErrorActionPreference = "Stop"
    $script:restarted = $false
    $script:coldStart = $false
    if (-not $Force) {
        try { $null = $script:ns.CurrentUser.Name; return $script:ns } catch { }
    }
    $script:coldStart = -not [bool](Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:ol) } catch { }
    $script:ol = New-Object -ComObject Outlook.Application
    $script:ns = $script:ol.GetNamespace("MAPI")
    try { $script:ns.Logon("", "", $false, $false) } catch { }
    $script:restarted = $true
    return $script:ns
}

function Get-MailItem($nsx, [string]$id, [string]$mb) {
    try {
        if ($script:storeIds.ContainsKey($mb)) { return $nsx.GetItemFromID($id, $script:storeIds[$mb]) }
        return $nsx.GetItemFromID($id)
    } catch { return $null }
}

# The Inbox of every mailbox (own + shared)
function Get-Targets($nsx) {
    $targets = New-Object System.Collections.ArrayList
    $seen = @{}
    $found = New-Object System.Collections.ArrayList
    foreach ($store in $nsx.Stores) {
        try { [void]$found.Add(@([string]$store.DisplayName, $store.GetDefaultFolder(6))) } catch { }
    }
    foreach ($addr in $SHARED_MAILBOXES) {
        try {
            $rcp = $nsx.CreateRecipient($addr)
            [void]$rcp.Resolve()
            if ($rcp.Resolved) { [void]$found.Add(@($addr, $nsx.GetSharedDefaultFolder($rcp, 6))) }
        } catch { }
    }
    foreach ($f in $found) {
        try { $id = [string]$f[1].EntryID } catch { continue }
        if ($seen.ContainsKey($id)) { continue }
        $seen[$id] = $true
        [void]$targets.Add([PSCustomObject]@{ Mailbox = $f[0]; Folder = $f[1] })
    }
    return , $targets
}

function Walk-Folders($fld, $mailbox, $path, $out) {
    $name = [string]$fld.Name
    if ($SKIP_FOLDERS -contains $name) { return }
    if ($path) { $fpath = $path + "\" + $name } else { $fpath = $name }
    [void]$out.Add([PSCustomObject]@{ Mailbox = $mailbox; Path = $fpath; Folder = $fld })
    try { foreach ($sub in $fld.Folders) { Walk-Folders $sub $mailbox $fpath $out } } catch { }
}

function New-Row($itm, $mailbox, $path) {
    return [PSCustomObject]@{
        Subject     = [string]$itm.Subject
        SenderName  = [string]$itm.SenderName
        SenderEmail = (Get-SenderSmtp $itm)
        Received    = $itm.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
        Mailbox     = $mailbox
        Folder      = $path
        EntryID     = [string]$itm.EntryID
        Seq         = 0
    }
}

function Open-Mail([string]$id, [string]$mb) {
    $ok = $false
    $msg = "Could not open this email."
    try {
        if ($Demo) { return @{ ok = $false; msg = "Demo mode: there is no real Outlook email behind this row." } }
        if ($id -notmatch '^[0-9A-Fa-f]{10,600}$') { return @{ ok = $false; msg = "This email has no valid Outlook ID." } }
        $item = $null
        $nsx = $null
        try { $nsx = Get-OutlookNs } catch { $nsx = $null; $msg = "Could not start Outlook: " + $_.Exception.Message }
        if ($nsx) {
            $item = Get-MailItem $nsx $id $mb
            if (-not $item -and -not $script:restarted) {
                # connection may be stale (Outlook was closed): reconnect once and try again
                try { $nsx = Get-OutlookNs -Force; $item = Get-MailItem $nsx $id $mb } catch { $item = $null; $msg = "Could not start Outlook: " + $_.Exception.Message }
            }
            if (-not $item -and $script:restarted -and $script:coldStart) {
                # Outlook was just started: give it time to load the mailboxes
                $deadline = (Get-Date).AddSeconds(45)
                while (-not $item -and (Get-Date) -lt $deadline) {
                    Start-Sleep -Milliseconds 1500
                    $item = Get-MailItem $nsx $id $mb
                }
            }
        }
        if ($item) {
            $item.Display(); $ok = $true; $msg = ""
            try { $item.GetInspector.Activate() } catch { }
        }
        elseif ($nsx) { $msg = "Outlook could not find that email. It may have been moved or deleted." }
    } catch { $msg = "Outlook error: " + $_.Exception.Message }
    return @{ ok = $ok; msg = $msg }
}

# ------------------- auto-refresh: pick up new mail ----------------
function Add-NewRow($row) {
    if ($row.EntryID -and -not $script:known.Add($row.EntryID)) { return $false }
    $script:seq++
    $row.Seq = $script:seq
    [void]$script:allRows.Add($row)
    [void]$script:newRows.Add($row)
    return $true
}

function Sync-NewMail {
    $added = 0
    if ($Demo) {
        $subs = @("Payment advice 4500", "Invoice INV-7781 query", "Automatic reply: Out of office", "Statement of account", "PO 90012 confirmation")
        for ($k = 0; $k -lt (Get-Random -Minimum 1 -Maximum 4); $k++) {
            $row = [PSCustomObject]@{
                Subject = ($subs | Get-Random); SenderName = "Demo Sender"; SenderEmail = "demo@example.com"
                Received = (Get-Date).ToString("yyyy-MM-dd HH:mm"); Mailbox = "Demo mailbox"; Folder = "Inbox"
                EntryID = ([Guid]::NewGuid().ToString("N") + "00000000"); Seq = 0
            }
            if (Add-NewRow $row) { $added++ }
        }
    } else {
        $cycle = Get-Date
        try {
            # Outlook was closed: start it again (in the background) so auto-sync keeps working.
            # Retried at most once a minute while Outlook is still loading.
            $wasClosed = -not [bool](Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
            if ($wasClosed) {
                if (([DateTime]::Now - $script:lastRelaunch).TotalSeconds -lt 60) { return }
                $script:lastRelaunch = [DateTime]::Now
                Write-Host ("{0}  Outlook is closed - starting it again..." -f (Get-Date -Format "HH:mm:ss"))
                $nsx = Get-OutlookNs -Force
                $script:folderCounts.Clear()   # re-check every folder once it is back
                Start-Sleep -Seconds 5
            } else {
                $nsx = Get-OutlookNs
            }
            # Outlook can download mail late (account catching up, was offline) and the email keeps its
            # ORIGINAL received time, so "since the last check" would miss it. Instead: when a folder's
            # item count changed, look back $LOOKBACK_DAYS days and keep whatever is not known yet.
            $cutoff = $cycle.AddDays(-$LOOKBACK_DAYS)   # no Restrict(): its date text is parsed by Windows regional settings
            $folders = New-Object System.Collections.ArrayList
            foreach ($t in (Get-Targets $nsx)) { Walk-Folders $t.Folder $t.Mailbox "" $folders }
            foreach ($f in $folders) {
                Pump-Http
                try {
                    $key = $f.Mailbox + "|" + $f.Path
                    $cnt = [int]$f.Folder.Items.Count
                    if ($script:folderCounts.ContainsKey($key) -and $script:folderCounts[$key] -eq $cnt) { continue }
                    $items = $f.Folder.Items
                    $items.Sort("[ReceivedTime]", $true)
                    $itm = $items.GetFirst()
                    while ($itm) {
                        try {
                            if ($itm.ReceivedTime -lt $cutoff) { break }
                            if ($itm.Class -eq 43 -and -not $script:known.Contains([string]$itm.EntryID)) {
                                if (Add-NewRow (New-Row $itm $f.Mailbox $f.Path)) { $added++ }
                            }
                        } catch { }
                        $itm = $items.GetNext()
                    }
                    $script:folderCounts[$key] = $cnt
                } catch { }
            }
        } catch {
            Write-Host ("Auto-refresh skipped: " + $_.Exception.Message)
            return
        }
    }
    $status.syncSeq = $script:seq
    $status.syncTotal = $status.syncTotal + $added
    $status.syncedAt = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
    if ($added -gt 0) { Write-Host ("{0}  {1} new email(s) synced" -f (Get-Date -Format "HH:mm:ss"), $added) }
}

# =============================== START ===============================
Write-Host "MJ Dashboard backend v$VERSION"

# 1. certificate (only with -Https)
if ($Https) { $script:cert = Ensure-Cert }

# 2. Outlook (exit code 2 = let the .bat retry with 64-bit PowerShell)
if (-not $Demo) {
    try {
        $ol = New-Object -ComObject Outlook.Application
    } catch {
        if (-not $FinalAttempt) { exit 2 }
        Fail "Could not start Outlook. Open Outlook once, then run again."
    }
    $ns = $ol.GetNamespace("MAPI")
}

# 3. HTTPS server (before the scan, so the page can show live progress)
try {
    $l = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $Port)
    $l.Start()
    $script:listener = $l
} catch {
    Write-Host ("ERROR: port $Port is busy. The backend is probably already running in another window - close that one first.")
    exit 1
}
Write-Host ("Listening on " + $(if ($script:cert) { "https" } else { "http" }) + "://localhost:$Port (this PC only)")

# 4. open the hosted dashboard
if ($DashboardUrl -and -not $NoBrowser) { try { Start-Process $DashboardUrl } catch { } }

try {
    if ($Demo) {
        $status.mailboxes = @("Demo mailbox")
        $status.state = "scanning"; $status.message = "Reading emails"; $status.total = 600
        $subs = @("Payment advice 4500", "Invoice INV-7781 query", "Automatic reply: Out of office", "Statement of account", "PO 90012 confirmation",
                  "Remittance for June", "Weekly newsletter - 20% off", "Accepted: Supplier review meeting", "Overdue invoice reminder", "Bank details change request")
        $demoRows = New-Object System.Collections.ArrayList
        for ($k = 0; $k -lt 600; $k++) {
            [void]$demoRows.Add([PSCustomObject]@{
                Subject = ($subs[$k % $subs.Count] + " " + (1000 + $k)); SenderName = "Sender " + ($k % 17); SenderEmail = "sender" + ($k % 17) + "@example.com"
                Received = (Get-Date).AddDays(-($k % 90)).AddHours(-($k % 9)).ToString("yyyy-MM-dd HH:mm"); Mailbox = "Demo mailbox"; Folder = "Inbox"
                EntryID = ("{0:x40}" -f $k); Seq = 0 })
            if ($k % 50 -eq 0) { $status.processed = $k; Pump-Http -Force; Start-Sleep -Milliseconds 250 }
        }
        foreach ($r in $demoRows) { [void]$script:known.Add($r.EntryID) }
        $sorted = @($demoRows | Sort-Object Received -Descending)
        $processed = 600
    } else {
        $targets = Get-Targets $ns
        if ($targets.Count -eq 0) { Fail "No mailbox Inbox found in this Outlook profile." }

        $status.mailboxes = @($targets | ForEach-Object { $_.Mailbox })
        $status.message = "Counting emails"
        Push-Status

        # ---- first pass: count emails so the progress bar is real ----
        $records = New-Object System.Collections.ArrayList
        $script:grandTotal = 0

        function Collect-Folder($fld, $mailbox, $path) {
            $name = [string]$fld.Name
            if ($SKIP_FOLDERS -contains $name) { return }
            if ($path) { $fpath = $path + "\" + $name } else { $fpath = $name }

            $items = $null
            $cnt = 0
            try { $items = $fld.Items; $cnt = [int]$items.Count } catch { $items = $null; $cnt = 0 }

            if ($items -and $cnt -gt 0) {
                [void]$records.Add([PSCustomObject]@{ Mailbox = $mailbox; Path = $fpath; Items = $items; Count = $cnt })
                $script:grandTotal += $cnt
            }
            if (-not $script:storeIds.ContainsKey($mailbox)) { try { $script:storeIds[$mailbox] = [string]$fld.StoreID } catch { } }
            $status.current = $mailbox + " > " + $fpath
            $status.total = $script:grandTotal
            Push-Status

            try { foreach ($sub in $fld.Folders) { Collect-Folder $sub $mailbox $fpath } } catch { }
        }

        foreach ($t in $targets) { Collect-Folder $t.Folder $t.Mailbox "" }

        $total = $script:grandTotal
        Write-Host ("Found {0} emails across {1} mailbox(es). Collecting..." -f $total, $targets.Count)

        $status.state = "scanning"
        $status.message = "Reading emails"
        $status.total = $total
        Push-Status

        # ---- second pass: collect every email ----
        $script:lastSync = Get-Date
        $processed = 0
        $lastPush = Get-Date

        foreach ($rec in $records) {
            $status.current = $rec.Mailbox + " > " + $rec.Path
            foreach ($itm in $rec.Items) {
                $processed++
                try {
                    if ($itm.Class -eq 43) {
                        $row = New-Row $itm $rec.Mailbox $rec.Path
                        [void]$script:allRows.Add($row)
                        if ($row.EntryID) { [void]$script:known.Add($row.EntryID) }
                    }
                } catch { }

                if (([DateTime]::Now - $lastPush).TotalMilliseconds -ge 250) {
                    $status.processed = $processed
                    Pump-Http
                    Write-Host ("`rScanning: {0} / {1} emails" -f $processed, $total) -NoNewline
                    $lastPush = Get-Date
                }
            }
        }
        $sorted = @($script:allRows | Sort-Object Received -Descending)
    }

    # ---- finish ----
    $script:allRows.Clear()
    foreach ($r in $sorted) { [void]$script:allRows.Add($r) }

    $csvName = ""
    if (-not $Demo -and $sorted.Count -gt 0) {
        $csvName = "EmailSubjects_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv"
        try { $sorted | Select-Object Subject, SenderName, SenderEmail, Received, Mailbox, Folder, EntryID | Export-Csv -Path (Join-Path $CsvFolder $csvName) -NoTypeInformation -Encoding UTF8 }
        catch { $csvName = ""; Write-Host ("Could not save the CSV: " + $_.Exception.Message) }
    }
    $status.csv = $csvName
    $status.csvFolder = $CsvFolder
    $status.processed = $processed
    $status.current = ""
    $status.syncedAt = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
    $status.message = "Done"
    $status.state = "done"

    $elapsed = ((Get-Date) - $startedAt).TotalSeconds
    Write-Host ""
    Write-Host ("DONE. {0} emails collected in {1:N0}s." -f $sorted.Count, $elapsed)
    if ($csvName) { Write-Host ("Saved: " + (Join-Path $CsvFolder $csvName)) }
    Write-Host ""
    Write-Host "Keep this window open: the dashboard reads from it, 'Open in Outlook' uses it,"
    Write-Host "and it checks for new mail every $SyncSeconds seconds. Close it (or press Ctrl+C) when finished."

    # ---- serve + auto-refresh until the window is closed ----
    $script:lastSync = if ($script:lastSync) { $script:lastSync } else { Get-Date }
    $nextSync = (Get-Date).AddSeconds($SyncSeconds)
    while ($true) {
        Pump-Http -Force
        if ((Get-Date) -ge $nextSync) {
            Sync-NewMail
            $nextSync = (Get-Date).AddSeconds($SyncSeconds)
        }
        Start-Sleep -Milliseconds 100
    }
}
catch {
    Fail ("Scan stopped: " + $_.Exception.Message)
}
finally {
    try { if ($script:listener) { $script:listener.Stop() } } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ol) } catch { }
}
::END backend.ps1
