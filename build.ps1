# Builds docs/ (what GitHub Pages serves):
#   docs/index.html                   the dashboard
#   docs/MJ_Dashboard_Backend.bat     self-extracting local backend (downloaded from the page)
param([string]$PagesUrl = "https://r3u-software.github.io/mj-dashboard/")
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$out = Join-Path $root "docs"
New-Item -ItemType Directory -Force $out | Out-Null
$crlf = { param($t) ($t -replace "`r`n", "`n") -replace "`n", "`r`n" }

$origin = ([Uri]$PagesUrl).GetLeftPart([UriPartial]::Authority)
$ps1 = & $crlf ([IO.File]::ReadAllText((Join-Path $root "src\backend.ps1"), [Text.Encoding]::UTF8))
$ps1 = $ps1.Replace('[string]$AllowOrigin = "https://r3u-software.github.io"', '[string]$AllowOrigin = "' + $origin + '"')

$header = @'
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
set "MJ_URL=@@URL@@"
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

'@
$header = (& $crlf $header).Replace("@@URL@@", $PagesUrl)
$bat = $header + "::BEGIN backend.ps1`r`n" + $ps1 + "::END backend.ps1`r`n"
[IO.File]::WriteAllText((Join-Path $out "MJ_Dashboard_Backend.bat"), $bat, (New-Object Text.UTF8Encoding($false)))

$html = & $crlf ([IO.File]::ReadAllText((Join-Path $root "src\index.html"), [Text.Encoding]::UTF8))
[IO.File]::WriteAllText((Join-Path $out "index.html"), $html, (New-Object Text.UTF8Encoding($false)))
Write-Host "Built docs/ for $PagesUrl (allowed origin $origin)"
