# pull-extensions.ps1
# One shortcut, all repos. Walks the whole list every run and pulls each one.
# Works for ANY git repo, not just Chrome extensions.
#
# Each repo has a Reload flag:
#   Reload = $true   -> it's a Chrome extension; reload Chrome after it updates
#   Reload = $false  -> plain repo; just pull, never touch Chrome
# Chrome is only reloaded/launched if a repo with Reload=$true actually updated.
#
# A repo may also have an optional OnUpdate script block. It runs ONLY when that
# repo actually changed (updated or freshly cloned), and receives the repo's
# local path as its one argument. Used here to rebuild the Word macro template
# from its text source after a pull. Failures in a hook are caught and reported;
# they never abort the rest of the run.
# Before any hook runs, a running Word is closed so the rebuild can overwrite the
# locked My_Macros.dotm: open documents are saved first via COM and Word is quit
# cleanly (no lost work), then any leftover windowless/orphaned WINWORD.EXE that a
# COM quit can't reach is force-closed.
#
# Reload behavior (only for Reload=$true repos that changed):
#   - Chrome open   -> briefly focuses Chrome, sends the Extensions Reloader
#                      keyboard command, then hands focus back. No tab is opened.
#   - Chrome closed -> launches Chrome (a fresh start loads updated files itself)
#
# Why a hotkey and not http://reload.extensions: on the current Manifest V3
# extension that URL trigger is unreliable - the background service worker sleeps
# after ~30s and (because MV3 can't block the request) just navigates to the dead
# URL, so it "frequently does nothing / just refreshes the page." A keyboard
# command reliably wakes the worker, so the reload actually fires every run.
#
# Requires the "Extensions Reloader" extension (by Arik W). ONE-TIME SETUP: at
# chrome://extensions/shortcuts set its "Reload all extensions in development"
# command to the combo in $ReloadHotkey below. Do NOT leave it on the default
# Alt+Shift+R - that collides with Chrome's built-in Reading Mode (which is why
# the old hotkey misfired).
# Note: manifest.json changes are NOT picked up by the reload - those need a
# real reload from chrome://extensions or a Chrome restart.

# ============================ CONFIG ============================
# Url = repo's git URL, Path = local folder, Reload = is it a Chrome extension?
$Repos = @(
    @{ Url = "https://github.com/zrcoderre-ux/pdf-viewer.git";   Path = "C:\Users\ZCoderre\Chrome Extensions\PDF Viewer"; Reload = $true }
    @{ Url = "https://github.com/zrcoderre-ux/Cross-Opener.git";   Path = "C:\Users\ZCoderre\Chrome Extensions\Cross-Opener"; Reload = $true }
    @{ Url = "https://github.com/zrcoderre-ux/PDF-Linker.git"; Path = "C:\Users\ZCoderre\Apps\PDF Linker"; Reload = $false }
    @{ Url = "https://github.com/zrcoderre-ux/workup-search.git"; Path = "C:\Users\ZCoderre\Apps\Workup Search"; Reload = $false }
    @{ URL="https://github.com/zrcoderre-ux/E-Court.git"; Path = "C:\Users\ZCoderre\Chrome Extensions\E-Court"; Reload = $true }
    @{ URL="https://github.com/zrcoderre-ux/Claude.git"; Path = "C:\Users\ZCoderre\Chrome Extensions\Claude"; Reload = $true }

    # This tool itself: C:\Users\ZCoderre\scripts is a clone of this repo, so the
    # pull keeps pull-extensions.ps1/.bat current. The update applies on the NEXT
    # run (PowerShell has already loaded the copy that's currently executing).
    # Reload=$false - it never touches Chrome.
    @{ Url = "https://github.com/zrcoderre-ux/Github-Refresh.git"; Path = "C:\Users\ZCoderre\scripts"; Reload = $false }

    # Word macro template. Reload=$false (it never touches Chrome). On update,
    # the OnUpdate hook rebuilds My_Macros.dotm from src and hot-swaps it into a
    # running Word. SET THE TWO VALUES BELOW: the repo URL and your local Path.
    @{
        Url    = "https://github.com/zrcoderre-ux/My-Macros.git"
        Path   = "C:\Users\ZCoderre\Apps\My Macros"
        Reload = $false
        OnUpdate = {
            param($RepoPath)
            $importer = Join-Path $RepoPath "build\Import-Macros.ps1"
            if (Test-Path $importer) { & $importer }
            else { Write-Warning "  [macros] build\Import-Macros.ps1 not found in $RepoPath" }
        }
    }

)

# Master switch for the Chrome step. $false = never reload/launch Chrome at all.
$ReloadExtensions = $true

# The keyboard shortcut you assigned to "Extensions Reloader" at
# chrome://extensions/shortcuts (see the one-time setup note at the top). Pick any
# free combo EXCEPT Alt+Shift+R (Chrome Reading Mode). If you change it in Chrome,
# change both values here to match.
#   SendKeys codes:  ^ = Ctrl   + = Shift   % = Alt   (letters are literal)
$ReloadHotkey      = "%+e"          # Alt+Shift+E
$ReloadHotkeyLabel = "Alt+Shift+E"  # human-readable, only shown in messages
# ================================================================

function Get-ChromePath {
    $candidates = @(
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "$env:LocalAppData\Google\Chrome\Application\chrome.exe"
    )
    $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
}

function Invoke-ExtensionReload {
    $proc = Get-Process chrome -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowTitle -ne '' } |
            Select-Object -First 1

    if ($proc) {
        # Chrome open: fire the Extensions Reloader keyboard command. Delivering a
        # command wakes the MV3 service worker, so the reload actually runs. We
        # briefly focus Chrome to send the keys, then hand focus back to whatever
        # window had it so this doesn't steal your place.
        try {
            Add-Type -AssemblyName System.Windows.Forms
            if (-not ('FgWin' -as [type])) {
                Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class FgWin {
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@
            }
            $prev = [FgWin]::GetForegroundWindow()

            $wshell = New-Object -ComObject WScript.Shell
            $null = $wshell.AppActivate($proc.Id)   # focus Chrome
            Start-Sleep -Milliseconds 500
            [System.Windows.Forms.SendKeys]::SendWait($ReloadHotkey)
            Start-Sleep -Milliseconds 200

            if ($prev -ne [IntPtr]::Zero) { [void][FgWin]::SetForegroundWindow($prev) }   # restore focus
            Write-Host "Sent reload hotkey ($ReloadHotkeyLabel) to Chrome." -ForegroundColor DarkGray
        } catch {
            Write-Host "Couldn't send the reload hotkey - reload manually ($ReloadHotkeyLabel or the toolbar button)." -ForegroundColor Yellow
        }
    } else {
        # Chrome closed: a fresh start already loads the updated files.
        $chrome = Get-ChromePath
        try {
            if ($chrome) { Start-Process $chrome } else { Start-Process "chrome" }
            Write-Host "Chrome was closed - launched it; the fresh start already loaded the update." -ForegroundColor DarkGray
        } catch {
            Write-Host "Couldn't launch Chrome - open it manually to pick up the update." -ForegroundColor Yellow
        }
    }
}

function Close-Word {
    # An open Word keeps My_Macros.dotm locked, which blocks the template rebuild.
    # First gracefully save + quit a running Word via COM so nothing is lost (you
    # usually close it yourself; this is the safety net). Then force-close any
    # leftover windowless / orphaned WINWORD.EXE that a COM quit can't reach.
    $acted = $false

    # 1) Save open documents, then quit, via COM (Windows PowerShell / .NET
    #    Framework only - the .bat launches powershell.exe, so this is fine).
    $word = $null
    try { $word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application") }
    catch { }   # GetActiveObject throws when no Word is running - nothing to do
    if ($word) {
        try {
            $word.DisplayAlerts = 0   # wdAlertsNone - never stop on a save dialog
            foreach ($doc in @($word.Documents)) {
                try {
                    if ($doc.Path -ne '') { $doc.Save() }   # save docs already on disk
                    else { $doc.Saved = $true }             # unnamed scratch doc: let Word quit
                } catch { Write-Host "  Couldn't save '$($doc.Name)': $_" -ForegroundColor Yellow }
            }
            $word.Quit()
            $acted = $true
            Write-Host "  Saved open documents and closed Word." -ForegroundColor DarkGray
        } catch {
            Write-Host "  Couldn't cleanly quit Word: $_" -ForegroundColor Yellow
        } finally {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
            $word = $null
        }
    }

    # 2) Force-close any windowless Word still lingering (orphans COM can't reach).
    $bg = Get-Process WINWORD -ErrorAction SilentlyContinue |
          Where-Object { $_.MainWindowHandle -eq 0 }
    foreach ($p in $bg) {
        try {
            Stop-Process -Id $p.Id -Force -ErrorAction Stop
            $acted = $true
            Write-Host "  Closed background Word (PID $($p.Id))." -ForegroundColor DarkGray
        } catch {
            Write-Host "  Couldn't close background Word (PID $($p.Id)): $_" -ForegroundColor Yellow
        }
    }

    if ($acted) { Start-Sleep -Milliseconds 500 }   # give Windows a moment to release the lock
}

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Host "ERROR: git is not installed or not on PATH." -ForegroundColor Red
    Write-Host "Install it from https://git-scm.com/download/win then re-run." -ForegroundColor Red
    exit 1
}

$updated = @(); $current = @(); $failed = @()
$reloadNeeded = $false
$postActions = @()   # queued OnUpdate hooks: @{ Name=...; Path=...; Action=... }

foreach ($repo in $Repos) {
    $url = $repo.Url; $path = $repo.Path; $name = Split-Path $path -Leaf
    # Missing Reload key defaults to $true (preserves old behavior).
    $wantsReload = if ($repo.ContainsKey('Reload')) { [bool]$repo.Reload } else { $true }
    Write-Host ""
    Write-Host "==> $name" -ForegroundColor Cyan

    try {
        if (Test-Path (Join-Path $path ".git")) {
            $before = git -C $path rev-parse HEAD 2>$null
            git -C $path pull --ff-only
            if ($LASTEXITCODE -ne 0) {
                $failed += $name
                Write-Host "FAILED - left untouched, moving on" -ForegroundColor Yellow
                continue
            }
            $after = git -C $path rev-parse HEAD 2>$null
            if ($before -ne $after) {
                $updated += $name; if ($wantsReload) { $reloadNeeded = $true }
                if ($repo.ContainsKey('OnUpdate')) { $postActions += @{ Name = $name; Path = $path; Action = $repo.OnUpdate } }
                Write-Host "UPDATED" -ForegroundColor Green
            } else {
                $current += $name; Write-Host "already current" -ForegroundColor DarkGray
            }
        } else {
            $parent = Split-Path $path -Parent
            if (-not (Test-Path $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
            git clone $url $path
            if ($LASTEXITCODE -eq 0) {
                $updated += $name; if ($wantsReload) { $reloadNeeded = $true }
                if ($repo.ContainsKey('OnUpdate')) { $postActions += @{ Name = $name; Path = $path; Action = $repo.OnUpdate } }
                Write-Host "CLONED" -ForegroundColor Green
            } else {
                $failed += $name; Write-Host "FAILED to clone" -ForegroundColor Yellow
            }
        }
    } catch {
        $failed += $name
        Write-Host "ERROR: $_" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "----------------------------------------"
Write-Host ("Updated: {0}   Already current: {1}   Failed: {2}" -f $updated.Count, $current.Count, $failed.Count)
if ($updated.Count) { Write-Host ("  Updated -> " + ($updated -join ", ")) -ForegroundColor Green }
if ($failed.Count)  { Write-Host ("  Failed  -> " + ($failed  -join ", ")) -ForegroundColor Yellow }

# Run any post-update hooks (e.g. rebuild the Word macro template). Each is
# isolated so a failure reports and moves on without aborting the run. Save + quit
# any running Word first so a locked My_Macros.dotm can't block the rebuild.
if ($postActions.Count) { Close-Word }
foreach ($pa in $postActions) {
    Write-Host ""
    Write-Host "==> post-update: $($pa.Name)" -ForegroundColor Cyan
    try { & $pa.Action $pa.Path }
    catch { Write-Host "  hook ERROR: $_" -ForegroundColor Yellow }
}

Write-Host ""
if ($ReloadExtensions -and $reloadNeeded) {
    Write-Host "Applying the update in Chrome (reload hotkey $ReloadHotkeyLabel) ..." -ForegroundColor Cyan
    Invoke-ExtensionReload
} elseif ($updated.Count -gt 0) {
    Write-Host "Updated - nothing needed a Chrome reload." -ForegroundColor DarkGray
} else {
    Write-Host "Nothing updated - Chrome left alone." -ForegroundColor DarkGray
}
