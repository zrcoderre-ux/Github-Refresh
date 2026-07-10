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
#
# Reload behavior (only for Reload=$true repos that changed):
#   - Chrome open   -> sends Extensions Reloader hotkey (Alt+Shift+R), no tab
#   - Chrome closed -> launches Chrome (a fresh start loads updated files itself)
#
# Requires the "Extensions Reloader" extension (by Arik W) with its command set
# to Alt+Shift+R (default - verify at chrome://extensions/shortcuts).
# Note: manifest.json changes are NOT picked up by the reload - those need a
# real reload from chrome://extensions or a Chrome restart.

# ============================ CONFIG ============================
# Url = repo's git URL, Path = local folder, Reload = is it a Chrome extension?
$Repos = @(
    @{ Url = "https://github.com/zrcoderre-ux/pdf-viewer.git";   Path = "C:\Users\ZCoderre\Chrome Extensions\PDF Viewer"; Reload = $true }
    @{ Url = "https://github.com/zrcoderre-ux/Cross-Opener.git";   Path = "C:\Users\ZCoderre\Chrome Extensions\Cross-Opener"; Reload = $true }
    @{ Url = "https://github.com/zrcoderre-ux/PDF-Linker.git"; Path = "C:\Users\ZCoderre\Apps\PDF Linker"; Reload = $false }
    @{ Url = "https://github.com/zrcoderre-ux/PDF-Redactor.git"; Path = "C:\Users\ZCoderre\Apps\PDF Redactor"; Reload = $false }
    @{ Url = "https://github.com/zrcoderre-ux/workup-search.git"; Path = "C:\Users\ZCoderre\Apps\Workup Search"; Reload = $false }
    @{ URL="https://github.com/zrcoderre-ux/E-Court.git"; Path = "C:\Users\ZCoderre\Chrome Extensions\E-Court"; Reload = $true }
    @{ URL="https://github.com/zrcoderre-ux/Claude.git"; Path = "C:\Users\ZCoderre\Chrome Extensions\Claude"; Reload = $true }

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
        try {
            Add-Type -AssemblyName System.Windows.Forms
            $wshell = New-Object -ComObject WScript.Shell
            $null = $wshell.AppActivate($proc.Id)   # focus Chrome
            Start-Sleep -Milliseconds 600
            [System.Windows.Forms.SendKeys]::SendWait("%+r")   # Alt+Shift+R
            Write-Host "Sent reload hotkey to Chrome." -ForegroundColor DarkGray
        } catch {
            Write-Host "Couldn't send the reload hotkey - reload manually (Alt+Shift+R or the toolbar button)." -ForegroundColor Yellow
        }
    } else {
        $chrome = Get-ChromePath
        try {
            if ($chrome) { Start-Process $chrome } else { Start-Process "chrome" }
            Write-Host "Chrome was closed - launched it; the fresh start already loaded the update." -ForegroundColor DarkGray
        } catch {
            Write-Host "Couldn't launch Chrome - open it manually to pick up the update." -ForegroundColor Yellow
        }
    }
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
# isolated so a failure reports and moves on without aborting the run.
foreach ($pa in $postActions) {
    Write-Host ""
    Write-Host "==> post-update: $($pa.Name)" -ForegroundColor Cyan
    try { & $pa.Action $pa.Path }
    catch { Write-Host "  hook ERROR: $_" -ForegroundColor Yellow }
}

Write-Host ""
if ($ReloadExtensions -and $reloadNeeded) {
    Write-Host "Applying the update in Chrome (don't grab focus for ~1s) ..." -ForegroundColor Cyan
    Invoke-ExtensionReload
} elseif ($updated.Count -gt 0) {
    Write-Host "Updated - nothing needed a Chrome reload." -ForegroundColor DarkGray
} else {
    Write-Host "Nothing updated - Chrome left alone." -ForegroundColor DarkGray
}
