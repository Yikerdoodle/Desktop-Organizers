<#
INSTALL - Desktop Organizers
============================
Right-click > "Run with PowerShell", or:  powershell -ExecutionPolicy Bypass -File .\Install.ps1

Asks before each optional step (press Enter to accept the [Y] default). Nothing is deleted.

  -Yes           accept every default without asking (all steps except "all accounts", see below)
  -AllAccounts   also run the Spotlight-icon hider for every account on the PC. Windows will show its
                 Administrator prompt (UAC); nothing PC-wide happens unless you approve it
  -DryRun        show what would happen, change nothing
  -Uninstall     remove the shortcuts and the sign-in check (your files, Inbox and logs are kept)
  -Target <dir>  where the tools are installed (default below)
#>
param([switch]$Yes, [switch]$AllAccounts, [switch]$DryRun, [switch]$Uninstall,
      [string]$Target = 'C:\Users\Public\Documents\Desktop system')
$ErrorActionPreference = 'Stop'

$Repo    = $PSScriptRoot
$Desktop = [Environment]::GetFolderPath('Desktop')
$Inbox   = 'C:\Users\Public\Inbox'
$Hider   = Join-Path $Target 'Hide Spotlight icon.ps1'
$PS      = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"

function Say($m, $c = 'Gray') { Write-Host "  $m" -ForegroundColor $c }
function Ask([string]$q, [bool]$default = $true) {
    if ($Yes) { return $default }
    $hint = if ($default) { 'Y/n' } else { 'y/N' }
    $a = Read-Host "  $q [$hint]"
    if ([string]::IsNullOrWhiteSpace($a)) { return $default }
    return $a.Trim().ToLower().StartsWith('y')
}
function Do-It([string]$what, [scriptblock]$action) {
    if ($DryRun) { Say "would: $what" DarkGray; return }
    & $action; Say $what Green
}
function New-Shortcut([string]$path, [string]$exe, [string]$args_, [string]$icon, [string]$desc) {
    $s = (New-Object -ComObject WScript.Shell).CreateShortcut($path)
    $s.TargetPath = $exe; $s.Arguments = $args_; $s.WorkingDirectory = $Target
    $s.IconLocation = $icon; $s.Description = $desc; $s.Save()
}
# Runs the hider for every account through Windows' own Administrator prompt (UAC)
function Run-Hider-AllAccounts([string]$extra = '') {
    $inner = "& '$Hider' -AllUsers $extra; Read-Host 'Done. Press Enter to close'"
    Start-Process $PS -Verb RunAs -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', $inner
}

Write-Host "`n  DESKTOP ORGANIZERS $(if ($Uninstall) { '- UNINSTALL' } else { '- INSTALL' })$(if ($DryRun) { ' (dry run: nothing is changed)' })`n" -ForegroundColor Cyan

# ============================================================ uninstall
if ($Uninstall) {
    if (Ask 'Remove the "Tidy desktop" and "Sort my Inbox" desktop shortcuts?') {
        foreach ($n in 'Tidy desktop.lnk', 'Sort my Inbox.lnk') {
            $p = Join-Path $Desktop $n
            if (Test-Path -LiteralPath $p) { Do-It "removed $n" { Remove-Item -LiteralPath $p } }
        }
    }
    if (Ask 'Show the Spotlight "Learn about this image" icon again for this account?') {
        if (Test-Path $Hider) { Do-It 'icon setting cleared for this account' { & $Hider -Remove -Quiet } }
    }
    if ($AllAccounts -or (Ask 'Also undo it for EVERY account (asks Windows for Administrator permission)?' $false)) {
        if (Test-Path $Hider) { Do-It 'asked Windows to undo it for every account' { Run-Hider-AllAccounts '-Remove' } }
    }
    Say "`nYour files, the Inbox and the logs were left alone. Delete '$Target' yourself if you want them gone." Yellow
    return
}

# ============================================================ install
# 1. files
$resolvedTarget = Resolve-Path -LiteralPath $Target -ErrorAction SilentlyContinue
if (-not $resolvedTarget -or (Resolve-Path $Repo).Path -ne $resolvedTarget.Path) {
    Do-It "copied the tools into $Target" {
        New-Item -ItemType Directory -Force "$Target\inbox-sorter" | Out-Null
        foreach ($f in 'Tidy desktop.ps1', 'Hide Spotlight icon.ps1', 'Get-DesktopIcons.ps1') { Copy-Item -LiteralPath "$Repo\$f" $Target -Force }
        foreach ($f in 'inbox_sorter.py', 'facts.py', 'ocr-helper.ps1', 'config.example.json') { Copy-Item -LiteralPath "$Repo\inbox-sorter\$f" "$Target\inbox-sorter" -Force }
    }
} else { Say "tools are already in $Target" DarkGray }
# your own settings file is never overwritten
if (-not (Test-Path "$Target\inbox-sorter\config.json") -and (Test-Path "$Target\inbox-sorter\config.example.json")) {
    Do-It 'created inbox-sorter\config.json from the example (edit it to describe your folders)' {
        Copy-Item "$Target\inbox-sorter\config.example.json" "$Target\inbox-sorter\config.json" }
}

# 2. Inbox + shortcuts
if (Ask "Create the Inbox folder ($Inbox)?") {
    Do-It "Inbox ready: $Inbox" { New-Item -ItemType Directory -Force $Inbox | Out-Null }
}
if (Ask 'Put "Tidy desktop" and "Sort my Inbox" shortcuts on your desktop?') {
    Do-It 'desktop shortcuts created' {
        New-Shortcut "$Desktop\Tidy desktop.lnk" $PS "-NoProfile -ExecutionPolicy Bypass -File `"$Target\Tidy desktop.ps1`"" `
            "$env:SystemRoot\System32\cleanmgr.exe,0" 'Moves loose files on the desktop into the Inbox. Never deletes anything.'
        New-Shortcut "$Desktop\Sort my Inbox.lnk" "$env:SystemRoot\py.exe" "-3 `"$Target\inbox-sorter\inbox_sorter.py`"" `
            "$env:SystemRoot\System32\imageres.dll,157" 'Suggests where each thing in the Inbox should go. Moves nothing until you say yes.'
    }
}

# 3. Python add-ons for reading PDFs / pictures
if (Get-Command py -ErrorAction SilentlyContinue) {
    if (Ask 'Install the optional Python add-ons that let Sort my Inbox read PDF pages and pictures (pypdf, pillow)?') {
        Do-It 'pypdf and pillow installed (for your account only)' { & py -3 -m pip install --user --disable-pip-version-check pypdf pillow | Out-Null }
    }
} else { Say 'Python 3 was not found, so Sort my Inbox is not usable yet (install it from python.org, then run this again).' Yellow }

# 4. Windows Spotlight "Learn about this image" desktop icon
if (Ask 'Hide Windows Spotlight''s "Learn about this image" desktop icon for THIS account?') {
    Do-It 'Spotlight icon hidden for this account' { & $Hider -Quiet }
}
if ($AllAccounts -or (Ask 'Also hide it for EVERY account on this PC, including new ones? (Windows will ask for Administrator permission)' $false)) {
    Do-It 'asked Windows for Administrator permission to cover every account (approve the prompt)' { Run-Hider-AllAccounts }
}

Say "`nDone. Double-click ""Tidy desktop"" any time. To uninstall: .\Install.ps1 -Uninstall" Cyan
