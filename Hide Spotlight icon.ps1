<#
HIDE SPOTLIGHT ICON
===================
Windows Spotlight (the rotating "Background" photos) puts a "Learn about this image" icon on the desktop.
This hides it. Windows has a built-in per-user switch for it (HideDesktopIcons, for the icon CLSID below);
nothing is deleted or uninstalled, and Spotlight itself keeps working.

  (no switches)  your own account: hide it now, and re-check every time you sign in
  -AllUsers      needs an Administrator PowerShell. Does the above for EVERY account on this PC, including
                 accounts that aren't signed in, plus the template for accounts created later, and
                 makes every sign-in re-apply it (HKLM "Run" entry), so it also covers anyone who
                 switches to Spotlight later.
  -Remove        undo (your account, or everyone with -AllUsers): shows the icon again, drops the sign-in check
  -Status        just report; changes nothing
  -DryRun        with -AllUsers: show what it would do; changes nothing

The setting is applied whether or not Spotlight is on right now: it has no effect on other backgrounds,
and this way the icon can't pop up the day someone switches to Spotlight.
#>
param([switch]$AllUsers, [switch]$Remove, [switch]$Status, [switch]$DryRun, [switch]$Quiet)
$ErrorActionPreference = 'Stop'

$Clsid   = '{2cc5ca98-6485-489a-920e-b3e88a6ccce3}'     # CLSID_DesktopSpotlight in shell32.dll
$Keys    = 'NewStartPanel', 'ClassicStartMenu'
$Sub     = 'Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons'
$RunKey  = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
$RunName = 'HideSpotlightIcon'
$Me      = $MyInvocation.MyCommand.Path

function Say($m, $c = 'Gray') { if (-not $Quiet) { Write-Host "  $m" -ForegroundColor $c } }

# ---- one user's registry hive (a "root" like HKCU:\ or Registry::HKEY_USERS\<SID>\)
function Set-HideFlag([string]$Root, [bool]$Hide) {
    foreach ($k in $Keys) {
        $path = "$Root$Sub\$k"
        if ($Hide) {
            if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
            New-ItemProperty -Path $path -Name $Clsid -Value 1 -PropertyType DWord -Force | Out-Null
        } elseif (Test-Path $path) {
            Remove-ItemProperty -Path $path -Name $Clsid -ErrorAction SilentlyContinue
        }
    }
}
function Get-HideFlag([string]$Root) {
    [bool](Get-ItemProperty -Path "$Root$Sub\NewStartPanel" -Name $Clsid -ErrorAction SilentlyContinue).$Clsid
}
function Test-Spotlight([string]$Root) {   # BackgroundType 3 = Windows Spotlight
    (Get-ItemProperty -Path "${Root}Software\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers" `
        -Name BackgroundType -ErrorAction SilentlyContinue).BackgroundType -eq 3
}
function Refresh-Desktop {                 # makes Explorer re-read the setting right away (no restart)
    try {
        Add-Type -Namespace HideSpot -Name Native -MemberDefinition `
            '[System.Runtime.InteropServices.DllImport("shell32.dll")] public static extern void SHChangeNotify(int e, int f, System.IntPtr a, System.IntPtr b);' -ErrorAction Stop
    } catch { }
    try { [HideSpot.Native]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero) } catch { }
}

# ---- your own account
if (-not $AllUsers) {
    if ($Status) {
        Say ("Icon hidden for this account: {0}   |   Spotlight is the background: {1}" -f (Get-HideFlag 'HKCU:\'), (Test-Spotlight 'HKCU:\'))
        return
    }
    Set-HideFlag 'HKCU:\' (-not $Remove)
    Refresh-Desktop
    Say ($(if ($Remove) { 'Spotlight icon will show again.' } else { 'Spotlight "Learn about this image" icon hidden.' })) Green
    return
}

# ---- everyone (needs admin)
$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $admin -and -not $DryRun -and -not $Status) {
    Say 'Doing this for every account changes PC-wide settings, so it has to run as Administrator.' Yellow
    Say 'Right-click Start > Windows PowerShell (Admin), then paste:' Yellow
    Say "& `"$Me`" -AllUsers" Cyan
    exit 2
}

$profiles = Get-CimInstance Win32_UserProfile | Where-Object { -not $_.Special -and $_.LocalPath -like 'C:\Users\*' }
$targets = @()
foreach ($p in $profiles) {
    $hive = Join-Path $p.LocalPath 'NTUSER.DAT'
    if (Test-Path -LiteralPath $hive -ErrorAction SilentlyContinue) { $targets += [pscustomobject]@{ Name = Split-Path $p.LocalPath -Leaf; Sid = $p.SID; Hive = $hive } }
}
$targets += [pscustomobject]@{ Name = '(template for new accounts)'; Sid = $null; Hive = 'C:\Users\Default\NTUSER.DAT' }

$done = 0
foreach ($t in $targets) {
    $loaded = $t.Sid -and (Test-Path "Registry::HKEY_USERS\$($t.Sid)")
    $how = if ($loaded) { 'signed in' } else { 'loading its settings file' }
    if ($Status) {
        if ($loaded) { Say ("{0,-30} hidden: {1}   spotlight: {2}" -f $t.Name, (Get-HideFlag "Registry::HKEY_USERS\$($t.Sid)\"), (Test-Spotlight "Registry::HKEY_USERS\$($t.Sid)\")) }
        else { Say ("{0,-30} (not signed in - checked at sign-in)" -f $t.Name) DarkGray }
        continue
    }
    if ($DryRun) { Say ("would {0} for {1} ({2})" -f $(if ($Remove) { 'show the icon again' } else { 'hide the icon' }), $t.Name, $how); continue }
    try {
        if ($loaded) {
            Set-HideFlag "Registry::HKEY_USERS\$($t.Sid)\" (-not $Remove)
        } else {
            $null = & reg.exe load HKU\HideSpotTmp $t.Hive 2>&1
            if ($LASTEXITCODE -ne 0) { throw 'settings file is in use or not readable' }
            try { Set-HideFlag 'Registry::HKEY_USERS\HideSpotTmp\' (-not $Remove) }
            finally { [gc]::Collect(); [gc]::WaitForPendingFinalizers(); $null = & reg.exe unload HKU\HideSpotTmp 2>&1 }
        }
        $done++; Say ("{0,-30} done ({1})" -f $t.Name, $how) Green
    } catch { Say ("{0,-30} SKIPPED: {1}" -f $t.Name, $_.Exception.Message) Yellow }
}
if ($Status) {
    Say ("Sign-in check installed: {0}" -f [bool](Get-ItemProperty $RunKey -Name $RunName -ErrorAction SilentlyContinue))
    return
}
if ($DryRun) { Say ("would {0} the sign-in check (HKLM Run: $RunName)" -f $(if ($Remove) { 'remove' } else { 'install' })); return }

if ($Remove) { Remove-ItemProperty $RunKey -Name $RunName -ErrorAction SilentlyContinue }
else { Set-ItemProperty $RunKey -Name $RunName -Value "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$Me`" -Quiet" }
Say ("{0} account setting(s) {1}; sign-in check {2}." -f $done, $(if ($Remove) { 'cleared' } else { 'applied' }), $(if ($Remove) { 'removed' } else { 'installed' })) Green
Refresh-Desktop
