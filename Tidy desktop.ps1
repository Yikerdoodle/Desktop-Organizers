# TIDY DESKTOP - one click puts the desktop back in order.
#
# The rule (always the same, no surprises):
#   STAYS on the desktop : numbered folders ("1 - Inbox", "7 - Whatever") and shortcuts (.lnk / .url)
#   MOVES into the Inbox  : every other file or folder (Inbox = C:\Users\Public\Inbox, "1 - Inbox" shortcut)
# Nothing is ever deleted. Every move is written to tidy-log.csv.
$ErrorActionPreference = 'Stop'
$D     = 'C:\Users\Public\Desktop'
$Sys   = 'C:\Users\Public\Documents\Desktop system'
$Inbox = 'C:\Users\Public\Inbox'   # the desktop has a "1 - Inbox" shortcut pointing here
$Log   = Join-Path $Sys 'tidy-log.csv'
$Guide = Join-Path $Sys 'Desktop guide.txt'

$Host.UI.RawUI.WindowTitle = 'Tidy desktop'
# Keep Windows Spotlight's "Learn about this image" icon off the desktop (see "Hide Spotlight icon.ps1")
try { & (Join-Path $Sys 'Hide Spotlight icon.ps1') -Quiet } catch { }
New-Item -ItemType Directory -Path $Inbox -Force | Out-Null
if (-not (Test-Path -LiteralPath $Log)) { 'When,From,To' | Set-Content -LiteralPath $Log -Encoding UTF8 }

$moved = @(); $skipped = @()
foreach ($item in @(Get-ChildItem -LiteralPath $D -Force)) {
    $a = $item.Attributes
    if (($a -band [IO.FileAttributes]::Hidden) -or ($a -band [IO.FileAttributes]::System)) { continue }
    if ($item.PSIsContainer -and $item.Name -match '^\d+ - ') { continue }
    if (-not $item.PSIsContainer -and $item.Extension -in '.lnk', '.url') { continue }

    # Never overwrite: "notes.txt" -> "notes (2).txt" if the Inbox already has one
    $base = if ($item.PSIsContainer) { $item.Name } else { $item.BaseName }
    $ext  = if ($item.PSIsContainer) { '' } else { $item.Extension }
    $dest = Join-Path $Inbox $item.Name
    for ($i = 2; Test-Path -LiteralPath $dest; $i++) { $dest = Join-Path $Inbox "$base ($i)$ext" }

    try {
        Move-Item -LiteralPath $item.FullName -Destination $dest
        ('{0},"{1}","{2}"' -f (Get-Date -Format 'yyyy-MM-dd HH:mm'), $item.FullName, $dest) | Add-Content -LiteralPath $Log -Encoding UTF8
        $moved += $item.Name
    } catch {
        $skipped += $item.Name
    }
}

Write-Host ''
if ($moved.Count -eq 0 -and $skipped.Count -eq 0) {
    Write-Host '  Desktop is already tidy. Nice.' -ForegroundColor Green
} else {
    if ($moved.Count) {
        Write-Host "  Moved $($moved.Count) thing(s) into 1 - Inbox:" -ForegroundColor Green
        $moved | ForEach-Object { Write-Host "    - $_" }
    }
    if ($skipped.Count) {
        Write-Host "`n  Left alone (probably open in an app - close it and click Tidy again):" -ForegroundColor Yellow
        $skipped | ForEach-Object { Write-Host "    - $_" }
    }
}

# Inbox check-in: gentle, not naggy
$inboxItems = @(Get-ChildItem -LiteralPath $Inbox -Force | Where-Object { -not ($_.Attributes -band [IO.FileAttributes]::Hidden) })
Write-Host ''
if ($inboxItems.Count -eq 0) {
    Write-Host '  Inbox: empty.' -ForegroundColor Green
} else {
    Write-Host "  Inbox: $($inboxItems.Count) item(s)." -NoNewline
    if ($inboxItems.Count -ge 10) {
        Write-Host '  Worth a 5-minute sort soon (guide has the steps).' -ForegroundColor Yellow
    } else { Write-Host '' }
}

Write-Host @'

  WHERE THINGS LIVE
    1 - Inbox ............ new stuff lands here, sort it weekly
    2-5 .................. active projects + personal
    6 - Games ............ game shortcuts + Switch screenshots link
    9 - Archive .......... finished stuff (drag it here when done)
    Switch captures ...... C:\Users\Public\Pictures\Nintendo Switch

'@ -ForegroundColor DarkGray

if ($inboxItems.Count -gt 0) {
    Write-Host '  Press S to sort the Inbox now, G for the guide, any other key to close (auto-closes in 30s).' -ForegroundColor Cyan
} else {
    Write-Host '  Press G to open the full guide, or any other key to close (auto-closes in 30s).' -ForegroundColor Cyan
}
$deadline = (Get-Date).AddSeconds(30)
try {
    while ((Get-Date) -lt $deadline) {
        if ([Console]::KeyAvailable) {
            $k = [Console]::ReadKey($true)
            if ($k.Key -eq 'G') { Start-Process notepad.exe -ArgumentList "`"$Guide`"" }
            if ($k.Key -eq 'S') { Start-Process "$env:SystemRoot\py.exe" -ArgumentList "-3 `"$Sys\inbox-sorter\inbox_sorter.py`"" }
            break
        }
        Start-Sleep -Milliseconds 200
    }
} catch { }  # no interactive console (e.g. run from a script) - just finish
