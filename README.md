# Desktop Organizers

Small Windows tools that keep a desktop tidy with as little thinking as possible.
Built for people (like me) who have ADHD/autism and want fixed rules instead of daily decisions.
Everything is local and nothing is ever deleted without a yes. Moves are logged and undoable.

| Tool | What it does |
|---|---|
| **Tidy desktop.ps1** | One click: numbered folders and shortcuts stay on the desktop, everything else goes into an Inbox. Never deletes or overwrites. |
| **inbox-sorter/** | "Sort my Inbox": suggests where each Inbox item belongs and moves it only when you say yes. Plain code first (rules, duplicate detection, text/name matching), then a local LLM (Ollama) only for the leftovers. Reads PDFs (incl. scanned, via Windows OCR), Office files, pictures, videos, zips and more. |
| **Hide Spotlight icon.ps1** | Hides Windows Spotlight's "Learn about this image" desktop icon, for your account or every account (`-AllUsers`, as Administrator). Uses Windows' own per-user `HideDesktopIcons` switch; takes effect instantly. |
| **Get-DesktopIcons.ps1** | Lists the icons Explorer is really showing on the desktop (handy for testing). |

## Setup
- Windows 10/11, PowerShell 5.1, Python 3.10+ (`pip install --user pypdf pillow` for PDF reading).
- Paths are set at the top of each script (`C:\Users\Public\Desktop`, Inbox `C:\Users\Public\Inbox`).
- Copy `inbox-sorter/config.example.json` to `config.json` and describe your own folders; the descriptions are what the local AI reads.

## Hide Spotlight icon
```powershell
& ".\Hide Spotlight icon.ps1"                 # this account, now
& ".\Hide Spotlight icon.ps1" -AllUsers       # every account (run as Administrator)
& ".\Hide Spotlight icon.ps1" -Status         # report only
& ".\Hide Spotlight icon.ps1" -Remove         # undo
```
The icon is `CLSID_DesktopSpotlight` (`{2cc5ca98-6485-489a-920e-b3e88a6ccce3}`). `-AllUsers` also writes the
setting into the settings file of accounts that aren't signed in and into the template for new accounts, and
adds an HKLM `Run` entry so each sign-in re-applies it.
