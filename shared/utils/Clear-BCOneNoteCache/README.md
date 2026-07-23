# Clear-BCOneNoteCache

**Version:** 1.0.0  
**Author:** BriComp IT Consulting Services — [bricomp.com](https://bricomp.com)

Clears stale local cache folders for the Microsoft OneNote desktop app (M365 / Office 16.0) to resolve freezing, keystroke lag, and sync thrash issues. Notebooks stored in OneDrive or SharePoint are unaffected — OneNote re-syncs everything from the cloud on next launch.

---

## Requirements

- PowerShell 5.1 or later
- Microsoft OneNote desktop (M365 Apps / Office 2016 or later — `16.0` build)
- **Not applicable** to the OneNote UWP (Microsoft Store) app

No elevation is required when running in the user's own context. Elevation (or SYSTEM context via RMM) is required when targeting another user's profile with `-Username`.

---

## What Gets Cleared

| Folder | Purpose | Action |
|---|---|---|
| `cache\` | Primary sync and keystroke cache | ✅ Cleared |
| `FullTextSearchIndex\` | Local search index | ✅ Cleared |
| `MasterIndex\` | Notebook index | ✅ Cleared |
| `ServerListings\` | Cached notebook server/location data | ✅ Cleared |
| `Backup\` | Local section backups | 🚫 Preserved |

All four cleared folders are rebuilt automatically by OneNote on next launch. No notebook content is stored in these paths.

---

## Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `-Username` | String | No | Target a specific local user profile (e.g. `jsmith`). If omitted, auto-detects the logged-on user. |
| `-Relaunch` | Switch | No | Relaunch OneNote after the cache is cleared. Only effective in user context, not when running as SYSTEM. |
| `-WhatIf` | Switch | No | Dry run — shows what would be removed without making any changes. |

---

## Usage

### Basic — clear cache for the current user

```powershell
.\Clear-BCOneNoteCache.ps1
```

### Clear and relaunch OneNote

```powershell
.\Clear-BCOneNoteCache.ps1 -Relaunch
```

### Target a specific user profile (elevated / RMM)

```powershell
.\Clear-BCOneNoteCache.ps1 -Username jsmith
```

### Dry run first

```powershell
.\Clear-BCOneNoteCache.ps1 -WhatIf
```

---

## Running via RMM or SCCM

When executed as SYSTEM (e.g. via ScreenConnect elevated shell, SCCM, Intune remediation, or similar), the script automatically detects the interactively logged-on user via WMI and targets that user's profile.

If auto-detection is unreliable — for example on a multi-session host or when no user is interactively logged in — supply `-Username` explicitly:

```powershell
.\Clear-BCOneNoteCache.ps1 -Username jsmith
```

> **Note:** `-Relaunch` cannot inject a process into a user's desktop session when running as SYSTEM. If you use `-Relaunch` in that context, the script will warn and skip the launch step. Have the user relaunch OneNote manually, or handle relaunch through your RMM's user-context execution method.

---

## When to Use This Script

This script is the right call when OneNote exhibits any of the following:

- Keystroke lag — input is accepted, then OneNote freezes, then the keyboard buffer flushes and freezes again in a repeating cycle
- Freezing during or immediately after sync activity
- Slow launch or hang on open
- The `cache\` folder contains files with very old modification dates (months or years old), indicating orphaned entries from deleted, moved, or migrated notebooks

---

## What This Script Does Not Fix

- **Notebook corruption** — if freezing is isolated to a specific notebook or section, the issue may be in the `.one` file itself rather than the local cache
- **OneDrive sync errors** — if OneDrive/OneDrive for Business is in an error state, clearing the OneNote cache will not resolve it; fix the sync client first
- **Office installation issues** — if OneNote crashes immediately on launch after the cache is cleared, consider an Office Quick Repair (`winget upgrade --id Microsoft.Office` or via Control Panel)

---

## License

MIT License — see [LICENSE](LICENSE) for details.
