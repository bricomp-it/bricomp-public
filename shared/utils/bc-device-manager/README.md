# BriComp Device Manager

**Version**: 3.0.0
**Author**: Brian R. Ricks, BriComp Computers, LLC — support@bricomp.com
**License**: MIT (see LICENSE)
**Requires**: PowerShell 7+ (`pwsh`)

---

## Overview

`Invoke-BCDeviceManager.ps1` is a WPF GUI tool for managing Windows devices
in bulk. Designed for IT admins and MSP consultants who need to run remote
actions across multiple endpoints quickly.

Key capabilities:
- **Pre-validation** — Check Ping, WinRM port, and Auth for all targets before taking action
- **BitLocker management** — Check status and bulk disable across selected devices
- **Remote reboot** — Force-restart selected computers with confirmation dialog
- **Flexible targeting** — Load from CSV or enter device names manually; DNS suffix auto-appended to short names
- **Smart selection** — Toolbar helpers to check/uncheck all, uncheck offline, select BitLocker-on devices
- **Right-click actions** — Remove individual devices, check/uncheck highlighted rows via context menu
- **WPF credential dialog** — Proper GUI credential prompt when not using current user context
- **Live logging** — Every action writes to a timestamped log file; "Show log" tails it live

---

## Quick Start

### Double-click from Explorer (recommended)

```
Double-click: BCDeviceManager.bat
```

The `.bat` file handles everything:
- Checks if PowerShell 7 is installed — offers winget install or download link if not
- Requests Administrator privileges automatically
- Launches the WPF GUI with no visible console window

### From command line

```powershell
# Basic launch
pwsh -File .\Invoke-BCDeviceManager.ps1

# Pre-load a CSV and set DNS suffix
pwsh -File .\Invoke-BCDeviceManager.ps1 -CsvPath .\computers.csv -DnsSuffix corp.local
```

---

## Files

| File | Purpose |
|------|---------|
| `Invoke-BCDeviceManager.ps1` | The application — WPF GUI, requires PS7 |
| `BCDeviceManager.bat` | Double-click launcher — PS7 detection, elevation, no-console launch |
| `sample-computers.csv` | Example CSV input format (short names or FQDNs) |
| `LICENSE` | MIT License |

---

## Parameters

| Parameter | Description | Default |
|-----------|-------------|---------|
| `-CsvPath` | CSV file to pre-load on startup | *(none)* |
| `-ComputerNameColumn` | Column name in CSV for device names | `COMPUTERNAME` |
| `-DnsSuffix` | DNS suffix appended to short names (e.g. `corp.local`) | *(none)* |
| `-LogRoot` | Directory for log output | Script directory |

---

## Using the GUI

### Adding devices

- Type a name in the **Enter device name...** box and press **Enter** or click **+**
- Click **Load CSV...** to browse for a CSV file
- CSV can contain short names (`SERVER01`) or FQDNs (`SERVER01.corp.local`)
- Set the **DNS Suffix** field to automatically expand short names — `SERVER01` becomes
  `SERVER01.corp.local` before any connection is attempted

### Credentials

- **Use current user** (checked) — uses your logged-in session, no prompt
- **Use current user** (unchecked) — shows a WPF credential dialog before running any action;
  pre-fills `DOMAIN\currentuser`, you only need to enter the password

### Actions (toolbar row 1)

| Button | What it does | Confirmation? |
|--------|-------------|---------------|
| Pre-validate | Tests Ping → WinRM port (5985) → WMI auth | No |
| BitLocker status | Queries BitLocker protection state via WMI | No |
| Disable BitLocker | Disables BitLocker on all encrypted volumes | Yes — lists devices |
| Reboot | Force-restarts selected devices | Yes — lists devices |
| Export CSV | Saves results grid to CSV file | Save dialog |
| Show log | Opens live-tailing log viewer in a new pwsh window | No |

### Selection helpers (toolbar row 2)

| Button | What it does |
|--------|-------------|
| All | Checks every device |
| None | Unchecks every device |
| Uncheck offline | Unchecks devices where Ping = OFFLINE — run after Pre-validate |
| Select BitLocker on | Checks only devices with BitLocker enabled — run after BitLocker status |

### Right-click context menu

Right-click any row (or multi-select with Ctrl/Shift then right-click) for:
- **Remove device** — removes highlighted devices from the list entirely
- **Check selected rows** — checks the checkbox for highlighted rows
- **Uncheck selected rows** — unchecks the checkbox for highlighted rows

### Status dots (left sidebar)

| Color | Meaning |
|-------|---------|
| Gray | Not yet validated |
| Green | Fully reachable — Ping OK, WinRM open, WMI auth OK |
| Yellow | Partially reachable — Ping OK but WinRM closed or auth failed |
| Red | Offline — no Ping response |

---

## Pre-validation results explained

| Ping | WSMan | Auth | Meaning |
|------|-------|------|---------|
| OK | OK | OK | Fully reachable, ready for all actions |
| OK | OK | FAIL | WinRM open but auth failed — check credentials or UAC policy |
| OK | CLOSED | — | Machine is up but WinRM not enabled — run `Enable-PSRemoting` on target |
| OFFLINE | — | — | No Ping — machine is off, unreachable, or firewall blocking ICMP |

**Common fixes:**

- **WSMan = CLOSED** on domain machines: GPO may be blocking WinRM. Check
  `Computer Configuration → Windows Settings → Security Settings → Windows Firewall`
  for rules blocking port 5985.
- **WSMan = CLOSED** on Entra-joined machines: WinRM is not enabled by default.
  Deploy an Intune remediation script running `Enable-PSRemoting -Force -SkipNetworkProfileCheck`.
- **Auth = FAIL** on on-prem servers: Try explicit domain admin credentials (uncheck
  "Use current user"). If that works, the issue is UAC token filtering — fix with:
  ```powershell
  Set-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' `
      -Name LocalAccountTokenFilterPolicy -Value 1 -Type DWord
  ```
- **Auth = FAIL** and using IP address: Switch to FQDN — Kerberos doesn't work with IPs.

---

## Typical workflows

### Validate and reboot only reachable machines
1. Load CSV → **Pre-validate**
2. Click **Uncheck offline** (deselects all OFFLINE machines automatically)
3. Click **Reboot** → confirm

### Find and disable BitLocker across a fleet
1. Load CSV → **Pre-validate** → **Uncheck offline**
2. Click **BitLocker status** (runs only on checked/reachable machines)
3. Click **Select BitLocker on** (selects only machines with BitLocker active)
4. Click **Disable BitLocker** → confirm

---

## Requirements

- **PowerShell 7+** — `BCDeviceManager.bat` installs automatically if missing
- **Run as Administrator** — `BCDeviceManager.bat` handles elevation automatically
- **WinRM enabled on targets** — run `Enable-PSRemoting -Force` on each remote device
- **Network access** — port 5985 (WinRM HTTP) must be reachable from your machine

---

## Logs

Every session creates a timestamped folder:

```
Logs_20260626_152526\
    BCDeviceManager.log
```

Log location shown in the status bar. Click **Show log** to open a live-tailing viewer.

---

## Change Log

### 3.0.0 — 2026-06-26
- Full WPF/XAML GUI (replaces console menu)
- Two-row toolbar: actions on row 1, selection helpers on row 2
- Right-click context menu: remove device, check/uncheck selected rows
- WPF credential dialog — works correctly under `-WindowStyle Hidden` launcher
- Sequential per-device execution with hard .NET timeouts (no blocking)
- WinRM check via `TcpClient` (3s timeout) instead of `Test-WSMan` (no timeout)
- WMI auth check via `ManagementScope` (5s timeout) instead of `Invoke-Command`
- `BCDeviceManager.bat` — single double-click entry point with PS7 detection,
  winget auto-install offer, and UAC elevation
- DNS suffix placeholder and descriptive label in sidebar
- Status counts in status bar (OK · warn · offline)
- PUBLISH: true — MIT licensed, community-shareable

### 2.0.0 — 2025-12-13
- PS7 parallel execution via `ForEach-Object -Parallel`
- Improved logging

### 1.0.0 — 2024
- Initial console menu release
