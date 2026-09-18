# Repair-BCDesktopConnectorSyncRoot

**Version:** 1.0.0  
**Author:** BriComp IT Consulting Services — [bricomp.com](https://bricomp.com)

Repairs Autodesk Desktop Connector "Unable to register drive" startup failures caused by orphaned Windows Cloud Files API sync root registrations — the failure mode left behind when a machine's user SID changes, such as an Entra-joined device migrated to on-premises Active Directory.

Creating a new Windows profile does **not** fix this. The stale registration lives in `HKLM` and is keyed by SID, not inside the user profile.

For the full root cause analysis, see [TROUBLESHOOTING.md](TROUBLESHOOTING.md).

---

## Symptoms

Desktop Connector starts, sits in the system tray for a few seconds, then disappears. Its log under `%LOCALAPPDATA%\Autodesk\Desktop Connector\Logs` shows:

```
Failed to register virtual file system for <name>
CloudFileApi.CloudFileApiException: Access is denied.
    (Exception from HRESULT: 0x80070005 E_ACCESSDENIED)
   at Windows.Storage.Provider.StorageProviderSyncRootManager.Register(...)

Tray Shutdown - ExitSource = "Unable to register drive"
```

Sign-in, token refresh and cloud API calls all succeed in the log. The failure is entirely local.

---

## Requirements

- PowerShell 5.1 or later
- Windows 10 or Windows 11
- Must run **elevated** — the repair modifies `HKLM`. This is enforced by `#Requires -RunAsAdministrator`, so PowerShell refuses to start the script in a non-elevated session.
- Autodesk Desktop Connector installed (the script exits cleanly if it is not)

---

## What It Does

| Step | Action |
|---|---|
| 1 | Confirms Desktop Connector is installed. Exits without changes if not. |
| 2 | Inventories every `DesktopConnector!*` sync root and classifies each one |
| 3 | Exits without changes if nothing qualifies for removal |
| 4 | Exports the entire `SyncRootManager` key to a `.reg` backup — aborts if the backup fails |
| 5 | Stops Desktop Connector |
| 6 | Removes the stale registrations and their leftover Explorer namespace entries |
| 7 | Restarts the machine after a two minute warning to the signed-in user |

### Classification

| Verdict | Meaning | Action |
|---|---|---|
| `ORPHANED` | The owning SID has no profile on this machine. This is the registration that blocks the live account. | 🗑️ Removed |
| `INCOMPLETE` | Live SID, but no claimed path — debris from a failed registration attempt. | 🗑️ Removed |
| `HEALTHY` | Live SID with a valid claimed path. | ✅ Left alone |

An `INCOMPLETE` registration is only removed when an `ORPHANED` one is also present. On its own it may simply mean Desktop Connector is mid-registration, which is not a fault.

Only keys beginning with `DesktopConnector!` are ever touched. OneDrive and every other cloud storage provider are left untouched.

---

## Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `-DryRun` | Switch | No | Report what would be removed and exit. No changes, no restart. |
| `-NoRestart` | Switch | No | Apply the repair but leave the restart to you. The repair does not take effect until the machine restarts. |
| `-RestartDelaySeconds` | Int | No | Seconds before the restart. Default `120`. |
| `-LogPath` | String | No | Override the log and backup directory. Default `C:\ProgramData\BriComp\DCSyncRootFix`. |

---

## Usage

### Preview first — always do this on a new environment

```powershell
.\Repair-BCDesktopConnectorSyncRoot.ps1 -DryRun
```

### Repair and restart

```powershell
.\Repair-BCDesktopConnectorSyncRoot.ps1
```

### Repair, leave the restart to your RMM

```powershell
.\Repair-BCDesktopConnectorSyncRoot.ps1 -NoRestart
```

### For non-technical operators

Keep `Run-Repair.cmd` in the same folder as the script. The operator right-clicks it and chooses **Run as administrator** — it self-elevates, runs the repair and reboots. Nothing to read, nothing to type.

```
Run-Repair.cmd              Repair and restart
Run-Repair.cmd /dryrun      Preview only
Run-Repair.cmd /norestart   Repair, no restart
```

> **Note:** the launcher calls PowerShell with `-ExecutionPolicy Bypass` so a freshly downloaded copy runs without prompting an operator who cannot answer a trust prompt. The script is Authenticode-signed, so if your environment prefers signature enforcement, change `Bypass` to `RemoteSigned` or `AllSigned` in the launcher.

---

## Running via RMM or SCCM

Call the script directly and act on the exit code:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "Repair-BCDesktopConnectorSyncRoot.ps1" -NoRestart
```

| Exit code | Meaning |
|---|---|
| `0` | No action needed — healthy, or Desktop Connector not installed |
| `1` | Error — collect the log |
| `2` | Not running elevated |
| `3` | Changes applied — restart pending or underway |

The script writes to `HKLM` and to every loaded user hive, so it works correctly when invoked as SYSTEM as well as in an elevated user session.

---

## Logs and Rollback

Both land in `C:\ProgramData\BriComp\DCSyncRootFix`:

```
DCSyncRootFix-<COMPUTER>-<timestamp>.log       what it found and what it did
SyncRootManager-<COMPUTER>-<timestamp>.reg     pre-change registry backup
```

To roll back, run the `.reg` backup for that machine and restart.

---

## After the Repair

Have the user sign in and launch Desktop Connector — it should stay in the tray. They will need to re-add their project subscriptions.

No cached file data is lost. On an affected machine there is none to lose, because Desktop Connector never successfully mounted the drive in the first place.

---

## What This Script Does Not Fix

- **Authentication or licensing failures** — if the log shows sign-in or token errors, the problem is not sync root registration
- **A corrupt Desktop Connector installation** — if registration still fails after the repair, uninstall, clear `%LOCALAPPDATA%\Autodesk\Desktop Connector`, and reinstall
- **Overlapping sync roots that are both legitimate** — if two live accounts genuinely claim the same workspace path, that needs a decision, not a script
- **Broken permissions on the workspace folder** — the Cloud Files API also returns Access Denied when the provider lacks `WRITE_DATA` or `WRITE_DAC` on the sync root folder. Check `icacls` on the workspace path if the repair does not resolve it.

---

## License

MIT License — see [LICENSE](LICENSE) for details.
