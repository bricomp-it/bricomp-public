# BriComp IT Consulting Services — Public Scripts

Public PowerShell scripts and tools from **BriComp IT Consulting Services**,
a Microsoft-focused IT consulting and managed services provider based in
Gilbert, Arizona.

These scripts are used daily in production MSP environments managing
Microsoft 365, Intune, SCCM/ConfigMgr, Teams, Azure, and Windows infrastructure.
We publish them here for the broader IT community.

---

## Available Tools

### [BriComp Device Manager](https://github.com/bricomp-it/bricomp-public/blob/main/shared/utils/bc-device-manager)

A WPF GUI tool for managing Windows devices in bulk. Designed for IT admins
and MSP consultants who need to run remote actions across multiple endpoints
quickly — without writing scripts for every task.

**Key features:**

- Pre-validate WinRM reachability (Ping / WinRM port / WMI auth) across all targets
- BitLocker status check and bulk disable
- Force reboot with confirmation dialog
- Load targets from CSV or enter manually — DNS suffix auto-appended to short names
- Smart selection helpers: uncheck offline, select BitLocker-on devices
- Right-click context menu to remove individual devices
- Live timestamped log with one-click tailing viewer
- WPF credential dialog — works correctly under hidden console launch
- Double-click `.bat` launcher handles PS7 detection, auto-install, and UAC elevation

**Requirements:** PowerShell 7+, Windows, run as Administrator

**[Full documentation and usage guide](https://github.com/bricomp-it/bricomp-public/blob/main/shared/utils/bc-device-manager/README.md)**

---

### [Get-BCUserSignInLogs](https://github.com/bricomp-it/bricomp-public/blob/main/shared/m365/Get-BCUserSignInLogs)

Queries Entra ID sign-in logs via Microsoft Graph (beta endpoint) for a
specified user and time window. Accepts local time input and converts to
UTC automatically, so you don't have to do the math during an investigation.

**Key features:**

- Uses `Get-MgBetaAuditLogSignIn` — the beta endpoint is required for
  `signInEventTypes` filtering on non-interactive sign-ins
- `-SuccessOnly` / `-FailuresOnly` switches (mutually exclusive) to narrow results
- CSV export, auto-named to the working directory by default
- Installs required Graph module at runtime if missing

**Requirements:** PowerShell 5.1+, Microsoft.Graph.Beta module, and a
Microsoft Graph permission that grants sign-in log read access (e.g.
`AuditLog.Read.All`)

**[Full documentation and usage guide](https://github.com/bricomp-it/bricomp-public/blob/main/shared/m365/Get-BCUserSignInLogs/README.md)**

---

### [Invoke-BCDirectoryObjectMapper](https://github.com/bricomp-it/bricomp-public/blob/main/shared/utils/Invoke-BCDirectoryObjectMapper)

Correlates and maps device and user Object IDs between on-premises Active
Directory and Entra ID. Built for hybrid-join environments where tracing a
single identity across both directories otherwise means several manual
lookups.

**Key features:**

- Devices mode and Users mode
- Source/Target AD + Entra lookups
- Confirmation prompt before clearing results on mode switch
- Prompts to install RSAT if the AD module isn't present

**Requirements:** RSAT Active Directory PowerShell module (prompts to
install if missing), plus a Microsoft Graph or Azure AD PowerShell module
for the Entra ID side

**[Full documentation and usage guide](https://github.com/bricomp-it/bricomp-public/blob/main/shared/utils/Invoke-BCDirectoryObjectMapper/README.md)**

---

### [Clear-BCOneNoteCache](https://github.com/bricomp-it/bricomp-public/blob/main/shared/utils/Clear-BCOneNoteCache)

Clears stale local cache folders for the Microsoft OneNote desktop app to
resolve freezing, keystroke lag, and sync thrash issues. Notebooks stored
in OneDrive or SharePoint are unaffected — OneNote re-syncs everything from
the cloud on next launch.

**Key features:**

- Clears `cache\`, `FullTextSearchIndex\`, `MasterIndex\`, and `ServerListings\`
  — all rebuilt automatically by OneNote on next launch
- Preserves `Backup\` (local section backups)
- Auto-detects the logged-on user when run as SYSTEM (RMM/SCCM/Intune),
  or accepts an explicit `-Username`
- `-Relaunch` to reopen OneNote after clearing; `-WhatIf` for a dry run

**Requirements:** PowerShell 5.1+, Microsoft OneNote desktop (M365 Apps /
Office 2016+, build `16.0`) — not applicable to the OneNote UWP (Store) app

**[Full documentation and usage guide](https://github.com/bricomp-it/bricomp-public/blob/main/shared/utils/Clear-BCOneNoteCache/README.md)**

---

### [Export-BCUserGroupMembership](https://github.com/bricomp-it/bricomp-public/blob/main/shared/ad/Export-BCUserGroupMembership)

Exports Active Directory user accounts and group memberships to CSV. A
scriptable, parameterized replacement for the DumpSec "Dump Users as Table"
function — useful for access control audits, compliance reporting, and
baseline documentation.

**Key features:**

- **PerRow** format (default) — one row per group membership, ideal for
  filtering and pivot analysis in Excel
- **Joined** format — one row per user with all group names delimited by
  semicolons, matching the classic DumpSec layout
- `-SecurityGroupsOnly` switch to exclude distribution lists from output
- Group name cache built before user enumeration — significantly faster
  than per-membership AD lookups on large domains
- Accounts with no expiration date report as `Never` rather than blank

**Requirements:** PowerShell 5.1+, Active Directory module
(`RSAT-AD-PowerShell`), read access to the target domain

**[Full documentation and usage guide](https://github.com/bricomp-it/bricomp-public/blob/main/shared/ad/Export-BCUserGroupMembership/README.md)**

---

### [Add-BCMailboxDomainAlias](https://github.com/bricomp-it/bricomp-public/blob/main/shared/exchange/Add-BCMailboxDomainAlias)

Adds a domain-based SMTP alias to Exchange recipients that have automatic
email address policy management turned off, built from each recipient's
own Exchange Alias attribute. Checks both on-premises mailboxes and hybrid
remote mailboxes by default, so nothing gets missed in a hybrid deployment.

**Key features:**

- Checks both `Get-Mailbox` (on-premises) and `Get-RemoteMailbox` (hybrid)
  — skips a source with a warning instead of failing if a cmdlet isn't available
- Idempotent — safe to re-run; recipients that already have the target
  address are skipped and logged as `NoChangeNeeded`
- Full `-WhatIf`/`-Confirm` support, with a summary that reconciles
  (Attempted = Changed + Would change + No change needed + Failed)
- Per-recipient error handling — one failure doesn't stop the batch
- Transcript logging on every run, including dry runs
- `-PassThru` for exporting full per-recipient results to CSV

**Requirements:** Exchange Management Shell session (on-premises Exchange
Server, including Exchange Server SE); at least one of `Get-Mailbox` or
`Get-RemoteMailbox` must be available

**[Full documentation and usage guide](https://github.com/bricomp-it/bricomp-public/blob/main/shared/exchange/Add-BCMailboxDomainAlias/README.md)**

---

## Usage

### Download a script

Click any `.ps1` file, then click the **Raw** button and save, or clone the repo:

```
git clone https://github.com/bricomp-it/bricomp-public.git
```

### Run BriComp Device Manager

1. Download the `shared/utils/bc-device-manager/` folder
2. Double-click `BCDeviceManager.bat`
3. The launcher checks for PowerShell 7, handles elevation, and opens the GUI

Or from a PS7 prompt:

```
pwsh -File .\Invoke-BCDeviceManager.ps1
pwsh -File .\Invoke-BCDeviceManager.ps1 -CsvPath .\computers.csv -DnsSuffix corp.local
```

---

## Script Signing

All published scripts are Authenticode-signed with the BriComp Computers, LLC
code signing certificate (DigiCert Trusted G4, expires 2028-08-10).

To verify a signature before running:

```
Get-AuthenticodeSignature .\Invoke-BCDeviceManager.ps1 | Select-Object Status, SignerCertificate
```

Expected output: `Status: Valid` signed by `CN="BriComp Computers, LLC"`.

---

## License

All scripts in this repository are released under the **MIT License** — see [LICENSE](https://github.com/bricomp-it/bricomp-public/blob/main/LICENSE) for details. You are free to use, modify, and distribute
these scripts in your own environments.

---

## About BriComp IT Consulting Services

BriComp IT Consulting Services is a Microsoft-focused MSP and IT consulting
firm based in Gilbert, Arizona. We specialize in:

- Microsoft 365 and Azure administration
- SCCM / ConfigMgr and Intune endpoint management
- Teams voice (Direct Routing, E911, Call Queues)
- PKI and certificate services
- Palo Alto / Panorama network security
- PowerShell automation and tooling

**Website**: [bricomp.com](https://bricomp.com) **Contact**: <support@bricomp.com> **GitHub**: [github.com/bricomp-it](https://github.com/bricomp-it)

---

*Scripts are provided as-is. Always test in a non-production environment first.*
