# BriComp IT Consulting Services — Public Scripts

Public PowerShell scripts and tools from **BriComp IT Consulting Services**,
a Microsoft-focused IT consulting and managed services provider based in
Gilbert, Arizona.

These scripts are used daily in production MSP environments managing
Microsoft 365, Intune, SCCM/ConfigMgr, Teams, Azure, and Windows infrastructure.
We publish them here for the broader IT community.

---

## Available Tools

### [BriComp Device Manager](shared/utils/bc-device-manager/)

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

**[Full documentation and usage guide](shared/utils/bc-device-manager/README.md)**

---

### [Get-BCUserSignInLogs](Get-BCUserSignInLogs/)

A PowerShell script for pulling Entra ID sign-in logs for a specific user within
a local time window via Microsoft Graph. Designed for incident response, access
reviews, and user activity investigations in managed Microsoft 365 environments.

**Key features:**
- Accepts local time input — UTC conversion handled automatically
- Captures both **interactive** and **non-interactive** sign-ins (token refreshes,
  silent auth) via the Microsoft Graph Beta endpoint
- Auto-installs required Graph modules if not present
- Exports results to CSV — auto-named by UPN and timestamp if no path specified
- `-SuccessOnly` to show successful sign-ins only; `-FailuresOnly` to show failed sign-ins only
- Output includes sign-in type, app, client, IP, location, status, device, CA policies,
  and correlation ID

**Requirements:** PowerShell 7+, Entra ID P1/P2, Security Reader or Reports Reader role

**[Full documentation and usage guide](Get-BCUserSignInLogs/README.md)**

---

## Usage

### Download a script

Click any `.ps1` file, then click the **Raw** button and save, or clone the repo:

```powershell
git clone https://github.com/bricomp-it/bricomp-public.git
```

### Run BriComp Device Manager

1. Download the `shared/utils/bc-device-manager/` folder
2. Double-click `BCDeviceManager.bat`
3. The launcher checks for PowerShell 7, handles elevation, and opens the GUI

Or from a PS7 prompt:

```powershell
pwsh -File .\Invoke-BCDeviceManager.ps1
pwsh -File .\Invoke-BCDeviceManager.ps1 -CsvPath .\computers.csv -DnsSuffix corp.local
```

### Run Get-BCUserSignInLogs

```powershell
# All sign-ins (interactive + non-interactive) with auto-named CSV
.\Get-BCUserSignInLogs.ps1 -UserPrincipalName jsmith@contoso.com `
    -LocalStartTime "2026-06-25 08:00" -LocalEndTime "2026-06-25 18:00"

# Successful sign-ins only, exported to a specific path
.\Get-BCUserSignInLogs.ps1 -UserPrincipalName jsmith@contoso.com `
    -LocalStartTime "2026-06-25 08:00" -LocalEndTime "2026-06-25 18:00" `
    -SuccessOnly -CsvPath C:\Temp\jsmith_signins.csv

# Failed sign-ins only
.\Get-BCUserSignInLogs.ps1 -UserPrincipalName jsmith@contoso.com `
    -LocalStartTime "2026-06-25 08:00" -LocalEndTime "2026-06-25 18:00" `
    -FailuresOnly
```

---

## Script Signing

All published scripts are Authenticode-signed with the BriComp Computers, LLC
code signing certificate (DigiCert Trusted G4, expires 2028-08-10).

To verify a signature before running:

```powershell
Get-AuthenticodeSignature .\Invoke-BCDeviceManager.ps1 | Select-Object Status, SignerCertificate
```

Expected output: `Status: Valid` signed by `CN="BriComp Computers, LLC"`.

---

## License

All scripts in this repository are released under the **MIT License** — see
[LICENSE](LICENSE) for details. You are free to use, modify, and distribute
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

**Website**: [bricomp.com](https://bricomp.com)
**Contact**: support@bricomp.com
**GitHub**: [github.com/bricomp-it](https://github.com/bricomp-it)

---

*Scripts are provided as-is. Always test in a non-production environment first.*
