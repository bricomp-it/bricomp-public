# Invoke-BCDirectoryObjectMapper

**Version:** 2.1.1
**Author:** BriComp IT Consulting Services — [bricomp.com](https://bricomp.com)

Correlates and maps device and user Object IDs between on-premises Active
Directory and Entra ID. Built for hybrid-join environments, where tracing a
single identity or device across both directories otherwise means several
separate manual lookups.

---

## Key Features

- **Devices mode and Users mode** — switch between correlating device
  objects or user objects
- **Source/Target AD + Entra lookups** — resolve an object's identifiers
  across both directories in either direction
- **Confirmation prompt before clearing results** when switching modes, so
  in-progress lookups aren't lost by accident
- Prompts to install RSAT if the Active Directory PowerShell module isn't
  already present

---

## Requirements

> **Please verify before relying on this section** — requirements below are
> inferred from the script's behavior (RSAT install prompt, dual AD/Entra
> lookups) rather than confirmed against the current script header. Update
> this section with the exact PowerShell version and module list before
> treating it as authoritative.

- PowerShell 5.1 or later (not yet confirmed against script requirements)
- RSAT `ActiveDirectory` PowerShell module, for on-premises AD lookups
  (script will prompt to install if missing)
- A Microsoft Graph or AzureAD PowerShell module, for the Entra ID side —
  exact module and required permissions not yet confirmed

---

## Usage

```powershell
.\Invoke-BCDirectoryObjectMapper.ps1
```

Run from an elevated PowerShell session if RSAT installation is needed.
