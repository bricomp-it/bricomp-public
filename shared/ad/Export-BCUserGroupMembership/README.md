# Export-BCUserGroupMembership

PowerShell script for exporting Active Directory user accounts and group memberships
to CSV. Intended as a scriptable, parameterized replacement for the DumpSec
"Dump Users as Table" function.

## Usage

```powershell
# Per-row format, all groups (default)
.\Export-BCUserGroupMembership.ps1 -Server dc01.contoso.com

# Per-row format, security groups only
.\Export-BCUserGroupMembership.ps1 -Server dc01.contoso.com -SecurityGroupsOnly

# Joined format (one row per user, groups delimited by semicolons)
.\Export-BCUserGroupMembership.ps1 -Server dc01.contoso.com -OutputFormat Joined

# Joined format, security groups only, custom output path
.\Export-BCUserGroupMembership.ps1 -Server dc01.contoso.com -OutputFormat Joined -SecurityGroupsOnly -OutputPath C:\Audit\Users.csv
```

## Parameters

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| Server | String | Yes | | FQDN or hostname of the domain controller to query |
| OutputPath | String | No | .\DumpUsers_YYYYMMDD.csv | Full path for the output CSV |
| OutputFormat | String | No | PerRow | PerRow or Joined (see below) |
| SecurityGroupsOnly | Switch | No | | Restrict output to security groups; exclude distribution lists |

## Output Formats

**PerRow** (default) — one row per group membership. Best for audit work where you
need to filter or pivot by group in Excel.

**Joined** — one row per user with all group names joined by semicolons. Matches the
classic DumpSec single-row-per-user layout.

## Output Fields

| Field | Description |
|---|---|
| UserName | sAMAccountName |
| Groups | Group name, or all group names joined by semicolons (Joined format) |
| FullName | Display name |
| AccountType | Always "User" |
| Comment | Description field |
| PswdCanBeChanged | True if the user is permitted to change their password |
| PswdLastSetTime | When the password was last set |
| PswdRequired | True if a password is required |
| PswdExpires | True if the password is subject to expiration |
| AcctDisabled | True if the account is disabled |
| AcctLockedOut | True if the account is currently locked out |
| AcctExpiresTime | Account expiration date, or "Never" if no expiration is set |

## Prerequisites

- PowerShell 5.1 or later
- Active Directory module (`RSAT-AD-PowerShell`)
- Read access to the target domain (Domain Users is sufficient for most fields;
  `CannotChangePassword` requires read access to user object ACLs)

## Author

BriComp IT Consulting Services — https://bricomp.com
