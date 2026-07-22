# Get-BCUserSignInLogs.ps1

Retrieves Entra ID sign-in logs for a specific user within a local time window
via Microsoft Graph. Automatically converts local time to UTC for the API query
and converts results back to local time for display. Captures both **interactive**
and **non-interactive** sign-ins.

## Prerequisites

- PowerShell 7.0 or higher
- An account with the **Security Reader** or **Reports Reader** role (or Global Admin)
  in the target Entra ID tenant
- Entra ID **P1 or P2** licensing in the target tenant

## Module Installation

The script will automatically install the required Graph modules if not present:

- `Microsoft.Graph.Authentication`
- `Microsoft.Graph.Beta.Reports`

If you prefer to install manually ahead of time:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Beta.Reports -Scope CurrentUser -Force
```

## Graph API Consent

On first run, the script will prompt you to authenticate and grant consent for
the `AuditLog.Read.All` permission scope. This is a one-time step per account
per tenant.

> **Note:** If your tenant requires admin consent for API permissions, a Global
> Admin will need to grant consent before non-admin accounts can use the script.
> This can be done in the Entra admin center under:
> **App registrations → Microsoft Graph PowerShell → API permissions → Grant admin consent**

## Usage

```powershell
# All sign-ins (interactive + non-interactive, success and failure) with auto-named CSV
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

## Parameters

| Parameter | Required | Description |
|---|---|---|
| `-UserPrincipalName` | Yes | UPN of the target user (e.g. `jsmith@contoso.com`) |
| `-LocalStartTime` | Yes | Start of window in local time (e.g. `"2026-06-25 08:00"`) |
| `-LocalEndTime` | Yes | End of window in local time (e.g. `"2026-06-25 18:00"`) |
| `-CsvPath` | No | Full path for CSV export. Defaults to current directory with auto-generated filename |
| `-SuccessOnly` | No | Limits results to successful sign-ins only (StatusCode eq 0). Cannot be combined with `-FailuresOnly` |
| `-FailuresOnly` | No | Limits results to failed sign-ins only (StatusCode ne 0). Cannot be combined with `-SuccessOnly` |

## Output

Results are written to both the console and a CSV file containing:

| Column | Description |
|---|---|
| `LocalTime` | Sign-in time converted to the local timezone of the machine running the script |
| `UtcTime` | Raw UTC timestamp from Graph |
| `SignInType` | `interactiveUser` or `nonInteractiveUser` |
| `UserPrincipalName` | UPN of the signing-in user |
| `AppDisplayName` | Application that was accessed |
| `ClientAppUsed` | Client type (Browser, Mobile Apps, MAPI, etc.) |
| `IPAddress` | Source IP address |
| `Location` | City and country resolved from IP |
| `StatusCode` | `0` = success; non-zero = failure |
| `StatusDetail` | Failure reason if applicable |
| `DeviceDetail` | Device display name if available |
| `CorrelationId` | Correlation ID for cross-referencing with other logs |
| `ConditionalAccess` | Applied Conditional Access policy names |

## Notes

- Sign-in log retention in Entra ID is **30 days** for interactive sign-ins (P1/P2)
- Non-interactive sign-ins (token refreshes, silent auth) can be significantly
  noisier than interactive — use `-SuccessOnly` to reduce noise if needed
- The beta Graph endpoint is required for `signInEventTypes` filtering; this is
  expected and stable for production use despite the beta label
- Time conversion uses the timezone of the machine running the script — results
  may differ if run from a machine set to a different timezone (e.g. a UTC-based server)
- The `AuditLog.Read.All` consent granted during first run is tied to the
  **Microsoft Graph PowerShell** enterprise app in your tenant

## Version History

| Version | Date | Notes |
|---|---|---|
| 1.1.0 | 2026-06-29 | Added `-FailuresOnly` switch; mutual exclusion check for `-SuccessOnly`/`-FailuresOnly` |
| 1.0.0 | 2026-06-29 | Initial release |
