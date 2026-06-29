#Requires -Version 7.0

<#
.SYNOPSIS
    Retrieves Entra ID sign-in logs for a user within a local time window.
.DESCRIPTION
    Queries /auditLogs/signIns via Microsoft Graph Beta, converting the caller's
    local time window to UTC for the API filter, then converts results back to
    local time for display. Captures both interactive and non-interactive sign-ins.
    Outputs to console and CSV.
.PARAMETER UserPrincipalName
    UPN of the target user (e.g. jsmith@contoso.com)
.PARAMETER LocalStartTime
    Start of window in local time (e.g. "2026-06-25 08:00")
.PARAMETER LocalEndTime
    End of window in local time (e.g. "2026-06-25 18:00")
.PARAMETER CsvPath
    Optional. Full path for CSV export. If not specified, exports to current
    directory with an auto-generated filename.
.PARAMETER SuccessOnly
    If specified, limits results to successful sign-ins (StatusCode eq 0).
    Cannot be combined with -FailuresOnly.
.PARAMETER FailuresOnly
    If specified, limits results to failed sign-ins (StatusCode ne 0).
    Cannot be combined with -SuccessOnly.
.NOTES
    Author  : BriComp IT Consulting Services
    Version : 1.1.0
    Requires: PowerShell 7.0+, Entra ID P1/P2, Security Reader or Reports Reader role
.EXAMPLE
    .\Get-BCUserSignInLogs.ps1 -UserPrincipalName jsmith@contoso.com `
        -LocalStartTime "2026-06-25 08:00" -LocalEndTime "2026-06-25 18:00"

.EXAMPLE
    .\Get-BCUserSignInLogs.ps1 -UserPrincipalName jsmith@contoso.com `
        -LocalStartTime "2026-06-25 08:00" -LocalEndTime "2026-06-25 18:00" `
        -SuccessOnly -CsvPath C:\Temp\jsmith_signins.csv

.EXAMPLE
    .\Get-BCUserSignInLogs.ps1 -UserPrincipalName jsmith@contoso.com `
        -LocalStartTime "2026-06-25 08:00" -LocalEndTime "2026-06-25 18:00" `
        -FailuresOnly
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]$UserPrincipalName,

    [Parameter(Mandatory)]
    [datetime]$LocalStartTime,

    [Parameter(Mandatory)]
    [datetime]$LocalEndTime,

    [string]$CsvPath,

    [switch]$SuccessOnly,
    [switch]$FailuresOnly
)

# --- Parameter validation ---
if ($SuccessOnly -and $FailuresOnly) {
    Write-Error "Cannot specify both -SuccessOnly and -FailuresOnly."
    exit 1
}

# --- Module check / auto-install ---
$requiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Beta.Reports'
)

foreach ($mod in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "Module '$mod' not found. Installing..."
        Install-Module $mod -Scope CurrentUser -Force -ErrorAction Stop
    }
    Import-Module $mod -ErrorAction Stop
}

function Write-BCLog {
    param([string]$Message, [string]$Level = 'INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "[$ts][$Level] $Message"
}

# --- Default CSV path if not specified ---
if (-not $CsvPath) {
    $CsvPath = ".\$($UserPrincipalName -replace '@','_')_signins_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    Write-BCLog "No CsvPath specified, defaulting to $CsvPath"
}

# --- Convert local window to UTC for Graph filter ---
$utcStart = $LocalStartTime.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
$utcEnd   = $LocalEndTime.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')

Write-BCLog "Local window : $LocalStartTime → $LocalEndTime"
Write-BCLog "UTC window   : $utcStart → $utcEnd"

# --- Connect to Graph ---
Write-BCLog "Connecting to Microsoft Graph..."
Connect-MgGraph -Scopes 'AuditLog.Read.All' -NoWelcome

# --- Build OData filter ---
# Beta endpoint required for signInEventTypes filtering.
# Captures both interactive and non-interactive user sign-ins.
$filter = "userPrincipalName eq '$UserPrincipalName' " +
          "and createdDateTime ge $utcStart " +
          "and createdDateTime le $utcEnd " +
          "and (signInEventTypes/any(t: t eq 'interactiveUser') " +
          "or signInEventTypes/any(t: t eq 'nonInteractiveUser'))"

Write-BCLog "OData filter : $filter"

# --- Query Graph (beta endpoint required for signInEventTypes support) ---
Write-BCLog "Querying /beta/auditLogs/signIns ..."
try {
    $signIns = Get-MgBetaAuditLogSignIn -Filter $filter -All -ErrorAction Stop
} catch {
    Write-BCLog "Graph query failed: $_" -Level 'ERROR'
    exit 1
}

if (-not $signIns) {
    Write-BCLog "No sign-in records found for $UserPrincipalName in the specified window." -Level 'WARN'
    exit 0
}

Write-BCLog "Retrieved $($signIns.Count) record(s)."

# --- Post-process: add local time column ---
$results = $signIns | ForEach-Object {
    [PSCustomObject]@{
        LocalTime         = [System.TimeZoneInfo]::ConvertTimeFromUtc($_.CreatedDateTime, [System.TimeZoneInfo]::Local)
        UtcTime           = $_.CreatedDateTime
        SignInType        = ($_.SignInEventTypes -join ', ')
        UserPrincipalName = $_.UserPrincipalName
        AppDisplayName    = $_.AppDisplayName
        ClientAppUsed     = $_.ClientAppUsed
        IPAddress         = $_.IpAddress
        Location          = "$($_.Location.City), $($_.Location.CountryOrRegion)"
        StatusCode        = $_.Status.ErrorCode
        StatusDetail      = $_.Status.FailureReason
        DeviceDetail      = $_.DeviceDetail.DisplayName
        CorrelationId     = $_.CorrelationId
        ConditionalAccess = ($_.AppliedConditionalAccessPolicies.DisplayName -join '; ')
    }
} | Sort-Object LocalTime

if ($SuccessOnly) {
    $results = $results | Where-Object { $_.StatusCode -eq 0 }
    Write-BCLog "Filtered to $($results.Count) successful sign-in(s)."
}
elseif ($FailuresOnly) {
    $results = $results | Where-Object { $_.StatusCode -ne 0 }
    Write-BCLog "Filtered to $($results.Count) failed sign-in(s)."
}

# --- Output ---
$results | Format-Table -AutoSize -Property LocalTime, SignInType, AppDisplayName, ClientAppUsed, IPAddress, Location, StatusCode, StatusDetail

$results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
Write-BCLog "Exported to $CsvPath"
