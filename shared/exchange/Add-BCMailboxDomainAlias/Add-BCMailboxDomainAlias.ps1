<#
.SYNOPSIS
    Adds a domain-based alias address to on-premises Exchange mailboxes and hybrid
    remote mailboxes that are exempt from automatic email address policy management.

.DESCRIPTION
    Finds recipients with EmailAddressPolicyEnabled set to $false (recipients whose
    SMTP addresses are not kept in sync automatically by email address policies) and
    adds a secondary SMTP alias built from each recipient's Exchange Alias attribute
    plus a supplied domain, e.g. jsmith@alias.contoso.com.

    Both recipient sources are checked by default:
      - On-premises mailboxes, via Get-Mailbox / Set-Mailbox.
      - Hybrid remote mailboxes (on-premises AD objects whose actual mailbox is
        hosted in Exchange Online), via Get-RemoteMailbox / Set-RemoteMailbox.
    If either cmdlet is not available in the current session, that source is skipped
    with a warning rather than failing the whole run. The script only throws if
    neither source is available.

    The script is idempotent. Recipients that already have the target address are
    detected and logged as "no change needed" rather than being touched again, so
    it is safe to re-run against the same environment.

.PARAMETER Domain
    The domain to append to each recipient's Alias to build the new SMTP address.
    Example: "alias.contoso.com" produces "jsmith@alias.contoso.com".

.PARAMETER RecipientTypeDetails
    Restricts which on-premises mailbox recipient types are evaluated (passed to
    Get-Mailbox -RecipientTypeDetails). Defaults to UserMailbox.

.PARAMETER RemoteRecipientTypeDetails
    Restricts which hybrid remote mailbox recipient types are evaluated. Get-RemoteMailbox
    has no built-in -RecipientTypeDetails parameter, so this is applied client-side
    against each result's RecipientTypeDetails property. Defaults to RemoteUserMailbox.
    Other common values include RemoteRoomMailbox, RemoteEquipmentMailbox,
    RemoteSharedMailbox, and RemoteTeamMailbox -- add them here if your environment
    needs those covered too.

.PARAMETER LogPath
    Path to a transcript log file. Defaults to a timestamped file in the current
    user's Documents folder.

.PARAMETER PassThru
    Also return the per-recipient result objects to the pipeline, in addition to
    printing the summary.

.EXAMPLE
    .\Add-BCMailboxDomainAlias.ps1 -Domain "alias.contoso.com" -WhatIf

    Reviews what would change, across both mailboxes and remote mailboxes, without
    making any modifications.

.EXAMPLE
    .\Add-BCMailboxDomainAlias.ps1 -Domain "alias.contoso.com"

    Adds the alias address to every eligible recipient and prints a summary.

.EXAMPLE
    .\Add-BCMailboxDomainAlias.ps1 -Domain "alias.contoso.com" -RemoteRecipientTypeDetails "RemoteUserMailbox","RemoteSharedMailbox" -PassThru |
        Export-Csv -Path "C:\Temp\AliasResults.csv" -NoTypeInformation

    Also covers hybrid shared mailboxes and exports the full per-recipient result set.

.NOTES
    Author:    BriComp Computers, LLC
    License:   MIT (see LICENSE in this folder)
    Requires:  Must be run from an Exchange Management Shell session (on-premises
               Exchange Server, including Exchange Server SE). At least one of
               Get-Mailbox or Get-RemoteMailbox must be available. This is a runtime
               prerequisite, not a bundled dependency -- no Microsoft code is
               redistributed by this script.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "Medium")]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Domain,

    [Parameter(Mandatory = $false)]
    [string[]]$RecipientTypeDetails = @("UserMailbox"),

    [Parameter(Mandatory = $false)]
    [string[]]$RemoteRecipientTypeDetails = @("RemoteUserMailbox"),

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Join-Path -Path ([Environment]::GetFolderPath("MyDocuments")) -ChildPath ("Add-BCMailboxDomainAlias_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))),

    [Parameter(Mandatory = $false)]
    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Add-BCAliasToRecipient {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        $Recipient,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Mailbox", "RemoteMailbox")]
        [string]$SourceType,

        [Parameter(Mandatory = $true)]
        [string]$Domain
    )

    $newAddress = "{0}@{1}" -f $Recipient.Alias, $Domain
    $proxyAddress = "smtp:{0}" -f $newAddress

    $existingAddresses = $Recipient.EmailAddresses | ForEach-Object { $_.ToString() -replace "^smtp:", "" }
    $alreadyPresent = $existingAddresses | Where-Object { $_ -ieq $newAddress }

    $status = $null
    $detail = $null

    if ($alreadyPresent) {
        $status = "NoChangeNeeded"
        $detail = "Address already present"
        Write-Host ("[SKIP] {0} ({1}): {2} already present" -f $Recipient.DisplayName, $SourceType, $newAddress)
    }
    else {
        if ($PSCmdlet.ShouldProcess($Recipient.DisplayName, "Add proxy address $newAddress")) {
            try {
                # Exchange Management Shell connects via implicit remoting, so $Recipient came
                # back as a deserialized object. Passing $Recipient.Identity directly can fail
                # with "Cannot process argument transformation ... Deserialized...ADObjectId"
                # because there is no conversion path for the deserialized type. The recipient's
                # GUID is a plain string that every Exchange recipient cmdlet accepts as an
                # -Identity value, so use that instead.
                if ($SourceType -eq "Mailbox") {
                    Set-Mailbox -Identity $Recipient.Guid.ToString() -EmailAddresses @{Add = $proxyAddress } -ErrorAction Stop
                }
                else {
                    Set-RemoteMailbox -Identity $Recipient.Guid.ToString() -EmailAddresses @{Add = $proxyAddress } -ErrorAction Stop
                }
                $status = "Changed"
                $detail = "Address added"
                Write-Host ("[ADD]  {0} ({1}): {2} added" -f $Recipient.DisplayName, $SourceType, $newAddress) -ForegroundColor Green
            }
            catch {
                $status = "Failed"
                $detail = $_.Exception.Message
                Write-Warning ("[FAIL] {0} ({1}): {2} - {3}" -f $Recipient.DisplayName, $SourceType, $newAddress, $detail)
            }
        }
        else {
            $status = "WhatIf"
            $detail = "Would add address"
        }
    }

    [pscustomobject]@{
        DisplayName        = $Recipient.DisplayName
        Source             = $SourceType
        Alias              = $Recipient.Alias
        PrimarySmtpAddress = $Recipient.PrimarySmtpAddress
        NewAddress         = $newAddress
        Status             = $status
        Detail             = $detail
    }
}

$hasGetMailbox = [bool](Get-Command -Name Get-Mailbox -ErrorAction SilentlyContinue)
$hasGetRemoteMailbox = [bool](Get-Command -Name Get-RemoteMailbox -ErrorAction SilentlyContinue)

if (-not $hasGetMailbox -and -not $hasGetRemoteMailbox) {
    throw "Neither Get-Mailbox nor Get-RemoteMailbox is available. Run this script from an Exchange Management Shell session."
}

# Start-Transcript honors ShouldProcess/WhatIfPreference and is silently skipped when this
# script is invoked with -WhatIf. Stop-Transcript does not support ShouldProcess, so it
# would then error trying to stop a transcript that never started. Force Start-Transcript
# to run for real on every invocation, dry run included, so a log is always captured.
Start-Transcript -Path $LogPath -Append -WhatIf:$false | Out-Null

$results = [System.Collections.Generic.List[pscustomobject]]::new()
$attempted = 0
$changed = 0
$noChangeNeeded = 0
$failed = 0
$wouldChange = 0
$localAttempted = 0
$remoteAttempted = 0

try {
    $targets = [System.Collections.Generic.List[pscustomobject]]::new()

    if ($hasGetMailbox) {
        Write-Host ("Discovering mailboxes with EmailAddressPolicyEnabled = `$false (RecipientTypeDetails: {0})..." -f ($RecipientTypeDetails -join ", "))
        Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails $RecipientTypeDetails |
            Where-Object { $_.EmailAddressPolicyEnabled -eq $false } |
            ForEach-Object { $targets.Add([pscustomobject]@{ Recipient = $_; SourceType = "Mailbox" }) }
    }
    else {
        Write-Warning "Get-Mailbox is not available in this session. Skipping on-premises mailboxes."
    }

    if ($hasGetRemoteMailbox) {
        Write-Host ("Discovering remote mailboxes with EmailAddressPolicyEnabled = `$false (RecipientTypeDetails: {0})..." -f ($RemoteRecipientTypeDetails -join ", "))
        Get-RemoteMailbox -ResultSize Unlimited |
            Where-Object { $_.EmailAddressPolicyEnabled -eq $false -and $RemoteRecipientTypeDetails -contains $_.RecipientTypeDetails.ToString() } |
            ForEach-Object { $targets.Add([pscustomobject]@{ Recipient = $_; SourceType = "RemoteMailbox" }) }
    }
    else {
        Write-Warning "Get-RemoteMailbox is not available in this session. Skipping hybrid remote mailboxes."
    }

    if ($targets.Count -eq 0) {
        Write-Host "No matching recipients found. Nothing to do."
    }

    foreach ($target in $targets) {
        $attempted++
        if ($target.SourceType -eq "Mailbox") { $localAttempted++ } else { $remoteAttempted++ }

        $result = Add-BCAliasToRecipient -Recipient $target.Recipient -SourceType $target.SourceType -Domain $Domain
        $results.Add($result)

        switch ($result.Status) {
            "Changed" { $changed++ }
            "NoChangeNeeded" { $noChangeNeeded++ }
            "Failed" { $failed++ }
            "WhatIf" { $wouldChange++ }
        }
    }
}
finally {
    try {
        Stop-Transcript | Out-Null
    }
    catch {
        # Nothing was transcribing (should not happen now that Start-Transcript is
        # forced above), so there is nothing to stop. Safe to ignore.
    }
}

Write-Host ""
Write-Host "===== Summary ====="
Write-Host ("Attempted:        {0} (Mailbox: {1}, RemoteMailbox: {2})" -f $attempted, $localAttempted, $remoteAttempted)
Write-Host ("Changed:          {0}" -f $changed)
Write-Host ("Would change:     {0}{1}" -f $wouldChange, $(if ($WhatIfPreference) { " (WhatIf run -- no changes made)" } else { "" }))
Write-Host ("No change needed: {0}" -f $noChangeNeeded)
Write-Host ("Failed:           {0}" -f $failed)
Write-Host ("Log:              {0}" -f $LogPath)

if ($PassThru) {
    $results
}
