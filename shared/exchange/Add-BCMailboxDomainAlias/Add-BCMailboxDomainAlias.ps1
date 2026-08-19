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

# SIG # Begin signature block
# MIIobgYJKoZIhvcNAQcCoIIoXzCCKFsCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBNQ89hcwhzqB7l
# PT/zdM2lYkoBHh2prbxnpXv4eq/QpaCCIWswggWNMIIEdaADAgECAhAOmxiO+dAt
# 5+/bUOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBa
# Fw0zMTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lD
# ZXJ0IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBAL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3E
# MB/zG6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKy
# unWZanMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsF
# xl7sWxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU1
# 5zHL2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJB
# MtfbBHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObUR
# WBf3JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6
# nj3cAORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxB
# YKqxYxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5S
# UUd0viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+x
# q4aLT8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIB
# NjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwP
# TzAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMC
# AYYweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0
# aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENB
# LmNybDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0Nc
# Vec4X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnov
# Lbc47/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65Zy
# oUi0mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFW
# juyk1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPF
# mCLBsln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9z
# twGpn1eqXijiuZQwggawMIIEmKADAgECAhAIrUCyYNKcTJ9ezam9k67ZMA0GCSqG
# SIb3DQEBDAUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRy
# dXN0ZWQgUm9vdCBHNDAeFw0yMTA0MjkwMDAwMDBaFw0zNjA0MjgyMzU5NTlaMGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcgUlNBNDA5NiBTSEEzODQg
# MjAyMSBDQTEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDVtC9C0Cit
# eLdd1TlZG7GIQvUzjOs9gZdwxbvEhSYwn6SOaNhc9es0JAfhS0/TeEP0F9ce2vnS
# 1WcaUk8OoVf8iJnBkcyBAz5NcCRks43iCH00fUyAVxJrQ5qZ8sU7H/Lvy0daE6ZM
# swEgJfMQ04uy+wjwiuCdCcBlp/qYgEk1hz1RGeiQIXhFLqGfLOEYwhrMxe6TSXBC
# Mo/7xuoc82VokaJNTIIRSFJo3hC9FFdd6BgTZcV/sk+FLEikVoQ11vkunKoAFdE3
# /hoGlMJ8yOobMubKwvSnowMOdKWvObarYBLj6Na59zHh3K3kGKDYwSNHR7OhD26j
# q22YBoMbt2pnLdK9RBqSEIGPsDsJ18ebMlrC/2pgVItJwZPt4bRc4G/rJvmM1bL5
# OBDm6s6R9b7T+2+TYTRcvJNFKIM2KmYoX7BzzosmJQayg9Rc9hUZTO1i4F4z8ujo
# 7AqnsAMrkbI2eb73rQgedaZlzLvjSFDzd5Ea/ttQokbIYViY9XwCFjyDKK05huzU
# tw1T0PhH5nUwjewwk3YUpltLXXRhTT8SkXbev1jLchApQfDVxW0mdmgRQRNYmtwm
# KwH0iU1Z23jPgUo+QEdfyYFQc4UQIyFZYIpkVMHMIRroOBl8ZhzNeDhFMJlP/2NP
# TLuqDQhTQXxYPUez+rbsjDIJAsxsPAxWEQIDAQABo4IBWTCCAVUwEgYDVR0TAQH/
# BAgwBgEB/wIBADAdBgNVHQ4EFgQUaDfg67Y7+F8Rhvv+YXsIiGX0TkIwHwYDVR0j
# BBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0
# cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0
# cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8E
# PDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVz
# dGVkUm9vdEc0LmNybDAcBgNVHSAEFTATMAcGBWeBDAEDMAgGBmeBDAEEATANBgkq
# hkiG9w0BAQwFAAOCAgEAOiNEPY0Idu6PvDqZ01bgAhql+Eg08yy25nRm95RysQDK
# r2wwJxMSnpBEn0v9nqN8JtU3vDpdSG2V1T9J9Ce7FoFFUP2cvbaF4HZ+N3HLIvda
# qpDP9ZNq4+sg0dVQeYiaiorBtr2hSBh+3NiAGhEZGM1hmYFW9snjdufE5BtfQ/g+
# lP92OT2e1JnPSt0o618moZVYSNUa/tcnP/2Q0XaG3RywYFzzDaju4ImhvTnhOE7a
# brs2nfvlIVNaw8rpavGiPttDuDPITzgUkpn13c5UbdldAhQfQDN8A+KVssIhdXNS
# y0bYxDQcoqVLjc1vdjcshT8azibpGL6QB7BDf5WIIIJw8MzK7/0pNVwfiThV9zeK
# iwmhywvpMRr/LhlcOXHhvpynCgbWJme3kuZOX956rEnPLqR0kq3bPKSchh/jwVYb
# KyP/j7XqiHtwa+aguv06P0WmxOgWkVKLQcBIhEuWTatEQOON8BUozu3xGFYHKi8Q
# xAwIZDwzj64ojDzLj4gLDb879M4ee47vtevLt/B3E+bnKD+sEq6lLyJsQfmCXBVm
# zGwOysWGw/YmMwwHS6DTBwJqakAwSEs0qFEgu60bhQjiWQ1tygVQK+pKHJ6l/aCn
# HwZ05/LWUpD9r4VIIflXO7ScA+2GRfS0YW6/aOImYIbqyK+p/pQd52MbOoZWeE4w
# gga0MIIEnKADAgECAhANx6xXBf8hmS5AQyIMOkmGMA0GCSqGSIb3DQEBCwUAMGIx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBH
# NDAeFw0yNTA1MDcwMDAwMDBaFw0zODAxMTQyMzU5NTlaMGkxCzAJBgNVBAYTAlVT
# MRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1
# c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwggIi
# MA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQC0eDHTCphBcr48RsAcrHXbo0Zo
# dLRRF51NrY0NlLWZloMsVO1DahGPNRcybEKq+RuwOnPhof6pvF4uGjwjqNjfEvUi
# 6wuim5bap+0lgloM2zX4kftn5B1IpYzTqpyFQ/4Bt0mAxAHeHYNnQxqXmRinvuNg
# xVBdJkf77S2uPoCj7GH8BLuxBG5AvftBdsOECS1UkxBvMgEdgkFiDNYiOTx4OtiF
# cMSkqTtF2hfQz3zQSku2Ws3IfDReb6e3mmdglTcaarps0wjUjsZvkgFkriK9tUKJ
# m/s80FiocSk1VYLZlDwFt+cVFBURJg6zMUjZa/zbCclF83bRVFLeGkuAhHiGPMvS
# GmhgaTzVyhYn4p0+8y9oHRaQT/aofEnS5xLrfxnGpTXiUOeSLsJygoLPp66bkDX1
# ZlAeSpQl92QOMeRxykvq6gbylsXQskBBBnGy3tW/AMOMCZIVNSaz7BX8VtYGqLt9
# MmeOreGPRdtBx3yGOP+rx3rKWDEJlIqLXvJWnY0v5ydPpOjL6s36czwzsucuoKs7
# Yk/ehb//Wx+5kMqIMRvUBDx6z1ev+7psNOdgJMoiwOrUG2ZdSoQbU2rMkpLiQ6bG
# RinZbI4OLu9BMIFm1UUl9VnePs6BaaeEWvjJSjNm2qA+sdFUeEY0qVjPKOWug/G6
# X5uAiynM7Bu2ayBjUwIDAQABo4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB/wIBADAd
# BgNVHQ4EFgQU729TSunkBnx6yuKQVvYv1Ensy04wHwYDVR0jBBgwFoAU7NfjgtJx
# XWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUF
# BwMIMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGln
# aWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5j
# b20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJo
# dHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNy
# bDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQEL
# BQADggIBABfO+xaAHP4HPRF2cTC9vgvItTSmf83Qh8WIGjB/T8ObXAZz8OjuhUxj
# aaFdleMM0lBryPTQM2qEJPe36zwbSI/mS83afsl3YTj+IQhQE7jU/kXjjytJgnn0
# hvrV6hqWGd3rLAUt6vJy9lMDPjTLxLgXf9r5nWMQwr8Myb9rEVKChHyfpzee5kH0
# F8HABBgr0UdqirZ7bowe9Vj2AIMD8liyrukZ2iA/wdG2th9y1IsA0QF8dTXqvcnT
# mpfeQh35k5zOCPmSNq1UH410ANVko43+Cdmu4y81hjajV/gxdEkMx1NKU4uHQcKf
# ZxAvBAKqMVuqte69M9J6A47OvgRaPs+2ykgcGV00TYr2Lr3ty9qIijanrUR3anzE
# wlvzZiiyfTPjLbnFRsjsYg39OlV8cipDoq7+qNNjqFzeGxcytL5TTLL4ZaoBdqbh
# OhZ3ZRDUphPvSRmMThi0vw9vODRzW6AxnJll38F0cuJG7uEBYTptMSbhdhGQDpOX
# gpIUsWTjd6xpR6oaQf/DJbg3s6KCLPAlZ66RzIg9sC+NJpud/v4+7RWsWCiKi9EO
# LLHfMR2ZyJ/+xhCx9yHbxtl5TPau1j/1MIDpMPx0LckTetiSuEtQvLsNz3Qbp7wG
# WqbIiOWCnb5WqxL3/BAPvIXKUjPSxyZsq8WhbaM2tszWkPZPubdcMIIG7TCCBNWg
# AwIBAgIQCoDvGEuN8QWC0cR2p5V0aDANBgkqhkiG9w0BAQsFADBpMQswCQYDVQQG
# EwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0
# IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0Ex
# MB4XDTI1MDYwNDAwMDAwMFoXDTM2MDkwMzIzNTk1OVowYzELMAkGA1UEBhMCVVMx
# FzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBTSEEy
# NTYgUlNBNDA5NiBUaW1lc3RhbXAgUmVzcG9uZGVyIDIwMjUgMTCCAiIwDQYJKoZI
# hvcNAQEBBQADggIPADCCAgoCggIBANBGrC0Sxp7Q6q5gVrMrV7pvUf+GcAoB38o3
# zBlCMGMyqJnfFNZx+wvA69HFTBdwbHwBSOeLpvPnZ8ZN+vo8dE2/pPvOx/Vj8Tch
# TySA2R4QKpVD7dvNZh6wW2R6kSu9RJt/4QhguSssp3qome7MrxVyfQO9sMx6ZAWj
# FDYOzDi8SOhPUWlLnh00Cll8pjrUcCV3K3E0zz09ldQ//nBZZREr4h/GI6Dxb2Uo
# yrN0ijtUDVHRXdmncOOMA3CoB/iUSROUINDT98oksouTMYFOnHoRh6+86Ltc5zjP
# KHW5KqCvpSduSwhwUmotuQhcg9tw2YD3w6ySSSu+3qU8DD+nigNJFmt6LAHvH3KS
# uNLoZLc1Hf2JNMVL4Q1OpbybpMe46YceNA0LfNsnqcnpJeItK/DhKbPxTTuGoX7w
# JNdoRORVbPR1VVnDuSeHVZlc4seAO+6d2sC26/PQPdP51ho1zBp+xUIZkpSFA8vW
# doUoHLWnqWU3dCCyFG1roSrgHjSHlq8xymLnjCbSLZ49kPmk8iyyizNDIXj//cOg
# rY7rlRyTlaCCfw7aSUROwnu7zER6EaJ+AliL7ojTdS5PWPsWeupWs7NpChUk555K
# 096V1hE0yZIXe+giAwW00aHzrDchIc2bQhpp0IoKRR7YufAkprxMiXAJQ1XCmnCf
# gPf8+3mnAgMBAAGjggGVMIIBkTAMBgNVHRMBAf8EAjAAMB0GA1UdDgQWBBTkO/zy
# Me39/dfzkXFjGVBDz2GM6DAfBgNVHSMEGDAWgBTvb1NK6eQGfHrK4pBW9i/USezL
# TjAOBgNVHQ8BAf8EBAMCB4AwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwgZUGCCsG
# AQUFBwEBBIGIMIGFMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5j
# b20wXQYIKwYBBQUHMAKGUWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydFRydXN0ZWRHNFRpbWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNy
# dDBfBgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGln
# aUNlcnRUcnVzdGVkRzRUaW1lU3RhbXBpbmdSU0E0MDk2U0hBMjU2MjAyNUNBMS5j
# cmwwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEB
# CwUAA4ICAQBlKq3xHCcEua5gQezRCESeY0ByIfjk9iJP2zWLpQq1b4URGnwWBdEZ
# D9gBq9fNaNmFj6Eh8/YmRDfxT7C0k8FUFqNh+tshgb4O6Lgjg8K8elC4+oWCqnU/
# ML9lFfim8/9yJmZSe2F8AQ/UdKFOtj7YMTmqPO9mzskgiC3QYIUP2S3HQvHG1FDu
# +WUqW4daIqToXFE/JQ/EABgfZXLWU0ziTN6R3ygQBHMUBaB5bdrPbF6MRYs03h4o
# bEMnxYOX8VBRKe1uNnzQVTeLni2nHkX/QqvXnNb+YkDFkxUGtMTaiLR9wjxUxu2h
# ECZpqyU1d0IbX6Wq8/gVutDojBIFeRlqAcuEVT0cKsb+zJNEsuEB7O7/cuvTQasn
# M9AWcIQfVjnzrvwiCZ85EE8LUkqRhoS3Y50OHgaY7T/lwd6UArb+BOVAkg2oOvol
# /DJgddJ35XTxfUlQ+8Hggt8l2Yv7roancJIFcbojBcxlRcGG0LIhp6GvReQGgMgY
# xQbV1S3CrWqZzBt1R9xJgKf47CdxVRd/ndUlQ05oxYy2zRWVFjF7mcr4C34Mj3oc
# CVccAvlKV9jEnstrniLvUxxVZE/rptb7IRE2lskKPIJgbaP5t2nGj/ULLi49xTcB
# ZU8atufk+EMF/cWuiC7POGT75qaL6vdCvHlshtjdNXOCIUjsarfNZzCCB3kwggVh
# oAMCAQICEAPvwdvfaByOuGfVs03RjH4wDQYJKoZIhvcNAQELBQAwaTELMAkGA1UE
# BhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2Vy
# dCBUcnVzdGVkIEc0IENvZGUgU2lnbmluZyBSU0E0MDk2IFNIQTM4NCAyMDIxIENB
# MTAeFw0yNTA4MTEwMDAwMDBaFw0yODA4MTAyMzU5NTlaMIGAMQswCQYDVQQGEwJV
# UzEQMA4GA1UECBMHQXJpem9uYTEQMA4GA1UEBxMHR2lsYmVydDEfMB0GA1UEChMW
# QnJpQ29tcCBDb21wdXRlcnMsIExMQzELMAkGA1UECxMCSVQxHzAdBgNVBAMTFkJy
# aUNvbXAgQ29tcHV0ZXJzLCBMTEMwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIK
# AoICAQDmORZ7VshQu610p+5XvpZwuBtGdXz/uDjhb08OBTFpxU+N/J7VrADg5IQr
# l2uow0RaFYPoWAp3tc9sszl9PL1moO0XjIhJngsIaqwZsqu9EO7kIFYZAqE5ziKP
# vipbdtKIw+Bo0BaMUwg5KuwU9Dp+BntlnlPcU00zECuyM/T+VCV+WxpFT1dXIj00
# chaNNSfzvhNc25HapytutgUpurgpTQ4zRpBJ3IhROJWb3yOJ8gcGaIUdqW49RpUg
# 9tuYVxtekZN+1Twl3hBwn8stZ+CD8vaUFARNs6WWgWX+trD9JHIQaxQz9DO0oPVz
# TKPWpJdUQNaPWU3x3hKZgF1nObXQEK31dmwVDlJPKQ/JV9wnkT9bDl4JNILgpJHR
# JYE9oszu0+mSUsjwglN10hrnhcEE6avZIUCN6zGrilVCkWee/mspUVTo0Oz/eaHV
# MnOVt/FH23zq6iMjfKY0bUcRUHE3EiT5bCVyLjxkfQgBlrZsyhcX3lLQ00bccIbA
# DsY+WS2peQKS1CKQAw19fAGct0HmFDWaOKqC6UBl710NH/HUs3K/QGHqWub+wFY5
# ypNC+WDXVd6Klm4CqbzAtWnunrnClbF8GW0418pVVK/syJroSYikyRF7R6dGbQo6
# C2/89l/S5g+dNmmzPBO1I5l4t9X2W4FQByx4b9W1gkHjQ+GcIQIDAQABo4ICAzCC
# Af8wHwYDVR0jBBgwFoAUaDfg67Y7+F8Rhvv+YXsIiGX0TkIwHQYDVR0OBBYEFMBR
# VSnudhF0J79q4bGDuXVuvTkDMD4GA1UdIAQ3MDUwMwYGZ4EMAQQBMCkwJwYIKwYB
# BQUHAgEWG2h0dHA6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAOBgNVHQ8BAf8EBAMC
# B4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwgbUGA1UdHwSBrTCBqjBToFGgT4ZNaHR0
# cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0Q29kZVNpZ25p
# bmdSU0E0MDk2U0hBMzg0MjAyMUNBMS5jcmwwU6BRoE+GTWh0dHA6Ly9jcmw0LmRp
# Z2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNENvZGVTaWduaW5nUlNBNDA5NlNI
# QTM4NDIwMjFDQTEuY3JsMIGUBggrBgEFBQcBAQSBhzCBhDAkBggrBgEFBQcwAYYY
# aHR0cDovL29jc3AuZGlnaWNlcnQuY29tMFwGCCsGAQUFBzAChlBodHRwOi8vY2Fj
# ZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRDb2RlU2lnbmluZ1JT
# QTQwOTZTSEEzODQyMDIxQ0ExLmNydDAJBgNVHRMEAjAAMA0GCSqGSIb3DQEBCwUA
# A4ICAQB0dH/YkZqUjUXmYUn/948bMbrlKzTTbdLNQ/9fcPBqdXdqQ/m6c6Bb98Ti
# x9Jmk9Vp5rtqD3E47pRSQ9G9Cn+2WgandklNhYOo3fZIPO7WHA/0hB9cu9bfBxO5
# vV78jR2xow+WLs096OYJaKHMiU8mT498Db2NHLSv6+FeGxf8WZf/Ujp3nrkedlCa
# iMI2wTzbjMBHfQTGYZTAG6Nic1PIpJRs9x817QtAsyhQqPqlxi/J6tui0aMUYnJK
# fZiYJN74v+ANvNqJ9lPIivhc6k90ishsOtb2u8Ol4CDtDjkOSLnqlmU6FL0VygEJ
# g6HafKyspnIx5u+pqG8Icq5LJ0Tk9U6O/1riJlWqo5GYEPRacNXjuAxHXrp5iSoB
# JzabJwiAvZEHYDtfkAt1obbRaVb39ghaGgX/hpgnN3vgagPRCh4zhq5KE0Y/iXUA
# QDXzWndaKL2ScQSj7w5KGsHxDETg2VgrlDfZJ9e4szEX8R2TKmRRoZofBAa4STHq
# s2EYudfauWMV5jFT5d8ux1nUAz8pRLlPUWogJbhBLrtxjH8trjK2+cv/173j4JtO
# hCKjYp6a8wX3ZP3u6x1JdCJ9p0QlA5OpL5IXiFzPxPTB5HJ9W4G5Aq7/bYOI2lH9
# uVPA6OJ+HigaBjL2tsX1D+up2iWVZSEj89P9tvu6sTRYLkDX5zGCBlkwggZVAgEB
# MH0waTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMUEwPwYD
# VQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IENvZGUgU2lnbmluZyBSU0E0MDk2IFNI
# QTM4NCAyMDIxIENBMQIQA+/B299oHI64Z9WzTdGMfjANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCCCyYMmG5a5YLg5sivIMZQymzTQp2Q3LfkvBjd+IWZxDDANBgkqhkiG9w0B
# AQEFAASCAgCTT8GEYzn9l691eXgt4Z7LxQGmU9Q/tkWDKdG8bkZ3fPyiPowQElr2
# K2l/sAkXAUKtzRfib1A/U1wa0Rnz5Je2DzKoJtLDKIikOSy7soBxej85taVfpUQt
# dkOud4eu9KwQq/2nbnbnCzbMlAItJV+yCpawMewT+TK2LMXWJYFvS1IDi5kj4pT3
# u/dPaesfeaibXomtMXvpv+0PaJXi/RzAaxtAF0FB8iFAXlcEaWhEOPctRXGmK3P3
# YDdd5RMlaTvIw2ZM3Ba54XUeTwRtX/lrKA0g2CrkXAxF9Vpa5ZNX5P0xb2+TIzBu
# kSadKDvU0K4qcrd6+55MWD7I8GznMzCURrAOa5/a8f8sLPqNlVSx4yWNMNZd/p16
# 15LWqwprS5yz5kcMTP46gG22IP4as8r6fJ+hkjH/Nx1t5gIZqKbRH+uwjLHOD2wF
# 8+6Ek96fHbRUZkXvCmBopyfAnDibP+f6n0VU90bdtf7TmSGDiQAyqo0Jj4PUFQhX
# vt6SYYHxZfQEjuG1DIgpH6dYnbcPh9ijF2Rl21ckGNjfqiVLwq0WWWn1QKZOw8e3
# dyX72MUsEMK0nRhOVfil6Xv5zJ5jsdwenzKquRAKnhMIhAIkKtj5foaeYr2ZFyp5
# 5lf02jSKY3IWkzfj9WwrjmnEs5mH4uRbKeMs9ROEvebySzjBaODeOaGCAyYwggMi
# BgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQK
# Ew5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBU
# aW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHE
# dqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcB
# MBwGCSqGSIb3DQEJBTEPFw0yNjA4MTkxODE2MDdaMC8GCSqGSIb3DQEJBDEiBCAN
# 0pxwH+uwZO+mRrhqpjdzQ50IemGV813hOdstm5H0wjANBgkqhkiG9w0BAQEFAASC
# AgAfW4v7woVD41KAKR+axdQksYQg1Weq5r8oxEnmTnB+yGRp+d3GLL4d0KVr5zmN
# Q/dcgBHgnwhrIXydIsSFzjYc/CZnGncgpg00QCHObye4PZ9LrhBdTiUHW/1+yBhm
# GYSocx/R8M6LLj89TKdGm0ihGVvVnvt3MYCs2NZ1sxaLHBkF2iknx2sO2ACvcSK1
# O7iP1S4iLUzgfdo3mj1CM6+qxC3WMOvPszS87HAqea0OHxzxKgeKIb9YrkxS9mBr
# rlTOSQP2wFtETkj3BTTSO3HxQoVzlbePckiGWAQ4xjt/7IMxmdnk4/DSr2cq0uCE
# NikmA624X4g1RuTOBzpMHAgK8T4qby3OgEiZsdYyqA5s9OSQjtKmSh8efKDTDFwV
# hbMUT0p4fFiXUxouErvVVt+6sfpBypsnVolykb3iE9qrMeBUHe+1q/Y/kICpUh81
# B76BMPW2X2yww+krGR5BeLer1XVhUSaXLw3IP8XmqjR7bPO5bFJvPMpDTvkqAiXl
# VWvvlWXQ9MEQcyVrOFnelDnNDJZPjoMdJIeBaOfiKuJhLiJKHYPopzP7BdWreL60
# Lf6Zd2ZOOqyN/PdhD7WAUhP9dyBEGuK7z8CHhaDXSCuqFXB3W/FcR7Gz7xsKOLjX
# E/PzKJKx6F2rBOW1rwDnNJwQbHismv0HKZHDyEGgkbQOCw==
# SIG # End signature block
