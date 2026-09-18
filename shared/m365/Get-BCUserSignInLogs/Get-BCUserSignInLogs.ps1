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

# SIG # Begin signature block
# MIIobgYJKoZIhvcNAQcCoIIoXzCCKFsCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDUXSodVcYiYcPe
# EjNK8nA4IPaNFnLeXNEkhXxfcXNRmKCCIWswggWNMIIEdaADAgECAhAOmxiO+dAt
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
# AwIBAgIQCE/cM09+RU7bww+P+ZIYNTANBgkqhkiG9w0BAQsFADBpMQswCQYDVQQG
# EwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0
# IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0Ex
# MB4XDTI2MDgwNTAwMDAwMFoXDTM3MTEwNDIzNTk1OVowYzELMAkGA1UEBhMCVVMx
# FzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBTSEEy
# NTYgUlNBNDA5NiBUaW1lc3RhbXAgUmVzcG9uZGVyIDIwMjYgMTCCAiIwDQYJKoZI
# hvcNAQEBBQADggIPADCCAgoCggIBALZ7pvLJ/s1K+NSbTGWz/TjGMPh8CQ6RucZC
# Lv5anHzWJjF/NWJrFIhy24fcpKXlgRiky4WAawDfU3YP0BMxt9l3Dm5oCG5Z69Aq
# EN1kgHg2epx+l+lZBcmJCcN0ASURML5uFIS80sZsDwO3BSkUxDjLJhBI+qiZP3ai
# xAC/qEGLjsBNlLol9VZ7pfGEXiMlneJIC5/YKuizVzNFKZZEeoy/0B8Zm+nzKBgS
# WG52lCO1w+nCg6XpCtklTJXeIg283hw7TmmsZXR+SMbjbrEOvZ3fP2VxIgeR28Y9
# 0ZStd3F9VuA5RVynb/whITPAo9b75Zr4Ta6Mj3URm26QZYMn/FnbuTegcoRcFEZ9
# FOqM5T6MTdtr/n74lIT/ug0eeOzmZ6QTFg33otX+bFRsIolvykE1jive4PuESaT8
# zzVeFWDAMDtozNgLctkGD1ZjkEyZtJrLl5ya0m5doH/ScpaZCZVl6pNUOCybMc/k
# xC6EAmSJY24L0yYKD1Nkddsnb/ItVKi/2nXpQNMu1PT5prW83vV8d67WowuUs0Hd
# Y4H8AMLGvdL/WHEj3ZnqMqAQQP9u3Ai9t+5eQ02GDwy0ODjdzi0xlp70W+ow63/0
# ++YDEX1M0iwgUHwbrJvfpklkZQvw3+kv3vUPItdwroczk9icflf55W1zOEKAcJVA
# IXpcMCU9AgMBAAGjggGVMIIBkTAMBgNVHRMBAf8EAjAAMB0GA1UdDgQWBBQUyWOK
# MC7USvtulPPm40B+9ezN4jAfBgNVHSMEGDAWgBTvb1NK6eQGfHrK4pBW9i/USezL
# TjAOBgNVHQ8BAf8EBAMCB4AwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwgZUGCCsG
# AQUFBwEBBIGIMIGFMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5j
# b20wXQYIKwYBBQUHMAKGUWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydFRydXN0ZWRHNFRpbWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNy
# dDBfBgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGln
# aUNlcnRUcnVzdGVkRzRUaW1lU3RhbXBpbmdSU0E0MDk2U0hBMjU2MjAyNUNBMS5j
# cmwwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEB
# CwUAA4ICAQCNxTphHp1SCt+ZrAmAfn0oQLFr0mLywSLaDXQIENoyKqxrFbJblzCV
# P/pkXmwXOdrOpWygLzlT12os5ipDCy35RBCg2UMeApEtrfGhz45F4Wt4WGdNdIbR
# Wt3YTYJmpR+b7lr4d7Uwn+H600u4D7RnOGf8Wj4UNgAdZkfHhHv1mx9EVh71SJel
# cEN/oORSjXzdjfw1iZH9d8Nh/thn6hH23d+VsPAr6GAYyzSA02nXD1nYLI7Ijmiv
# +xLCiYC41DSFYL3GhTiy0PxpawPtGRyaBVGzq+UiTfM8pD7KVyF5aQyWP4KhVGUU
# Tnmm/RlYJoW3TiXA/+t0YcT2oRVBm3JETjajHug2AL+v5jhtKVnd3D0rbHXEu27o
# +Q8p4sEWPMqKDB+qbceb6T/6WcwTwXmQ9lOCLLYcsQeSWmvKqzpAec9etE14jOQA
# zLKWdE3w/TCaKtLRaRT7LCkRYVnhA2D73FLje1O5b3HR5eHs0NzU/+xX7NbEdcof
# y0W3Wdwd1XOqtlpg/JgwtKfZM5dqO94lbUveOiJBI+xZEbGRsMNbXmMREUTgu+Oc
# a7Y73MPWcslIx2VhkSKSXjDbD6rgg39H5Mh7QfieAIjWagkJNt68Yfim6cjEzVSi
# LSeZfdkr5dtFPTW6jATlWJdYeeDRGCyatf8R1hSjzSvdN8yWQPT9gzCCB3kwggVh
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
# BDEiBCCTsWIKmehROkJoRttgljz177V97PLBSbSNx1msyMmGWjANBgkqhkiG9w0B
# AQEFAASCAgCJy6+nwCA6grZXPzh2cnZ4mdlO5ELuJCsv0nwbL3QoyfIr0b1cSTfK
# DLpfh+BeSz6Mgz9Qwr5aD+ncxPEIXb5PmKpOa7oZHD4HS/t0LPgvQ+jLDtN5uJNW
# d7kCOZEELlvld2we/dr1x6GSBKi5FE/sG8D5iAv1d64aVXiS+dg9xiPd/Iyh/xZn
# +i7Xdoj89K4dx/DlxnXpmJuC+uLbpQIG2c5OMWuM/iAx9MM7fFfe7gxPuEPMg5Kd
# GyOePfDtsPeZyq2IRt5QbBN5sE1V7r6KESKPCArjWQVCQN0A0tXX75Z8xJECOW53
# swYEOPc+YRuBPVzzjCw2NTukNqvPipIJZnMdxvSrxEhUR/jSnz40ig2GKAl264Pt
# uqDLDtaVqzf6CRL7TL4oilfIcJXVEZhdlRMke9ozIldYZ8XUKzkxZT3uJMAOQ3KO
# eUn44hJaAssz45gBWEl1PT1Y1FnnqZxopBahn4W3QuXOiFOHOMMrtv8INnSox7Z4
# MrEua0Qr0Dj+XiMtp8DzncTjkRDxXLx4rE/PkSHitlF7EXpBGOJZ/gJX44V831uP
# zTuXNTi+A1OCVYb8pVkIakQOaiSkiUB/fC+yBgGx4irdOer8tNP6FehiExz1rU4K
# qbZJVf25Ob7uyZ6ir29PVDKpE9b+ZsUrl727NZjzXRH7eoglFrISR6GCAyYwggMi
# BgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQK
# Ew5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBU
# aW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAhP3DNPfkVO28MP
# j/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcB
# MBwGCSqGSIb3DQEJBTEPFw0yNjA5MTgxNjE3MDBaMC8GCSqGSIb3DQEJBDEiBCC3
# CzDusRNvk6yawgHNRdfhC5udB2lUTgo4OLhpriUH7DANBgkqhkiG9w0BAQEFAASC
# AgBdwSB4ZD4t0ZIIZdnJuywA62qYmDnibsTJcasz+tvWHWSr2ElAyLn9JOCZLsbQ
# F9fZAXD+k6OFCV7Dud2W5juOPZcXt1zlb4EkQAJ9dES4icmrvFXwbR5mtIoFrIzj
# lVvx6+e4Djvu0gpDEHFK2h+AX0mpzgRGugz5+NzCUs44uYB1d8nCL3VG9bFhq6Jk
# T7gyOoIGt6s/Hu4B2fjbXNZqxkrWvPva8XfnkY/EVrTu4ZH4nZiambOtbhorHfmz
# +yveicpU7zifGT+FlRO6dyD1i6aFXDa/N76QC/wFo3WSISQg9nMwWXvQEeWNJYM9
# htRlfUkTAxn3oKtjpJzQ8LYAkJQyUygIi0iDvD7jkWjEtsTEhOOY+bwegA+jtDmF
# BWAQowh4t1+RFLppAlkoipYR57ZfnwNnC8sb0GBT9BJ84xf4DVCXuDmtKYPUf5IH
# Eur2KS5rQU6SEnebryray+4sqHo8U0C+TLKr3a3UprEcSiqFcUAseKtfNSta4Uve
# cIoE6TSlmHfdFjeAN5IkWDQ6Pqu9LMWsLFftyw+Y68Ay/3++5mfF4e9Yij6nuqRD
# Oh6rQSYzN7CqWxOyfNhWTVvJTRi5D/eyBFvU0iqL4VnUMK0miWdljpiBtjp/iMvp
# hs5a+lgivWTuOpl4jL/wuavspxCH8Fv4nV2i9cZI1ym6Gg==
# SIG # End signature block
