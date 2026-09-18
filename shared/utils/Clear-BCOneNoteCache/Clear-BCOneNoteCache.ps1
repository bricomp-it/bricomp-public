#Requires -Version 5.1
# PUBLISH: true

<#
.SYNOPSIS
    Clears the local OneNote desktop cache to resolve freezing and sync issues.

.DESCRIPTION
    Stops OneNote if running, removes stale local cache folders under the
    OneNote 16.0 profile directory, and optionally relaunches the application.

    Folders cleared:
        cache\                  Primary sync and keystroke cache (most common culprit)
        FullTextSearchIndex\    Local search index; rebuilt automatically on launch
        MasterIndex\            Notebook index; rebuilt automatically on launch
        ServerListings\         Cached notebook server/location data; rebuilt on launch

    The Backup\ folder is intentionally preserved.

    Safe to run when notebooks are stored in OneDrive or SharePoint -- no
    notebook data lives in these cache paths. OneNote re-syncs everything
    from the cloud on next launch.

    When run as SYSTEM (e.g. via RMM, SCCM, or ScreenConnect elevated shell),
    the script auto-detects the interactively logged-on user. Supply -Username
    explicitly if auto-detection is ambiguous or if targeting a specific profile.

    Note: -Relaunch only functions when running in user context. When run as
    SYSTEM, OneNote cannot be launched into the user desktop session.

.PARAMETER Username
    Optional. Target a specific local user profile by username (e.g. 'jsmith').
    If omitted, defaults to the currently logged-on interactive user. Required
    when multiple users have active sessions or auto-detection fails.

.PARAMETER Relaunch
    Optional. Relaunches OneNote after the cache is cleared. Only effective
    when running in user context, not as SYSTEM.

.PARAMETER WhatIf
    Shows what would be removed without making any changes.

.EXAMPLE
    .\Clear-BCOneNoteCache.ps1
    Clears the OneNote cache for the currently logged-on user. OneNote is
    stopped if running.

.EXAMPLE
    .\Clear-BCOneNoteCache.ps1 -Relaunch
    Clears the cache and relaunches OneNote when complete.

.EXAMPLE
    .\Clear-BCOneNoteCache.ps1 -Username jsmith
    Clears the cache targeting the profile at C:\Users\jsmith. Useful when
    running elevated or via RMM and auto-detection is unreliable.

.EXAMPLE
    .\Clear-BCOneNoteCache.ps1 -WhatIf
    Dry run -- shows what would be removed without making changes.

.NOTES
    Version:    1.0.0
    Author:     BriComp IT Consulting Services
    Website:    https://bricomp.com
    Created:    2026-07-01

    Tested against M365 Apps for Enterprise (OneNote 16.0).
    Not applicable to the OneNote UWP (Store) app.
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(HelpMessage = 'Local username whose profile to target (e.g. jsmith).')]
    [string]$Username,

    [Parameter(HelpMessage = 'Relaunch OneNote after clearing the cache.')]
    [switch]$Relaunch
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region -- Helpers ------------------------------------------------------------

function Write-BCLog {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'OK')]
        [string]$Level = 'INFO'
    )
    $ts     = Get-Date -Format 'HH:mm:ss'
    $prefix = switch ($Level) {
        'INFO'  { '[INFO ]' }
        'WARN'  { '[WARN ]' }
        'ERROR' { '[ERROR]' }
        'OK'    { '[ OK  ]' }
    }
    $color  = switch ($Level) {
        'INFO'  { 'Cyan'   }
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red'    }
        'OK'    { 'Green'  }
    }
    Write-Host "$ts $prefix $Message" -ForegroundColor $color
}

#endregion

#region -- Resolve target user profile ----------------------------------------

Write-BCLog '================================================'
Write-BCLog ' Clear-BCOneNoteCache  v1.0.0'
Write-BCLog ' BriComp IT Consulting Services'
Write-BCLog '================================================'
Write-BCLog ''

$ProfilePath = $null

if ($PSBoundParameters.ContainsKey('Username')) {
    $ProfilePath = "C:\Users\$Username"
    if (-not (Test-Path $ProfilePath)) {
        Write-BCLog "Profile path not found: $ProfilePath" -Level ERROR
        exit 1
    }
    Write-BCLog "Targeting explicit user profile: $ProfilePath"
}
else {
    $currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

    if ($currentIdentity -notmatch '^(NT AUTHORITY\\|SYSTEM$)') {
        $ProfilePath = $env:USERPROFILE
        Write-BCLog "Running in user context: $currentIdentity"
    }
    else {
        Write-BCLog 'Running as SYSTEM -- attempting to detect logged-on user...' -Level WARN
        try {
            $loggedOnFull = (Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).UserName
            if ($loggedOnFull -match '\\(.+)$') {
                $resolvedUser = $Matches[1]
                $ProfilePath  = "C:\Users\$resolvedUser"
                Write-BCLog "Detected logged-on user: $resolvedUser" -Level OK
            }
            else {
                Write-BCLog 'No interactive user session detected. Supply -Username explicitly.' -Level ERROR
                exit 1
            }
        }
        catch {
            Write-BCLog "WMI query failed: $($_.Exception.Message)" -Level ERROR
            Write-BCLog 'Supply -Username explicitly and re-run.' -Level ERROR
            exit 1
        }
    }
}

if (-not (Test-Path $ProfilePath)) {
    Write-BCLog "Resolved profile path does not exist: $ProfilePath" -Level ERROR
    exit 1
}

Write-BCLog "Profile path: $ProfilePath" -Level OK

#endregion

#region -- Locate OneNote 16.0 base directory ---------------------------------

$OneNoteBase = Join-Path $ProfilePath 'AppData\Local\Microsoft\OneNote\16.0'

Write-BCLog ''
Write-BCLog "OneNote cache base: $OneNoteBase"

if (-not (Test-Path $OneNoteBase)) {
    Write-BCLog 'OneNote 16.0 cache directory not found.' -Level WARN
    Write-BCLog 'OneNote may not be installed, or has never been launched for this user.' -Level WARN
    exit 0
}

$TargetFolders = @(
    'cache'
    'FullTextSearchIndex'
    'MasterIndex'
    'ServerListings'
)

#endregion

#region -- Stop OneNote -------------------------------------------------------

Write-BCLog ''
Write-BCLog 'Checking for running OneNote processes...'

$oneNoteProcs = Get-Process -Name 'ONENOTE' -ErrorAction SilentlyContinue

if ($oneNoteProcs) {
    Write-BCLog "Found $($oneNoteProcs.Count) OneNote process(es) -- stopping." -Level WARN
    if ($PSCmdlet.ShouldProcess('ONENOTE.EXE', 'Stop-Process -Force')) {
        $oneNoteProcs | Stop-Process -Force
        Start-Sleep -Seconds 3
        $stillRunning = Get-Process -Name 'ONENOTE' -ErrorAction SilentlyContinue
        if ($stillRunning) {
            Write-BCLog 'OneNote process still running after stop attempt. Proceeding anyway.' -Level WARN
        }
        else {
            Write-BCLog 'OneNote stopped successfully.' -Level OK
        }
    }
}
else {
    Write-BCLog 'OneNote is not running.' -Level OK
}

#endregion

#region -- Clear cache folders ------------------------------------------------

Write-BCLog ''
Write-BCLog 'Clearing cache folders...'

$totalItems  = 0
$totalErrors = 0

foreach ($folder in $TargetFolders) {
    $fullPath = Join-Path $OneNoteBase $folder

    if (-not (Test-Path $fullPath)) {
        Write-BCLog "  [$folder] Not present -- skipping."
        continue
    }

    $items = Get-ChildItem -Path $fullPath -Recurse -Force -ErrorAction SilentlyContinue
    $count = ($items | Measure-Object).Count

    if ($count -eq 0) {
        Write-BCLog "  [$folder] Already empty -- skipping." -Level OK
        continue
    }

    Write-BCLog "  [$folder] $count item(s) found."

    if ($PSCmdlet.ShouldProcess($fullPath, "Remove all contents")) {
        try {
            Remove-Item -Path (Join-Path $fullPath '*') -Recurse -Force -ErrorAction Stop
            Write-BCLog "  [$folder] Cleared $count item(s)." -Level OK
            $totalItems += $count
        }
        catch {
            Write-BCLog "  [$folder] Error: $($_.Exception.Message)" -Level ERROR
            $totalErrors++
        }
    }
}

#endregion

#region -- Summary ------------------------------------------------------------

Write-BCLog ''
Write-BCLog '------------------------------------------------'

if ($totalErrors -eq 0) {
    Write-BCLog "Cache cleared successfully. Total items removed: $totalItems" -Level OK
}
else {
    Write-BCLog "Cache clear completed with $totalErrors error(s). Items removed: $totalItems" -Level WARN
    Write-BCLog 'Review output above for details.' -Level WARN
}

#endregion

#region -- Optional relaunch --------------------------------------------------

if ($Relaunch) {
    Write-BCLog ''

    $currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    if ($currentIdentity -match '^(NT AUTHORITY\\|SYSTEM$)') {
        Write-BCLog '-Relaunch specified but cannot launch OneNote from SYSTEM context.' -Level WARN
        Write-BCLog 'Ask the user to relaunch OneNote manually.' -Level WARN
    }
    else {
        $candidatePaths = @(
            'C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE'
            'C:\Program Files (x86)\Microsoft Office\root\Office16\ONENOTE.EXE'
        )
        $oneNoteBin = $candidatePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

        if ($oneNoteBin) {
            Write-BCLog "Relaunching OneNote: $oneNoteBin"
            if ($PSCmdlet.ShouldProcess($oneNoteBin, 'Start-Process')) {
                Start-Process -FilePath $oneNoteBin
                Write-BCLog 'OneNote launched.' -Level OK
            }
        }
        else {
            Write-BCLog 'OneNote executable not found at expected M365 paths. Skipping relaunch.' -Level WARN
            Write-BCLog 'Launch OneNote manually.' -Level WARN
        }
    }
}

Write-BCLog ''
Write-BCLog 'Done.'

#endregion

# SIG # Begin signature block
# MIIobgYJKoZIhvcNAQcCoIIoXzCCKFsCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB2jfjYNT5Ggm/1
# BMNAlxotGtWMNqmKyXmF1cKpFImzmKCCIWswggWNMIIEdaADAgECAhAOmxiO+dAt
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
# BDEiBCDM0LA29m7445v30yNyJfttEvxwADaRaJnHG0F3n4KOODANBgkqhkiG9w0B
# AQEFAASCAgDOdwtgkfrmm/s1la8VfNroADexhXhXA1PQWb+7l9oZ7dC3U/9HAagS
# PvDI3Zw+ZvILtUUFODxb1/2TgS+qAq23dYqSlVo/qizLrdWpSdNv5FskuO0vCfvq
# BfTsyBgt/guQVEC8dvsVz/PD7c9FiyQ/r8Min4OR9WEHc2vAaa7gv8qt5kBSaHsN
# rwgYmmCKcG1xqAOvqm6lYBX0nDapPo54ZaVefdUw2C16GTAg3PirHcYf1o2i3Kuf
# ExAHJYuAG+D2uD2lrD4oSvfhtSmbsNKVEpWTmkE04e1DyEzvebT1Z6FiToZBZG63
# 3gwT08oY6L97tt6yCTw+W5sHSCwMB1c0MXhncAeqsky5uIywTjKYEyxa/HKlCJ58
# zOgNc8yxrQ4WQaLt7zJG0eWPSIeGEtcRzeleT/1hrMy/Rj7TKXyIWJexQmsT7M5s
# KRJCv+ld1t/4RZMmPzMbCNorMG91ZkSWQN0Q44qLF295fiDjnH4aJBZ0AWIduhoS
# k3EGbhh4HPk6ReUuN4GHYoCsVABkWbYyv+qp4oAlOjygzCuE6Hcm8O9Iyp4abwfK
# 3ikPdZcDO+IM1XIK6aqqWEwykUifO/5ViVTkyFB6dnSWd1xYun/aaIWA31aOjTwu
# VmUHkF+Ql2llbSk/SoUA4AETaCGBjo02Uw/sg9bfqcJwiZBel9Ou8qGCAyYwggMi
# BgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQK
# Ew5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBU
# aW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAhP3DNPfkVO28MP
# j/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcB
# MBwGCSqGSIb3DQEJBTEPFw0yNjA5MTgxNjE3MDRaMC8GCSqGSIb3DQEJBDEiBCA2
# ZjB0dyfjPsSL3IUVPjrPdTxURVMlX8VERiGJD/uAOjANBgkqhkiG9w0BAQEFAASC
# AgCy5v5/m80UaZFAZx7WgOpuzfkQGfEzmyxLjHHaG3NNzg0ejRQh83z7PhdJzBrs
# mjr1wT6r3fDHynMgSb+BL6SthQVIHNUIFoHXyrp1ZYHej+bGmDtlQJuNLdw0BoM3
# hK+jd1GlICv7sTjahbdlOfX6E318QWrxQgerjwQ7B6UU5P06MsnYwnrSH24eYJ3G
# k7s8OylWUM1lznUkm+XgMc+t0zevE6Tc6YbsBOBqA08j6cn2Ksu0VOtMaFanKi69
# zC3PW/l5JteGsspfgyFADUqRM+fOKx7MGz6ROqwvM/08PDHl5MgYavJSjC+TkRIK
# gJzNYAylhQnIyjg7WzYegq1C9U+kTPfBu+6VR1m+JKa2Dy/KQbYBVVrJLCNQ4qkA
# LCyhYl7a3A2ND0r3T8yN/H3Cd7VXUjpp2cVlVrpu+GcCYgjy41BG4EvQfa7xL7/Y
# eZXMwyaMgujM0/wTOl3XAw7YrEiujdbTs8aVmk4jRw0yc+rRFyou/HAPFN+MgIzU
# OcJ5c8k8f5m4RMCGFkAiVokUbllvkgfw4K/eW/+ZFimJrW0E8xGwkZfTH63ANXbe
# CiWA/RFYE3mzj1OvAI1iDetZNjsYjvj3cClxq0YfGTfansrXfKsPKSruybrbEa1L
# uUO4LxdOZhx+iGDMDXTemTGQLXWVBwv+Ag5BWlniedPEHQ==
# SIG # End signature block
