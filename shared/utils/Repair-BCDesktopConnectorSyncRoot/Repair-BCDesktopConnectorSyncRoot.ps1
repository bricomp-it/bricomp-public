# PUBLISH: true
#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Repairs Autodesk Desktop Connector "Unable to register drive" startup failures
    caused by orphaned Cloud Files API sync root registrations.

.DESCRIPTION
    Desktop Connector registers its virtual drive through the Windows Cloud Files API.
    Every registration is stored under:

        HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\SyncRootManager

    and is named using the pattern:

        <Provider>!<Base64 AccountId>!<User SID>!User

    Windows does not permit two sync root trees to overlap. When a machine is migrated
    in a way that changes the user's SID - an Entra-joined device moved to on-premises
    Active Directory, a cross-domain migration, or a rebuilt account - the previous
    registration remains behind and keeps its claim on the workspace folder. The live
    account's registration is then refused and Desktop Connector exits at startup with:

        Failed to register virtual file system for <name>
        CloudFileApi.CloudFileApiException: Access is denied.
            (Exception from HRESULT: 0x80070005 E_ACCESSDENIED)
        Tray Shutdown - ExitSource = "Unable to register drive"

    Creating a new Windows profile does not resolve this. The registrations live in
    HKLM and are keyed by SID, not stored inside the user profile.

    This script inventories every Desktop Connector sync root, classifies each one,
    and removes only those that are demonstrably stale:

        ORPHANED    The owning SID has no profile on this machine. This is the
                    registration that blocks the live account.      Removed.
        INCOMPLETE  Live SID, but the registration has no claimed path - debris
                    from a failed registration attempt.             Removed.
        HEALTHY     Live SID with a valid claimed path.              Left alone.

    An INCOMPLETE registration is only removed when an ORPHANED one is also present,
    since on its own it may simply mean Desktop Connector is mid-registration.

    The entire SyncRootManager key is exported to a .reg backup before any change is
    made, and the script aborts if that backup cannot be written. Only keys beginning
    with "DesktopConnector!" are ever touched - OneDrive and every other cloud storage
    provider are left untouched.

    On a healthy machine the script makes no changes and does not restart.

.PARAMETER DryRun
    Report what would be removed and make no changes. No restart. Run this first.

.PARAMETER NoRestart
    Apply the repair but do not restart. The repair does not take effect until the
    machine is restarted.

.PARAMETER RestartDelaySeconds
    Seconds to wait before restarting, so a signed-in user can save their work.
    Default 120.

.PARAMETER LogPath
    Override the log and backup directory.
    Default C:\ProgramData\BriComp\DCSyncRootFix

.EXAMPLE
    .\Repair-BCDesktopConnectorSyncRoot.ps1 -DryRun

    Reports what would be removed and exits. Makes no changes.

.EXAMPLE
    .\Repair-BCDesktopConnectorSyncRoot.ps1

    Repairs the machine and restarts it after a two minute warning.

.EXAMPLE
    .\Repair-BCDesktopConnectorSyncRoot.ps1 -NoRestart

    Repairs the machine and leaves the restart to you or your RMM.

.NOTES
    Version:    1.0.0
    Author:     BriComp IT Consulting Services
    Website:    https://bricomp.com
    Created:    2026-09-17

    Must be run elevated. Safe to run against healthy machines - they are no-ops.

    Exit codes:
        0   No action needed (healthy, or Desktop Connector not installed)
        1   Error
        2   Not running elevated
        3   Changes applied - restart pending or underway

    Verified against Desktop Connector 2027.2.2.3 on Windows 11 23H2, on a device
    migrated from Entra join to on-premises Active Directory.
#>

[CmdletBinding()]
param(
    [switch] $DryRun,
    [switch] $NoRestart,
    [int]    $RestartDelaySeconds = 120,
    [string] $LogPath = 'C:\ProgramData\BriComp\DCSyncRootFix'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$SyncRootBase   = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\SyncRootManager'
$SyncRootBaseNp = 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\SyncRootManager'
$ProfileListKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
$ProviderPrefix = 'DesktopConnector!'

$script:Stamp      = Get-Date -Format 'yyyyMMdd-HHmmss'
$script:LogFile    = Join-Path $LogPath ("DCSyncRootFix-{0}-{1}.log" -f $env:COMPUTERNAME, $script:Stamp)
$script:Changed    = $false
$script:ErrorCount = 0

#--------------------------------------------------------------------------------------
# Logging
#--------------------------------------------------------------------------------------
function Write-Log {
    param(
        [Parameter(Mandatory)][string] $Message,
        [ValidateSet('INFO','WARN','ERROR','ACTION','SKIP')][string] $Level = 'INFO'
    )
    $line = '{0}  [{1,-6}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    switch ($Level) {
        'ERROR'  { Write-Host $line -ForegroundColor Red }
        'WARN'   { Write-Host $line -ForegroundColor Yellow }
        'ACTION' { Write-Host $line -ForegroundColor Cyan }
        'SKIP'   { Write-Host $line -ForegroundColor DarkGray }
        default  { Write-Host $line }
    }
    try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 } catch { }
}

#--------------------------------------------------------------------------------------
# Preflight
#--------------------------------------------------------------------------------------
try { New-Item -Path $LogPath -ItemType Directory -Force | Out-Null } catch { }

Write-Log "=============================================================="
Write-Log "Desktop Connector sync root repair"
Write-Log ("Computer : {0}" -f $env:COMPUTERNAME)
Write-Log ("Running as: {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
Write-Log ("Mode      : {0}" -f $(if ($DryRun) { 'DRY RUN - no changes' } else { 'REMEDIATE' }))
Write-Log ("Log file  : {0}" -f $script:LogFile)
Write-Log "=============================================================="

$identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Log "Not running elevated. Re-run as administrator." 'ERROR'
    exit 2
}

# Is Desktop Connector present at all?
$dcInstalled = $false
foreach ($root in @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall')) {
    if (Test-Path -LiteralPath $root) {
        $hit = Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue |
               ForEach-Object { (Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue).DisplayName } |
               Where-Object { $_ -like '*Desktop Connector*' }
        if ($hit) { $dcInstalled = $true; break }
    }
}
if (-not $dcInstalled -and (Test-Path -LiteralPath $SyncRootBase)) {
    if (Get-ChildItem -LiteralPath $SyncRootBase -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -like "$ProviderPrefix*" }) { $dcInstalled = $true }
}
if (-not $dcInstalled) {
    Write-Log "Autodesk Desktop Connector is not installed on this machine. Nothing to do." 'SKIP'
    exit 0
}

if (-not (Test-Path -LiteralPath $SyncRootBase)) {
    Write-Log "SyncRootManager key does not exist. Nothing to do." 'SKIP'
    exit 0
}

#--------------------------------------------------------------------------------------
# Gather context: which SIDs actually have a profile on this machine
#--------------------------------------------------------------------------------------
$knownSids = @{}
foreach ($pf in @(Get-ChildItem -LiteralPath $ProfileListKey -ErrorAction SilentlyContinue)) {
    $sid  = $pf.PSChildName
    $path = $null
    try { $path = (Get-ItemProperty -LiteralPath $pf.PSPath -ErrorAction Stop).ProfileImagePath } catch { }
    $knownSids[$sid] = $path
}
Write-Log ("Profiles present on this machine: {0}" -f $knownSids.Count)
foreach ($s in $knownSids.Keys) { Write-Log ("   {0}  ->  {1}" -f $s, $knownSids[$s]) }

# Loaded user hives, for shell namespace cleanup
$loadedHives = @()
foreach ($hv in @(Get-ChildItem -LiteralPath 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue)) {
    if ($hv.PSChildName -match '^S-1-(5-21|12-1)-' -and $hv.PSChildName -notlike '*_Classes') {
        $loadedHives += $hv.PSChildName
    }
}
Write-Log ("Loaded user hives: {0}" -f $(if ($loadedHives.Count) { $loadedHives -join ', ' } else { '(none)' }))

#--------------------------------------------------------------------------------------
# Classify every DesktopConnector sync root
#--------------------------------------------------------------------------------------
$candidates = @()

$dcKeys = @(Get-ChildItem -LiteralPath $SyncRootBase -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -like "$ProviderPrefix*" })

foreach ($dcKey in $dcKeys) {

    $keyName = $dcKey.PSChildName
    $keyPath = Join-Path $SyncRootBase $keyName

    # SID is the segment that looks like a SID
    $sid = ($keyName -split '!' | Where-Object { $_ -match '^S-1-\d+' } | Select-Object -First 1)
    if (-not $sid) {
        Write-Log ("Cannot parse a SID out of key '{0}' - leaving it alone." -f $keyName) 'WARN'
        continue
    }

    # Read the claimed workspace paths and the shell namespace CLSID
    $paths = @()
    $usrPath = Join-Path $keyPath 'UserSyncRoots'
    if (Test-Path -LiteralPath $usrPath) {
        try {
            $k = Get-Item -LiteralPath $usrPath
            foreach ($vn in $k.GetValueNames()) {
                $v = $k.GetValue($vn)
                if ($v) { $paths += [string]$v }
            }
        } catch { }
    }

    $clsid = $null
    try { $clsid = (Get-ItemProperty -LiteralPath $keyPath -ErrorAction Stop).NamespaceCLSID } catch { }

    $hasProfile = $knownSids.ContainsKey($sid)

    # ---- Classification -------------------------------------------------------------
    # ORPHANED   : the SID owning this registration has no profile on this machine.
    #              Typical after a migration that changed the account's SID. It still
    #              holds the path claim, which blocks the live account. REMOVE.
    # INCOMPLETE : the SID is live, but the registration never finished (no claimed
    #              path). Debris from a failed Register call. REMOVE so DC rebuilds it.
    # HEALTHY    : live SID with a claimed path. LEAVE ALONE.
    if (-not $hasProfile) {
        $verdict = 'ORPHANED'
    } elseif ($paths.Count -eq 0) {
        $verdict = 'INCOMPLETE'
    } else {
        $verdict = 'HEALTHY'
    }

    $candidates += [pscustomobject]@{
        KeyName    = $keyName
        KeyPath    = $keyPath
        Sid        = $sid
        HasProfile = $hasProfile
        Paths      = $paths
        Clsid      = $clsid
        Verdict    = $verdict
    }
}

Write-Log "--------------------------------------------------------------"
Write-Log ("Desktop Connector sync roots found: {0}" -f $candidates.Count)
foreach ($c in $candidates) {
    Write-Log ("  [{0,-10}] {1}" -f $c.Verdict, $c.KeyName)
    Write-Log ("               SID    : {0}  (profile on box: {1})" -f $c.Sid, $c.HasProfile)
    Write-Log ("               Paths  : {0}" -f $(if ($c.Paths.Count) { $c.Paths -join '; ' } else { '(none)' }))
    Write-Log ("               CLSID  : {0}" -f $(if ($c.Clsid) { $c.Clsid } else { '(none)' }))
}
Write-Log "--------------------------------------------------------------"

$orphaned   = @($candidates | Where-Object { $_.Verdict -eq 'ORPHANED' })
$incomplete = @($candidates | Where-Object { $_.Verdict -eq 'INCOMPLETE' })

# An INCOMPLETE key on its own is not proof of the defect - Desktop Connector may simply
# be part-way through a first registration right now. Only clean those up when an
# ORPHANED key is also present, which is the actual fault signature.
if ($orphaned.Count -eq 0) {
    if ($incomplete.Count -gt 0) {
        Write-Log ("{0} incomplete registration(s) found, but no orphaned SID." -f $incomplete.Count) 'WARN'
        Write-Log "That is not the fault signature this script repairs - Desktop Connector may" 'WARN'
        Write-Log "simply be mid-registration. No changes made. Re-run later if the fault persists." 'WARN'
    } else {
        Write-Log "No orphaned registrations. This machine is healthy - no changes made." 'SKIP'
    }
    exit 0
}

$toRemove = @($orphaned + $incomplete)

Write-Log ("{0} registration(s) will be removed." -f $toRemove.Count) 'ACTION'

if ($DryRun) {
    Write-Log "DRY RUN - stopping here. No backup taken, no changes made, no restart." 'SKIP'
    exit 0
}

#--------------------------------------------------------------------------------------
# Back up before touching anything
#--------------------------------------------------------------------------------------
$backupFile = Join-Path $LogPath ("SyncRootManager-{0}-{1}.reg" -f $env:COMPUTERNAME, $script:Stamp)
Write-Log ("Backing up SyncRootManager to: {0}" -f $backupFile) 'ACTION'
$null = & reg.exe export "$SyncRootBaseNp" "$backupFile" /y 2>&1
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $backupFile)) {
    Write-Log "Registry backup FAILED. Aborting without making any changes." 'ERROR'
    exit 1
}
Write-Log ("Backup written ({0:N0} bytes)." -f (Get-Item -LiteralPath $backupFile).Length)

#--------------------------------------------------------------------------------------
# Stop Desktop Connector
#--------------------------------------------------------------------------------------
$procNames = @('DesktopConnector.Applications.Tray','DesktopConnector.Applications.Host','DesktopConnector')
foreach ($p in $procNames) {
    Get-Process -Name $p -ErrorAction SilentlyContinue | ForEach-Object {
        Write-Log ("Stopping process {0} (PID {1})" -f $_.Name, $_.Id) 'ACTION'
        try { Stop-Process -Id $_.Id -Force -ErrorAction Stop } catch { Write-Log ("  could not stop: {0}" -f $_.Exception.Message) 'WARN' }
    }
}
Start-Sleep -Seconds 3

#--------------------------------------------------------------------------------------
# Remove the offending registrations and their shell namespace entries
#--------------------------------------------------------------------------------------
foreach ($c in $toRemove) {

    Write-Log ("Removing [{0}] {1}" -f $c.Verdict, $c.KeyName) 'ACTION'

    # 1. Shell namespace entries, in every loaded user hive (Explorer may have created
    #    the node under a different user than the one the sync root is keyed to).
    if ($c.Clsid) {
        foreach ($hive in $loadedHives) {
            $targets = @(
                "Registry::HKEY_USERS\$hive\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\$($c.Clsid)",
                "Registry::HKEY_USERS\$hive\Software\Classes\CLSID\$($c.Clsid)",
                "Registry::HKEY_USERS\${hive}_Classes\CLSID\$($c.Clsid)"
            )
            foreach ($t in $targets) {
                if (Test-Path -LiteralPath $t) {
                    try {
                        Remove-Item -LiteralPath $t -Recurse -Force -ErrorAction Stop
                        Write-Log ("   removed {0}" -f $t) 'ACTION'
                    } catch {
                        Write-Log ("   FAILED to remove {0} : {1}" -f $t, $_.Exception.Message) 'ERROR'
                        $script:ErrorCount++
                    }
                }
            }
            # Hidden-desktop-icon value, if it exists
            $hdi = "Registry::HKEY_USERS\$hive\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
            if (Test-Path -LiteralPath $hdi) {
                try {
                    $props = Get-ItemProperty -LiteralPath $hdi -ErrorAction Stop
                    if ($props.PSObject.Properties.Name -contains $c.Clsid) {
                        Remove-ItemProperty -LiteralPath $hdi -Name $c.Clsid -Force -ErrorAction Stop
                        Write-Log ("   removed HideDesktopIcons value {0}" -f $c.Clsid) 'ACTION'
                    }
                } catch { }
            }
        }
        # Machine-wide class registration for this specific sync root, if present
        foreach ($t in @("HKLM:\SOFTWARE\Classes\CLSID\$($c.Clsid)",
                         "HKLM:\SOFTWARE\WOW6432Node\Classes\CLSID\$($c.Clsid)")) {
            if (Test-Path -LiteralPath $t) {
                try {
                    Remove-Item -LiteralPath $t -Recurse -Force -ErrorAction Stop
                    Write-Log ("   removed {0}" -f $t) 'ACTION'
                } catch {
                    Write-Log ("   FAILED to remove {0} : {1}" -f $t, $_.Exception.Message) 'ERROR'
                    $script:ErrorCount++
                }
            }
        }
    } else {
        Write-Log "   no NamespaceCLSID recorded - skipping shell namespace cleanup" 'SKIP'
    }

    # 2. The sync root registration itself
    try {
        Remove-Item -LiteralPath $c.KeyPath -Recurse -Force -ErrorAction Stop
        Write-Log ("   removed sync root key {0}" -f $c.KeyName) 'ACTION'
        $script:Changed = $true
    } catch {
        Write-Log ("   FAILED to remove sync root key {0} : {1}" -f $c.KeyName, $_.Exception.Message) 'ERROR'
        $script:ErrorCount++
    }
}

#--------------------------------------------------------------------------------------
# Wrap up
#--------------------------------------------------------------------------------------
Write-Log "--------------------------------------------------------------"
if ($script:ErrorCount -gt 0) {
    Write-Log ("Completed with {0} error(s). Review this log before restarting." -f $script:ErrorCount) 'ERROR'
}
if (-not $script:Changed) {
    Write-Log "No registry keys were actually removed." 'WARN'
    exit 1
}

Write-Log ("Backup of the original registry state: {0}" -f $backupFile)
Write-Log "To roll back: run the .reg backup above, then restart."

if ($NoRestart) {
    Write-Log "NoRestart specified. The fix takes effect after the next restart." 'WARN'
    exit 3
}

Write-Log ("Restarting in {0} seconds." -f $RestartDelaySeconds) 'ACTION'
$msg = "IT maintenance: Autodesk Desktop Connector is being repaired. This computer will restart in $([math]::Round($RestartDelaySeconds/60,1)) minute(s). Please save your work."
& shutdown.exe /r /t $RestartDelaySeconds /c "$msg" /d p:4:1 | Out-Null
exit 3

# SIG # Begin signature block
# MIIobgYJKoZIhvcNAQcCoIIoXzCCKFsCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC3aEnJGe7TtUxo
# robtK5szfzZOYwV137dWseSJqjQufKCCIWswggWNMIIEdaADAgECAhAOmxiO+dAt
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
# BDEiBCBc+B7r5TU/RkDHiE8G1WY7L0BpV6r2yEz1I4vP3j+duzANBgkqhkiG9w0B
# AQEFAASCAgDe98g6OFPI6AcNhEas+Ic4PY+Z70NLnVAoCRDjqHqYbE0me8djY/+b
# JV7ob2UfFQyl7dNUOgwvwRR7qdxluZGrEIT7FiDugZ5nazk2av61B0Rctodjnn47
# 7q1IYSp7ppexlo9ieavkiUtjHnp3mP2LV1ti0GdiiDiBcQelGKkw83A15Z/AWiOy
# 87DzpQi84A4AcGZbVyumI7G8Ob7oRSg5ZhzyE8B8spagKhZJMWDYTW0JHaB9VVY1
# tVFteCl3NrKs589SsU58z1n31fAnbcBvmllLLPFYAPh1InPAjdRC6Ph0uicp3fA7
# Sk4g+vOF+XrTeRtxDT2Jjn9txEHg5CMCdyqvesXQtB9nv+jZwvqWyVNaDHumGu24
# wNk/F0rO+HkvQrfVwP7sa8OGaMYkcywQgu/auC2IwWGIAOUO2Sq+tu7IwgZ5cR7l
# iWPlX4edIYoheIvPPFWzCsKjQpFqkMYjKFjbxeREIUA+0S96LT3X+e1yP21uNhQk
# tTS3TplKpSKhauMq2oWcgs+T/+kYo+A6xOYFOYFuohJbKsjVVqo8YdyuF9fEi5yR
# 4NpfOTZWvJfA7SXYYl+gtddNqH/t1m9OQjpICCDq1a2+h0A9hD2y5OTEN0+TgwZW
# ROiHfp6vhLxvKClYbWjLau8gOKpPItjuXSGgjhYYJjSt6llP8tlRPKGCAyYwggMi
# BgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQK
# Ew5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBU
# aW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAhP3DNPfkVO28MP
# j/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcB
# MBwGCSqGSIb3DQEJBTEPFw0yNjA5MTgxNTUyMjBaMC8GCSqGSIb3DQEJBDEiBCBW
# e/GGd0SCmDr0RTospVpFDS2MJw8VialOJxfg7LuTNjANBgkqhkiG9w0BAQEFAASC
# AgCoJSB48FItNcERE5eJ9Ed0uKselqraN027kpWviBFl4+cbmAOmtZbfqIUXAKJA
# QmCAGv1corjjV2NhMs7u1OmFaFbLJYEJSxnL10PXRanQkdapdQdROeOk789jxXJo
# h2WgbtBz3BNBXOH3k9e8sKqZZgG6xKncgn5KuwSk8OtvVhrNC8eM0D9koJFxwmWu
# SpSzYLdMawg0VoWdxfsagpxFQFhN/KUOhjaYnzoiSttN158H+rpuWH6/VKK3PYCZ
# dkq7Vq6bTXaQgeysFQurPXG3Jzi0MqSloYCrK4HQ7ySCA87vL7uyBVzp2y9Oj6UJ
# 89UlIBs75ktHCXiEcDdM//ZW3Ok9W4gZLBys5F5Iuzjstx2D2QTEkF7J474Kk1Xo
# nh6xPI6hghgd2B3VUWm1BR7i9Gt49AhelxP4Suktgn5czl3ENUZtAoWzkyVmbbMC
# gNhgZdZP3jGOOd13jPOdxouZcJFmfOr4kWMMMXfrBFGZeZuqgIafepGpPVODz2Xq
# yzLZxo1GfGZWWHW2pkI1alK2Qn7giWsjTYpxf0pLCgvXCRWraUQOQGNa6FAbYbQI
# o7h8eVu9gI3R5eYVyq4Qb7rYOJL+ynvZ97BMRSnzee7bqUbFHJgq+NtWHmQwsFJA
# 5dDCmfp2yucTuwfOCy1rZxPRg++fS3vRlqQZFGHFvwCWSA==
# SIG # End signature block
