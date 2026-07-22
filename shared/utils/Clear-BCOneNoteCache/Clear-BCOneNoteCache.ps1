#Requires -Version 5.1

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
