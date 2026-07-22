#Requires -Version 5.1
<#
.SYNOPSIS
    New-BCClient.ps1 — Scaffold a new client into the bricomp-scripts repository.

.DESCRIPTION
    Creates the full folder structure, module stub, and README for a new client,
    then patches BC.Utils.psm1 to register the new module in Import-BCModule's
    module map. Run from the root of the bricomp-scripts repo.

    What gets created:
        clients/<slug>/
            <PREFIX>.Functions.psm1     — client module stub
            README.md                   — client-specific notes
        clients/<slug>/sccm/            — placeholder for SCCM scripts
        clients/<slug>/intune/          — placeholder for Intune scripts
        clients/<slug>/reports/         — gitignored output folder

    What gets patched:
        shared/utils/BC.Utils.psm1      — adds entry to $moduleMap
        README.md (repo root)           — adds client to the structure table

.PARAMETER ClientName
    Full client name (e.g. 'Acme Corporation'). Used in headers and README.

.PARAMETER ShortCode
    2–5 character prefix used for script naming (e.g. 'ACM').
    Must be uppercase letters only. Will be used as: ACM.Functions.psm1

.PARAMETER Slug
    Folder name under clients/ (e.g. 'acme'). Defaults to ShortCode.ToLower().

.PARAMETER SiteCode
    SCCM site code if applicable (e.g. 'ACM'). Defaults to ShortCode.

.PARAMETER PrimaryContact
    Client's primary IT contact name. Optional — goes into the module header.

.PARAMETER RepoRoot
    Path to the repo root. Defaults to the current directory.

.PARAMETER WhatIf
    Preview all actions without writing any files.

.EXAMPLE
    .\New-BCClient.ps1 -ClientName 'Acme Corporation' -ShortCode 'ACM'

.EXAMPLE
    .\New-BCClient.ps1 -ClientName 'Northbrook Logistics' -ShortCode 'NBL' -Slug 'northbrook' -PrimaryContact 'Jane Smith' -WhatIf

.NOTES
    BriComp Computers, LLC
    After running, commit the new files and the BC.Utils.psm1 patch:
        git add -A
        git commit -m "feat(clients): scaffold <ClientName> (<ShortCode>)"
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$ClientName,

    [Parameter(Mandatory)]
    [ValidatePattern('^[A-Z]{2,5}$')]
    [string]$ShortCode,

    [string]$Slug,

    [string]$SiteCode,

    [string]$PrimaryContact = '',

    [string]$RepoRoot = (Get-Location).Path
)

$ErrorActionPreference = 'Stop'

#region --- Resolve defaults ---
if (-not $Slug)     { $Slug     = $ShortCode.ToLower() }
if (-not $SiteCode) { $SiteCode = $ShortCode }

$clientDir  = Join-Path $RepoRoot "clients\$Slug"
$moduleFile = Join-Path $clientDir "$ShortCode.Functions.psm1"
$readmeFile = Join-Path $clientDir 'README.md'
$utilsFile  = Join-Path $RepoRoot 'shared\utils\BC.Utils.psm1'
$rootReadme = Join-Path $RepoRoot 'README.md'
$date       = Get-Date -Format 'yyyy-MM-dd'
#endregion

#region --- Pre-flight checks ---
if (-not (Test-Path (Join-Path $RepoRoot '.git'))) {
    Write-Error "RepoRoot '$RepoRoot' does not appear to be a git repository. Run from the repo root or pass -RepoRoot."
}

if (Test-Path $clientDir) {
    Write-Warning "Client folder already exists: $clientDir"
    $confirm = Read-Host "Overwrite existing files? (y/N)"
    if ($confirm -notmatch '^[Yy]') {
        Write-Host "Aborted." -ForegroundColor Yellow
        exit 0
    }
}

# Check BC.Utils.psm1 exists and already has the module map
if (-not (Test-Path $utilsFile)) {
    Write-Error "BC.Utils.psm1 not found at: $utilsFile. Check -RepoRoot."
}

$utilsContent = Get-Content $utilsFile -Raw
if ($utilsContent -notmatch '\$moduleMap\s*=\s*@\{') {
    Write-Error "BC.Utils.psm1 does not contain the expected `$moduleMap block. Cannot auto-patch."
}

# Check for duplicate slug/shortcode
if ($utilsContent -match "'$($Slug.ToLower())'") {
    Write-Warning "BC.Utils.psm1 already has an entry for '$($Slug.ToLower())'. The patch step will be skipped."
    $skipPatch = $true
}
#endregion

#region --- Summary ---
Write-Host "`n=== New-BCClient.ps1 ===" -ForegroundColor Cyan
Write-Host "Client name  : $ClientName"
Write-Host "Short code   : $ShortCode"
Write-Host "Slug         : $Slug"
Write-Host "Site code    : $SiteCode"
Write-Host "Contact      : $(if ($PrimaryContact) { $PrimaryContact } else { '(not specified)' })"
Write-Host "Client dir   : $clientDir"
Write-Host "Module file  : $moduleFile"
Write-Host ""
if ($WhatIfPreference) {
    Write-Host "[WHATIF] No files will be written." -ForegroundColor Yellow
    Write-Host ""
}
#endregion

#region --- Create folder structure ---
$dirs = @(
    $clientDir,
    (Join-Path $clientDir 'sccm'),
    (Join-Path $clientDir 'intune'),
    (Join-Path $clientDir 'reports')
)

foreach ($dir in $dirs) {
    if ($PSCmdlet.ShouldProcess($dir, 'Create directory')) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-Host "  [DIR]  $dir" -ForegroundColor DarkCyan
    }
}

# .gitkeep placeholders so git tracks the empty dirs
foreach ($subDir in @('sccm','intune')) {
    $keepFile = Join-Path $clientDir "$subDir\.gitkeep"
    if ($PSCmdlet.ShouldProcess($keepFile, 'Create .gitkeep')) {
        Set-Content -Path $keepFile -Value '' -Encoding UTF8
        Write-Host "  [FILE] $keepFile" -ForegroundColor DarkGray
    }
}
#endregion

#region --- Generate client module stub ---
$moduleStub = @"
#Requires -Version 5.1
<#
.SYNOPSIS
    $ShortCode.Functions.psm1 — BriComp client module for $ClientName.

.DESCRIPTION
    Client-specific functions for $ClientName.
    Add environment-specific functions here as the engagement grows.

    Load via:
        Import-BCModule -Name $($Slug.ToLower())

.NOTES
    BriComp Computers, LLC
    Client     : $ClientName
    Short code : $ShortCode
    Site code  : $SiteCode (SCCM)
    Created    : $date$(if ($PrimaryContact) { "`n    Contact    : $PrimaryContact" })
#>

#region --- Module-level config ---
`$script:${ShortCode}Config = @{
    SiteCode        = '$SiteCode'
    SiteServer      = ''                     # TODO: populate primary SCCM/site server
    TenantId        = ''                     # TODO: populate M365 tenant ID
    SCCMCredTarget  = '$($Slug.ToLower())-sccm-admin'
    IntuneCredTarget= '$($Slug.ToLower())-intune-svc'
}
#endregion

#region --- Connect-${ShortCode}Site ---
function Connect-${ShortCode}Site {
    <#
    .SYNOPSIS
        Initialize the SCCM PSDrive for the $ClientName site ($SiteCode`:).
    .EXAMPLE
        Connect-${ShortCode}Site
        Set-Location ${SiteCode}:
    #>
    [CmdletBinding()]
    param()

    if (-not (Get-Module ConfigurationManager -ErrorAction SilentlyContinue)) {
        `$cmPath = "`$env:SMS_ADMIN_UI_PATH\..\ConfigurationManager.psd1"
        if (-not (Test-Path `$cmPath)) {
            Write-Error "ConfigurationManager module not found. Run from a machine with the SCCM Admin Console installed."
            return
        }
        Import-Module `$cmPath -ErrorAction Stop
    }

    if (-not (Get-PSDrive -Name `$script:${ShortCode}Config.SiteCode -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name `$script:${ShortCode}Config.SiteCode ``
                    -PSProvider CMSite ``
                    -Root `$script:${ShortCode}Config.SiteServer ``
                    -ErrorAction Stop | Out-Null
        Write-BCLog "Connected to SCCM site `$(`$script:${ShortCode}Config.SiteCode) on `$(`$script:${ShortCode}Config.SiteServer)" -Level INFO
    } else {
        Write-BCLog "PSDrive `$(`$script:${ShortCode}Config.SiteCode): already mapped." -Level DEBUG
    }
}
Export-ModuleMember -Function Connect-${ShortCode}Site
#endregion

#region --- Get-${ShortCode}DeviceInfo ---
function Get-${ShortCode}DeviceInfo {
    <#
    .SYNOPSIS
        Query SCCM for a $ClientName device's co-management state and last check-in.
    .PARAMETER ComputerName
        One or more device names to query.
    .EXAMPLE
        Get-${ShortCode}DeviceInfo -ComputerName '${ShortCode}-LPT-0001'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [string[]]`$ComputerName
    )

    process {
        Connect-${ShortCode}Site
        `$prev = Get-Location
        Set-Location "`$(`$script:${ShortCode}Config.SiteCode):"
        try {
            foreach (`$name in `$ComputerName) {
                Get-CMDevice -Name `$name -Fast |
                    Select-Object Name, ResourceID, IsClient, CoManaged,
                                  LastActiveTime, PrimaryUser, DeviceOS
            }
        } finally {
            Set-Location `$prev
        }
    }
}
Export-ModuleMember -Function Get-${ShortCode}DeviceInfo
#endregion
"@

if ($PSCmdlet.ShouldProcess($moduleFile, 'Create client module stub')) {
    Set-Content -Path $moduleFile -Value $moduleStub -Encoding UTF8
    Write-Host "  [FILE] $moduleFile" -ForegroundColor Green
}
#endregion

#region --- Generate client README ---
$clientReadme = @"
# $ClientName ($ShortCode)

BriComp managed client. Environment notes and runbook links go here.

## Environment Summary

| Property      | Value                        |
|---------------|------------------------------|
| Short code    | $ShortCode                   |
| SCCM site     | $SiteCode                    |
| Site server   | *(TODO)*                     |
| M365 tenant   | *(TODO)*                     |
| Primary contact | $(if ($PrimaryContact) { $PrimaryContact } else { '*(TODO)*' }) |
| Onboarded     | $date                        |

## Module

```powershell
Import-BCModule -Name $($Slug.ToLower())
```

## Key Scripts

| Script | Purpose |
|--------|---------|
| `$ShortCode.Functions.psm1` | Core client functions (Connect, DeviceInfo) |

## Notes

*(Add client-specific runbook notes, known issues, and environment quirks here.)*
"@

if ($PSCmdlet.ShouldProcess($readmeFile, 'Create client README')) {
    Set-Content -Path $readmeFile -Value $clientReadme -Encoding UTF8
    Write-Host "  [FILE] $readmeFile" -ForegroundColor Green
}
#endregion

#region --- Patch BC.Utils.psm1 ---
if (-not $skipPatch) {
    # Find the last client entry line in $moduleMap and insert after it
    # Pattern: any line like   'xyz'  = 'clients\...'
    $newEntry = "        '$($Slug.ToLower())'  = 'clients\\$Slug\\$ShortCode.Functions.psm1'"

    # Insert before the closing of the Client modules block
    # We look for the 'bricomp' entry (always last) and insert before it
    $insertAfterPattern = "        'bricomp' = 'clients\\bricomp\\BC.Internal.psm1'"

    if ($utilsContent -match [regex]::Escape($insertAfterPattern)) {
        $patchedContent = $utilsContent.Replace(
            $insertAfterPattern,
            "$newEntry`n$insertAfterPattern"
        )

        if ($PSCmdlet.ShouldProcess($utilsFile, "Patch `$moduleMap with '$($Slug.ToLower())'")) {
            Set-Content -Path $utilsFile -Value $patchedContent -Encoding UTF8 -NoNewline
            Write-Host "  [PATCH] BC.Utils.psm1 — added '$($Slug.ToLower())' to `$moduleMap" -ForegroundColor Magenta
        }
    } else {
        Write-Warning "Could not locate insertion point in BC.Utils.psm1. Add this line manually to `$moduleMap:"
        Write-Warning "  $newEntry"
    }
}
#endregion

#region --- Summary ---
Write-Host "`n--- Done ---" -ForegroundColor Green
if (-not $WhatIfPreference) {
    Write-Host @"

Next steps:
  1. Populate `$script:${ShortCode}Config in $ShortCode.Functions.psm1
       (SiteServer, TenantId, etc.)
  2. Add client-specific functions to the module
  3. Commit:
       git add -A
       git commit -m "feat(clients): scaffold $ClientName ($ShortCode)"
  4. Test:
       Import-BCModule -Name $($Slug.ToLower())
       Connect-${ShortCode}Site
"@
}
#endregion
