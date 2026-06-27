#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Invoke-BCSignPublishable.ps1
    Find and Authenticode-sign all scripts marked for public release.

.DESCRIPTION
    Scans the bricomp-scripts repo for any .ps1 or .psm1 file containing
    '# PUBLISH: true' and signs each one with the BriComp code signing
    certificate. Safe to re-run - already-valid signatures are skipped
    unless -Force is specified.

    Prerequisites:
      - BriComp code signing certificate imported into the certificate store
      - Certificate must have Enhanced Key Usage: Code Signing (1.3.6.1.5.5.7.3.3)
      - Run as Administrator

    Workflow:
      1. Run this script to sign all publishable scripts
      2. Verify signatures with -Audit switch first if unsure
      3. Commit the signed files: git add -A; git commit -m "chore: sign publishable scripts"
      4. Push and tag to trigger the public sync GitHub Action

.PARAMETER RepoRoot
    Path to the bricomp-scripts repo root. Defaults to C:\Dev\bricomp-scripts.

.PARAMETER CertThumbprint
    Specific certificate thumbprint to use. If omitted, the script auto-selects
    the first valid BriComp code signing cert from the store.

.PARAMETER CertStore
    Certificate store to search. Default: CurrentUser\My
    Options: CurrentUser\My, LocalMachine\My

.PARAMETER TimestampServer
    RFC 3161 timestamp server URL. Default: http://timestamp.digicert.com
    Timestamping means the signature remains valid after cert expiry.

.PARAMETER Audit
    Scan and report signature status without signing anything. Use this first
    to preview what would be signed.

.PARAMETER Force
    Re-sign files that already have a valid signature. Default: skip valid sigs.

.PARAMETER WhatIf
    Show what would be signed without actually signing.

.EXAMPLE
    # Preview what would be signed
    .\Invoke-BCSignPublishable.ps1 -Audit

    # Sign all publishable scripts (auto-select cert)
    .\Invoke-BCSignPublishable.ps1

    # Sign with a specific cert thumbprint
    .\Invoke-BCSignPublishable.ps1 -CertThumbprint 'AB12CD34EF56...'

    # Force re-sign everything (e.g. after cert renewal)
    .\Invoke-BCSignPublishable.ps1 -Force

    # Preview without signing
    .\Invoke-BCSignPublishable.ps1 -WhatIf

.NOTES
    BriComp Computers, LLC
    Script version : 1.0.0
    Created        : 2026-06-26

    After signing, commit the changes before tagging for public release:
        git add -A
        git commit -m "chore: sign publishable scripts - $(Get-Date -Format yyyy-MM-dd)"
        git tag publish/shared/utils/DeviceMenu -m "Publish DeviceMenu v2.0.0"
        git push origin --tags
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$RepoRoot        = 'C:\Dev\bricomp-scripts',
    [string]$CertThumbprint  = '',
    [string]$CertStore       = 'CurrentUser\My',
    [string]$TimestampServer = 'http://timestamp.digicert.com',
    [switch]$Audit,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$ScriptVersion         = '1.0.0'

#region --- Logging ---
function Write-SignLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','OK')][string]$Level = 'INFO'
    )
    $ts    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = switch ($Level) {
        'OK'    { 'Green'  }
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red'    }
        default { 'Cyan'   }
    }
    Write-Host "[$ts] [$Level] $Message" -ForegroundColor $color
}
#endregion

#region --- Pre-flight ---
Write-SignLog "Invoke-BCSignPublishable.ps1 v$ScriptVersion" -Level INFO
Write-SignLog "Repo root : $RepoRoot"
Write-SignLog "Cert store: Cert:\$CertStore"
Write-SignLog "Mode      : $(if ($Audit) { 'AUDIT (no changes)' } elseif ($WhatIfPreference) { 'WHATIF' } else { 'SIGN' })"
Write-Host ""

if (-not (Test-Path $RepoRoot)) {
    Write-SignLog "Repo root not found: $RepoRoot" -Level ERROR
    exit 1
}

if (-not (Test-Path (Join-Path $RepoRoot '.git'))) {
    Write-SignLog "Not a git repository: $RepoRoot" -Level ERROR
    exit 1
}
#endregion

#region --- Certificate selection ---
function Get-BCSIgningCert {
    param([string]$Store, [string]$Thumbprint)

    $storePath = "Cert:\$Store"
    $certs     = Get-ChildItem $storePath -ErrorAction Stop |
                     Where-Object {
                         $_.EnhancedKeyUsageList.ObjectId -contains '1.3.6.1.5.5.7.3.3' -and
                         $_.NotAfter -gt (Get-Date) -and
                         $_.HasPrivateKey
                     }

    if ($Thumbprint) {
        $cert = $certs | Where-Object { $_.Thumbprint -eq $Thumbprint }
        if (-not $cert) {
            Write-SignLog "Certificate with thumbprint '$Thumbprint' not found in $storePath or not valid for code signing." -Level ERROR
            exit 1
        }
        return $cert
    }

    # Auto-select: prefer BriComp subject
    $cert = $certs | Where-Object { $_.Subject -match 'BriComp' } | Sort-Object NotAfter -Descending | Select-Object -First 1

    if (-not $cert) {
        # Fall back to any valid code signing cert
        $cert = $certs | Sort-Object NotAfter -Descending | Select-Object -First 1
    }

    if (-not $cert) {
        Write-SignLog "No valid code signing certificate found in $storePath." -Level ERROR
        Write-SignLog "Import your BriComp code signing cert from the thumb drive first." -Level ERROR
        Write-SignLog "Then re-run this script." -Level ERROR
        exit 1
    }

    return $cert
}

if (-not $Audit) {
    $signingCert = Get-BCSIgningCert -Store $CertStore -Thumbprint $CertThumbprint
    Write-SignLog "Using certificate: $($signingCert.Subject)" -Level OK
    Write-SignLog "  Thumbprint : $($signingCert.Thumbprint)"
    Write-SignLog "  Expires    : $($signingCert.NotAfter.ToString('yyyy-MM-dd'))"
    Write-SignLog "  Timestamp  : $TimestampServer"
    Write-Host ""
}
#endregion

#region --- Find publishable scripts ---
Write-SignLog "Scanning repo for '# PUBLISH: true' markers..."

$publishable = Get-ChildItem -Path $RepoRoot -Recurse -Include '*.ps1','*.psm1' |
    Where-Object { $_.FullName -notmatch '\\\.git\\' } |
    Where-Object {
        $firstLines = Get-Content $_.FullName -TotalCount 5 -ErrorAction SilentlyContinue
        $firstLines -match '# PUBLISH: true'
    }

if (-not $publishable) {
    Write-SignLog "No scripts marked '# PUBLISH: true' found." -Level WARN
    exit 0
}

Write-SignLog "Found $($publishable.Count) publishable script(s):" -Level OK
$publishable | ForEach-Object {
    Write-Host "  $($_.FullName.Replace($RepoRoot, ''))" -ForegroundColor DarkCyan
}
Write-Host ""
#endregion

#region --- Sign or audit ---
$results = @()

foreach ($file in $publishable) {
    $relativePath = $file.FullName.Replace($RepoRoot + '\', '')

    # Check current signature status
    $sig    = Get-AuthenticodeSignature -FilePath $file.FullName
    $status = $sig.Status

    $result = [PSCustomObject]@{
        File           = $relativePath
        CurrentStatus  = $status
        Action         = 'None'
        NewStatus      = $status
        Error          = ''
    }

    if ($Audit) {
        # Audit mode - just report
        $color = switch ($status) {
            'Valid'        { 'Green'  }
            'NotSigned'    { 'Yellow' }
            'HashMismatch' { 'Red'    }
            default        { 'Yellow' }
        }
        Write-Host ("  {0,-60} {1}" -f $relativePath, $status) -ForegroundColor $color
        $result.Action = 'Audit'
        $results += $result
        continue
    }

    # Skip already-valid signatures unless -Force
    if ($status -eq 'Valid' -and -not $Force) {
        Write-SignLog "SKIP (already signed): $relativePath" -Level INFO
        $result.Action = 'Skipped'
        $results += $result
        continue
    }

    # Sign
    if ($PSCmdlet.ShouldProcess($relativePath, 'Authenticode sign')) {
        try {
            $signResult = Set-AuthenticodeSignature `
                -FilePath        $file.FullName `
                -Certificate     $signingCert `
                -TimestampServer $TimestampServer `
                -HashAlgorithm   'SHA256' `
                -ErrorAction     Stop

            if ($signResult.Status -eq 'Valid') {
                Write-SignLog "SIGNED : $relativePath" -Level OK
                $result.Action    = 'Signed'
                $result.NewStatus = 'Valid'
            } else {
                Write-SignLog "FAILED : $relativePath - Status: $($signResult.Status)" -Level ERROR
                $result.Action    = 'Failed'
                $result.NewStatus = $signResult.Status
                $result.Error     = $signResult.StatusMessage
            }
        } catch {
            Write-SignLog "ERROR  : $relativePath - $($_.Exception.Message)" -Level ERROR
            $result.Action = 'Error'
            $result.Error  = $_.Exception.Message
        }
    } else {
        $result.Action = 'WhatIf'
    }

    $results += $result
}
#endregion

#region --- Summary ---
Write-Host ""
Write-SignLog "=== Summary ===" -Level INFO

$signed  = @($results | Where-Object { $_.Action -eq 'Signed'  })
$skipped = @($results | Where-Object { $_.Action -eq 'Skipped' })
$failed  = @($results | Where-Object { $_.Action -in 'Failed','Error' })

if (-not $Audit) {
    Write-SignLog "Signed  : $($signed.Count)" -Level OK
    Write-SignLog "Skipped : $($skipped.Count) (already valid)" -Level INFO
    if ($failed.Count -gt 0) {
        Write-SignLog "Failed  : $($failed.Count)" -Level ERROR
        $failed | ForEach-Object { Write-SignLog "  $($_.File): $($_.Error)" -Level ERROR }
    }

    if ($signed.Count -gt 0 -and $failed.Count -eq 0) {
        Write-Host ""
        Write-SignLog "All signatures applied successfully." -Level OK
        Write-Host ""
        Write-Host "Next steps:" -ForegroundColor Cyan
        Write-Host "  1. Review signed files: git diff --stat" -ForegroundColor Cyan
        Write-Host "  2. Commit: git add -A; git commit -m `"chore: sign publishable scripts - $(Get-Date -Format yyyy-MM-dd)`"" -ForegroundColor Cyan
        Write-Host "  3. Tag and push to trigger public sync:" -ForegroundColor Cyan
        Write-Host "       git tag publish/<path>/<script> -m `"Publish <script> v<version>`"" -ForegroundColor Cyan
        Write-Host "       git push origin --tags" -ForegroundColor Cyan
    }
} else {
    $validCount     = @($results | Where-Object { $_.CurrentStatus -eq 'Valid'     }).Count
    $unsignedCount  = @($results | Where-Object { $_.CurrentStatus -eq 'NotSigned' }).Count
    $mismatchCount  = @($results | Where-Object { $_.CurrentStatus -eq 'HashMismatch' }).Count
    Write-SignLog "Valid      : $validCount" -Level OK
    Write-SignLog "Not signed : $unsignedCount" -Level WARN
    Write-SignLog "Mismatch   : $mismatchCount" -Level $(if ($mismatchCount -gt 0) { 'ERROR' } else { 'INFO' })
    Write-Host ""
    if ($unsignedCount -gt 0 -or $mismatchCount -gt 0) {
        Write-SignLog "Run without -Audit to sign." -Level INFO
    }
}
#endregion
