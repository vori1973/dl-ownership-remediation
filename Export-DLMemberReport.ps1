#Requires -Version 7.2
<#
.SYNOPSIS
    Exports distribution group members to a CSV for ownership review.

.DESCRIPTION
    Step 1 of the DL ownership remediation workflow.

    Fetches all distribution groups (or a filtered subset) from Exchange Online,
    enumerates their members, and writes a flat CSV with one row per member.
    Each row has IsOwner=0 by default.

    The customer edits the CSV, sets IsOwner=1 for the users they want as owners,
    then runs Set-DLOwners.ps1 to apply the changes.

    Efficiency notes:
    - All groups are fetched in a single bulk call; individual lookups are avoided.
    - Members are fetched with -ResultSize Unlimited per group.
    - Groups with no members still produce one placeholder row (no member data).
    - Groups that fail enumeration are logged separately.

    Typical workflow:
      1. Export ownerless groups from Exchange Online to a CSV (Name, manageby, RecipientTypeDetails).
      2. Run this script with -InputCsvPath pointing to that CSV.
      3. Open the output CSV, filter MemberIsEligible=TRUE, set IsOwner=1 for desired owners.
      4. Save the modified CSV and run Set-DLOwners.ps1.

.PARAMETER InputCsvPath
    Optional. Path to a CSV with a Name column listing distribution group names to process.
    Typical input: the ownerless-groups export (e.g. NoOwners2.csv).
    If omitted, all distribution groups in the tenant are processed.

.PARAMETER OutputCsvPath
    Path for the output member report CSV.
    Defaults to .\DLMemberReport_<timestamp>.csv in the current directory.

.PARAMETER AppId
    Entra app registration client ID (application / client ID GUID).
    Required for app-based authentication. Pair with -TenantId and one of:
    -ClientSecret (primary), -CertificateThumbprint, or -CertificatePath.

.PARAMETER TenantId
    Tenant ID (GUID) or primary domain (e.g. contoso.onmicrosoft.com).
    Required with -AppId.

.PARAMETER ClientSecret
    Client secret for the app registration.
    Primary credential method — use when no certificate infrastructure is available.

.PARAMETER CertificateThumbprint
    Thumbprint of a certificate installed in the local machine certificate store
    (Cert:\LocalMachine\My or Cert:\CurrentUser\My).
    Optional alternative to -ClientSecret; preferred from a security standpoint.

.PARAMETER CertificatePath
    Path to a .pfx certificate file.
    Optional alternative when the certificate is not installed in the local store.
    Pair with -CertificatePassword if the .pfx is password-protected.

.PARAMETER CertificatePassword
    Plain-text password for the .pfx file specified in -CertificatePath.

.PARAMETER OwnerlessOnly
    When specified, skips any group that already has at least one owner in ManagedBy.
    Works with both -InputCsvPath (filter the provided list) and without it (scan the
    entire tenant and automatically narrow to ownerless groups only).

    This eliminates the need to pre-filter a CSV manually:
      .\Export-DLMemberReport.ps1 -OwnerlessOnly ...
    is equivalent to: export all → filter ManagedBy empty → re-run.

.PARAMETER AdminUPN
    UPN of the admin account for interactive (delegated) authentication.
    Fallback when no app-based auth params are supplied.
    If omitted and no app-based params are provided, the script will prompt.

.EXAMPLE
    # Tenant-wide ownerless scan — no input CSV needed
    . .\config.ps1
    .\Export-DLMemberReport.ps1 -OwnerlessOnly `
        -AppId $cfg.AppId -TenantId $cfg.TenantId -ClientSecret $cfg.ClientSecret

.EXAMPLE
    # Filter a provided CSV to ownerless groups only
    . .\config.ps1
    .\Export-DLMemberReport.ps1 -InputCsvPath .\AllGroups.csv -OwnerlessOnly `
        -AppId $cfg.AppId -TenantId $cfg.TenantId -ClientSecret $cfg.ClientSecret

.EXAMPLE
    # App-based auth with certificate thumbprint
    . .\config.ps1
    .\Export-DLMemberReport.ps1 -InputCsvPath .\test\NoOwners2.csv `
        -AppId $cfg.AppId -TenantId $cfg.TenantId -CertificateThumbprint $cfg.CertificateThumbprint

.EXAMPLE
    # Interactive fallback (delegated auth)
    .\Export-DLMemberReport.ps1 -InputCsvPath .\test\NoOwners2.csv -AdminUPN admin@contoso.com
#>

[CmdletBinding()]
param(
    [string]$InputCsvPath,

    [string]$OutputCsvPath = ".\DLMemberReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    # Skip groups that already have at least one owner
    [switch]$OwnerlessOnly,

    # App-based auth (primary)
    [string]$AppId,
    [string]$TenantId,
    [string]$ClientSecret,

    # Certificate alternatives (optional)
    [string]$CertificateThumbprint,
    [string]$CertificatePath,
    [string]$CertificatePassword,

    # Interactive fallback
    [string]$AdminUPN
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─── Helpers ─────────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "[$ts] [$Level] $Message" -ForegroundColor $(switch ($Level) {
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red'    }
        'SUCCESS' { 'Green'  }
        default   { 'Cyan'   }
    })
}

function Connect-ToExchangeOnline {
    param([string]$UPN)

    # Skip if a session is already open
    try { $null = Get-OrganizationConfig -ErrorAction Stop; Write-Log 'Already connected to Exchange Online.' 'SUCCESS'; return } catch {}

    # Module check — install if missing, warn if below v3
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Log 'ExchangeOnlineManagement module not found. Installing...' 'WARN'
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module ExchangeOnlineManagement -Force

    $modVersion = (Get-Module ExchangeOnlineManagement).Version
    if ($modVersion -lt [version]'3.0.0') {
        Write-Log "ExchangeOnlineManagement v$modVersion detected — v3.0+ required for REST cmdlets." 'WARN'
        Write-Log 'Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force' 'WARN'
    }

    if ($AppId -and $TenantId) {
        if ($ClientSecret) {
            Write-Log "Connecting to Exchange Online (app — client secret | AppId: $AppId | Org: $TenantId)..."
            Connect-ExchangeOnline -AppId $AppId `
                -ClientSecret (ConvertTo-SecureString $ClientSecret -AsPlainText -Force) `
                -Organization $TenantId -ShowBanner:$false
        } elseif ($CertificateThumbprint) {
            Write-Log "Connecting to Exchange Online (app — cert thumbprint | AppId: $AppId | Org: $TenantId | Cert: $CertificateThumbprint)..."
            Connect-ExchangeOnline -AppId $AppId `
                -CertificateThumbprint $CertificateThumbprint `
                -Organization $TenantId -ShowBanner:$false
        } elseif ($CertificatePath) {
            Write-Log "Connecting to Exchange Online (app — cert file | AppId: $AppId | Org: $TenantId | Path: $CertificatePath)..."
            $certPass = if ($CertificatePassword) { ConvertTo-SecureString $CertificatePassword -AsPlainText -Force } else { $null }
            Connect-ExchangeOnline -AppId $AppId `
                -CertificatePath $CertificatePath -CertificatePassword $certPass `
                -Organization $TenantId -ShowBanner:$false
        } else {
            throw 'AppId and TenantId provided but no credential specified. Add -ClientSecret, -CertificateThumbprint, or -CertificatePath.'
        }
    } else {
        if (-not $UPN) { $UPN = Read-Host 'Enter admin UPN for Exchange Online (interactive)' }
        Write-Log "Connecting to Exchange Online as $UPN (interactive)..."
        Connect-ExchangeOnline -UserPrincipalName $UPN -ShowProgress $false -ShowBanner:$false
    }

    Write-Log "Connected (module v$modVersion)." 'SUCCESS'
}

# ─── Entry ───────────────────────────────────────────────────────────────────

Connect-ToExchangeOnline -UPN $AdminUPN

#region --- Load all distribution groups in one bulk call ---

Write-Log 'Fetching all distribution groups...'
$allDLs = Get-DistributionGroup -ResultSize Unlimited
Write-Log "  Found $($allDLs.Count) distribution groups in tenant"

# Build a case-insensitive lookup by both Name and DisplayName
$dlLookup = [System.Collections.Generic.Dictionary[string, object]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
foreach ($dl in $allDLs) {
    if (-not $dlLookup.ContainsKey($dl.Name))        { $dlLookup[$dl.Name]        = $dl }
    if (-not $dlLookup.ContainsKey($dl.DisplayName)) { $dlLookup[$dl.DisplayName] = $dl }
}

#endregion

#region --- Determine target groups ---

if ($InputCsvPath) {
    if (-not (Test-Path $InputCsvPath)) {
        Write-Error "Input CSV not found: $InputCsvPath"
        exit 1
    }
    $inputRows   = Import-Csv -Path $InputCsvPath
    $targetDLs   = [System.Collections.Generic.List[object]]::new()
    $notFound    = [System.Collections.Generic.List[string]]::new()

    foreach ($row in $inputRows) {
        if ($dlLookup.ContainsKey($row.Name)) {
            $targetDLs.Add($dlLookup[$row.Name])
        } else {
            $notFound.Add($row.Name)
        }
    }

    Write-Log "  Matched $($targetDLs.Count) / $($inputRows.Count) groups from $InputCsvPath"
    if ($notFound.Count -gt 0) {
        Write-Log "  $($notFound.Count) group(s) from input CSV not found in Exchange Online" 'WARN'
        $notFound | ForEach-Object { Write-Log "    NOT FOUND: $_" 'WARN' }
    }
} else {
    $targetDLs = $allDLs
    Write-Log "  No input CSV — processing all $($targetDLs.Count) distribution groups"
}

# -OwnerlessOnly: drop any group that already has at least one ManagedBy entry.
# ManagedBy is already present on every object from the bulk Get-DistributionGroup call —
# no extra API calls needed.
if ($OwnerlessOnly) {
    $beforeFilter = $targetDLs.Count
    $targetDLs    = @($targetDLs | Where-Object { -not $_.ManagedBy -or $_.ManagedBy.Count -eq 0 })
    $skipped      = $beforeFilter - $targetDLs.Count
    Write-Log "  -OwnerlessOnly: $($targetDLs.Count) ownerless groups to process ($skipped skipped — already have owners)"

    if ($targetDLs.Count -eq 0) {
        Write-Log 'No ownerless groups found — nothing to export.' 'WARN'
        exit 0
    }
}

#endregion

#region --- Enumerate members ---

$results  = [System.Collections.Generic.List[pscustomobject]]::new()
$errList  = [System.Collections.Generic.List[pscustomobject]]::new()
$total    = $targetDLs.Count
$i        = 0
$sw       = [System.Diagnostics.Stopwatch]::StartNew()

Write-Log "Enumerating members for $total groups..."

foreach ($dl in $targetDLs) {
    $i++
    $elapsed  = $sw.Elapsed.TotalSeconds
    $rate     = if ($elapsed -gt 0 -and $i -gt 1) { ($i - 1) / $elapsed } else { 0.5 }
    $secLeft  = [int](($total - $i) / $rate)
    $eta      = [timespan]::FromSeconds($secLeft).ToString('mm\:ss')

    Write-Progress -Activity 'Exporting DL Members' `
        -Status "[$i/$total] $($dl.DisplayName)  |  ETA $eta" `
        -PercentComplete ([int](($i / $total) * 100))

    # EXO v3 REST returns ManagedBy as Azure AD object GUIDs, not display names or SMTP.
    # Resolve each identity to a PrimarySmtpAddress via Get-Recipient so we can:
    #   1. Match reliably against member.PrimarySmtpAddress (GUID match never works)
    #   2. Store a human-readable value in the CurrentManagedBy column instead of raw GUIDs
    $managedByRaw = if ($dl.ManagedBy -and $dl.ManagedBy.Count -gt 0) {
        @($dl.ManagedBy | ForEach-Object { $_.ToString() })
    } else { @() }

    $ownerSmtpList = [System.Collections.Generic.List[string]]::new()
    $ownerSmtpSet  = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($ownerIdentity in $managedByRaw) {
        try {
            $r = Get-Recipient -Identity $ownerIdentity -ErrorAction SilentlyContinue
            if ($r -and $r.PrimarySmtpAddress) {
                $ownerSmtpList.Add($r.PrimarySmtpAddress)
                [void]$ownerSmtpSet.Add($r.PrimarySmtpAddress)
            } else {
                $ownerSmtpList.Add($ownerIdentity)   # keep raw if resolution fails
            }
        } catch {
            $ownerSmtpList.Add($ownerIdentity)
        }
    }

    $currentManagedBy = $ownerSmtpList -join '; '

    try {
        $members = Get-DistributionGroupMember -Identity $dl.PrimarySmtpAddress `
                       -ResultSize Unlimited -ErrorAction Stop

        if ($members.Count -eq 0) {
            # Emit a placeholder row so empty groups appear in the report
            $results.Add([pscustomobject]@{
                GroupName           = $dl.DisplayName
                GroupEmail          = $dl.PrimarySmtpAddress
                GroupType           = $dl.RecipientTypeDetails
                CurrentManagedBy    = $currentManagedBy
                MemberDisplayName   = ''
                MemberEmail         = ''
                MemberRecipientType = ''
                MemberIsEligible    = ''
                IsOwner             = 0
            })
        } else {
            foreach ($member in $members) {
                # Eligible = can be set as ManagedBy (requires a mailbox, not a contact or nested DL)
                $eligible = $member.RecipientTypeDetails -in @(
                    'UserMailbox', 'SharedMailbox', 'LinkedMailbox', 'TeamMailbox'
                )

                # Pre-populate IsOwner=1 if this member is already a ManagedBy owner.
                # Resolved set contains SMTP addresses — match is unambiguous.
                $isCurrentOwner = $ownerSmtpSet.Count -gt 0 -and
                                  $member.PrimarySmtpAddress -and
                                  $ownerSmtpSet.Contains($member.PrimarySmtpAddress)

                $results.Add([pscustomobject]@{
                    GroupName           = $dl.DisplayName
                    GroupEmail          = $dl.PrimarySmtpAddress
                    GroupType           = $dl.RecipientTypeDetails
                    CurrentManagedBy    = $currentManagedBy
                    MemberDisplayName   = $member.DisplayName
                    MemberEmail         = $member.PrimarySmtpAddress
                    MemberRecipientType = $member.RecipientTypeDetails
                    MemberIsEligible    = if ($eligible) { 'TRUE' } else { 'FALSE' }
                    IsOwner             = if ($isCurrentOwner) { 1 } else { 0 }
                })
            }
        }
    } catch {
        $errList.Add([pscustomobject]@{
            GroupName = $dl.DisplayName
            GroupEmail = $dl.PrimarySmtpAddress
            Error     = $_.Exception.Message
        })
        Write-Log "  [$i/$total] FAILED: $($dl.DisplayName) — $($_.Exception.Message)" 'WARN'
    }
}

Write-Progress -Activity 'Exporting DL Members' -Completed
$sw.Stop()

#endregion

#region --- Export ---

$results | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8

Write-Log "Completed in $($sw.Elapsed.ToString('mm\:ss'))" 'SUCCESS'
Write-Log "  Groups processed      : $total"
Write-Log "  Total rows            : $($results.Count)"
Write-Log "  Eligible members      : $(@($results | Where-Object { $_.MemberIsEligible -eq 'TRUE' }).Count)"
Write-Log "  Pre-populated IsOwner : $(@($results | Where-Object { $_.IsOwner -eq 1 }).Count) (already owners — preserved)"
if ($errList.Count -gt 0) { Write-Log "  Errors                : $($errList.Count)" 'WARN' }
Write-Log "  Output CSV       : $OutputCsvPath"

if ($errList.Count -gt 0) {
    $errPath = $OutputCsvPath -replace '\.csv$', '_errors.csv'
    $errList | Export-Csv -Path $errPath -NoTypeInformation -Encoding UTF8
    Write-Log "  Error log        : $errPath" 'WARN'
}

Write-Log 'Next steps: open the CSV, filter MemberIsEligible=TRUE, set IsOwner=1 for desired owners, then run Set-DLOwners.ps1'

#endregion
