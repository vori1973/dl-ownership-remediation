#Requires -Version 7.2
<#
.SYNOPSIS
    Applies ownership assignments from a modified DL member report CSV.

.DESCRIPTION
    Step 2 of the DL ownership remediation workflow.

    Reads the CSV produced by Export-DLMemberReport.ps1 after the customer has
    set IsOwner=1 for the desired owners. For each group, sets ManagedBy to the
    specified owners.

    Validations performed before any change:
    - Only rows where IsOwner=1 AND MemberIsEligible=TRUE are applied.
      Rows with MemberIsEligible=FALSE and IsOwner=1 are flagged as warnings.
    - Groups with no eligible owners in the CSV are skipped.

    Mode:
    - Append (default): The owners in the CSV are added to any existing ManagedBy value.
      Use this when groups already have owners you want to keep.
    - Replace: ManagedBy is set to exactly the owners in the CSV.
      Use this to overwrite all existing owners.

    Use -WhatIf to preview all changes without applying them.
    Results are written to a timestamped audit log CSV.

.PARAMETER InputCsvPath
    Path to the modified member report CSV (IsOwner column updated by customer).
    Must contain columns: GroupEmail, MemberEmail, MemberIsEligible, IsOwner.

.PARAMETER Mode
    Append (default) — adds CSV-designated owners to the existing ManagedBy list.
    Replace          — replaces ManagedBy entirely with the CSV-designated owners.

.PARAMETER Notify
    If specified, sends an email to each newly assigned owner informing them of their new role.
    Requires app-based auth with -ClientSecret and the app to have the Mail.Send Graph permission.
    Pair with -NotificationFrom to set the sender address.

.PARAMETER NotificationFrom
    Sender address for owner notification emails (requires a licensed mailbox).
    Defaults to -AdminUPN when not specified. Only used when -Notify is set.

.PARAMETER NotificationSubject
    Subject line for notification emails. Supports the same placeholders as the body template:
    {{GroupName}}, {{GroupEmail}}, {{OwnerEmail}}.
    Defaults to: "You have been added as owner of distribution list: {{GroupName}}"

.PARAMETER NotificationTemplatePath
    Path to an external HTML file used as the email body.
    Supports the following placeholders, replaced per email:
        {{GroupName}}  — display name of the distribution list
        {{GroupEmail}} — primary SMTP address of the distribution list
        {{OwnerEmail}} — email address of the recipient (the newly assigned owner)
    If omitted, the built-in template is used.
    Copy notification-template.html from the script directory as a starting point.

.PARAMETER AuditLogPath
    Path for the audit log CSV. Defaults to .\DLOwnerAssignment_<timestamp>.csv.

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
    Thumbprint of a certificate installed in the local machine certificate store.
    Optional alternative to -ClientSecret; preferred from a security standpoint.

.PARAMETER CertificatePath
    Path to a .pfx certificate file.
    Optional alternative when the certificate is not installed in the local store.
    Pair with -CertificatePassword if the .pfx is password-protected.

.PARAMETER CertificatePassword
    Plain-text password for the .pfx file specified in -CertificatePath.

.PARAMETER AdminUPN
    UPN of the admin account for interactive (delegated) authentication.
    Fallback when no app-based auth params are supplied.

.EXAMPLE
    # Preview with app-based auth (always run -WhatIf first)
    .\Set-DLOwners.ps1 -InputCsvPath .\DLMemberReport_modified.csv `
        -AppId 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' -TenantId 'contoso.onmicrosoft.com' `
        -ClientSecret 'your-secret-value' -WhatIf

.EXAMPLE
    # Apply changes — client secret (primary)
    .\Set-DLOwners.ps1 -InputCsvPath .\DLMemberReport_modified.csv `
        -AppId 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' -TenantId 'contoso.onmicrosoft.com' `
        -ClientSecret 'your-secret-value'

.EXAMPLE
    # Apply changes — certificate thumbprint
    .\Set-DLOwners.ps1 -InputCsvPath .\DLMemberReport_modified.csv `
        -AppId 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' -TenantId 'contoso.onmicrosoft.com' `
        -CertificateThumbprint 'AABBCCDDEEFF...'

.EXAMPLE
    # Append to existing owners
    .\Set-DLOwners.ps1 -InputCsvPath .\DLMemberReport_modified.csv -Mode Append `
        -AppId 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' -TenantId 'contoso.onmicrosoft.com' `
        -ClientSecret 'your-secret-value'
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [ValidateSet('Replace', 'Append')]
    [string]$Mode = 'Append',

    [switch]$Notify,
    [string]$NotificationFrom,
    [string]$NotificationSubject = 'You have been added as owner of distribution list: {{GroupName}}',
    [string]$NotificationTemplatePath,

    [string]$AuditLogPath = ".\DLOwnerAssignment_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

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

function Get-GraphToken {
    # Returns an access token via client-credentials flow; $null when credentials are absent
    # (caller falls back to Send-MgUserMail SDK path).
    if (-not ($AppId -and $TenantId -and $ClientSecret)) { return $null }
    try {
        $body = @{
            grant_type    = 'client_credentials'
            client_id     = $AppId
            client_secret = $ClientSecret
            scope         = 'https://graph.microsoft.com/.default'
        }
        $resp = Invoke-RestMethod `
            -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
            -Method POST -Body $body -ContentType 'application/x-www-form-urlencoded' `
            -ErrorAction Stop
        Write-Log 'Graph token acquired for notification emails.' 'SUCCESS'
        return $resp.access_token
    } catch {
        Write-Log "Could not acquire Graph token — notifications disabled. $($_.Exception.Message)" 'WARN'
        return $null
    }
}

function Send-OwnerNotification {
    param(
        [string]$OwnerEmail,
        [string]$GroupName,
        [string]$GroupEmail,
        [string]$FromEmail,
        [string]$GraphToken,
        [string]$Template,
        [string]$Subject
    )

    $expandedSubject = $Subject `
        -replace '\{\{GroupName\}\}',  $GroupName `
        -replace '\{\{GroupEmail\}\}', $GroupEmail `
        -replace '\{\{OwnerEmail\}\}', $OwnerEmail

    $htmlBody = $Template `
        -replace '\{\{GroupName\}\}',  $GroupName `
        -replace '\{\{GroupEmail\}\}', $GroupEmail `
        -replace '\{\{OwnerEmail\}\}', $OwnerEmail

    if ($GraphToken) {
        # App-based path: Graph REST API with client-credentials token
        $payload = @{
            message        = @{
                subject      = $expandedSubject
                body         = @{ contentType = 'HTML'; content = $htmlBody }
                toRecipients = @(@{ emailAddress = @{ address = $OwnerEmail } })
            }
            saveToSentItems = $false
        } | ConvertTo-Json -Depth 10

        try {
            Invoke-RestMethod `
                -Uri "https://graph.microsoft.com/v1.0/users/$FromEmail/sendMail" `
                -Method POST -Body $payload -ContentType 'application/json' `
                -Headers @{ Authorization = "Bearer $GraphToken" } `
                -ErrorAction Stop
            Write-Log "    Notification sent → $OwnerEmail" 'INFO'
        } catch {
            Write-Log "    Notification FAILED → $OwnerEmail — $($_.Exception.Message)" 'WARN'
        }
    } elseif (Get-Command Send-MgUserMail -ErrorAction SilentlyContinue) {
        # Interactive/delegated path: Microsoft.Graph SDK
        try {
            $msg = @{
                Subject      = $expandedSubject
                Body         = @{ ContentType = 'HTML'; Content = $htmlBody }
                ToRecipients = @(@{ EmailAddress = @{ Address = $OwnerEmail } })
            }
            Send-MgUserMail -UserId $FromEmail -Message $msg -SaveToSentItems:$false -ErrorAction Stop
            Write-Log "    Notification sent → $OwnerEmail" 'INFO'
        } catch {
            Write-Log "    Notification FAILED → $OwnerEmail — $($_.Exception.Message)" 'WARN'
        }
    } else {
        Write-Log "    Notification skipped → $OwnerEmail (no Graph token and Microsoft.Graph module unavailable)" 'WARN'
    }
}

# ─── Entry ───────────────────────────────────────────────────────────────────

Connect-ToExchangeOnline -UPN $AdminUPN

# ─── Notification setup ───────────────────────────────────────────────────────
$graphToken           = $null
$notificationTemplate = $null
if ($Notify) {
    if (-not $NotificationFrom -and $AdminUPN) { $NotificationFrom = $AdminUPN }
    if (-not $NotificationFrom) {
        Write-Error '-NotificationFrom (or -AdminUPN) is required when using -Notify.'
        exit 1
    }
    if ($WhatIfPreference) {
        Write-Log '-Notify: WhatIf mode — notification emails will not actually be sent.' 'WARN'
    } elseif ($AppId -and $TenantId -and $ClientSecret) {
        $graphToken = Get-GraphToken
    } elseif (Get-Command Send-MgUserMail -ErrorAction SilentlyContinue) {
        Write-Log '-Notify: no -ClientSecret — using Microsoft.Graph module (Send-MgUserMail).' 'INFO'
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
        if (-not $ctx) {
            Write-Log 'Connecting to Microsoft Graph (Mail.Send scope)...' 'INFO'
            Connect-MgGraph -Scopes 'Mail.Send' -NoWelcome -ErrorAction Stop
        }
    } else {
        Write-Log '-Notify requires either -ClientSecret (Graph REST) or the Microsoft.Graph module.' 'WARN'
        Write-Log '  Install with: Install-Module Microsoft.Graph.Users.Actions -Scope CurrentUser' 'WARN'
        $Notify = $false
    }

    # Validate the sender mailbox exists — Graph sendMail returns 404 if the sender
    # is not a real, licensed user mailbox (e.g. a DL, contact, or non-existent address).
    if ($Notify -and -not $WhatIfPreference) {
        try {
            $null = Get-Mailbox -Identity $NotificationFrom -ErrorAction Stop
            Write-Log "Notification sender mailbox verified: $NotificationFrom" 'INFO'
        } catch {
            Write-Error "-NotificationFrom '$NotificationFrom' was not found or is not a user mailbox. Provide a licensed mailbox that exists in Exchange Online (e.g. your admin UPN). $($_.Exception.Message)"
            exit 1
        }
    }

    # Load email template
    if ($NotificationTemplatePath) {
        if (-not (Test-Path $NotificationTemplatePath)) {
            Write-Error "Notification template not found: $NotificationTemplatePath"
            exit 1
        }
        $notificationTemplate = Get-Content $NotificationTemplatePath -Raw -Encoding UTF8
        Write-Log "Notification template loaded: $NotificationTemplatePath" 'INFO'
    } else {
        $notificationTemplate = @'
<html><body>
<p>Hi,</p>
<p>You have been assigned as an <strong>owner</strong> of the following distribution list:</p>
<table cellpadding="4" style="border-collapse:collapse">
  <tr><td><strong>Display name</strong></td><td>{{GroupName}}</td></tr>
  <tr><td><strong>Email address</strong></td><td>{{GroupEmail}}</td></tr>
</table>
<p>As an owner you can manage this list&rsquo;s membership and settings in Outlook or the Exchange admin center.</p>
<p><em>This message was sent automatically by the DL Ownership Remediation tool.</em></p>
</body></html>
'@
    }
}

#region --- Load and validate input CSV ---

if (-not (Test-Path $InputCsvPath)) {
    Write-Error "Input CSV not found: $InputCsvPath"
    exit 1
}

$allRows = Import-Csv -Path $InputCsvPath

$requiredColumns = @('GroupEmail', 'MemberEmail', 'MemberIsEligible', 'IsOwner')
$missingColumns  = @($requiredColumns | Where-Object { $_ -notin $allRows[0].PSObject.Properties.Name })
if ($missingColumns.Count -gt 0) {
    Write-Error "Input CSV is missing required columns: $($missingColumns -join ', ')"
    exit 1
}

# Separate valid owner rows from problem rows
$ownerRows  = @($allRows | Where-Object { $_.IsOwner -eq '1' -and $_.MemberIsEligible -eq 'TRUE' -and $_.MemberEmail -ne '' })
$ineligible = @($allRows | Where-Object { $_.IsOwner -eq '1' -and $_.MemberIsEligible -ne 'TRUE' })

if ($ineligible.Count -gt 0) {
    Write-Log "$($ineligible.Count) row(s) have IsOwner=1 but MemberIsEligible is not TRUE — these will be skipped:" 'WARN'
    $ineligible | ForEach-Object {
        Write-Log "  SKIPPED  $($_.GroupEmail)  ←  $($_.MemberEmail)  ($($_.MemberRecipientType))" 'WARN'
    }
}

if ($ownerRows.Count -eq 0) {
    Write-Log 'No eligible owner rows found (IsOwner=1 AND MemberIsEligible=TRUE). Nothing to do.' 'WARN'
    exit 0
}

# Group by GroupEmail
$groupedByDL = $ownerRows | Group-Object -Property GroupEmail
Write-Log "Input: $($ownerRows.Count) eligible owner assignment(s) across $($groupedByDL.Count) group(s)"

#endregion

#region --- Apply ownership changes ---

$auditLog = [System.Collections.Generic.List[pscustomobject]]::new()
$i        = 0
$total    = $groupedByDL.Count

foreach ($group in $groupedByDL) {
    $i++
    $dlEmail     = $group.Name
    $newOwners   = $group.Group | Select-Object -ExpandProperty MemberEmail
    $groupName   = $group.Group[0].GroupName

    Write-Progress -Activity 'Setting DL Owners' `
        -Status "[$i/$total] $groupName" `
        -PercentComplete ([int](($i / $total) * 100))

    try {
        $dl = Get-DistributionGroup -Identity $dlEmail -ErrorAction Stop

        # Append mode: resolve existing ManagedBy GUIDs to SMTP first, then merge.
        # EXO v3 stores ManagedBy as object GUIDs — passing raw GUIDs mixed with
        # SMTP addresses to Set-DistributionGroup resolves inconsistently.
        $finalOwners = if ($Mode -eq 'Append' -and $dl.ManagedBy -and $dl.ManagedBy.Count -gt 0) {
            $resolvedExisting = foreach ($ownerIdentity in $dl.ManagedBy) {
                try {
                    $r = Get-Recipient -Identity $ownerIdentity.ToString() -ErrorAction SilentlyContinue
                    if ($r -and $r.PrimarySmtpAddress) { $r.PrimarySmtpAddress }
                    else { $ownerIdentity.ToString() }
                } catch { $ownerIdentity.ToString() }
            }
            @($resolvedExisting) + @($newOwners) |
                Where-Object { $_ } | Sort-Object -Unique
        } else {
            $newOwners
        }

        $ownerList = $finalOwners -join '; '
        $action    = if ($WhatIfPreference) { 'WhatIf' } else { 'Applied' }

        if ($PSCmdlet.ShouldProcess(
            "$groupName ($dlEmail)",
            "Set ManagedBy = $ownerList"
        )) {
            Set-DistributionGroup -Identity $dlEmail -ManagedBy $finalOwners -ErrorAction Stop
            $action = 'Applied'
        }

        $auditLog.Add([pscustomobject]@{
            Timestamp   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            GroupName   = $groupName
            GroupEmail  = $dlEmail
            Mode        = $Mode
            NewOwners   = $ownerList
            Status      = $action
            Error       = ''
        })

        Write-Log "  [$i/$total] $action  $groupName  →  $ownerList" $(if ($action -eq 'Applied') { 'SUCCESS' } else { 'INFO' })

        if ($action -eq 'Applied' -and $Notify) {
            foreach ($owner in $newOwners) {
                Send-OwnerNotification -OwnerEmail $owner -GroupName $groupName `
                    -GroupEmail $dlEmail -FromEmail $NotificationFrom `
                    -GraphToken $graphToken -Template $notificationTemplate `
                    -Subject $NotificationSubject
            }
        }

    } catch {
        $errMsg = $_.Exception.Message
        Write-Log "  [$i/$total] FAILED  $groupName — $errMsg" 'ERROR'
        if ($errMsg -match 'write scope|write scopes') {
            Write-Log '  HINT: service principal lacks Exchange write rights. Run once as Exchange admin:' 'WARN'
            Write-Log '    $sp = Get-ServicePrincipal -Identity ''<AppDisplayName>''' 'WARN'
            Write-Log "    New-ManagementRoleAssignment -App `$sp.ObjectId -Role 'Mail Recipients'" 'WARN'
        }

        $auditLog.Add([pscustomobject]@{
            Timestamp   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            GroupName   = $groupName
            GroupEmail  = $dlEmail
            Mode        = $Mode
            NewOwners   = ($newOwners -join '; ')
            Status      = 'Failed'
            Error       = $errMsg
        })
    }
}

Write-Progress -Activity 'Setting DL Owners' -Completed

#endregion

#region --- Audit log and summary ---

$auditLog | Export-Csv -Path $AuditLogPath -NoTypeInformation -Encoding UTF8 -WhatIf:$false

$applied = @($auditLog | Where-Object { $_.Status -eq 'Applied' }).Count
$failed  = @($auditLog | Where-Object { $_.Status -eq 'Failed'  }).Count
$whatif  = @($auditLog | Where-Object { $_.Status -eq 'WhatIf'  }).Count

Write-Log "=== Summary ==="                                                   'SUCCESS'
Write-Log "  Groups processed : $total"
Write-Log "  Applied          : $applied"                                    $(if ($applied -gt 0) { 'SUCCESS' } else { 'INFO' })
if ($failed -gt 0)  { Write-Log "  Failed           : $failed"              'ERROR'   }
if ($whatif -gt 0)  { Write-Log "  WhatIf (preview) : $whatif"             'WARN'    }
if ($ineligible.Count -gt 0) { Write-Log "  Ineligible skipped: $($ineligible.Count)" 'WARN' }
Write-Log "  Audit log        : $AuditLogPath"

#endregion
