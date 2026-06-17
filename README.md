# Distribution Group Ownership Remediation

Two-script workflow for identifying and assigning owners to ownerless Exchange Online
distribution groups at scale.

---

## Overview

| Step | Script | What it does |
|------|--------|-------------|
| 1 | `Export-DLMemberReport.ps1` | Exports all members of target groups to a CSV with `IsOwner=0` |
| — | *(customer edits CSV)* | Sets `IsOwner=1` for each user they want as owner |
| 2 | `Set-DLOwners.ps1` | Reads the edited CSV and applies `ManagedBy` in Exchange Online |

---

## Step 1 — Configure Credentials

Copy the sample config file and fill in your values:

```powershell
Copy-Item .\config.sample.ps1 .\config.ps1
# Open config.ps1 in your editor and fill in AppId, TenantId, ClientSecret
```

`config.ps1` is gitignored — it will never be committed. The sample file
(`config.sample.ps1`) is committed as a template with placeholder values.

**Dot-source before running either script:**

```powershell
. .\config.ps1          # loads $cfg into the session
```

All examples in this document assume `$cfg` is loaded.

---

## Prerequisites

### PowerShell

- **Version**: PowerShell **7.2 or later** required (`#Requires -Version 7.2`)
- Verify: `$PSVersionTable.PSVersion`

### Exchange Online Management Module

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
# or update if already installed
Update-Module ExchangeOnlineManagement
```

Minimum supported version: **3.0** (EXO v3, REST-based cmdlets).  
Verify: `Get-Module ExchangeOnlineManagement -ListAvailable`

---

## Authentication

Both scripts support three authentication methods, checked in this priority order:

| Priority | Method | Parameters | When to use |
|----------|--------|-----------|-------------|
| 1 | App — client secret | `-AppId` `-TenantId` `-ClientSecret` | Primary. No cert infrastructure needed. |
| 2 | App — cert thumbprint | `-AppId` `-TenantId` `-CertificateThumbprint` | Preferred security-wise; cert already in local store. |
| 3 | App — cert file | `-AppId` `-TenantId` `-CertificatePath` (`-CertificatePassword`) | Cert as a .pfx file on disk. |
| 4 | Interactive (delegated) | `-AdminUPN` | Fallback; prompts for MFA in a browser window. |

---

## Required App Registration (Methods 1–3)

### Create the app

1. **Entra ID → App registrations → New registration**
   - Name: e.g. `DL-Governance-Automation`
   - Supported account types: **Single tenant**
2. Copy the **Application (client) ID** → this is your `-AppId`
3. Copy the **Directory (tenant) ID** → this is your `-TenantId`

### Add API permissions

| API | Permission | Type | When required |
|-----|-----------|------|--------------|
| Office 365 Exchange Online | `Exchange.ManageAsApp` | Application | Always (core functionality) |
| Microsoft Graph | `Mail.Send` | Application | Only when using `-Notify` to email assigned owners |

After adding: click **Grant admin consent for \<tenant\>**.

### Assign Exchange directory role (Entra)

The API permission above is not enough on its own. The service principal also needs
an Exchange role assigned in Entra ID:

1. **Entra ID → Roles and administrators → Exchange Recipient Administrator**
2. **Add assignment** → search for the app name → select the **service principal** entry

> Do not assign to the app registration — assign to the **enterprise application**
> (service principal) that Entra creates automatically.

### Assign Exchange Online RBAC role (required for write operations)

The Entra directory role alone covers read operations and admin-center UI access.
For app-only `Set-DistributionGroup` (and other write cmdlets), the service principal
also needs an Exchange Online management **role** (not a role group) assigned directly.

> **Common mistake:** `Recipient Management` is a *role group*, not a role.  
> `New-ManagementRoleAssignment -App` requires a management role name.  
> The correct role for `Set-DistributionGroup` is **`Mail Recipients`**.

Run this **once** as an Exchange admin.

> **Important:** `-App` requires the service principal **Object ID**, not the Application (client) ID.  
> Find it: Entra ID → Enterprise applications → your app → Overview → Object ID  
> Or via PowerShell:  
> `Get-MgServicePrincipal -Filter "AppId eq '<AppId>'" | Select-Object Id, DisplayName`

```powershell
Connect-ExchangeOnline -UserPrincipalName admin@contoso.com

# 1. Find (or register) the Exchange service principal
$sp = Get-ServicePrincipal -Identity 'YourAppDisplayName' -ErrorAction SilentlyContinue
if (-not $sp) {
    # First time — register the app in Exchange Online
    $sp = New-ServicePrincipal -AppId '<AppId>' `
                               -ObjectId '<EntraServicePrincipalObjectId>' `
                               -DisplayName 'YourAppDisplayName'
}

# 2. Assign the role using $sp.ObjectId
New-ManagementRoleAssignment -App $sp.ObjectId -Role 'Mail Recipients'
```

Verify it was applied:
```powershell
Get-ManagementRoleAssignment -RoleAssignee $sp.ObjectId | Select Name, Role
```

> Without this step, `Set-DistributionGroup` will fail with:  
> *"The operation failed because it's out of the current user's write scope."*

**Role vs role group — quick reference**

| Name | Type | Use with |
|------|------|----------|
| `Mail Recipients` | Management role ✓ | `New-ManagementRoleAssignment -App` |
| `Recipient Management` | Role group ✗ | Cannot be used with `-App` |
| `Organization Management` | Role group ✗ | Cannot be used with `-App` |

---

## Credential Setup (per method)

### Method 1 — Client secret

1. App registration → **Certificates & secrets → New client secret**
2. Copy the secret **Value** immediately (shown once)
3. Pass it as `-ClientSecret 'your-value'`

### Method 2 — Certificate thumbprint

Generate a self-signed cert and import it to the local store, then upload the public key:

```powershell
# Generate and export (run once on the machine that will run the scripts)
$cert = New-SelfSignedCertificate -Subject 'CN=DL-Governance' `
            -CertStoreLocation 'Cert:\LocalMachine\My' `
            -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 `
            -NotAfter (Get-Date).AddYears(2)
Export-Certificate -Cert $cert -FilePath 'C:\certs\dl-governance.cer'
Write-Host "Thumbprint: $($cert.Thumbprint)"
```

Upload `dl-governance.cer` to: **App registration → Certificates & secrets → Certificates → Upload**.  
Pass the thumbprint as `-CertificateThumbprint 'AABBCC...'`

### Method 3 — Certificate file (.pfx)

```powershell
# Export the cert with its private key to a .pfx
Export-PfxCertificate -Cert $cert -FilePath 'C:\certs\dl-governance.pfx' `
    -Password (Read-Host 'PFX password' -AsSecureString)
```

Pass as `-CertificatePath 'C:\certs\dl-governance.pfx' -CertificatePassword 'pfx-password'`

---

## Required Permissions (Interactive / Delegated)

If using `-AdminUPN` (Method 4), the account must hold one of these Exchange Online roles:

| Role | Notes |
|------|-------|
| **Exchange Recipient Administrator** | Minimum required — sufficient for both scripts |
| Exchange Administrator | Broader scope; more than needed |
| Global Administrator | Use only if no Exchange-scoped admin exists |

### Cmdlets and minimum role

| Cmdlet | Script | Minimum role |
|--------|--------|-------------|
| `Get-OrganizationConfig` | Both | View-Only Organization Management |
| `Get-DistributionGroup` | Export | View-Only Recipients |
| `Get-DistributionGroupMember` | Export | View-Only Recipients |
| `Set-DistributionGroup -ManagedBy` | Set-DLOwners | Recipient Management |

---

## Input File Format

Both scripts accept a CSV produced by a prior Exchange Online export. The expected
input for **Step 1** is the ownerless-groups export:

```
Name,manageby,RecipientTypeDetails
(All) AIDHC,,MailUniversalDistributionGroup
(IS) Alerts,,MailUniversalDistributionGroup
VPN Access,,MailUniversalSecurityGroup
...
```

The `Name` column is used to match against Exchange Online. The `manageby` column
may be empty (ownerless groups). `RecipientTypeDetails` is informational.

### RecipientTypeDetails — what to expect

| Value | Type | Ownership priority |
|-------|------|--------------------|
| `MailUniversalDistributionGroup` | Standard distribution list | Normal |
| `MailUniversalSecurityGroup` | Mail-enabled security group — also controls resource access (SharePoint, VPN, apps) | **High** — wrong owner can grant/revoke access |
| `RoomList` | Room resource container for Outlook room finder | Low — not a communication DL; ownership has no operational impact |

---

## Script Reference

### `Export-DLMemberReport.ps1`

Exports group members to a flat CSV for customer review.

```powershell
. .\config.ps1   # load credentials

# Process only the ownerless groups — client secret (primary)
.\Export-DLMemberReport.ps1 `
    -InputCsvPath .\NoOwners2.csv `
    -AppId $cfg.AppId -TenantId $cfg.TenantId -ClientSecret $cfg.ClientSecret

# Same, using certificate thumbprint instead
.\Export-DLMemberReport.ps1 `
    -InputCsvPath          .\NoOwners2.csv `
    -AppId $cfg.AppId -TenantId $cfg.TenantId `
    -CertificateThumbprint $cfg.CertificateThumbprint

# Process ALL distribution groups in the tenant
.\Export-DLMemberReport.ps1 `
    -AppId $cfg.AppId -TenantId $cfg.TenantId -ClientSecret $cfg.ClientSecret

# Write output to a specific path
.\Export-DLMemberReport.ps1 `
    -InputCsvPath  .\NoOwners2.csv `
    -OutputCsvPath C:\Reports\DLMembers.csv `
    -AppId $cfg.AppId -TenantId $cfg.TenantId -ClientSecret $cfg.ClientSecret
```

**Parameters**

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-InputCsvPath` | No | — | CSV with `Name` column. If omitted, all tenant DLs are processed. |
| `-OwnerlessOnly` | No | — | Skip any group that already has at least one owner. Combine with no `-InputCsvPath` for a full tenant ownerless scan with no pre-filtering step. |
| `-OutputCsvPath` | No | `.\DLMemberReport_<timestamp>.csv` | Output file path |
| `-AppId` | No* | — | Entra app client ID (required for app-based auth) |
| `-TenantId` | No* | — | Tenant ID or domain (required with `-AppId`) |
| `-ClientSecret` | No* | — | Client secret — primary credential |
| `-CertificateThumbprint` | No | — | Cert thumbprint in local store — cert alternative |
| `-CertificatePath` | No | — | Path to .pfx file — cert alternative |
| `-CertificatePassword` | No | — | Password for `-CertificatePath` .pfx |
| `-AdminUPN` | No* | Prompted | Interactive fallback when no app-based params supplied |

**Output CSV columns**

| Column | Description |
|--------|-------------|
| `GroupName` | Display name |
| `GroupEmail` | Primary SMTP address |
| `GroupType` | `MailUniversalDistributionGroup` / `MailUniversalSecurityGroup` / `RoomList` |
| `CurrentManagedBy` | Existing owners resolved to SMTP (semicolon-separated; blank for ownerless groups) |
| `MemberDisplayName` | Member display name |
| `MemberEmail` | Member primary SMTP |
| `MemberRecipientType` | `UserMailbox`, `SharedMailbox`, `MailContact`, etc. |
| `MemberIsEligible` | `TRUE` = can be set as owner; `FALSE` = contact or nested group |
| `IsOwner` | Pre-populated `1` if already an owner, `0` otherwise. **Set to `1` for each user you want as owner.** |

> Groups with zero members still appear as one placeholder row with blank member columns.

**`IsOwner` pre-population behavior**

Members who are already owners of the group are automatically exported with `IsOwner=1`.
This means you only need to edit rows where ownership should change:

| Row in CSV | `IsOwner` on export | Action needed |
|---|---|---|
| Existing owner who is a group member | `1` (pre-populated) | Leave as-is to keep them |
| Existing owner who is **not** a member | Not in CSV | Use `-Mode Append` to preserve them |
| New candidate you want to promote | `0` | Change to `1` |
| Member you want to remove as owner | `1` (pre-populated) | Change to `0` |

---

### `Set-DLOwners.ps1`

Reads the customer-edited CSV and applies `ManagedBy`.

```powershell
. .\config.ps1   # load credentials

# Always preview first (no writes)
.\Set-DLOwners.ps1 `
    -InputCsvPath .\DLMemberReport_modified.csv `
    -AppId $cfg.AppId -TenantId $cfg.TenantId -ClientSecret $cfg.ClientSecret -WhatIf

# Apply (Append mode — default, adds owners without removing existing ones)
.\Set-DLOwners.ps1 `
    -InputCsvPath .\DLMemberReport_modified.csv `
    -AppId $cfg.AppId -TenantId $cfg.TenantId -ClientSecret $cfg.ClientSecret

# Replace mode — overwrites ManagedBy with exactly the IsOwner=1 rows
.\Set-DLOwners.ps1 `
    -InputCsvPath .\DLMemberReport_modified.csv -Mode Replace `
    -AppId $cfg.AppId -TenantId $cfg.TenantId -ClientSecret $cfg.ClientSecret

# Apply and notify each newly assigned owner by email.
# Always include -AppId/-TenantId/-ClientSecret with -Notify — without them the
# script falls back to an interactive Microsoft Graph browser auth popup.
# -NotificationFrom must be a real licensed user mailbox (not a DL or alias);
# the script validates this upfront and exits with a clear error if not found.
.\Set-DLOwners.ps1 `
    -InputCsvPath .\DLMemberReport_modified.csv `
    -Notify -NotificationFrom $cfg.AdminUPN `
    -AppId $cfg.AppId -TenantId $cfg.TenantId -ClientSecret $cfg.ClientSecret
```

**Parameters**

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-InputCsvPath` | **Yes** | — | The customer-edited member report CSV |
| `-Mode` | No | `Append` | `Append` = add `IsOwner=1` rows to whoever is already in ManagedBy (default — safe for most scenarios). `Replace` = set ManagedBy to exactly the `IsOwner=1` rows. |
| `-Notify` | No | — | Send an HTML notification email to each newly assigned owner. Requires app-based auth with `-ClientSecret` and the `Mail.Send` Microsoft Graph application permission. |
| `-NotificationFrom` | No | `-AdminUPN` value | Sender address for notification emails. **Must be a real, licensed user mailbox** in Exchange Online — DLs, contacts, and non-existent addresses cause a 404 error. Defaults to `-AdminUPN` when not set. Only used with `-Notify`. |
| `-NotificationSubject` | No | `You have been added as owner of distribution list: {{GroupName}}` | Subject line for notification emails. Supports the same `{{GroupName}}`, `{{GroupEmail}}`, `{{OwnerEmail}}` placeholders as the body template. |
| `-NotificationTemplatePath` | No | Built-in template | Path to a custom HTML file used as the email body. Supports `{{GroupName}}`, `{{GroupEmail}}`, and `{{OwnerEmail}}` placeholders. Copy `notification-template.html` as a starting point. |
| `-AuditLogPath` | No | `.\DLOwnerAssignment_<timestamp>.csv` | Audit log output path |
| `-AppId` | No* | — | Entra app client ID (required for app-based auth) |
| `-TenantId` | No* | — | Tenant ID or domain (required with `-AppId`) |
| `-ClientSecret` | No* | — | Client secret — primary credential |
| `-CertificateThumbprint` | No | — | Cert thumbprint in local store — cert alternative |
| `-CertificatePath` | No | — | Path to .pfx file — cert alternative |
| `-CertificatePassword` | No | — | Password for `-CertificatePath` .pfx |
| `-AdminUPN` | No* | Prompted | Interactive fallback when no app-based params supplied |
| `-WhatIf` | No | — | Preview all changes without writing to Exchange Online |

**Validation rules applied before any write**

- Only rows where `IsOwner=1` AND `MemberIsEligible=TRUE` are applied
- Rows with `IsOwner=1` but `MemberIsEligible=FALSE` (contacts, nested DLs) are logged as warnings and skipped
- Groups with no eligible owners in the CSV are skipped entirely

**Adding multiple owners**

A group can have multiple owners — set `IsOwner=1` on as many member rows as needed
for the same group. All of them will be applied as owners in one `Set-DistributionGroup` call.

Choosing the right mode:

| Scenario | Mode | Why |
|---|---|---|
| General use / groups already have some owners | `Append` (default) | Merges CSV owners with existing ManagedBy; no one is accidentally removed |
| Existing owner is **not** a group member | `Append` (default) | They won't appear in the CSV; Append preserves whoever is already in ManagedBy |
| Full control — you want the final list to be exactly what's in the CSV | `Replace` | Export pre-populates existing owners with `IsOwner=1`; add new ones the same way; only `IsOwner=1` rows end up in ManagedBy |
| Remove a specific owner | `Replace` | Set their `IsOwner` to `0` (or leave as `0`); only `IsOwner=1` rows are applied |

> **Append is the safe default.** Use `Replace` only when you want full control over the final
> ManagedBy list and have verified the export captured all existing owners you intend to keep.

**Customising the notification email**

The email body is driven by an HTML template with three placeholders:

| Placeholder | Replaced with |
|-------------|--------------|
| `{{GroupName}}` | Display name of the distribution list |
| `{{GroupEmail}}` | Primary SMTP address of the distribution list |
| `{{OwnerEmail}}` | Email address of the recipient (the new owner) |

Copy the included `notification-template.html` and edit freely — branding, wording, layout — then point `-NotificationTemplatePath` at your copy. Use `-NotificationSubject` to customise the subject line. The script itself never needs to be touched.

```powershell
.\Set-DLOwners.ps1 -InputCsvPath .\DLMemberReport_modified.csv `
    -Notify -NotificationFrom $cfg.AdminUPN `
    -NotificationSubject 'Action required: you are now owner of {{GroupName}}' `
    -NotificationTemplatePath .\my-notification-template.html `
    -AppId $cfg.AppId -TenantId $cfg.TenantId -ClientSecret $cfg.ClientSecret
```

**Audit log columns**

| Column | Description |
|--------|-------------|
| `Timestamp` | UTC datetime of the operation |
| `GroupName` / `GroupEmail` | Identifies the group |
| `Mode` | `Replace` or `Append` |
| `NewOwners` | Semicolon-separated list of applied owners |
| `Status` | `Applied` / `Failed` / `WhatIf` |
| `Error` | Error message if `Status=Failed` |

---

## Full Workflow Example

### Option A — you have a pre-filtered ownerless groups CSV

```powershell
# 1. Load credentials (once per session)
. .\config.ps1

# 2. Export members — only the groups in the CSV
.\Export-DLMemberReport.ps1 -InputCsvPath .\NoOwners2.csv `
    -AppId $cfg.AppId -TenantId $cfg.TenantId -ClientSecret $cfg.ClientSecret

# 3. Open DLMemberReport_<timestamp>.csv in Excel
#    - Filter MemberIsEligible = TRUE
#    - Filter GroupType != RoomList  (skip room lists)
#    - Set IsOwner = 1 for desired owners per group
#    - Save as DLMemberReport_modified.csv

# 4. Preview (no writes)
.\Set-DLOwners.ps1 -InputCsvPath .\DLMemberReport_modified.csv `
    -AppId $cfg.AppId -TenantId $cfg.TenantId -ClientSecret $cfg.ClientSecret -WhatIf

# 5a. Apply without notifications
.\Set-DLOwners.ps1 -InputCsvPath .\DLMemberReport_modified.csv `
    -AppId $cfg.AppId -TenantId $cfg.TenantId -ClientSecret $cfg.ClientSecret

# 5b. Apply and email each newly assigned owner (requires Mail.Send Graph permission on the app)
.\Set-DLOwners.ps1 -InputCsvPath .\DLMemberReport_modified.csv `
    -Notify -NotificationFrom $cfg.AdminUPN `
    -AppId $cfg.AppId -TenantId $cfg.TenantId -ClientSecret $cfg.ClientSecret
```

### Option B — scan the whole tenant, skip groups with owners automatically

```powershell
# 1. Load credentials (once per session)
. .\config.ps1

# 2. Export — no CSV needed, -OwnerlessOnly filters automatically
.\Export-DLMemberReport.ps1 -OwnerlessOnly `
    -AppId $cfg.AppId -TenantId $cfg.TenantId -ClientSecret $cfg.ClientSecret

# 3. Open DLMemberReport_<timestamp>.csv in Excel
#    - Filter MemberIsEligible = TRUE
#    - Filter GroupType != RoomList  (skip room lists)
#    - Set IsOwner = 1 for desired owners per group
#    - Save as DLMemberReport_modified.csv

# 4. Preview (no writes)
.\Set-DLOwners.ps1 -InputCsvPath .\DLMemberReport_modified.csv `
    -AppId $cfg.AppId -TenantId $cfg.TenantId -ClientSecret $cfg.ClientSecret -WhatIf

# 5a. Apply without notifications
.\Set-DLOwners.ps1 -InputCsvPath .\DLMemberReport_modified.csv `
    -AppId $cfg.AppId -TenantId $cfg.TenantId -ClientSecret $cfg.ClientSecret

# 5b. Apply and email each newly assigned owner (requires Mail.Send Graph permission on the app)
.\Set-DLOwners.ps1 -InputCsvPath .\DLMemberReport_modified.csv `
    -Notify -NotificationFrom $cfg.AdminUPN `
    -AppId $cfg.AppId -TenantId $cfg.TenantId -ClientSecret $cfg.ClientSecret
```

---

## Output Files Summary

| File | Produced by | Purpose |
|------|------------|---------|
| `DLMemberReport_<ts>.csv` | Export script | Customer fills in `IsOwner=1` |
| `DLMemberReport_<ts>_errors.csv` | Export script | Groups that failed member enumeration |
| `DLOwnerAssignment_<ts>.csv` | Set script | Audit trail of every change (or WhatIf preview) |
