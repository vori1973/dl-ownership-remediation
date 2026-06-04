# config.sample.ps1
# ------------------
# Copy this file to config.ps1 and fill in your values.
# config.ps1 is gitignored — never commit real credentials.
#
# Usage:
#   . .\config.ps1
#   .\Export-DLMemberReport.ps1 -InputCsvPath .\NoOwners2.csv `
#       -AppId $cfg.AppId -TenantId $cfg.TenantId -ClientSecret $cfg.ClientSecret

$cfg = @{
    # ── App-based authentication ─────────────────────────────────────────────
    # Required for all three app-based methods (Methods 1–3).
    AppId    = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'   # Entra app client ID
    TenantId = 'contoso.onmicrosoft.com'                 # Tenant ID or primary domain

    # ── Method 1: Client secret (primary) ────────────────────────────────────
    # Entra app registration → Certificates & secrets → New client secret
    ClientSecret = 'your-client-secret-value'

    # ── Method 2: Certificate thumbprint (optional — preferred security) ─────
    # Run once to generate:
    #   $cert = New-SelfSignedCertificate -Subject 'CN=DL-Governance' `
    #               -CertStoreLocation 'Cert:\LocalMachine\My' `
    #               -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 `
    #               -NotAfter (Get-Date).AddYears(2)
    #   Export-Certificate -Cert $cert -FilePath 'C:\certs\dl-governance.cer'
    # Upload the .cer to App registration → Certificates & secrets → Certificates
    # CertificateThumbprint = 'AABBCCDDEEFF0011223344556677889900AABBCC'

    # ── Method 3: Certificate file (optional) ────────────────────────────────
    # CertificatePath     = 'C:\certs\dl-governance.pfx'
    # CertificatePassword = 'pfx-password'               # omit if not password-protected

    # ── Interactive fallback (Method 4) ──────────────────────────────────────
    # Used only when no AppId/TenantId are supplied.
    # AdminUPN = 'admin@contoso.com'
}
