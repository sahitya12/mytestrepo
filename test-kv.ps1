<# ====================== CONFIGURE THESE VALUES ====================== #>

# Excel file path (.xlsx/.xlsm), with worksheet "KVSecrets" and column "SECRET_NAME"
$FilePath  = "C:\path\to\resource_sanitychecks.xlsx"

# Custodian suffix used in subscription names (e.g., CSM -> matches *_ADHCSM)
$adh_group = "CSM"

# Output directory (leave blank to create 'kv-scan-local' next to the Excel file)
$OutputDir = ""

# Include raw secret values in outputs? ($true = include, $false = mask)
$IncludeSecretValues = $true

# --- Service Principal (SPN) credentials for Azure login ---
$TenantId     = "<tenant-guid>"
$ClientId     = "<appId-guid>"
$ClientSecret = "<sp-secret>"

# Also log in Azure CLI? (not required for this script)
$AlsoAzCliLogin = $false

<# ========================= DO NOT EDIT BELOW ======================== #>

# ---------------- Modules (import only) ----------------
try { Import-Module Az.Accounts -ErrorAction Stop } catch { throw "Az.Accounts not found. Install once: Install-Module Az -Scope CurrentUser" }
try { Import-Module Az.KeyVault  -ErrorAction Stop } catch { throw "Az.KeyVault not found. Install once: Install-Module Az -Scope CurrentUser" }
try { Import-Module ImportExcel  -ErrorAction Stop } catch { throw "ImportExcel not found. Install once: Install-Module ImportExcel -Scope CurrentUser" }

# ---------------- Validate & resolve input path ----------------
if ([string]::IsNullOrWhiteSpace($FilePath)) { throw "FilePath is required and cannot be empty." }
if (-not (Test-Path -LiteralPath $FilePath)) { throw "Workbook not found at: $FilePath" }

$ext = [IO.Path]::GetExtension($FilePath)
if ($ext -notin @(".xlsx",".xlsm")) {
    throw "Unsupported Excel format '$ext'. Please save as .xlsx or .xlsm."
}
$FilePath = (Resolve-Path -LiteralPath $FilePath).Path

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path -Path (Split-Path -Parent $FilePath) -ChildPath 'kv-scan-local'
}
if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

# ---------------- SPN login & context scoping ----------------
if ([string]::IsNullOrWhiteSpace($ClientId) -or [string]::IsNullOrWhiteSpace($ClientSecret) -or [string]::IsNullOrWhiteSpace($TenantId)) {
    throw "ClientId / ClientSecret / TenantId must be set in the CONFIG section."
}
$secure = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$creds  = New-Object System.Management.Automation.PSCredential ($ClientId, $secure)

try {
    Connect-AzAccount -ServicePrincipal -Tenant $TenantId -Credential $creds -ErrorAction Stop | Out-Null
    Set-AzContext -Tenant $TenantId -ErrorAction Stop | Out-Null
} catch {
    throw "Connect-AzAccount (SPN) failed: $($_.Exception.Message)"
}

if ($AlsoAzCliLogin) {
    try { az login --service-principal -u $ClientId -p $ClientSecret --tenant $TenantId 1>$null }
    catch { Write-Warning "az login failed: $($_.Exception.Message)" }
}

# ---------------- Load secrets from Excel (fixed worksheet & column) ----------------
$WorksheetName    = 'KVSecrets'
$SecretColumnName = 'SECRET_NAME'   # exact header expected in row 1

try {
    $data = Import-Excel -Path $FilePath -WorksheetName $WorksheetName -ErrorAction Stop
} catch {
    throw "Failed to read sheet '$WorksheetName' from '$FilePath': $($_.Exception.Message)"
}

if (-not ($data | Get-Member -Name $SecretColumnName -MemberType NoteProperty)) {
    $cols = if ($data) { ($data[0].psobject.Properties.Name -join ', ') } else { "<no columns>" }
    throw "Column '$SecretColumnName' not found in sheet '$WorksheetName'. Columns seen: $cols"
}

$SecretNames = $data.$SecretColumnName |
    Where-Object { $_ -and $_.ToString().Trim() } |
    ForEach-Object { $_.ToString().Trim() } |
    Sort-Object -Unique

if (-not $SecretNames -or $SecretNames.Count -eq 0) {
    throw "No secret names found under column '$SecretColumnName' in sheet '$WorksheetName'."
}
Write-Host "Loaded $($SecretNames.Count) secret(s) from '$WorksheetName' ('$SecretColumnName')." -ForegroundColor Cyan

# ---------------- Pick subscriptions (THIS tenant) by pattern *_ADH<adh_group> ----------------
$suffix = "_ADH$adh_group"
$allSubs = Get-AzSubscription -TenantId $TenantId -ErrorAction Stop
if (-not $allSubs) { throw "No subscriptions found in tenant $TenantId for the given SPN." }

$subs = $allSubs | Where-Object { $_.Name -like "*$suffix" }
if (-not $subs) {
    $avail = ($allSubs.Name -join '; ')
    throw "No subscriptions matched '*$suffix'. Available in tenant: $avail"
}
Write-Host "Scanning $($subs.Count) subscription(s) matching '*$suffix' (tenant $TenantId)..." -ForegroundColor Cyan

# ---------------- Helpers ----------------
function Render-Value {
    param([string]$Value, [bool]$Include)
    if ($Include) { return $Value }
    if ([string]::IsNullOrEmpty($Value)) { return "" }
    if ($Value.Length -le 6) { return "***masked***" }
    return "{0}{1}{2}" -f $Value.Substring(0,2),'***masked***',$Value.Substring($Value.Length-2,2)
}

# ---------------- Scan ----------------
$results = New-Object System.Collections.Generic.List[object]

foreach ($sub in $subs) {
    Write-Host "`n=== Subscription: $($sub.Name) [$($sub.Id)] ===" -ForegroundColor Green
    Set-AzContext -SubscriptionId $sub.Id -Tenant $TenantId -ErrorAction Stop | Out-Null

    $vaults = @()
    try { $vaults = Get-AzKeyVault -ErrorAction Stop }
    catch {
        Write-Warning "  Could not list Key Vaults: $($_.Exception.Message)"
        $vaults = @()
    }

    if (-not $vaults -or $vaults.Count -eq 0) {
        foreach ($secName in $SecretNames) {
            $results.Add([pscustomobject]@{
                SubscriptionName = $sub.Name
                SubscriptionId   = $sub.Id
                VaultName        = "N/A (no key vaults)"
                SecretName       = $secName
                Exists           = "No"
                Value            = ""
                Version          = ""
                ContentType      = ""
                UpdatedOn        = ""
                MissingReason    = "No Key Vaults in subscription"
            })
        }
        continue
    }

    foreach ($kv in $vaults) {
        Write-Host "  -> Vault: $($kv.VaultName)"
        foreach ($secName in $SecretNames) {
            try {
                $sec = Get-AzKeyVaultSecret -VaultName $kv.VaultName -Name $secName -ErrorAction Stop
                $results.Add([pscustomobject]@{
                    SubscriptionName = $sub.Name
                    SubscriptionId   = $sub.Id
                    VaultName        = $kv.VaultName
                    SecretName       = $secName
                    Exists           = "Yes"
                    Value            = (Render-Value -Value $sec.SecretValueText -Include $IncludeSecretValues)
                    Version          = $sec.Version
                    ContentType      = $sec.ContentType
                    UpdatedOn        = $sec.Updated
                    MissingReason    = ""
                })
            } catch {
                $results.Add([pscustomobject]@{
                    SubscriptionName = $sub.Name
                    SubscriptionId   = $sub.Id
                    VaultName        = $kv.VaultName
                    SecretName       = $secName
                    Exists           = "No"
                    Value            = ""
                    Version          = ""
                    ContentType      = ""
                    UpdatedOn        = ""
                    MissingReason    = "Secret not found"
                })
            }
        }
    }
}

# ---------------- Outputs ----------------
$ts = (Get-Date).ToString('yyyyMMdd_HHmmss')
$csvPath = Join-Path $OutputDir "KV-Secret-Check-$($adh_group.ToUpper())-$ts.csv"

$results |
  Sort-Object SubscriptionName, VaultName, SecretName |
  Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "`nCSV saved: $csvPath" -ForegroundColor Cyan

# HTML (green/red)
$style = @"
table { border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px; }
th, td { border: 1px solid #ddd; padding: 6px 8px; }
th { background: #f2f2f2; text-align: left; }
tr:nth-child(even) { background: #fafafa; }
.good { color: #0b6a0b; font-weight: 600; }
.bad  { color: #a40000; font-weight: 600; }
.mono { font-family: Consolas, monospace; }
"@

$rowsHtml = ($results | Sort-Object SubscriptionName, VaultName, SecretName | ForEach-Object {
    $cls = if ($_.Exists -eq "Yes") { "good" } else { "bad" }
    "<tr>" +
    "<td>$($_.SubscriptionName)</td>" +
    "<td class='mono'>$($_.SubscriptionId)</td>" +
    "<td>$($_.VaultName)</td>" +
    "<td>$($_.SecretName)</td>" +
    "<td class='$cls'>$($_.Exists)</td>" +
    "<td class='mono'>$($_.Value)</td>" +
    "<td class='mono'>$($_.Version)</td>" +
    "<td>$($_.ContentType)</td>" +
    "<td>$($_.UpdatedOn)</td>" +
    "<td>$($_.MissingReason)</td>" +
    "</tr>"
}) -join "`n"

$html = @"
<!doctype html>
<html>
<head>
<meta charset="utf-8"/>
<title>Key Vault Secret Check — ADH $($adh_group.ToUpper())</title>
<style>$style</style>
</head>
<body>
<h2>Key Vault Secret Check — ADH $($adh_group.ToUpper())</h2>
<p>Run: $(Get-Date)</p>
<p><strong>Workbook:</strong> $([System.Web.HttpUtility]::HtmlEncode($FilePath))</p>
<table>
<thead>
<tr>
  <th>Subscription</th>
  <th>SubscriptionId</th>
  <th>KeyVault</th>
  <th>Secret</th>
  <th>Status</th>
  <th>Value</th>
  <th>Version</th>
  <th>ContentType</th>
  <th>UpdatedOn</th>
  <th>MissingReason</th>
</tr>
</thead>
<tbody>
$rowsHtml
</tbody>
</table>
</body>
</html>
"@

$htmlPath = Join-Path $OutputDir "KV-Secret-Check-$($adh_group.ToUpper())-$ts.html"
Set-Content -Path $htmlPath -Value $html -Encoding UTF8

Write-Host "HTML saved: $htmlPath" -ForegroundColor Cyan
Write-Host "`nDone." -ForegroundColor Green
