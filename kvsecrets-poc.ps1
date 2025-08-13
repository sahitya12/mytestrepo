<# ====================== CONFIGURE THESE VALUES ====================== #>

# Excel file path (sheet 'KVSecrets', column 'SECRET NAME')
$FilePath  = "C:\path\to\resource_sanitychecks.xlsx"

# Short custodian suffix used in subscription names (e.g., CSM -> matches *_ADHCSM)
$adh_group = "CSM"

# Output directory (if blank, a 'kv-scan-local' folder will be created next to the Excel)
$OutputDir = ""

# Include raw secret values in outputs? ($true = include, $false = mask)
$IncludeSecretValues = $true

# --- Service Principal credentials for SPN login ---
$ClientId     = "<appId-guid>"
$ClientSecret = "<sp-secret>"
$TenantId     = "<tenant-guid>"

# Also log in Azure CLI context? (usually not required) 
$AlsoAzCliLogin = $false

<# ========================= DO NOT EDIT BELOW ======================== #>

# ---------------- Modules (import only) ----------------
try { Import-Module Az -ErrorAction Stop }
catch { throw "Az module not found. Install once: Install-Module Az -Scope CurrentUser" }

try { Import-Module ImportExcel -ErrorAction Stop }
catch { throw "ImportExcel module not found. Install once: Install-Module ImportExcel -Scope CurrentUser" }

# ---------------- Validate & resolve input path ----------------
if ([string]::IsNullOrWhiteSpace($FilePath)) { throw "FilePath is required and cannot be empty." }

function Resolve-Workbook {
    param([string]$p)
    $candidates = @(
        $p,
        [IO.Path]::ChangeExtension($p,'xlsx'),
        [IO.Path]::ChangeExtension($p,'xls'),
        [IO.Path]::ChangeExtension($p,'xlx')
    ) | Where-Object { $_ -and $_.Trim() } | Select-Object -Unique
    foreach($c in $candidates){
        if(Test-Path -LiteralPath $c){ return (Resolve-Path -LiteralPath $c).Path }
    }
    throw "Workbook not found. Tried:`n - $($candidates -join "`n - ")"
}
$FilePath = Resolve-Workbook -p $FilePath

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path -Path (Split-Path -Parent $FilePath) -ChildPath 'kv-scan-local'
}
if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

# ---------------- SPN login ----------------
if ([string]::IsNullOrWhiteSpace($ClientId) -or [string]::IsNullOrWhiteSpace($ClientSecret) -or [string]::IsNullOrWhiteSpace($TenantId)) {
    throw "ClientId / ClientSecret / TenantId must be set in the CONFIG section."
}
$secure = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$creds  = New-Object System.Management.Automation.PSCredential ($ClientId, $secure)

try {
    Connect-AzAccount -ServicePrincipal -Tenant $TenantId -Credential $creds -ErrorAction Stop | Out-Null
} catch {
    throw "Connect-AzAccount (SPN) failed: $($_.Exception.Message)"
}

if ($AlsoAzCliLogin) {
    try { az login --service-principal -u $ClientId -p $ClientSecret --tenant $TenantId 1>$null }
    catch { Write-Warning "az login failed: $($_.Exception.Message)" }
}

# ---------------- Load secrets from Excel ----------------
$WorksheetName    = 'KVSecrets'
$SecretColumnName = 'SECRET NAME'

try {
    $data = Import-Excel -Path $FilePath -WorksheetName $WorksheetName -ErrorAction Stop
} catch {
    throw "Failed to read sheet '$WorksheetName' from '$FilePath': $($_.Exception.Message)"
}
if (-not ($data | Get-Member -Name $SecretColumnName -MemberType NoteProperty)) {
    throw "Column '$SecretColumnName' not found in sheet '$WorksheetName'."
}
$SecretNames = $data.$SecretColumnName |
    Where-Object { $_ -and $_.ToString().Trim() } |
    ForEach-Object { $_.ToString().Trim() } |
    Sort-Object -Unique
if (-not $SecretNames) { throw "No secret names found in column '$SecretColumnName'." }
Write-Host "Loaded $($SecretNames.Count) secrets from '$WorksheetName'." -ForegroundColor Cyan

# ---------------- Pick subscriptions by name (ends with *_ADH<adh_group>) ----------------
$suffix = "_ADH$adh_group"
$allSubs = Get-AzSubscription -ErrorAction Stop
$subs = $allSubs | Where-Object { $_.Name -like "*$suffix" }
if (-not $subs) { throw "No subscriptions matched '*$suffix'. Expected like: dev_...$suffix or prd_...$suffix" }
Write-Host "Scanning $($subs.Count) subscription(s) matching '*$suffix'..." -ForegroundColor Cyan

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
    Set-AzContext -SubscriptionId $sub.Id -ErrorAction Stop | Out-Null

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
