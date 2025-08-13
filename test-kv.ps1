<# ====================== CONFIGURE THESE VALUES ====================== #>

# Excel path (script auto-detects sheet & header row)
$FilePath  = "C:\path\to\resource_sanitychecks.xlsx"

# Custodian suffix used in subscription names (e.g., CSM -> matches *_ADHCSM)
$adh_group = "CSM"

# Output directory (leave blank to create 'kv-scan-local' beside the Excel)
$OutputDir = ""

# Include raw secret values in outputs? ($true = include, $false = mask)
$IncludeSecretValues = $true

# --- Service Principal credentials for SPN login ---
$ClientId     = "<appId-guid>"
$ClientSecret = "<sp-secret>"
$TenantId     = "<tenant-guid>"

# Optional CLI login too (not required for this script)
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
    throw "Unsupported Excel format '$ext'. Please save as .xlsx or .xlsm (ImportExcel cannot read .xls/.xlx)."
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

# ---------------- Robust sheet & header detection ----------------
function Normalize([string]$s){ return ($s -replace '\s+',' ' ).Trim().ToUpperInvariant() }

# candidates we accept for the header text:
$AcceptedHeaderPatterns = @(
    '^SECRET NAME$',
    '^SECRET\s*NAME$',
    '^SECRET_NAME$',
    '^SECRET$',
    '^SECRET\s*NAMES?$',
    '^SECRETNAME$',
    '^KV\s*SECRET\s*NAME$',
    '^KEYVAULTS?\s*SECRETS?$'   # last-resort broad header
)

$PreferredSheets = @('KVSecrets','KeyVaults Secrets','KeyVaults_Secrets','KeyVaultSecrets')
$WorksheetName   = $null
$HeaderRow       = $null
$SecretColumnName = $null

# Search up to first 40 rows for a matching header
function Try-FindHeader {
    param([string]$SheetName)
    for ($r=1; $r -le 40; $r++) {
        try {
            $probe = Import-Excel -Path $FilePath -WorksheetName $SheetName -StartRow $r -EndRow $r -ErrorAction Stop
            if (-not $probe) { continue }
            $props = $probe.psobject.Properties.Name
            foreach ($p in $props) {
                $np = Normalize $p
                foreach ($pat in $AcceptedHeaderPatterns) {
                    if ($np -match $pat) {
                        return @{ Found = $true; Row = $r; ColName = $p }
                    }
                }
            }
        } catch { }
    }
    return @{ Found = $false }
}

# 1) Try preferred sheets
foreach ($ws in $PreferredSheets) {
    $res = Try-FindHeader -SheetName $ws
    if ($res.Found) { $WorksheetName = $ws; $HeaderRow = $res.Row; $SecretColumnName = $res.ColName; break }
}

# 2) Fall back to scanning all sheets
if (-not $WorksheetName) {
    $info = Get-ExcelSheetInfo -Path $FilePath
    foreach ($entry in $info) {
        $res = Try-FindHeader -SheetName $entry.Name
        if ($res.Found) { $WorksheetName = $entry.Name; $HeaderRow = $res.Row; $SecretColumnName = $res.ColName; break }
    }
}

if (-not $WorksheetName -or -not $HeaderRow -or -not $SecretColumnName) {
    # Print columns we saw in row 1 for easier debugging
    try {
        $testSheet = ($PreferredSheets + (Get-ExcelSheetInfo -Path $FilePath | Select-Object -ExpandProperty Name)) | Select-Object -Unique | Select-Object -First 1
        $hdr1 = Import-Excel -Path $FilePath -WorksheetName $testSheet -StartRow 1 -EndRow 1 -ErrorAction SilentlyContinue
        $cols = if ($hdr1) { $hdr1.psobject.Properties.Name -join ', ' } else { "<no visible header in first row>" }
    } catch { $cols = "<unreadable>" }
    throw "No worksheet/header row found containing a 'SECRET NAME' column (case/space-insensitive).
Checked first 40 rows of each sheet.
First-row columns on a sample sheet: $cols"
}

# Load data using the discovered header row
try {
    $data = Import-Excel -Path $FilePath -WorksheetName $WorksheetName -StartRow $HeaderRow -ErrorAction Stop
} catch {
    throw "Failed to read sheet '$WorksheetName' (header at row $HeaderRow): $($_.Exception.Message)"
}

# Extract secret names (cast to string, trim, de-dup)
$SecretNames =
    $data |
    ForEach-Object {
        $v = $_.$SecretColumnName
        if ($null -ne $v) { $v.ToString().Trim() } else { $null }
    } |
    Where-Object { $_ } |
    Sort-Object -Unique

if (-not $SecretNames -or $SecretNames.Count -eq 0) {
    $cols = ($data[0].psobject.Properties.Name -join ', ')
    throw "Found header at row $HeaderRow (column '$SecretColumnName'), but no secret names under it.
Columns in that sheet: $cols"
}

# Diagnostics
Write-Host "Worksheet: $WorksheetName" -ForegroundColor Cyan
Write-Host "HeaderRow: $HeaderRow" -ForegroundColor Cyan
Write-Host "Secret column resolved to: '$SecretColumnName'" -ForegroundColor Cyan
Write-Host "First secrets: $((($SecretNames | Select-Object -First 10) -join ', '))" -ForegroundColor DarkCyan

# ---------------- Pick subscriptions (in THIS tenant) by name pattern *_ADH<adh_group> ----------------
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
