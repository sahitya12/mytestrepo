param(
    # Excel file with sheet 'KVSecrets' and column 'SECRET NAME'
    [Parameter(Mandatory = $true)][string]$FilePath,

    # Short custodian code appearing in subscription names as a suffix: *_ADH<adh_group>
    # e.g., CSM -> matches dev_..._ADHCSM and prd_..._ADHCSM
    [Parameter(Mandatory = $true)][string]$adh_group,

    # Optional: include raw secret values in CSV/HTML (set $false to mask)
    [bool]$IncludeSecretValues = $true,

    # Optional: output folder (defaults next to your Excel file)
    [string]$OutputDir
)

# ---------------- Modules ----------------
if (-not (Get-Module -ListAvailable -Name Az)) {
    Install-Module -Name Az -Force -AllowClobber -Scope CurrentUser -Confirm:$false
}
Import-Module Az -ErrorAction Stop

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser -Confirm:$false
}
Import-Module ImportExcel -ErrorAction Stop

# ---------------- Login check ----------------
try {
    $ctx = Get-AzContext -ErrorAction Stop
    if (-not $ctx.Account) { throw "No Az PowerShell context" }
} catch {
    Write-Host "No Az PowerShell login detected. Launching interactive login window..." -ForegroundColor Yellow
    # This will prompt you to choose an account/tenant if needed
    Connect-AzAccount -ErrorAction Stop | Out-Null
}

# ---------------- Resolve paths ----------------
if (-not (Test-Path -LiteralPath $FilePath)) {
    # Try common alternates if they passed without extension or with xlx/xls
    $try = @(
        $FilePath,
        [IO.Path]::ChangeExtension($FilePath,'xlsx'),
        [IO.Path]::ChangeExtension($FilePath,'xls'),
        [IO.Path]::ChangeExtension($FilePath,'xlx')
    ) | Select-Object -Unique
    $found = $try | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if (-not $found) {
        throw "Workbook not found. Tried:`n - " + ($try -join "`n - ")
    }
    $FilePath = $found
}

if (-not $OutputDir) {
    $OutputDir = Join-Path -Path (Split-Path -Parent $FilePath) -ChildPath 'kv-scan-local'
}
if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
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

if (-not $SecretNames -or $SecretNames.Count -eq 0) {
    throw "No secret names found in column '$SecretColumnName'."
}

Write-Host "Loaded $($SecretNames.Count) secrets from '$WorksheetName'." -ForegroundColor Cyan

# ---------------- Pick subscriptions by name ----------------
# Matches ANY subscription ending with _ADH<adh_group> (case-insensitive), e.g. *_ADHCSM
$suffix = "_ADH$adh_group"
$allSubs = Get-AzSubscription -ErrorAction Stop
$subs = $allSubs | Where-Object { $_.Name -like "*$suffix" }

if (-not $subs) {
    throw "No subscriptions matched '*$suffix'. Example expected: dev_...$suffix or prd_...$suffix"
}

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
    try {
        $vaults = Get-AzKeyVault -ErrorAction Stop
    } catch {
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

# ---------------- Write CSV ----------------
$ts = (Get-Date).ToString('yyyyMMdd_HHmmss')
$csvPath = Join-Path $OutputDir "KV-Secret-Check-$($adh_group.ToUpper())-$ts.csv"

$results |
  Sort-Object SubscriptionName, VaultName, SecretName |
  Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "`nCSV saved: $csvPath" -ForegroundColor Cyan

# ---------------- Write HTML (green/red) ----------------
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
