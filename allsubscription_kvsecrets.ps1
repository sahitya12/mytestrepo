# KeyVault_Secrets_Scan_All_ADH.ps1

# --- CONFIG: set these ---
$filePath  = "C:\Path\To\resource_sanitychecks.xlsx"
$outputDir = "C:\KVScanResults"

# Service principal for az login
$ClientId     = "<appId-guid>"
$ClientSecret = "<sp-secret>"
$TenantId     = "<tenant-guid>"

# --- Imports (no installs) ---
Import-Module ImportExcel -ErrorAction Stop

# Ensure output folder
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# Login (Azure CLI context since your script uses 'az ...')
az login --service-principal -u $ClientId -p $ClientSecret --tenant $TenantId | Out-Null

Write-Host "Reading input secret list from $filePath..." -ForegroundColor Cyan
$secretList = Import-Excel -Path $filePath -WorksheetName "KeyVaults Secrets" -ErrorAction Stop |
              Select-Object -ExpandProperty "SECRET NAME" -Unique
if (-not $secretList) { Write-Error "No secrets found in 'KeyVaults Secrets'."; exit 1 }

Write-Host "Fetching subscriptions..." -ForegroundColor Cyan
$subscriptions = az account list --query "[?contains(name, '_ADH')].{Name:name,Id:id}" -o json | ConvertFrom-Json
if (-not $subscriptions) { Write-Error "No subscriptions found ending with _ADH..."; exit 1 }

$htmlRows = @()
$excelPath = Join-Path $outputDir "KeyVault_Secrets_Scan_All_ADH.xlsx"
$firstSheet = $true

foreach ($sub in $subscriptions) {
    $subName = $sub.Name
    $subId   = $sub.Id
    # last 3 letters as sheet name, keep Excel-safe
    $suffix  = ($subName.Substring([Math]::Max(0, $subName.Length-3))) -replace '[:\\/?*\[\]]',''
    if ([string]::IsNullOrWhiteSpace($suffix)) { $suffix = "SUB" }

    Write-Host "`nScanning subscription: $subName ($subId)" -ForegroundColor Yellow
    az account set --subscription $subId | Out-Null

    $kvList = az keyvault list --query "[].name" -o tsv
    if (-not $kvList) {
        Write-Warning "No KeyVaults found in $subName."
        $kvList = @()
    }

    $results = @()

    foreach ($kv in $kvList) {
        foreach ($secretName in $secretList) {
            $secretValue = ""
            $status = "Missing ❌"
            $color  = "#f8d7da"
            try {
                $sv = az keyvault secret show --vault-name $kv --name $secretName --query "value" -o tsv 2>$null
                if ($sv) {
                    $secretValue = $sv
                    $status = "Exists ✅"
                    $color  = "#d4edda"
                }
            } catch { }

            $results += [pscustomobject]@{
                "Subscription Name" = $subName
                "KeyVault Name"     = $kv
                "Secret Name"       = $secretName
                "Secret Value"      = $secretValue
                "Status"            = $status
            }

            $htmlRows += "<tr style='background-color:$color;'>
                <td>$subName</td>
                <td>$kv</td>
                <td>$secretName</td>
                <td>$secretValue</td>
                <td>$status</td>
            </tr>"
        }
    }

    # Write Excel — create on first sheet, then append
    if ($results.Count -gt 0) {
        if ($firstSheet -and -not (Test-Path $excelPath)) {
            $results | Export-Excel -Path $excelPath -WorksheetName $suffix -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter
            $firstSheet = $false
        } else {
            $results | Export-Excel -Path $excelPath -WorksheetName $suffix -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter -Append
        }
    }
}

# HTML report
$htmlReport = @"
<html>
<head>
<title>KeyVault Secrets Scan Report</title>
<style>
  table { border-collapse: collapse; width: 100%; font-family: Arial; font-size: 12px; }
  th, td { border: 1px solid #ddd; padding: 6px 8px; }
  th { background-color: #f2f2f2; }
</style>
</head>
<body>
<h2>KeyVault Secrets Scan Report</h2>
<table>
<tr>
  <th>Subscription Name</th><th>KeyVault Name</th><th>Secret Name</th><th>Secret Value</th><th>Status</th>
</tr>
$htmlRows
</table>
</body>
</html>
"@
$htmlPath = Join-Path $outputDir "KeyVault_Secrets_Scan_All_ADH.html"
Set-Content -Path $htmlPath -Value $htmlReport -Encoding UTF8

Write-Host "`nReports generated:" -ForegroundColor Green
Write-Host "Excel: $excelPath"
Write-Host "HTML : $htmlPath"
