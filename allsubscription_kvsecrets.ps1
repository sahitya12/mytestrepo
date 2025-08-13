# Requires Az.Accounts, Az.KeyVault, ImportExcel
# Install if missing:
# Install-Module Az -Force -AllowClobber -Scope CurrentUser
# Install-Module ImportExcel -Force -Scope CurrentUser

param(
    [string]$filePath = "C:\Path\To\resource_sanitychecks.xlsx",
    [string]$outputDir = "C:\KVScanResults"
)

# Create output directory if missing
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

Write-Host "Reading input secret list from $filePath..." -ForegroundColor Cyan
$secretList = Import-Excel -Path $filePath -WorksheetName "KeyVaults Secrets" |
              Select-Object -ExpandProperty "SECRET NAME" -Unique

if (-not $secretList) {
    Write-Error "No secrets found in 'KeyVaults Secrets' sheet."
    exit 1
}

Write-Host "Fetching subscriptions..." -ForegroundColor Cyan
$subscriptions = az account list --query "[?contains(name, '_ADH')].{Name:name,Id:id}" -o json | ConvertFrom-Json

if (-not $subscriptions) {
    Write-Error "No subscriptions found ending with _ADH..."
    exit 1
}

# Prepare HTML report content
$htmlRows = @()

foreach ($sub in $subscriptions) {
    $subName = $sub.Name
    $subId   = $sub.Id
    $suffix  = $subName.Substring($subName.Length - 3)

    Write-Host "`nScanning subscription: $subName ($subId)" -ForegroundColor Yellow
    az account set --subscription $subId

    $kvList = az keyvault list --query "[].name" -o tsv
    if (-not $kvList) {
        Write-Warning "No KeyVaults found in $subName."
        continue
    }

    $results = @()

    foreach ($kv in $kvList) {
        foreach ($secretName in $secretList) {
            try {
                $secretValue = az keyvault secret show --vault-name $kv --name $secretName --query "value" -o tsv 2>$null
                if ($secretValue) {
                    $status = "Exists ✅"
                    $color = "#d4edda" # green
                } else {
                    $status = "Missing ❌"
                    $color = "#f8d7da" # red
                }
            } catch {
                $status = "Missing ❌"
                $color = "#f8d7da"
                $secretValue = ""
            }

            $results += [pscustomobject]@{
                "Subscription Name" = $subName
                "KeyVault Name"     = $kv
                "Secret Name"       = $secretName
                "Secret Value"      = $secretValue
                "Status"            = $status
            }

            # Add HTML row
            $htmlRows += "<tr style='background-color:$color;'>
                <td>$subName</td>
                <td>$kv</td>
                <td>$secretName</td>
                <td>$secretValue</td>
                <td>$status</td>
            </tr>"
        }
    }

    # Export to Excel (sheet per suffix)
    $excelPath = Join-Path $outputDir "KeyVault_Secrets_Scan_All_ADH.xlsx"
    $results | Export-Excel -Path $excelPath -WorksheetName $suffix -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter
}

# Build HTML report
$htmlReport = @"
<html>
<head>
<title>KeyVault Secrets Scan Report</title>
<style>
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #ddd; padding: 8px; }
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
