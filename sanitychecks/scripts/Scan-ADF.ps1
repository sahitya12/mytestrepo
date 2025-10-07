param(
  [Parameter(Mandatory=$true)][string]$TenantId,
  [Parameter(Mandatory=$true)][string]$ClientId,
  [Parameter(Mandatory=$true)][string]$ClientSecret,
  [string]$adh_group="",
  [switch]$ScanAll,
  [string]$OutputDir="",
  [string]$BranchName=""
)
$ErrorActionPreference='Stop'
Import-Module Az.Accounts -ErrorAction Stop
Import-Module Az.DataFactory -ErrorAction Stop
Import-Module Az.Resources  -ErrorAction Stop

function Ensure-Dir([string]$p){ if([string]::IsNullOrWhiteSpace($p)){ $p = Join-Path (Get-Location) 'adf-out' } if(-not(Test-Path $p)){ New-Item -ItemType Directory -Path $p -Force | Out-Null } return $p }

$sec=ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$cred=[pscredential]::new($ClientId,$sec)
Connect-AzAccount -ServicePrincipal -Tenant $TenantId -Credential $cred | Out-Null

$OutputDir=Ensure-Dir $OutputDir
$stamp=(Get-Date).ToString('yyyyMMdd_HHmmss')
$outCsv=Join-Path $OutputDir "adf_scan_$stamp.csv"
$outHtml=Join-Path $OutputDir "adf_scan_$stamp.html"
$outJson=Join-Path $OutputDir "adf_scan_$stamp.json"
$rows=New-Object System.Collections.Generic.List[object]

$subs = if($ScanAll){ Get-AzSubscription | ? { $_.Name -match 'ADH' } } else { Get-AzSubscription | ? { $_.Name -match 'ADH' -and $_.Name -match [regex]::Escape($adh_group) } }

foreach($sub in $subs){
  Set-AzContext -Tenant $TenantId -SubscriptionId $sub.Id | Out-Null
  $factories = Get-AzDataFactoryV2 -ErrorAction SilentlyContinue
  if(-not $factories){ 
    $rows.Add([pscustomobject]@{ SubscriptionName=$sub.Name; FactoryRG=''; FactoryName=''; Exists='NO'; LinkedServices=''; IntegrationRuntimes='' })
    continue
  }
  foreach($f in $factories){
    $ls = Get-AzDataFactoryV2LinkedService -ResourceGroupName $f.ResourceGroupName -DataFactoryName $f.Name -ErrorAction SilentlyContinue
    $ir = Get-AzDataFactoryV2IntegrationRuntime -ResourceGroupName $f.ResourceGroupName -DataFactoryName $f.Name -ErrorAction SilentlyContinue
    $rows.Add([pscustomobject]@{
      SubscriptionName=$sub.Name; FactoryRG=$f.ResourceGroupName; FactoryName=$f.Name; Exists='YES'
      LinkedServices=($(if($ls){ ($ls | % { $_.Name }) -join '; ' }else{''}))
      IntegrationRuntimes=($(if($ir){ ($ir | % { "$($_.Name):$($_.Type)"} ) -join '; ' }else{''}))
    })
  }
}

$rows | Export-Csv $outCsv -NoTypeInformation -Encoding UTF8
($rows | ConvertTo-Html -Title "ADF Scan $stamp" -PreContent "<h2>ADF Scan ($BranchName)</h2>") | Set-Content -Path $outHtml -Encoding UTF8
$rows | ConvertTo-Json -Depth 6 | Set-Content -Path $outJson -Encoding UTF8
Write-Host "CSV:  $outCsv"
Write-Host "HTML: $outHtml"
Write-Host "JSON: $outJson"
