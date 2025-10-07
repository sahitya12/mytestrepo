param(
  [Parameter(Mandatory=$true)][string]$TenantId,
  [Parameter(Mandatory=$true)][string]$ClientId,
  [Parameter(Mandatory=$true)][string]$ClientSecret,
  [string]$adh_group = "",
  [switch]$ScanAll,
  [string]$OutputDir = "",
  [string]$BranchName = ""
)
$ErrorActionPreference='Stop'
Import-Module Az.Accounts -ErrorAction Stop
Import-Module Az.Resources -ErrorAction Stop

function Ensure-Dir([string]$p){ if([string]::IsNullOrWhiteSpace($p)){ $p = Join-Path (Get-Location) 'rg-tags-out' } if(-not(Test-Path $p)){ New-Item -ItemType Directory -Path $p -Force | Out-Null } return $p }

$sec = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$cred = [pscredential]::new($ClientId,$sec)
Connect-AzAccount -ServicePrincipal -Tenant $TenantId -Credential $cred | Out-Null

$OutputDir = Ensure-Dir $OutputDir
$stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
$outCsv  = Join-Path $OutputDir "rg_tags_$stamp.csv"
$outHtml = Join-Path $OutputDir "rg_tags_$stamp.html"
$outJson = Join-Path $OutputDir "rg_tags_$stamp.json"

$rows = New-Object System.Collections.Generic.List[object]
$subs = if($ScanAll){ Get-AzSubscription | ? { $_.Name -match '(?i)ADH' } } else { Get-AzSubscription | ? { $_.Name -match '(?i)ADH' -and $_.Name -match [regex]::Escape($adh_group) } }

foreach($s in $subs){
  Set-AzContext -Tenant $TenantId -SubscriptionId $s.Id | Out-Null
  $env = (if($s.Name -match '(?i)\b(prod|prd|production)\b'){'PRODUCTION'}else{'NONPRODUCTION'})
  $rgs = Get-AzResourceGroup -ErrorAction SilentlyContinue
  foreach($rg in $rgs){
    $flat = if($rg.Tags){ ($rg.Tags.GetEnumerator()|%{"$($_.Key)=$($_.Value)"}) -join '; ' } else { '' }
    $rows.Add([pscustomobject]@{
      SubscriptionName=$s.Name; SubscriptionId=$s.Id; Environment=$env
      ResourceGroup=$rg.ResourceGroupName; TagsFlat=$flat
    })
  }
}

$rows | Export-Csv $outCsv -NoTypeInformation -Encoding UTF8
($rows | ConvertTo-Html -Title "RG Tags $stamp" -PreContent "<h2>RG Tags ($BranchName)</h2>") | Set-Content -Path $outHtml -Encoding UTF8
$rows | ConvertTo-Json -Depth 5 | Set-Content -Path $outJson -Encoding UTF8
Write-Host "CSV:  $outCsv"
Write-Host "HTML: $outHtml"
Write-Host "JSON: $outJson"
