param(
  [Parameter(Mandatory=$true)][string]$TenantId,
  [Parameter(Mandatory=$true)][string]$ClientId,
  [Parameter(Mandatory=$true)][string]$ClientSecret,

  [switch]$ScanAll,
  [string]$CustodianFilter = "",
  [string[]]$SubscriptionIds = @(),

  [string]$OutputDir = "",
  [string]$TeamsWebhookUrl = "",
  [string[]]$RequiredTagKeys = @('owner','environment','costcenter'),
  [switch]$PerTagRows
)

$ErrorActionPreference='Stop'
Import-Module Az.Accounts -ErrorAction Stop
Import-Module Az.Resources -ErrorAction Stop

function EnvFromSub([string]$n){ if($n -match '(?i)\b(prod|production|prd)\b'){'PRODUCTION'}else{'NONPRODUCTION'} }
function OutDir([string]$d){ if([string]::IsNullOrWhiteSpace($d)){$d=Join-Path (Get-Location) 'rg-tags-out'}; if(-not(Test-Path $d)){New-Item -ItemType Directory -Path $d|Out-Null}; $d }
function TagFlat($t){ if(-not $t){return ''}; ($t.GetEnumerator()|%{"$($_.Key)=$($_.Value)"}) -join '; ' }
function HasTag($t,[string]$k){ if(-not $t){return $false}; foreach($x in $t.Keys){ if($x -ieq $k){return $true} }; $false }

$sec=ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$creds=New-Object System.Management.Automation.PSCredential($ClientId,$sec)
Connect-AzAccount -ServicePrincipal -Tenant $TenantId -Credential $creds | Out-Null

$subs=@()
if($SubscriptionIds -and $SubscriptionIds.Count -gt 0){
  foreach($sid in $SubscriptionIds){ $s=Get-AzSubscription -SubscriptionId $sid -ErrorAction SilentlyContinue; if($s){$subs+=$s} }
  if(-not $subs){ throw "None of the provided -SubscriptionIds were accessible." }
}else{
  $cand=Get-AzSubscription | ?{ $_.Name -match '(?i)ADH' }
  if($CustodianFilter){ $cand = $cand | ?{ $_.Name -match [regex]::Escape($CustodianFilter) } }
  if(-not $ScanAll.IsPresent -and -not $CustodianFilter){ throw "Provide -ScanAll or -CustodianFilter or -SubscriptionIds." }
  if(-not $cand){ throw "No subscriptions matched selection." }
  $subs=$cand
}

$OutputDir=OutDir $OutputDir
$stamp=(Get-Date).ToString('yyyyMMdd_HHmmss')
$csv=Join-Path $OutputDir "rg-tags_${stamp}.csv"
$html=Join-Path $OutputDir "rg-tags_${stamp}.html"

$rows=New-Object System.Collections.Generic.List[object]

foreach($sub in $subs){
  Set-AzContext -Tenant $TenantId -SubscriptionId $sub.Id | Out-Null
  $env=EnvFromSub $sub.Name
  $rgs=Get-AzResourceGroup -ErrorAction SilentlyContinue

  foreach($rg in $rgs){
    if($PerTagRows){
      if($rg.Tags -and $rg.Tags.Count -gt 0){
        foreach($k in $rg.Tags.Keys){
          $rows.Add([pscustomobject]@{SubscriptionName=$sub.Name;SubscriptionId=$sub.Id;Environment=$env;ResourceGroup=$rg.ResourceGroupName;TagName=$k;TagValue=$rg.Tags[$k]})
        }
      }else{
        $rows.Add([pscustomobject]@{SubscriptionName=$sub.Name;SubscriptionId=$sub.Id;Environment=$env;ResourceGroup=$rg.ResourceGroupName;TagName='';TagValue=''})
      }
    }else{
      $base=[ordered]@{
        SubscriptionName=$sub.Name;SubscriptionId=$sub.Id;Environment=$env;ResourceGroup=$rg.ResourceGroupName
        TagCount=($rg.Tags?$rg.Tags.Count:0);TagsFlat=(TagFlat $rg.Tags)
      }
      foreach($k in $RequiredTagKeys){ $base["Has_$($k)"]=(HasTag $rg.Tags $k) }
      $rows.Add([pscustomobject]$base)
    }
  }
}

$rows|Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
($rows|ConvertTo-Html -Title "RG Tags - $stamp" -PreContent "<h2>RG Tags - $stamp</h2>") | Set-Content -Path $html -Encoding UTF8
Write-Host "RG Tags CSV: $csv"; Write-Host "RG Tags HTML: $html"

if($TeamsWebhookUrl){
  $rgCount = ($rows | Select-Object -ExpandProperty ResourceGroup -Unique | Measure-Object).Count
  $subCount= ($rows | Select-Object -ExpandProperty SubscriptionId -Unique | Measure-Object).Count
  $summary="RG Tag Scan $stamp`nSubscriptions:$subCount`nRGs(unique):$rgCount`n$csv`n$html"
  try{Invoke-RestMethod -Method Post -Uri $TeamsWebhookUrl -ContentType 'application/json' -Body (@{text=$summary}|ConvertTo-Json)}catch{Write-Warning "Teams post failed: $_"}
}
