[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][string]$TenantId,
  [Parameter(Mandatory=$true)][string]$ClientId,
  [Parameter(Mandatory=$true)][string]$ClientSecret,
  [Parameter(Mandatory=$true)][string]$SecretCsvPath,
  [Parameter(Mandatory=$true)][string]$adh_group,
  [string]$OutputDir = "",
  [switch]$IncludeSecretValues,
  [string]$TeamsWebhookUrl = ""
)

$ErrorActionPreference = 'Stop'
Import-Module Az.Accounts -ErrorAction Stop
Import-Module Az.Resources -ErrorAction Stop
Import-Module Az.KeyVault -ErrorAction Stop

function HNorm([string]$s){ ($s -replace '[_\s]','').ToLowerInvariant() }
function LoadSecretCsv($p){
  if(-not(Test-Path $p)){ throw "CSV not found: $p" }
  $raw = Import-Csv $p
  $m=@{}; foreach($k in $raw[0].psobject.Properties.Name){ $m[(HNorm $k)]=$k }
  if(-not $m.ContainsKey('secretname')){ throw "CSV must have column SECRET_NAME" }
  $raw | ForEach-Object { $_.$($m['secretname']).Trim() } | Where-Object {$_} | Select-Object -Unique
}
function OutDir([string]$d){ if(-not $d){ $d = Join-Path (Get-Location) 'kv-secrets-out'}; if(-not(Test-Path $d)){ New-Item -ItemType Directory -Path $d|Out-Null}; $d }
function Clean([string]$s){ if(-not $s){return ''}; ([regex]::Replace($s,'\p{Cf}','')).Trim() }

$sec=ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$creds=[pscredential]::new($ClientId,$sec)
Connect-AzAccount -ServicePrincipal -Tenant $TenantId -Credential $creds | Out-Null

$subs = Get-AzSubscription | ? { $_.Name -match '(?i)ADH' -and $_.Name -match [regex]::Escape($adh_group) }
if(-not $subs){ throw "No ADH subscriptions for $adh_group" }

$expected = LoadSecretCsv $SecretCsvPath
$OutputDir = OutDir $OutputDir
$stamp=(Get-Date).ToString('yyyyMMdd_HHmmss')
$cmpCsv=Join-Path $OutputDir "kv-secrets_compare_${adh_group}_${stamp}.csv"
$invCsv=Join-Path $OutputDir "kv-secrets_inventory_${adh_group}_${stamp}.csv"
$json=Join-Path $OutputDir "kv-secrets_${adh_group}_${stamp}.json"
$html=Join-Path $OutputDir "kv-secrets_${adh_group}_${stamp}.html"

$cmpRows=@(); $invRows=@()

foreach($sub in $subs){
  Set-AzContext -Tenant $TenantId -SubscriptionId $sub.Id | Out-Null
  $vaults=Get-AzKeyVault -ErrorAction SilentlyContinue
  foreach($v in $vaults){
    $all=@(); try{ $all=Get-AzKeyVaultSecret -VaultName $v.VaultName -ErrorAction SilentlyContinue }catch{}
    foreach($s in $all){
      $invRows+= [pscustomobject]@{ SubscriptionName=$sub.Name; Vault=$v.VaultName; Secret=$s.Name; Enabled=$s.Enabled; Updated=$s.Updated }
    }
    foreach($sn in $expected){
      $exists=$false; $val=''; $err=''
      try{ $found=Get-AzKeyVaultSecret -VaultName $v.VaultName -Name $sn -ErrorAction Stop; $exists=$true; if($IncludeSecretValues){$val=$found.SecretValueText}}catch{$err=$_.Exception.Message}
      $cmpRows+= [pscustomobject]@{ SubscriptionName=$sub.Name; Vault=$v.VaultName; SecretName=$sn; Exists=$exists; Value=$val; Note=$err }
    }
  }
}

$cmpRows|Export-Csv $cmpCsv -NoTypeInformation
$invRows|Export-Csv $invCsv -NoTypeInformation
@{comparison=$cmpRows;inventory=$invRows}|ConvertTo-Json -Depth 6|Set-Content $json
($cmpRows|ConvertTo-Html -Title "KV Secrets $adh_group $stamp")|Set-Content $html

Write-Host "Secrets Compare: $cmpCsv"
Write-Host "Secrets Inventory: $invCsv"
