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
Import-Module Az.Resources -ErrorAction Stop
Import-Module Az.KeyVault -ErrorAction Stop

function Ensure-Dir([string]$p){ if([string]::IsNullOrWhiteSpace($p)){ $p = Join-Path (Get-Location) 'kv-networks-out' } if(-not(Test-Path $p)){ New-Item -ItemType Directory -Path $p -Force | Out-Null } return $p }

$sec=ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$cred=[pscredential]::new($ClientId,$sec)
Connect-AzAccount -ServicePrincipal -Tenant $TenantId -Credential $cred | Out-Null

$OutputDir=Ensure-Dir $OutputDir
$stamp=(Get-Date).ToString('yyyyMMdd_HHmmss')
$outCsv=Join-Path $OutputDir "kv_networks_$stamp.csv"

$subs = if($ScanAll){ Get-AzSubscription | ? { $_.Name -match 'ADH' } } else { Get-AzSubscription | ? { $_.Name -match 'ADH' -and $_.Name -match [regex]::Escape($adh_group) } }
$rows = New-Object System.Collections.Generic.List[object]

foreach($sub in $subs){
  Set-AzContext -Tenant $TenantId -SubscriptionId $sub.Id | Out-Null
  $vaults=Get-AzKeyVault -ErrorAction SilentlyContinue
  foreach($v in $vaults){
    $rows.Add([pscustomobject]@{
      SubscriptionName=$sub.Name; SubscriptionId=$sub.Id
      Vault=$v.VaultName; ResourceGroup=$v.ResourceGroupName
      PublicNetworkAccess=$v.PublicNetworkAccess
      DefaultAction=$v.NetworkAcls.DefaultAction
      IpRules=($v.NetworkAcls.IpRules.IpAddressRange -join ';')
      VnetRules=($v.NetworkAcls.VirtualNetworkRules.Id -join ';')
    })
  }
}

$rows | Export-Csv $outCsv -NoTypeInformation -Encoding UTF8
Write-Host "CSV: $outCsv"
