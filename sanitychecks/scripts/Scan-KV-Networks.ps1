param(
  [Parameter(Mandatory=$true)][string]$TenantId,
  [Parameter(Mandatory=$true)][string]$ClientId,
  [Parameter(Mandatory=$true)][string]$ClientSecret,
  [string]$adh_group="",
  [switch]$ScanAll,
  [string]$OutputDir=""
)

Import-Module Az.Accounts; Import-Module Az.Resources; Import-Module Az.KeyVault
$sec=ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$creds=[pscredential]::new($ClientId,$sec)
Connect-AzAccount -ServicePrincipal -Tenant $TenantId -Credential $creds|Out-Null

if($ScanAll){$subs=Get-AzSubscription|?{$_.Name -match 'ADH'}}else{$subs=Get-AzSubscription|?{$_.Name -match "ADH" -and $_.Name -match $adh_group}}

$OutputDir=if($OutputDir){$OutputDir}else{Join-Path (Get-Location) 'kv-networks-out'}
$stamp=(Get-Date).ToString('yyyyMMdd_HHmmss')
$csv=Join-Path $OutputDir "kv-networks_${stamp}.csv"

$rows=@()
foreach($sub in $subs){
  Set-AzContext -Tenant $TenantId -SubscriptionId $sub.Id|Out-Null
  $vaults=Get-AzKeyVault
  foreach($v in $vaults){
    $rows+=[pscustomobject]@{
      Sub=$sub.Name;Vault=$v.VaultName;RG=$v.ResourceGroupName
      PublicNetworkAccess=$v.PublicNetworkAccess
      DefaultAction=$v.NetworkAcls.DefaultAction
      IpRules=($v.NetworkAcls.IpRules.IpAddressRange -join ';')
      VnetRules=($v.NetworkAcls.VirtualNetworkRules.Id -join ';')
    }
  }
}
$rows|Export-Csv $csv -NoTypeInformation
Write-Host "Networks: $csv"
