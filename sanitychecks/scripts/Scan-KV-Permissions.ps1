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

$OutputDir=if($OutputDir){$OutputDir}else{Join-Path (Get-Location) 'kv-perms-out'}
$stamp=(Get-Date).ToString('yyyyMMdd_HHmmss')
$rbacCsv=Join-Path $OutputDir "kv_rbac_${stamp}.csv"
$polCsv=Join-Path $OutputDir "kv_accesspolicies_${stamp}.csv"

$rbac=@(); $pol=@()
foreach($sub in $subs){
  Set-AzContext -Tenant $TenantId -SubscriptionId $sub.Id|Out-Null
  $vaults=Get-AzKeyVault
  foreach($v in $vaults){
    $scope="/subscriptions/$($sub.Id)/resourceGroups/$($v.ResourceGroupName)/providers/Microsoft.KeyVault/vaults/$($v.VaultName)"
    $ra=Get-AzRoleAssignment -Scope $scope -ErrorAction SilentlyContinue
    foreach($a in $ra){ $rbac+=[pscustomobject]@{Sub=$sub.Name;Vault=$v.VaultName;Principal=$a.DisplayName;Role=$a.RoleDefinitionName;Scope=$a.Scope} }
    foreach($p in $v.AccessPolicies){ $pol+=[pscustomobject]@{Sub=$sub.Name;Vault=$v.VaultName;ObjectId=$p.ObjectId;SecretsPerms=($p.PermissionsToSecrets -join ',')} }
  }
}
$rbac|Export-Csv $rbacCsv -NoTypeInformation
$pol|Export-Csv $polCsv -NoTypeInformation
Write-Host "RBAC: $rbacCsv"
Write-Host "Policies: $polCsv"
