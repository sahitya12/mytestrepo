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

function Ensure-Dir([string]$p){ if([string]::IsNullOrWhiteSpace($p)){ $p = Join-Path (Get-Location) 'kv-perms-out' } if(-not(Test-Path $p)){ New-Item -ItemType Directory -Path $p -Force | Out-Null } return $p }

$sec=ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$cred=[pscredential]::new($ClientId,$sec)
Connect-AzAccount -ServicePrincipal -Tenant $TenantId -Credential $cred | Out-Null

$OutputDir=Ensure-Dir $OutputDir
$stamp=(Get-Date).ToString('yyyyMMdd_HHmmss')
$rbacCsv=Join-Path $OutputDir "kv_rbac_$stamp.csv"
$polCsv =Join-Path $OutputDir "kv_accesspolicies_$stamp.csv"

$subs = if($ScanAll){ Get-AzSubscription | ? { $_.Name -match 'ADH' } } else { Get-AzSubscription | ? { $_.Name -match 'ADH' -and $_.Name -match [regex]::Escape($adh_group) } }

$rbac=New-Object System.Collections.Generic.List[object]
$pol =New-Object System.Collections.Generic.List[object]

foreach($sub in $subs){
  Set-AzContext -Tenant $TenantId -SubscriptionId $sub.Id | Out-Null
  $vaults=Get-AzKeyVault -ErrorAction SilentlyContinue
  foreach($v in $vaults){
    $scope="/subscriptions/$($sub.Id)/resourceGroups/$($v.ResourceGroupName)/providers/Microsoft.KeyVault/vaults/$($v.VaultName)"
    $ra=Get-AzRoleAssignment -Scope $scope -ErrorAction SilentlyContinue
    foreach($a in $ra){
      $rbac.Add([pscustomobject]@{
        SubscriptionName=$sub.Name; SubscriptionId=$sub.Id; Vault=$v.VaultName; ResourceGroup=$v.ResourceGroupName
        Principal=$a.DisplayName; PrincipalId=$a.ObjectId; Role=$a.RoleDefinitionName; Scope=$a.Scope
      })
    }
    foreach($p in $v.AccessPolicies){
      $pol.Add([pscustomobject]@{
        SubscriptionName=$sub.Name; SubscriptionId=$sub.Id; Vault=$v.VaultName; ResourceGroup=$v.ResourceGroupName
        ObjectId=$p.ObjectId
        Secrets=($p.PermissionsToSecrets -join ',')
        Keys=($p.PermissionsToKeys -join ',')
        Certificates=($p.PermissionsToCertificates -join ',')
        Storage=($p.PermissionsToStorage -join ',')
      })
    }
  }
}

$rbac | Export-Csv $rbacCsv -NoTypeInformation -Encoding UTF8
$pol  | Export-Csv $polCsv  -NoTypeInformation -Encoding UTF8
Write-Host "RBAC CSV: $rbacCsv"
Write-Host "AccessPolicies CSV: $polCsv"
