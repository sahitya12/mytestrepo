param(
  [Parameter(Mandatory=$true)][string]$TenantId,
  [Parameter(Mandatory=$true)][string]$ClientId,
  [Parameter(Mandatory=$true)][string]$ClientSecret,
  [Parameter(Mandatory=$true)][string]$SecretCsvPath,
  [string]$OutputDir = ""
)

$ErrorActionPreference='Stop'
Import-Module Az.Accounts; Import-Module Az.Resources; Import-Module Az.KeyVault

$sec=ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$creds=[pscredential]::new($ClientId,$sec)
Connect-AzAccount -ServicePrincipal -Tenant $TenantId -Credential $creds|Out-Null

$subs=Get-AzSubscription|?{$_.Name -match 'ADH'}
$expected=(Import-Csv $SecretCsvPath|%{ $_.SECRET_NAME })

$OutputDir=if($OutputDir){$OutputDir}else{Join-Path (Get-Location) 'kv-secrets-out'}
$stamp=(Get-Date).ToString('yyyyMMdd_HHmmss')
$cmpCsv=Join-Path $OutputDir "kv-secrets_compare_ALLADH_${stamp}.csv"
$cmpRows=@()

foreach($sub in $subs){
  Set-AzContext -Tenant $TenantId -SubscriptionId $sub.Id|Out-Null
  $vaults=Get-AzKeyVault
  foreach($v in $vaults){
    foreach($sn in $expected){
      $exists=$false
      try{Get-AzKeyVaultSecret -VaultName $v.VaultName -Name $sn -ErrorAction Stop; $exists=$true}catch{}
      $cmpRows+=[pscustomobject]@{Subscription=$sub.Name;Vault=$v.VaultName;SecretName=$sn;Exists=$exists}
    }
  }
}
$cmpRows|Export-Csv $cmpCsv -NoTypeInformation
Write-Host "Secrets Compare: $cmpCsv"
