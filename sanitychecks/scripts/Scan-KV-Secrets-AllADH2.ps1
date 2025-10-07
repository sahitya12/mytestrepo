param(
  [Parameter(Mandatory=$true)][string]$TenantId,
  [Parameter(Mandatory=$true)][string]$ClientId,
  [Parameter(Mandatory=$true)][string]$ClientSecret,
  [Parameter(Mandatory=$true)][string]$SecretCsvPath,
  [string]$OutputDir="",
  [string]$BranchName=""
)
$ErrorActionPreference='Stop'
Import-Module Az.Accounts -ErrorAction Stop
Import-Module Az.KeyVault -ErrorAction Stop
Import-Module Az.Resources -ErrorAction Stop

function Ensure-Dir([string]$p){ if([string]::IsNullOrWhiteSpace($p)){ $p = Join-Path (Get-Location) 'kv-secrets-out' } if(-not(Test-Path $p)){ New-Item -ItemType Directory -Path $p -Force | Out-Null } return $p }

$sec=ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$cred=[pscredential]::new($ClientId,$sec)
Connect-AzAccount -ServicePrincipal -Tenant $TenantId -Credential $cred | Out-Null

$OutputDir=Ensure-Dir $OutputDir
$stamp=(Get-Date).ToString('yyyyMMdd_HHmmss')
$outCsv = Join-Path $OutputDir "kv_secrets_ALLADH_$stamp.csv"
$outHtml= Join-Path $OutputDir "kv_secrets_ALLADH_$stamp.html"
$outJson= Join-Path $OutputDir "kv_secrets_ALLADH_$stamp.json"

$exp = Import-Csv $SecretCsvPath | % { "$($_.SECRET_NAME)".Trim() } | ? { $_ -ne '' } | Select-Object -Unique
$subs = Get-AzSubscription | ? { $_.Name -match '(?i)ADH' }
$rows = New-Object System.Collections.Generic.List[object]

foreach($s in $subs){
  Set-AzContext -Tenant $TenantId -SubscriptionId $s.Id | Out-Null
  $vaults = Get-AzKeyVault -ErrorAction SilentlyContinue
  foreach($v in $vaults){
    foreach($name in $exp){
      $exists=$false
      try { $null = Get-AzKeyVaultSecret -VaultName $v.VaultName -Name $name -ErrorAction Stop; $exists=$true } catch {}
      $rows.Add([pscustomobject]@{
        SubscriptionName=$s.Name; Vault=$v.VaultName; ResourceGroup=$v.ResourceGroupName
        SecretName=$name; Exists=$(if($exists){'EXISTS'}else{'MISSING'})
      })
    }
  }
}

$rows | Export-Csv $outCsv -NoTypeInformation -Encoding UTF8
($rows | ConvertTo-Html -Title "KV Secrets ALLADH $stamp" -PreContent "<h2>KV Secrets ALLADH ($BranchName)</h2>") | Set-Content -Path $outHtml -Encoding UTF8
$rows | ConvertTo-Json -Depth 5 | Set-Content -Path $outJson -Encoding UTF8
Write-Host "CSV:  $outCsv"
Write-Host "HTML: $outHtml"
Write-Host "JSON: $outJson"
