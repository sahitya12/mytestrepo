param(
  [Parameter(Mandatory=$true)][string]$TenantId,
  [Parameter(Mandatory=$true)][string]$ClientId,
  [Parameter(Mandatory=$true)][string]$ClientSecret,
  [string]$adh_group = "",
  [ValidateSet('prd','nonprd')][string]$adh_subscription_type = 'nonprd',
  [Parameter(Mandatory=$true)][string]$AdlsCsvPath,
  [switch]$ScanAll,
  [string]$OutputDir="",
  [string]$BranchName=""
)
$ErrorActionPreference='Stop'
Import-Module Az.Accounts -ErrorAction Stop
Import-Module Az.Storage  -ErrorAction Stop
Import-Module Az.Resources -ErrorAction Stop
Import-Module Az.Resources -ErrorAction Stop

function Ensure-Dir([string]$p){ if([string]::IsNullOrWhiteSpace($p)){ $p = Join-Path (Get-Location) 'adls-out' } if(-not(Test-Path $p)){ New-Item -ItemType Directory -Path $p -Force | Out-Null } return $p }
function Normalize([string]$s){ ($s -replace '[_\s]','').ToLowerInvariant() }
function Map-CustShort([string]$cust){ 
  if([string]::IsNullOrWhiteSpace($cust)){ return '' }
  # change this rule to match your naming – here we use first 3 letters lowercased
  return ($cust.Substring(0, [Math]::Min(3,$cust.Length))).ToLower()
}
function Safe($x){ if($x){$x}else{''} }

# CSV columns required:
# ResourceGroupName,StorageAccountName,ContainerName,Identity,AccessPath,PermissionType,Type,Scope
if(-not (Test-Path $AdlsCsvPath)){ throw "CSV not found: $AdlsCsvPath" }
$raw = Import-Csv $AdlsCsvPath
if(-not $raw){ throw "CSV empty: $AdlsCsvPath" }

$sec=ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$cred=[pscredential]::new($ClientId,$sec)
Connect-AzAccount -ServicePrincipal -Tenant $TenantId -Credential $cred | Out-Null

$OutputDir=Ensure-Dir $OutputDir
$stamp=(Get-Date).ToString('yyyyMMdd_HHmmss')
$outCsv=Join-Path $OutputDir "adls_validate_$stamp.csv"
$outHtml=Join-Path $OutputDir "adls_validate_$stamp.html"
$outJson=Join-Path $OutputDir "adls_validate_$stamp.json"
$rows=New-Object System.Collections.Generic.List[object]

# choose subscriptions
$subs = if($ScanAll){ Get-AzSubscription | ? { $_.Name -match 'ADH' } } else { Get-AzSubscription | ? { $_.Name -match 'ADH' -and $_.Name -match [regex]::Escape($adh_group) } }

foreach($sub in $subs){
  Set-AzContext -Tenant $TenantId -SubscriptionId $sub.Id | Out-Null

  foreach($r in $raw){
    # Expand tokens in expected values
    $rgExp = "$($r.ResourceGroupName)".Replace('<Custodian>', $adh_group)
    $custShort = Map-CustShort $adh_group
    $saNameExp = "$($r.StorageAccountName)".Replace('<Cust>', $custShort)
    $saNameExp = $saNameExp.Replace('nonprd', $(if($adh_subscription_type -eq 'nonprd'){'nonprd'}else{'prd'}))
    $containerExp = "$($r.ContainerName)"
    if($adh_subscription_type -eq 'prd'){ $containerExp = 'prd' }

    # Lookup actuals
    $sa = Get-AzStorageAccount -ResourceGroupName $rgExp -Name $saNameExp -ErrorAction SilentlyContinue
    if(-not $sa){
      $rows.Add([pscustomobject]@{
        SubscriptionName=$sub.Name; ResourceGroup=$rgExp; StorageAccount=$saNameExp; Container=$containerExp
        AccessPath=$r.AccessPath; Identity=$r.Identity; Type=$r.Type; Scope=$r.Scope; PermissionType=$r.PermissionType
        Exists='SA_NOT_FOUND'; RBAC='N/A'; ACL='N/A'
      })
      continue
    }

    # RBAC (role assignments at account scope)
    $scope = $sa.Id
    $rbac = Get-AzRoleAssignment -Scope $scope -ErrorAction SilentlyContinue
    $rbacSummary = ($rbac | % { "$($_.DisplayName):$($_.RoleDefinitionName)" }) -join '; '

    # Storage context with OAuth (needs Storage Blob Data Reader at least)
    $ctx = $null
    try {
      $ctx = New-AzStorageContext -StorageAccountName $sa.StorageAccountName -UseConnectedAccount -ErrorAction Stop
    } catch {
      $ctx = $null
    }

    # Container check
    $containerExists='UNKNOWN'
    if($ctx){
      try {
        $cont = Get-AzStorageContainer -Name $containerExp -Context $ctx -ErrorAction Stop
        $containerExists = if($cont){'YES'}else{'NO'}
      } catch { $containerExists='NO' }
    } else { $containerExists='CTX_ERR' }

    # ACL check (best-effort): read ACL at AccessPath if context allowed
    $aclResult='UNABLE'
    if($ctx -and $containerExists -eq 'YES'){
      try {
        # ListPath returns basic entries; Az cmdlets expose ACL via Get-AzDataLakeGen2Item with -GetAccessControl
        $item = Get-AzDataLakeGen2Item -Context $ctx -FileSystem $containerExp -Path $r.AccessPath -ErrorAction Stop
        $aclText = ($item.ACL -join ',')
        $aclResult = if([string]::IsNullOrWhiteSpace($aclText)){'EMPTY_OR_NOACL'}else{$aclText}
      } catch {
        $aclResult = "READ_ERR"
      }
    }

    $rows.Add([pscustomobject]@{
      SubscriptionName=$sub.Name; ResourceGroup=$rgExp; StorageAccount=$saNameExp; Container=$containerExp
      AccessPath=$r.AccessPath; Identity=$r.Identity; Type=$r.Type; Scope=$r.Scope; PermissionType=$r.PermissionType
      Exists='OK'; RBAC=$rbacSummary; ContainerExists=$containerExists; ACL=$aclResult
    })
  }
}

$rows | Export-Csv $outCsv -NoTypeInformation -Encoding UTF8
($rows | ConvertTo-Html -Title "ADLS Validate $stamp" -PreContent "<h2>ADLS Validate ($BranchName)</h2>") | Set-Content -Path $outHtml -Encoding UTF8
$rows | ConvertTo-Json -Depth 6 | Set-Content -Path $outJson -Encoding UTF8
Write-Host "CSV:  $outCsv"
Write-Host "HTML: $outHtml"
Write-Host "JSON: $outJson"
