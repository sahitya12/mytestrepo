<# 
.SYNOPSIS
  Scan RG-level RBAC for subscriptions whose names contain both "ADH" and a given custodian (adh_group).
  Compares against expected permissions from two CSVs (PROD vs NONPROD).

.PARAMETER TenantId
  Entra tenant ID (GUID)

.PARAMETER ClientId
  Service Principal (App) ID (GUID)

.PARAMETER ClientSecret
  Service Principal secret

.PARAMETER ProdCsvPath
  CSV with columns: resource_group_name,role_definition_name,ad_group_name (PRODUCTION expectations)

.PARAMETER NonProdCsvPath
  CSV with columns: resource_group_name,role_definition_name,ad_group_name (NONPRODUCTION expectations)

.PARAMETER adh_group
  Custodian code (e.g., CSM). Only subscriptions with names containing ADH and this text are scanned.

.PARAMETER OutputDir
  Output folder (CSV/HTML). Default: ./infra-sanity-out

.PARAMETER TeamsWebhookUrl
  Optional Incoming Webhook URL for Teams notifications. If omitted/empty, notification is skipped.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][string]$TenantId,
  [Parameter(Mandatory=$true)][string]$ClientId,
  [Parameter(Mandatory=$true)][string]$ClientSecret,
  [Parameter(Mandatory=$true)][string]$ProdCsvPath,
  [Parameter(Mandatory=$true)][string]$NonProdCsvPath,
  [Parameter(Mandatory=$true)][string]$adh_group,
  [string]$OutputDir = "",
  [string]$TeamsWebhookUrl = ""
)

$ErrorActionPreference = 'Stop'

Import-Module Az.Accounts -ErrorAction Stop
Import-Module Az.Resources -ErrorAction Stop

function EnvFromSub([string]$n) {
  if ($n -match '(?i)\b(prod|production|prd)\b') { 'PRODUCTION' } else { 'NONPRODUCTION' }
}
function CustFromSub([string]$n) {
  if ($n -match '(?i)ADH([A-Za-z0-9_-]+)') {
    $v = $Matches[1] -replace '^[^A-Za-z0-9]+',''
    if ($v -match '^([A-Za-z0-9]+)') { return $Matches[1] }
    return $v
  }
  $null
}
function HNorm([string]$s) { ($s -replace '[_\s]','').ToLowerInvariant() }
function LoadCsv($p) {
  if (-not (Test-Path $p)) { throw "CSV not found: $p" }
  $raw = Import-Csv $p
  if (-not $raw) { throw "CSV empty: $p" }
  $m = @{}
  foreach ($k in $raw[0].psobject.Properties.Name) { $m[(HNorm $k)] = $k }
  foreach ($r in 'resourcegroupname','roledefinitionname','adgroupname') {
    if (-not $m.ContainsKey($r)) { throw "CSV '$p' missing column like '$r'" }
  }
  $out = @()
  foreach ($x in $raw) {
    $out += [pscustomobject]@{
      RawRG   = "$($x.$($m['resourcegroupname']))".Trim()
      RawRole = "$($x.$($m['roledefinitionname']))".Trim()
      RawGrp  = "$($x.$($m['adgroupname']))".Trim()
    }
  }
  $out
}
function Resolve-Group([string]$n) {
  if ([string]::IsNullOrWhiteSpace($n)) { return $null }
  $g = Get-AzADGroup -DisplayName $n -ErrorAction SilentlyContinue
  if (-not $g) {
    $g = Get-AzADGroup -SearchString $n -ErrorAction SilentlyContinue | ? { $_.DisplayName -eq $n } | Select-Object -First 1
  }
  $g
}
function OutDir([string]$d) {
  if ([string]::IsNullOrWhiteSpace($d)) { $d = Join-Path (Get-Location) 'infra-sanity-out' }
  if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d | Out-Null }
  $d
}

# ---- Login (Service Principal) ----
$sec   = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$creds = New-Object System.Management.Automation.PSCredential($ClientId,$sec)
Connect-AzAccount -ServicePrincipal -Tenant $TenantId -Credential $creds | Out-Null

# ---- Subscriptions: must contain ADH + adh_group in Name ----
$subs = Get-AzSubscription | ? { $_.Name -match '(?i)ADH' -and $_.Name -match [regex]::Escape($adh_group) }
if (-not $subs) { throw "No subscriptions with 'ADH' and '$adh_group' in the name were found." }

# ---- Load inputs ----
$prod = LoadCsv $ProdCsvPath
$nonp = LoadCsv $NonProdCsvPath

# ---- Outputs ----
$OutputDir = OutDir $OutputDir
$stamp     = (Get-Date).ToString('yyyyMMdd_HHmmss')
$permCsv   = Join-Path $OutputDir "permissions_${adh_group}_${stamp}.csv"
$permHtml  = Join-Path $OutputDir "permissions_${adh_group}_${stamp}.html"
$rows      = New-Object System.Collections.Generic.List[object]

# ---- Scan ----
foreach ($sub in $subs) {
  Set-AzContext -Tenant $TenantId -SubscriptionId $sub.Id | Out-Null

  $env      = EnvFromSub $sub.Name
  $expected = if ($env -eq 'PRODUCTION') { $prod } else { $nonp }

  $subCust = CustFromSub $sub.Name
  if (-not $subCust) { $subCust = $adh_group }

  $rgList = Get-AzResourceGroup -ErrorAction SilentlyContinue
  $rgMap  = @{}
  foreach ($r in $rgList) {
    if ($r.ResourceGroupName) { $rgMap[$r.ResourceGroupName.ToLowerInvariant()] = $r }
  }

  foreach ($e in $expected) {
    $inRG   = $e.RawRG
    $inRole = $e.RawRole
    $inGrp  = $e.RawGrp

    # Replace <Custodian>
    $rg = $inRG   -replace '<Custodian>', $subCust
    $rl = $inRole -replace '<Custodian>', $subCust
    $gp = $inGrp  -replace '<Custodian>', $subCust

    # Find RG
    $rgObj = $null
    $key   = ''
    if ($rg) { $key = $rg.ToLowerInvariant() }
    if ($key -and $rgMap.ContainsKey($key)) { $rgObj = $rgMap[$key] }

    if (-not $rgObj) {
      $rows.Add([pscustomobject]@{
        SubscriptionName     = $sub.Name
        SubscriptionId       = $sub.Id
        Environment          = $env
        InputResourceGroup   = $inRG
        ScannedResourceGroup = $rg
        RoleDefinition       = $rl
        InputAdGroup         = $inGrp
        ResolvedAdGroup      = $gp
        GroupObjectId        = ''
        RGStatus             = 'NOT_FOUND'
        PermissionStatus     = 'N/A_RG_NOT_FOUND'
        Status               = 'RG_NOT_FOUND'
        Details              = 'RG not found'
      })
      continue
    }

    # Resolve group
    $g = Resolve-Group $gp
    if (-not $g) {
      $rows.Add([pscustomobject]@{
        SubscriptionName     = $sub.Name
        SubscriptionId       = $sub.Id
        Environment          = $env
        InputResourceGroup   = $inRG
        ScannedResourceGroup = $rg
        RoleDefinition       = $rl
        InputAdGroup         = $inGrp
        ResolvedAdGroup      = $gp
        GroupObjectId        = ''
        RGStatus             = 'EXISTS'
        PermissionStatus     = 'N/A_GROUP_NOT_FOUND'
        Status               = 'GROUP_NOT_FOUND'
        Details              = 'Group not found'
      })
      continue
    }

    # Check RG-scope role assignment
    $scope = "/subscriptions/$($sub.Id)/resourceGroups/$rg"
    $ra    = Get-AzRoleAssignment -Scope $scope -ObjectId $g.Id -RoleDefinitionName $rl -ErrorAction SilentlyContinue

    if ($ra) {
      $rows.Add([pscustomobject]@{
        SubscriptionName     = $sub.Name
        SubscriptionId       = $sub.Id
        Environment          = $env
        InputResourceGroup   = $inRG
        ScannedResourceGroup = $rg
        RoleDefinition       = $rl
        InputAdGroup         = $inGrp
        ResolvedAdGroup      = $gp
        GroupObjectId        = $g.Id
        RGStatus             = 'EXISTS'
        PermissionStatus     = 'EXISTS'
        Status               = 'EXISTS'
        Details              = ''
      })
    }
    else {
      $rows.Add([pscustomobject]@{
        SubscriptionName     = $sub.Name
        SubscriptionId       = $sub.Id
        Environment          = $env
        InputResourceGroup   = $inRG
        ScannedResourceGroup = $rg
        RoleDefinition       = $rl
        InputAdGroup         = $inGrp
        ResolvedAdGroup      = $gp
        GroupObjectId        = $g.Id
        RGStatus             = 'EXISTS'
        PermissionStatus     = 'MISSING'
        Status               = 'MISSING'
        Details              = 'Role assignment missing at RG'
      })
    }
  }
}

# ---- Export ----
$rows | Export-Csv $permCsv -NoTypeInformation -Encoding UTF8
($rows | ConvertTo-Html -Title "RG Permissions - $adh_group - $stamp" -PreContent "<h2>RG Permissions - $adh_group - $stamp</h2>") | Set-Content -Path $permHtml -Encoding UTF8
Write-Host "Permissions CSV : $permCsv"
Write-Host "Permissions HTML: $permHtml"

# ---- Teams (optional) ----
if ($TeamsWebhookUrl) {
  $ok   = ($rows | ? { $_.PermissionStatus -eq 'EXISTS' }).Count
  $miss = ($rows | ? { $_.PermissionStatus -eq 'MISSING' }).Count
  $rgna = ($rows | ? { $_.PermissionStatus -eq 'N/A_RG_NOT_FOUND' }).Count
  $gpna = ($rows | ? { $_.PermissionStatus -eq 'N/A_GROUP_NOT_FOUND' }).Count
  $sum  = "RG Permissions ($adh_group) $stamp`nTotal:$($rows.Count) Exists:$ok Missing:$miss RG_NA:$rgna Group_NA:$gpna`n$permCsv`n$permHtml"
  try { Invoke-RestMethod -Method Post -Uri $TeamsWebhookUrl -ContentType 'application/json' -Body (@{ text = $sum } | ConvertTo-Json) }
  catch { Write-Warning "Teams post failed: $_" }
}
