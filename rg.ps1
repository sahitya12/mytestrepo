<#
Scans subscriptions (scoped to given tenant) whose names end with: _ADH<adh_group>
Example: if -adh_group "CSM" → matches "..._ADHCSM"

Excel worksheet layout (single sheet, e.g. rg_permissions):
  PRODUCTION block (A–C):
    A: resource_group_name
    B: role_definition_name
    C: ad_group_name      (may contain "<Custodian>")
  NONPRODUCTION block (F–H):
    F: resource_group_name
    G: role_definition_name
    H: ad_group_name      (may contain "<Custodian>")

Status values:
  EXISTS, MISSING, RG_NOT_FOUND, GROUP_NOT_FOUND
#>

param(
  [Parameter(Mandatory=$true)]
  [string]$tenant_id,                       # Tenant to scope to

  [Parameter(Mandatory=$true)]
  [string]$FilePath,                        # e.g. C:\temp\rg-permissions.xlsx

  [Parameter(Mandatory=$true)]
  [string]$WorksheetName,                   # e.g. "rg_permissions"

  [Parameter(Mandatory=$true)]
  [string]$adh_group,                       # e.g. "CSM"

  [switch]$UseDeviceLogin                   # optional device code login
)

# ---------------- Safety & modules ----------------
$ErrorActionPreference = 'Stop'

function Assert-Module {
  param([string]$Name)
  try { Import-Module $Name -ErrorAction Stop }
  catch { throw "PowerShell module '$Name' is missing. Install it first: Install-Module $Name -Scope CurrentUser" }
}

Assert-Module -Name Az.Accounts
Assert-Module -Name Az.Resources
Assert-Module -Name ImportExcel

# ---------------- Login & scope to tenant ----------------
if ($UseDeviceLogin) {
  Connect-AzAccount -Tenant $tenant_id -UseDeviceAuthentication | Out-Null
} else {
  Connect-AzAccount -Tenant $tenant_id | Out-Null
}

# ---------------- Helpers ----------------
function Get-EnvFromSubscriptionName {
  param([string]$Name)
  if ($Name -match '(?i)\b(prod|production)\b') { 'PRODUCTION' } else { 'NONPRODUCTION' }
}

function Ensure-Worksheet {
  param([string]$Path, [string]$Sheet)
  $sheets = Get-ExcelSheetInfo -Path $Path
  if (-not ($sheets | Where-Object { $_.Name -eq $Sheet })) {
    $all = ($sheets | Select-Object -ExpandProperty Name) -join ', '
    throw "Worksheet '$Sheet' not found in '$Path'. Sheets available: [$all]"
  }
}

function Read-RGMatrixFromExcel {
  param([string]$Path,[string]$Sheet)
  Ensure-Worksheet -Path $Path -Sheet $Sheet
  $rows = Import-Excel -Path $Path -WorksheetName $Sheet -NoHeader
  if (-not $rows -or $rows.Count -eq 0) { throw "No data read from worksheet '$Sheet'." }

  $prodStart = $null; $nonProdStart = $null
  for ($i=0; $i -lt $rows.Count; $i++) {
    $a = "$($rows[$i].P1)".Trim().ToLower()
    $f = "$($rows[$i].P6)".Trim().ToLower()
    if ($null -eq $prodStart -and $a -eq 'resource_group_name')    { $prodStart = $i }
    if ($null -eq $nonProdStart -and $f -eq 'resource_group_name') { $nonProdStart = $i }
    if ($prodStart -ne $null -and $nonProdStart -ne $null) { break }
  }
  if ($null -eq $prodStart -and $null -eq $nonProdStart) {
    throw "Header row not found. Expect 'resource_group_name' in column A and/or F."
  }

  $prod = @()
  if ($null -ne $prodStart) {
    for ($r=$prodStart+1; $r -lt $rows.Count; $r++) {
      $rg   = "$($rows[$r].P1)".Trim(); $role = "$($rows[$r].P2)".Trim(); $grp = "$($rows[$r].P3)".Trim()
      if ([string]::IsNullOrWhiteSpace($rg) -and [string]::IsNullOrWhiteSpace($role) -and [string]::IsNullOrWhiteSpace($grp)) { break }
      if ($rg) { $prod += [pscustomobject]@{ Environment='PRODUCTION'; ResourceGroup=$rg; Role=$role; Group=$grp } }
    }
  }

  $nonprod = @()
  if ($null -ne $nonProdStart) {
    for ($r=$nonProdStart+1; $r -lt $rows.Count; $r++) {
      $rg   = "$($rows[$r].P6)".Trim(); $role = "$($rows[$r].P7)".Trim(); $grp = "$($rows[$r].P8)".Trim()
      if ([string]::IsNullOrWhiteSpace($rg) -and [string]::IsNullOrWhiteSpace($role) -and [string]::IsNullOrWhiteSpace($grp)) { break }
      if ($rg) { $nonprod += [pscustomobject]@{ Environment='NONPRODUCTION'; ResourceGroup=$rg; Role=$role; Group=$grp } }
    }
  }

  [pscustomobject]@{ PRODUCTION=$prod; NONPRODUCTION=$nonprod }
}

function Resolve-Group {
  param([string]$DisplayName)
  $g = Get-AzADGroup -DisplayName $DisplayName -ErrorAction SilentlyContinue
  if (-not $g) {
    $g = Get-AzADGroup -SearchString $DisplayName -ErrorAction SilentlyContinue |
         Where-Object { $_.DisplayName -eq $DisplayName } | Select-Object -First 1
  }
  $g
}

# ---------------- Load requirements ----------------
$matrix = Read-RGMatrixFromExcel -Path $FilePath -Sheet $WorksheetName
$prodReqs    = $matrix.PRODUCTION
$nonprodReqs = $matrix.NONPRODUCTION
if (($prodReqs.Count + $nonprodReqs.Count) -eq 0) {
  throw "No requirements parsed. Check A–C (PRODUCTION) and F–H (NONPRODUCTION)."
}

# ---------------- Filter subscriptions: names ending with _ADH<adh_group> ----------------
$pattern = "(?i)_ADH" + [regex]::Escape($adh_group) + "$"
$subs = Get-AzSubscription | Where-Object { $_.Name -match $pattern }

if (-not $subs) {
  throw "No subscriptions found in tenant '$tenant_id' with names ending in '_ADH$adh_group'."
}

# ---------------- Scan ----------------
$result = New-Object System.Collections.Generic.List[object]

foreach ($sub in $subs) {
  # Keep tenant scoped
  Set-AzContext -Tenant $tenant_id -SubscriptionId $sub.Id | Out-Null

  # Determine env from subscription name
  $env  = Get-EnvFromSubscriptionName -Name $sub.Name
  $reqs = if ($env -eq 'PRODUCTION') { $prodReqs } else { $nonprodReqs }

  # Enumerate ALL RGs in this subscription (as requested)
  $rgList = Get-AzResourceGroup -ErrorAction SilentlyContinue
  $rgMap  = @{}  # name(lower) -> object
  foreach ($r in $rgList) { $rgMap[$r.ResourceGroupName.ToLowerInvariant()] = $r }

  foreach ($req in $reqs) {
    $rgName    = $req.ResourceGroup
    $roleName  = $req.Role
    $grpTempl  = $req.Group

    $adGroupName = if ($grpTempl) { $grpTempl -replace '<Custodian>', $adh_group } else { '' }

    # RG existence from enumerated map
    $rgKey = ($rgName ?? '').ToLowerInvariant()
    $rg    = $null
    if ($rgKey -and $rgMap.ContainsKey($rgKey)) { $rg = $rgMap[$rgKey] }

    if (-not $rg) {
      $result.Add([pscustomobject]@{
        SubscriptionName = $sub.Name
        SubscriptionId   = $sub.Id
        Environment      = $env
        ResourceGroup    = $rgName
        RoleDefinition   = $roleName
        AdGroupName      = $adGroupName
        GroupObjectId    = ''
        Status           = 'RG_NOT_FOUND'
        Details          = 'Resource group not found in this subscription'
      })
      continue
    }

    # Resolve Entra group
    $group = $null
    if (-not [string]::IsNullOrWhiteSpace($adGroupName)) {
      $group = Resolve-Group -DisplayName $adGroupName
    }
    if (-not $group) {
      $result.Add([pscustomobject]@{
        SubscriptionName = $sub.Name
        SubscriptionId   = $sub.Id
        Environment      = $env
        ResourceGroup    = $rgName
        RoleDefinition   = $roleName
        AdGroupName      = $adGroupName
        GroupObjectId    = ''
        Status           = 'GROUP_NOT_FOUND'
        Details          = 'Entra ID group not found'
      })
      continue
    }

    $scope = "/subscriptions/$($sub.Id)/resourceGroups/$rgName"

    # Check RG-scope role assignment
    $ra = Get-AzRoleAssignment -Scope $scope -ObjectId $group.Id -RoleDefinitionName $roleName -ErrorAction SilentlyContinue

    if ($ra) {
      $result.Add([pscustomobject]@{
        SubscriptionName = $sub.Name
        SubscriptionId   = $sub.Id
        Environment      = $env
        ResourceGroup    = $rgName
        RoleDefinition   = $roleName
        AdGroupName      = $adGroupName
        GroupObjectId    = $group.Id
        Status           = 'EXISTS'
        Details          = ''
      })
    } else {
      $result.Add([pscustomobject]@{
        SubscriptionName = $sub.Name
        SubscriptionId   = $sub.Id
        Environment      = $env
        ResourceGroup    = $rgName
        RoleDefinition   = $roleName
        AdGroupName      = $adGroupName
        GroupObjectId    = $group.Id
        Status           = 'MISSING'
        Details          = 'Role assignment not found at RG scope'
      })
    }
  }
}

# ---------------- Output ----------------
$stamp  = (Get-Date).ToString('yyyyMMdd_HHmmss')
$outDir = Split-Path -Path $FilePath -Parent
if ([string]::IsNullOrWhiteSpace($outDir)) { $outDir = (Get-Location).Path }
$outFile = Join-Path $outDir "rg-permissions-scan_${stamp}.csv"

$result | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8

# Console summary
$tot = $result.Count
$ok  = ($result | Where-Object { $_.Status -eq 'EXISTS' }).Count
$miss= ($result | Where-Object { $_.Status -eq 'MISSING' }).Count
$rgna= ($result | Where-Object { $_.Status -eq 'RG_NOT_FOUND' }).Count
$gna = ($result | Where-Object { $_.Status -eq 'GROUP_NOT_FOUND' }).Count

Write-Host ""
Write-Host "Scan complete for tenant $tenant_id:"
Write-Host "  Total checks:     $tot"
Write-Host "  Present:          $ok"
Write-Host "  Missing:          $miss"
Write-Host "  RG not found:     $rgna"
Write-Host "  Group not found:  $gna"
Write-Host "Report: $outFile"
