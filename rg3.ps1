<# ====================== CONFIGURE THESE VALUES ====================== #>

# Excel file path (.xlsx/.xlsm), with worksheet "rg_permissions"
$FilePath  = "C:\path\to\resource_sanitychecks.xlsx"

# Custodian suffix used in subscription names (e.g., CSM -> matches *_ADHCSM)
$adh_group = "CSM"

# Output directory (leave blank to create 'kv-scan-local' next to the Excel file)
$OutputDir = ""

# Include raw secret values in outputs? ($true = include, $false = mask)
# (Not used in this RBAC scan, kept for consistency)
$IncludeSecretValues = $true

# --- Service Principal (SPN) credentials for Azure login ---
$TenantId     = "<tenant-guid>"
$ClientId     = "<appId-guid>"
$ClientSecret = "<sp-secret>"

# Worksheet name for the RG permissions matrix
$worksheetname = "rg_permissions"

<# ========================= DO NOT EDIT BELOW ======================== #>

$ErrorActionPreference = 'Stop'

function Assert-Module {
  param([string]$Name)
  try { Import-Module $Name -ErrorAction Stop }
  catch { throw "PowerShell module '$Name' is missing. Install it first: Install-Module $Name -Scope CurrentUser" }
}

Assert-Module -Name Az.Accounts
Assert-Module -Name Az.Resources
Assert-Module -Name ImportExcel

# -------- Login (Service Principal) --------
try {
  $sec   = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
  $creds = New-Object System.Management.Automation.PSCredential($ClientId, $sec)
  Connect-AzAccount -ServicePrincipal -Tenant $TenantId -Credential $creds | Out-Null
} catch {
  throw "Failed SPN login. Verify TenantId/AppId/Secret. $_"
}

# -------- Helpers --------
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

  # Expect two side-by-side blocks:
  #   PROD: A–C  (A:resource_group_name, B:role_definition_name, C:ad_group_name)
  #   NONP: F–H  (F:resource_group_name, G:role_definition_name, H:ad_group_name)
  $rows = Import-Excel -Path $Path -WorksheetName $Sheet -NoHeader
  if (-not $rows -or $rows.Count -eq 0) { throw "No data read from worksheet '$Sheet'." }

  $prodStart    = $null
  $nonProdStart = $null

  for ($i=0; $i -lt $rows.Count; $i++) {
    $a = "$($rows[$i].P1)".Trim().ToLower()
    $f = "$($rows[$i].P6)".Trim().ToLower()
    if ($null -eq $prodStart    -and $a -eq 'resource_group_name') { $prodStart    = $i }
    if ($null -eq $nonProdStart -and $f -eq 'resource_group_name') { $nonProdStart = $i }
    if ($prodStart -ne $null -and $nonProdStart -ne $null) { break }
  }

  if ($null -eq $prodStart -and $null -eq $nonProdStart) {
    throw "Header row not found. Expect 'resource_group_name' in column A and/or F of '$Sheet'."
  }

  $prod = @()
  if ($null -ne $prodStart) {
    for ($r=$prodStart+1; $r -lt $rows.Count; $r++) {
      $rg   = "$($rows[$r].P1)".Trim()
      $role = "$($rows[$r].P2)".Trim()
      $grp  = "$($rows[$r].P3)".Trim()
      if ([string]::IsNullOrWhiteSpace($rg) -and [string]::IsNullOrWhiteSpace($role) -and [string]::IsNullOrWhiteSpace($grp)) { break }
      if ($rg) { $prod += [pscustomobject]@{ Environment='PRODUCTION'; ResourceGroup=$rg; Role=$role; Group=$grp } }
    }
  }

  $nonprod = @()
  if ($null -ne $nonProdStart) {
    for ($r=$nonProdStart+1; $r -lt $rows.Count; $r++) {
      $rg   = "$($rows[$r].P6)".Trim()
      $role = "$($rows[$r].P7)".Trim()
      $grp  = "$($rows[$r].P8)".Trim()
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

# -------- Load requirements --------
$matrix = Read-RGMatrixFromExcel -Path $FilePath -Sheet $worksheetname

# ---- Keep RAW values, then replace <Custodian> in ALL fields (RG, Role, Group) ----
$prodReqs = $matrix.PRODUCTION | ForEach-Object {
  $_ | Add-Member -NotePropertyName RawGroup -NotePropertyValue $_.Group -Force
  $_ | Add-Member -NotePropertyName RawRG    -NotePropertyValue $_.ResourceGroup -Force
  $_.ResourceGroup = ($_.ResourceGroup -replace '<Custodian>', $adh_group)
  $_.Role          = ($_.Role          -replace '<Custodian>', $adh_group)
  $_.Group         = ($_.Group         -replace '<Custodian>', $adh_group)
  $_
}

$nonprodReqs = $matrix.NONPRODUCTION | ForEach-Object {
  $_ | Add-Member -NotePropertyName RawGroup -NotePropertyValue $_.Group -Force
  $_ | Add-Member -NotePropertyName RawRG    -NotePropertyValue $_.ResourceGroup -Force
  $_.ResourceGroup = ($_.ResourceGroup -replace '<Custodian>', $adh_group)
  $_.Role          = ($_.Role          -replace '<Custodian>', $adh_group)
  $_.Group         = ($_.Group         -replace '<Custodian>', $adh_group)
  $_
}

if (($prodReqs.Count + $nonprodReqs.Count) -eq 0) {
  throw "No requirements parsed after placeholder replacement. Check sheet '$worksheetname'."
}

# -------- Filter subscriptions: names ending with _ADH<adh_group> (case-insensitive) --------
$pattern = "(?i)_ADH" + [regex]::Escape($adh_group) + "$"
$subs = Get-AzSubscription | Where-Object { $_.Name -match $pattern }

if (-not $subs) {
  throw "No subscriptions found in tenant '$TenantId' with names ending in '_ADH$adh_group'."
}

# -------- Scan --------
$result = New-Object System.Collections.Generic.List[object]

foreach ($sub in $subs) {
  Set-AzContext -Tenant $TenantId -SubscriptionId $sub.Id | Out-Null

  $env  = Get-EnvFromSubscriptionName -Name $sub.Name
  $reqs = if ($env -eq 'PRODUCTION') { $prodReqs } else { $nonprodReqs }

  # Enumerate all RGs in the subscription
  $rgList = Get-AzResourceGroup -ErrorAction SilentlyContinue
  $rgMap  = @{}
  foreach ($r in $rgList) {
    if ($r.ResourceGroupName) {
      $rgMap[$r.ResourceGroupName.ToString().ToLowerInvariant()] = $r
    }
  }

  foreach ($req in $reqs) {
    $rgName             = $req.ResourceGroup        # resolved RG
    $roleName           = $req.Role
    $adGroupName        = $req.Group                # resolved group
    $inputResourceRaw   = $req.RawRG                # raw RG from Excel
    $inputGroupRaw      = $req.RawGroup             # raw group from Excel

    # RG existence via preloaded map
    $rgKey = ''
    if ($rgName) { $rgKey = $rgName.ToString().ToLowerInvariant() }
    $rg = $null
    if ($rgKey -and $rgMap.ContainsKey($rgKey)) { $rg = $rgMap[$rgKey] }

    # If RG doesn't exist in subscription
    if (-not $rg) {
      $result.Add([pscustomobject]@{
        SubscriptionName       = $sub.Name
        SubscriptionId         = $sub.Id
        Environment            = $env
        InputResourceGroupRaw  = $inputResourceRaw
        ScannedResourceGroup   = $rgName
        RoleDefinition         = $roleName
        InputAdGroupRaw        = $inputGroupRaw
        AdGroupName            = ''
        GroupObjectId          = ''
        RGStatus               = 'NOT_FOUND'
        PermissionStatus       = 'N/A_RG_NOT_FOUND'
        Status                 = 'RG_NOT_FOUND'
        Details                = 'Resource group not found in this subscription'
      })
      continue
    }

    # Resolve Entra group
    $group = $null
    if ($adGroupName -and ($adGroupName.Trim()).Length -gt 0) {
      $group = Resolve-Group -DisplayName $adGroupName
    }

    if (-not $group) {
      $result.Add([pscustomobject]@{
        SubscriptionName       = $sub.Name
        SubscriptionId         = $sub.Id
        Environment            = $env
        InputResourceGroupRaw  = $inputResourceRaw
        ScannedResourceGroup   = $rgName
        RoleDefinition         = $roleName
        InputAdGroupRaw        = $inputGroupRaw
        AdGroupName            = $adGroupName
        GroupObjectId          = ''
        RGStatus               = 'EXISTS'
        PermissionStatus       = 'N/A_GROUP_NOT_FOUND'
        Status                 = 'GROUP_NOT_FOUND'
        Details                = 'Entra ID group not found'
      })
      continue
    }

    $scope = "/subscriptions/$($sub.Id)/resourceGroups/$rgName"
    $ra = Get-AzRoleAssignment -Scope $scope -ObjectId $group.Id -RoleDefinitionName $roleName -ErrorAction SilentlyContinue

    if ($ra) {
      $result.Add([pscustomobject]@{
        SubscriptionName       = $sub.Name
        SubscriptionId         = $sub.Id
        Environment            = $env
        InputResourceGroupRaw  = $inputResourceRaw
        ScannedResourceGroup   = $rgName
        RoleDefinition         = $roleName
        InputAdGroupRaw        = $inputGroupRaw
        AdGroupName            = $adGroupName
        GroupObjectId          = $group.Id
        RGStatus               = 'EXISTS'
        PermissionStatus       = 'EXISTS'
        Status                 = 'EXISTS'
        Details                = ''
      })
    } else {
      $result.Add([pscustomobject]@{
        SubscriptionName       = $sub.Name
        SubscriptionId         = $sub.Id
        Environment            = $env
        InputResourceGroupRaw  = $inputResourceRaw
        ScannedResourceGroup   = $rgName
        RoleDefinition         = $roleName
        InputAdGroupRaw        = $inputGroupRaw
        AdGroupName            = $adGroupName
        GroupObjectId          = $group.Id
        RGStatus               = 'EXISTS'
        PermissionStatus       = 'MISSING'
        Status                 = 'MISSING'
        Details                = 'Role assignment not found at RG scope'
      })
    }
  }
}

# -------- Output --------
$stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')

# Decide output directory
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  $excelDir = Split-Path -Path $FilePath -Parent
  if ([string]::IsNullOrWhiteSpace($excelDir)) { $excelDir = (Get-Location).Path }
  $OutputDir = Join-Path $excelDir 'kv-scan-local'
}

if (-not (Test-Path $OutputDir)) {
  New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

$outFile = Join-Path $OutputDir "rg-permissions-scan_${stamp}.csv"

$result | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8

# Console summary
$tot = $result.Count
$ok  = ($result | Where-Object { $_.Status -eq 'EXISTS' }).Count
$miss= ($result | Where-Object { $_.Status -eq 'MISSING' }).Count
$rgna= ($result | Where-Object { $_.Status -eq 'RG_NOT_FOUND' }).Count
$gna = ($result | Where-Object { $_.Status -eq 'GROUP_NOT_FOUND' }).Count

Write-Host ""
Write-Host "Scan complete for tenant $TenantId and subscriptions ending with _ADH$adh_group"
Write-Host "  Total checks:     $tot"
Write-Host "  Present:          $ok"
Write-Host "  Missing:          $miss"
Write-Host "  RG not found:     $rgna"
Write-Host "  Group not found:  $gna"
Write-Host "Report: $outFile"
