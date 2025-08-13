<# 
Purpose:
  Scan subscriptions → resource groups → RBAC role assignments and compare
  against a matrix kept in an Excel worksheet that has two side-by-side blocks:
    - PRODUCTION:     A:resource_group_name  B:role_definition_name  C:ad_group_name
    - NONPRODUCTION:  F:resource_group_name  G:role_definition_name  H:ad_group_name

  Group names in the sheet may contain the placeholder "<Custodian>", which will
  be replaced with the value of -AdhGroup (e.g., ADH_<Custodian>_KV -> ADH_CSM_KV).

Outputs:
  A CSV report with columns:
    SubscriptionName, SubscriptionId, Environment, ResourceGroup, RoleDefinition,
    AdGroupName, GroupObjectId, Status, Details
#>

param(
  [Parameter(Mandatory=$true)]
  [string]$FilePath,                        # e.g. C:\path\to\rg-permissions.xlsx

  [Parameter(Mandatory=$true)]
  [string]$WorksheetName,                   # e.g. "rg permissions"

  [Parameter(Mandatory=$true)]
  [string]$AdhGroup,                        # e.g. "CSM"

  [string]$SubscriptionNameLike = "",       # optional filter e.g. "*ADHCSM*" or "*PROD*"
  
  [switch]$UseDeviceLogin                    # use device code login instead of browser/SP
)

# ---------------- Safety & modules ----------------
$ErrorActionPreference = 'Stop'
try { Import-Module Az.Accounts -ErrorAction Stop } catch { Write-Error "Az.Accounts missing. Run: Install-Module Az -Scope CurrentUser"; exit 1 }
try { Import-Module Az.Resources -ErrorAction Stop } catch { Write-Error "Az.Resources missing. Run: Install-Module Az -Scope CurrentUser"; exit 1 }
try { Import-Module ImportExcel -ErrorAction Stop } catch { Write-Error "ImportExcel missing. Run: Install-Module ImportExcel -Scope CurrentUser"; exit 1 }

# ---------------- Login ----------------
if ($UseDeviceLogin) {
  Connect-AzAccount -UseDeviceAuthentication | Out-Null
} else {
  # Uses default interactive browser; if already logged in, this is a no-op
  Connect-AzAccount | Out-Null
}

# ---------------- Helpers ----------------
function Get-EnvFromSubscriptionName {
  param([string]$Name)
  if ($Name -match '(?i)\b(prod|production)\b') { return 'PRODUCTION' }
  else { return 'NONPRODUCTION' }
}

function Read-RGMatrixFromExcel {
  param(
    [string]$Path,
    [string]$Sheet
  )
  # We ingest the sheet WITHOUT headers so we can safely pull the two column-blocks.
  $rows = Import-Excel -Path $Path -WorksheetName $Sheet -NoHeader

  if (-not $rows -or $rows.Count -eq 0) {
    throw "No data read from worksheet '$Sheet'. Verify the sheet name and content."
  }

  # ImportExcel uses P1..Pn for columns when -NoHeader is used.
  # We will locate the header rows where the first column equals 'resource_group_name'
  # (for PRODUCTION block A-C => P1..P3) and where P6 equals 'resource_group_name'
  # (for NONPRODUCTION block F-H => P6..P8).
  # Then read until blank rows.
  $prodStart = ($rows | Select-Object P1,P2,P3 | 
                Select-Object @{n='val';e={$_.P1}} | 
                ForEach-Object -Begin {$i=0} -Process {
                  $script:i++
                  [pscustomobject]@{ idx=$i-1; val=$_.val }
                } | Where-Object { "$($_.val)".Trim().ToLower() -eq 'resource_group_name' } |
                Select-Object -First 1).idx

  $nonProdStart = ($rows | Select-Object P6,P7,P8 |
                   Select-Object @{n='val';e={$_.P6}} |
                   ForEach-Object -Begin {$j=0} -Process {
                     $script:j++
                     [pscustomobject]@{ idx=$j-1; val=$_.val }
                   } | Where-Object { "$($_.val)".Trim().ToLower() -eq 'resource_group_name' } |
                   Select-Object -First 1).idx

  if ($null -eq $prodStart -and $null -eq $nonProdStart) {
    throw "Could not find header rows. Expecting 'resource_group_name' in column A and/or F."
  }

  $prod = @()
  if ($null -ne $prodStart) {
    for ($r = $prodStart + 1; $r -lt $rows.Count; $r++) {
      $rg   = "$($rows[$r].P1)".Trim()
      $role = "$($rows[$r].P2)".Trim()
      $grp  = "$($rows[$r].P3)".Trim()
      if ([string]::IsNullOrWhiteSpace($rg) -and [string]::IsNullOrWhiteSpace($role) -and [string]::IsNullOrWhiteSpace($grp)) { break }
      if (-not [string]::IsNullOrWhiteSpace($rg)) {
        $prod += [pscustomobject]@{ Environment='PRODUCTION'; ResourceGroup=$rg; Role=$role; Group=$grp }
      }
    }
  }

  $nonprod = @()
  if ($null -ne $nonProdStart) {
    for ($r = $nonProdStart + 1; $r -lt $rows.Count; $r++) {
      $rg   = "$($rows[$r].P6)".Trim()
      $role = "$($rows[$r].P7)".Trim()
      $grp  = "$($rows[$r].P8)".Trim()
      if ([string]::IsNullOrWhiteSpace($rg) -and [string]::IsNullOrWhiteSpace($role) -and [string]::IsNullOrWhiteSpace($grp)) { break }
      if (-not [string]::IsNullOrWhiteSpace($rg)) {
        $nonprod += [pscustomobject]@{ Environment='NONPRODUCTION'; ResourceGroup=$rg; Role=$role; Group=$grp }
      }
    }
  }

  return [pscustomobject]@{
    PRODUCTION    = $prod
    NONPRODUCTION = $nonprod
  }
}

function Resolve-Group {
  param([string]$DisplayName)
  # Prefer exact-display name match; fall back to -SearchString if needed
  $g = Get-AzADGroup -DisplayName $DisplayName -ErrorAction SilentlyContinue
  if (-not $g) {
    $g = Get-AzADGroup -SearchString $DisplayName -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -eq $DisplayName } | Select-Object -First 1
  }
  return $g
}

# ---------------- Load matrix ----------------
$matrix = Read-RGMatrixFromExcel -Path $FilePath -Sheet $WorksheetName
$prodReqs    = $matrix.PRODUCTION
$nonprodReqs = $matrix.NONPRODUCTION

if (($prodReqs.Count + $nonprodReqs.Count) -eq 0) {
  throw "No requirements parsed from worksheet '$WorksheetName'. Check the columns A–C and F–H."
}

# ---------------- Subscriptions ----------------
$subs = Get-AzSubscription
if ($SubscriptionNameLike) {
  $subs = $subs | Where-Object { $_.Name -like $SubscriptionNameLike }
}

if (-not $subs) { throw "No subscriptions to scan. Adjust -SubscriptionNameLike or your access." }

# ---------------- Scan ----------------
$result = New-Object System.Collections.Generic.List[object]

foreach ($sub in $subs) {
  Set-AzContext -SubscriptionId $sub.Id | Out-Null

  $env = Get-EnvFromSubscriptionName -Name $sub.Name
  $reqs = if ($env -eq 'PRODUCTION') { $prodReqs } else { $nonprodReqs }

  foreach ($req in $reqs) {
    $rgName = $req.ResourceGroup
    $role   = $req.Role
    $grpTempl = $req.Group

    # Replace placeholder
    $adGroupName = $grpTempl -replace '<Custodian>', $AdhGroup

    # Validate RG exists
    $rg = Get-AzResourceGroup -Name $rgName -ErrorAction SilentlyContinue
    if (-not $rg) {
      $result.Add([pscustomobject]@{
        SubscriptionName = $sub.Name
        SubscriptionId   = $sub.Id
        Environment      = $env
        ResourceGroup    = $rgName
        RoleDefinition   = $role
        AdGroupName      = $adGroupName
        GroupObjectId    = ''
        Status           = 'RG_NOT_FOUND'
        Details          = 'Resource group not found in this subscription'
      })
      continue
    }

    # Resolve Group
    $group = Resolve-Group -DisplayName $adGroupName
    if (-not $group) {
      $result.Add([pscustomobject]@{
        SubscriptionName = $sub.Name
        SubscriptionId   = $sub.Id
        Environment      = $env
        ResourceGroup    = $rgName
        RoleDefinition   = $role
        AdGroupName      = $adGroupName
        GroupObjectId    = ''
        Status           = 'GROUP_NOT_FOUND'
        Details          = 'Entra ID group not found'
      })
      continue
    }

    $scope = "/subscriptions/$($sub.Id)/resourceGroups/$rgName"

    # Check role assignment
    $ra = Get-AzRoleAssignment -Scope $scope -ObjectId $group.Id -RoleDefinitionName $role -ErrorAction SilentlyContinue

    if ($ra) {
      $result.Add([pscustomobject]@{
        SubscriptionName = $sub.Name
        SubscriptionId   = $sub.Id
        Environment      = $env
        ResourceGroup    = $rgName
        RoleDefinition   = $role
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
        RoleDefinition   = $role
        AdGroupName      = $adGroupName
        GroupObjectId    = $group.Id
        Status           = 'MISSING'
        Details          = 'Role assignment not found at RG scope'
      })
    }
  }
}

# ---------------- Output ----------------
$stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
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

Write-Host "Scan complete:"
Write-Host "  Total checks:     $tot"
Write-Host "  Present:          $ok"
Write-Host "  Missing:          $miss"
Write-Host "  RG not found:     $rgna"
Write-Host "  Group not found:  $gna"
Write-Host ""
Write-Host "Report: $outFile"
