param(
  # ===== Auth (SPN) =====
  [Parameter(Mandatory=$true)][string]$TenantId,
  [Parameter(Mandatory=$true)][string]$ClientId,
  [Parameter(Mandatory=$true)][string]$ClientSecret,

  # ===== Subscription selection =====
  [switch]$ScanAll,               # scan all subs with "ADH" in name
  [string]$CustodianFilter = "",  # e.g., "CSM" -> subs with ADH and CSM
  [string[]]$SubscriptionIds = @(),  # optional explicit list (overrides filters)

  # ===== Output & options =====
  [string]$OutputDir = "",            # default: ./rg-tags-out
  [string]$TeamsWebhookUrl = "",      # optional incoming webhook
  [string[]]$RequiredTagKeys = @(),   # optionally check these exist per RG (case-insensitive)
  [switch]$PerTagRows                 # if set, output 1 row per tag instead of 1 per RG
)

$ErrorActionPreference = 'Stop'
Import-Module Az.Accounts -ErrorAction Stop
Import-Module Az.Resources -ErrorAction Stop

# ---------- helpers ----------
function Get-EnvFromSubscriptionName([string]$Name){
  if ($Name -match '(?i)\b(prod|production|prd)\b'){ 'PRODUCTION' } else { 'NONPRODUCTION' }
}
function Ensure-OutputDir([string]$dir){
  if ([string]::IsNullOrWhiteSpace($dir)){ $dir = Join-Path (Get-Location).Path 'rg-tags-out' }
  if (-not (Test-Path $dir)){ New-Item -ItemType Directory -Path $dir | Out-Null }
  $dir
}
function To-FlatTagString($tags){
  if (-not $tags){ return '' }
  ($tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
}
function Has-TagKey($tags, [string]$key){
  if (-not $tags){ return $false }
  foreach($k in $tags.Keys){ if ($k -ieq $key){ return $true } }
  return $false
}

# ---------- login ----------
$sec   = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$creds = New-Object System.Management.Automation.PSCredential($ClientId,$sec)
Connect-AzAccount -ServicePrincipal -Tenant $TenantId -Credential $creds | Out-Null

# ---------- pick subscriptions ----------
$subs = @()
if ($SubscriptionIds -and $SubscriptionIds.Count -gt 0){
  foreach($sid in $SubscriptionIds){
    $s = Get-AzSubscription -SubscriptionId $sid -ErrorAction SilentlyContinue
    if ($s){ $subs += $s }
  }
  if (-not $subs){ throw "None of the provided -SubscriptionIds were accessible." }
} else {
  $cand = Get-AzSubscription | Where-Object { $_.Name -match '(?i)ADH' }
  if ($CustodianFilter){ $cand = $cand | Where-Object { $_.Name -match [regex]::Escape($CustodianFilter) } }
  if (-not $ScanAll.IsPresent -and -not $CustodianFilter){
    throw "Provide -ScanAll to scan all 'ADH' subscriptions, or -CustodianFilter/-SubscriptionIds."
  }
  if (-not $cand){ throw "No subscriptions matched the selection (ADH + optional custodian filter)." }
  $subs = $cand
}

# ---------- outputs ----------
$OutputDir = Ensure-OutputDir $OutputDir
$stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
$csvPath  = Join-Path $OutputDir "rg-tags_${stamp}.csv"
$htmlPath = Join-Path $OutputDir "rg-tags_${stamp}.html"

$rows = New-Object System.Collections.Generic.List[object]

# ---------- scan ----------
foreach($sub in $subs){
  Set-AzContext -Tenant $TenantId -SubscriptionId $sub.Id | Out-Null
  $env = Get-EnvFromSubscriptionName $sub.Name

  $rgs = Get-AzResourceGroup -ErrorAction SilentlyContinue
  foreach($rg in $rgs){
    if ($PerTagRows){
      if ($rg.Tags -and $rg.Tags.Count -gt 0){
        foreach($k in $rg.Tags.Keys){
          $rows.Add([pscustomobject]@{
            SubscriptionName = $sub.Name
            SubscriptionId   = $sub.Id
            Environment      = $env
            ResourceGroup    = $rg.ResourceGroupName
            TagName          = $k
            TagValue         = $rg.Tags[$k]
          })
        }
      } else {
        # emit a row to show "no tags"
        $rows.Add([pscustomobject]@{
          SubscriptionName = $sub.Name
          SubscriptionId   = $sub.Id
          Environment      = $env
          ResourceGroup    = $rg.ResourceGroupName
          TagName          = ''
          TagValue         = ''
        })
      }
    } else {
      # one row per RG, flattened tags + optional required-key checks
      $flat = To-FlatTagString $rg.Tags

      # synthesize per-required-key presence columns (true/false)
      $presence = @{}
      foreach($key in $RequiredTagKeys){
        $presence["Has_$($key)"] = (Has-TagKey $rg.Tags $key)
      }

      $base = [ordered]@{
        SubscriptionName = $sub.Name
        SubscriptionId   = $sub.Id
        Environment      = $env
        ResourceGroup    = $rg.ResourceGroupName
        TagCount         = ($rg.Tags ? $rg.Tags.Count : 0)
        TagsFlat         = $flat
      }

      foreach($k in $presence.Keys){ $base[$k] = $presence[$k] }
      $rows.Add([pscustomobject]$base)
    }
  }
}

# ---------- export ----------
$rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
($rows | ConvertTo-Html -Title "RG Tags $stamp" -PreContent "<h2>Resource Group Tags - $stamp</h2>") |
  Set-Content -Path $htmlPath -Encoding UTF8

Write-Host "`nRG Tags CSV : $csvPath"
Write-Host "RG Tags HTML: $htmlPath"

# ---------- optional Teams summary ----------
if ($TeamsWebhookUrl -and $TeamsWebhookUrl.Trim()){
  try{
    $rgCount = ($rows | Select-Object -ExpandProperty ResourceGroup -Unique | Measure-Object).Count
    $subCount = ($rows | Select-Object -ExpandProperty SubscriptionId -Unique | Measure-Object).Count
    $summary  = "RG Tag Scan $stamp`nSubscriptions: $subCount`nRGs (unique): $rgCount`nMode: " + ($(if($PerTagRows){'PerTagRows'} else {'PerRG'}))
    $payload = @{ text = $summary } | ConvertTo-Json
    Invoke-RestMethod -Method Post -Uri $TeamsWebhookUrl -ContentType 'application/json' -Body $payload | Out-Null
    Write-Host "Teams notification posted."
  } catch { Write-Warning "Failed to post to Teams: $_" }
}
