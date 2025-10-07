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
Import-Module Az.Network  -ErrorAction Stop
Import-Module Az.Resources -ErrorAction Stop

function Ensure-Dir([string]$p){ if([string]::IsNullOrWhiteSpace($p)){ $p = Join-Path (Get-Location) 'vnet-out' } if(-not(Test-Path $p)){ New-Item -ItemType Directory -Path $p -Force | Out-Null } return $p }

$sec=ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$cred=[pscredential]::new($ClientId,$sec)
Connect-AzAccount -ServicePrincipal -Tenant $TenantId -Credential $cred | Out-Null

$OutputDir=Ensure-Dir $OutputDir
$stamp=(Get-Date).ToString('yyyyMMdd_HHmmss')
$outCsv=Join-Path $OutputDir "vnet_topology_$stamp.csv"
$outHtml=Join-Path $OutputDir "vnet_topology_$stamp.html"
$outJson=Join-Path $OutputDir "vnet_topology_$stamp.json"
$rows = New-Object System.Collections.Generic.List[object]

$subs = if($ScanAll){ Get-AzSubscription | ? { $_.Name -match 'ADH' } } else { Get-AzSubscription | ? { $_.Name -match 'ADH' -and $_.Name -match [regex]::Escape($adh_group) } }

foreach($sub in $subs){
  Set-AzContext -Tenant $TenantId -SubscriptionId $sub.Id | Out-Null
  $vnets = Get-AzVirtualNetwork -ErrorAction SilentlyContinue
  foreach($v in $vnets){
    $peerings = Get-AzVirtualNetworkPeering -VirtualNetworkName $v.Name -ResourceGroupName $v.ResourceGroupName -ErrorAction SilentlyContinue
    $subnets  = $v.Subnets
    foreach($sn in $subnets){
      $rtId = $sn.RouteTable.Id
      $rtName = ""
      $rtRoutes = ""
      if($rtId){
        try {
          $rt = Get-AzRouteTable -ResourceGroupName $v.ResourceGroupName -Name (Split-Path $rtId -Leaf) -ErrorAction Stop
          $rtName = $rt.Name
          $rtRoutes = ($rt.Routes | % { "$($_.AddressPrefix) -> $($_.NextHopType) $($_.NextHopIpAddress)" }) -join '; '
        } catch {
          $rtName = "NotAccessible"
        }
      }
      $peerFlat = ($peerings | % { "$($_.Name):$(if($_.PeeringState){$_.PeeringState}else{'?'}):$(if($_.RemoteVirtualNetwork.Id){(Split-Path $_.RemoteVirtualNetwork.Id -Leaf)}else{'?'})" }) -join '; '
      $rows.Add([pscustomobject]@{
        SubscriptionName=$sub.Name; SubscriptionId=$sub.Id
        VNetRG=$v.ResourceGroupName; VNetName=$v.Name; AddressSpace=($v.AddressSpace.AddressPrefixes -join ',')
        SubnetName=$sn.Name; SubnetPrefix=($sn.AddressPrefix); RouteTable=$rtName; RouteRules=$rtRoutes
        Peerings=$peerFlat
      })
    }
  }
}

$rows | Export-Csv $outCsv -NoTypeInformation -Encoding UTF8
($rows | ConvertTo-Html -Title "VNet Topology $stamp" -PreContent "<h2>VNet Topology ($BranchName)</h2>") | Set-Content -Path $outHtml -Encoding UTF8
$rows | ConvertTo-Json -Depth 6 | Set-Content -Path $outJson -Encoding UTF8
Write-Host "CSV:  $outCsv"
Write-Host "HTML: $outHtml"
Write-Host "JSON: $outJson"
