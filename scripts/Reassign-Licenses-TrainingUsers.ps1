<#
.SYNOPSIS
  Removes all licenses from all users except an exempt UPN, then assigns a selected license SKU to P01-P10.

.DESCRIPTION
  Phase A: For each targeted user, remove ALL currently assigned license SKUs.
  Phase B: Assign a selected SKU to training users P01..P10.

  Uses Microsoft Graph PowerShell.

IMPORTANT / RISK
  This can disrupt your tenant by removing licenses from users (mailbox access, Teams, SharePoint, etc.).
  Always run with -WhatIf first.

PREREQUISITES
  - Microsoft Graph PowerShell module (Microsoft.Graph)
  - Permissions/scopes (admin consent likely required): User.ReadWrite.All, Organization.Read.All
  - Role: typically Global Administrator or User Administrator

.EXAMPLE
  # Dry-run first (recommended)
  ./Reassign-Licenses-TrainingUsers.ps1 -SkuPartNumber ENTERPRISEPACK -WhatIf

.EXAMPLE
  # Execute against the tenant, excluding Albert only
  ./Reassign-Licenses-TrainingUsers.ps1 -SkuPartNumber ENTERPRISEPACK

.EXAMPLE
  # Execute and include guests too (not recommended)
  ./Reassign-Licenses-TrainingUsers.ps1 -SkuPartNumber ENTERPRISEPACK -IncludeGuests

.EXAMPLE
  # Execute and reset licenses for existing training users too, then assign
  ./Reassign-Licenses-TrainingUsers.ps1 -SkuPartNumber ENTERPRISEPACK -Start 1 -End 10
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
  # Pick ONE of these:
  [Parameter(ParameterSetName = 'ByPartNumber', Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]$SkuPartNumber,

  [Parameter(ParameterSetName = 'BySkuId', Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [Guid]$SkuId,

  [Parameter()]
  [ValidateNotNullOrEmpty()]
  [string]$ExemptUpn = 'albert.jeremy@6x7m1f.onmicrosoft.com',

  [Parameter()]
  [ValidateRange(1, 200)]
  [int]$Start = 1,

  [Parameter()]
  [ValidateRange(1, 200)]
  [int]$End = 10,

  [Parameter()]
  [ValidateNotNullOrEmpty()]
  [string]$UserPrefix = 'P',

  [Parameter()]
  [ValidateNotNullOrEmpty()]
  [string]$UpnSuffix = '6x7m1f.onmicrosoft.com',

  [Parameter()]
  [ValidateNotNullOrEmpty()]
  [string]$UsageLocation = 'ID',

  [Parameter()]
  [switch]$SetUsageLocationIfMissing,

  [Parameter()]
  [switch]$IncludeGuests,

  [Parameter()]
  [ValidateRange(0, 500)]
  [int]$ThrottleMs = 50,

  [Parameter()]
  [string]$LogCsvPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-GraphModuleInstalled {
  if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    throw "Microsoft.Graph module not found. Install it with: Install-Module Microsoft.Graph -Scope CurrentUser"
  }
}

function Connect-Graph {
  $scopes = @(
    'User.ReadWrite.All',
    'Organization.Read.All'
  )
  Write-Host "Connecting to Microsoft Graph with scopes: $($scopes -join ', ')" -ForegroundColor Cyan
  Connect-MgGraph -Scopes $scopes | Out-Null
}

function Resolve-SkuId {
  if ($PSCmdlet.ParameterSetName -eq 'BySkuId') {
    return $SkuId
  }

  $skus = Get-MgSubscribedSku -All
  $match = $skus | Where-Object { $_.SkuPartNumber -eq $SkuPartNumber } | Select-Object -First 1
  if (-not $match) {
    $available = ($skus | Sort-Object SkuPartNumber | Select-Object -ExpandProperty SkuPartNumber) -join ', '
    throw "SkuPartNumber '$SkuPartNumber' not found. Available: $available"
  }
  return [Guid]$match.SkuId
}

function Format-Id([int]$i) {
  return '{0}{1:00}' -f $UserPrefix, $i
}

function Get-TrainingUpns {
  $upns = @()
  for ($i = $Start; $i -le $End; $i++) {
    $id = Format-Id -i $i
    $upns += "$id@$UpnSuffix"
  }
  return $upns
}

function Get-TargetUsers {
  # Grab minimal properties; we'll query license details per-user.
  $props = @('id','userPrincipalName','displayName','userType','accountEnabled','usageLocation')
  $users = Get-MgUser -All -Property ($props -join ',')

  $filtered = $users | Where-Object {
    $_.UserPrincipalName -and
    $_.UserPrincipalName.ToLowerInvariant() -ne $ExemptUpn.ToLowerInvariant() -and
    ($IncludeGuests -or $_.UserType -eq 'Member')
  }

  return $filtered
}

function Get-UserSkuIdsToRemove([string]$userId) {
  $details = Get-MgUserLicenseDetail -UserId $userId
  $skuIdsToRemove = @($details | ForEach-Object { [Guid]$_.SkuId })

  return $skuIdsToRemove
}

function Test-UsageLocationMissing($currentUsageLocation) {
  if (-not $SetUsageLocationIfMissing) { return $false }
  if ($null -ne $currentUsageLocation -and "$currentUsageLocation".Trim().Length -gt 0) { return $false }
  return $true
}

function Set-UserLicense([string]$userId, [Guid]$targetSkuId) {
  $add = @(@{ SkuId = $targetSkuId })
  Set-MgUserLicense -UserId $userId -AddLicenses $add -RemoveLicenses @() | Out-Null
}

function Main {
  Test-GraphModuleInstalled
  Connect-Graph

  $targetSkuId = Resolve-SkuId
  Write-Host "Target license SKU: $targetSkuId" -ForegroundColor Green

  $log = New-Object System.Collections.Generic.List[object]

  # Phase A: remove all licenses from all users except exempt
  Write-Host "Phase A: removing ALL licenses from users (except $ExemptUpn)" -ForegroundColor Yellow
  $targets = Get-TargetUsers
  $total = $targets.Count

  $idx = 0
  foreach ($u in $targets) {
    $idx++
    $upn = $u.UserPrincipalName

    try {
      $skuIdsToRemove = Get-UserSkuIdsToRemove -userId $u.Id
      if (-not $skuIdsToRemove -or $skuIdsToRemove.Count -eq 0) {
        $log.Add([pscustomobject]@{ Phase = 'Remove'; UserPrincipalName = $upn; Action = 'NoLicenses'; Timestamp = (Get-Date).ToString('s') })
      } else {
        if ($PSCmdlet.ShouldProcess($upn, "Remove licenses: $($skuIdsToRemove -join ', ')")) {
          Set-MgUserLicense -UserId $u.Id -AddLicenses @() -RemoveLicenses $skuIdsToRemove | Out-Null
        }
        $log.Add([pscustomobject]@{ Phase = 'Remove'; UserPrincipalName = $upn; Action = 'Removed'; Timestamp = (Get-Date).ToString('s') })
      }
    } catch {
      $log.Add([pscustomobject]@{ Phase = 'Remove'; UserPrincipalName = $upn; Action = 'Error'; Error = $_.Exception.Message; Timestamp = (Get-Date).ToString('s') })
      Write-Warning "[$idx/$total] Failed removing licenses for ${upn}: $($_.Exception.Message)"
    }

    if ($ThrottleMs -gt 0) { Start-Sleep -Milliseconds $ThrottleMs }
  }

  # Phase B: assign target SKU to P01-P10
  Write-Host "Phase B: assigning license to training users $(Format-Id -i $Start)..$(Format-Id -i $End)" -ForegroundColor Cyan
  $trainingUpns = Get-TrainingUpns

  foreach ($upn in $trainingUpns) {
    try {
      $tu = Get-MgUser -UserId $upn -Property 'id,userPrincipalName,usageLocation'
      if (Test-UsageLocationMissing -currentUsageLocation $tu.UsageLocation) {
        if ($PSCmdlet.ShouldProcess($upn, "Set usageLocation='$UsageLocation'")) {
          Update-MgUser -UserId $tu.Id -UsageLocation $UsageLocation | Out-Null
        }
      }

      if ($PSCmdlet.ShouldProcess($upn, "Assign SKU $targetSkuId")) {
        Set-UserLicense -userId $tu.Id -targetSkuId $targetSkuId
      }
      $log.Add([pscustomobject]@{ Phase = 'Assign'; UserPrincipalName = $upn; Action = 'Assigned'; Timestamp = (Get-Date).ToString('s') })
    } catch {
      $log.Add([pscustomobject]@{ Phase = 'Assign'; UserPrincipalName = $upn; Action = 'Error'; Error = $_.Exception.Message; Timestamp = (Get-Date).ToString('s') })
      Write-Warning "Failed assigning license for ${upn}: $($_.Exception.Message)"
    }

    if ($ThrottleMs -gt 0) { Start-Sleep -Milliseconds $ThrottleMs }
  }

  Write-Host "Done." -ForegroundColor Green

  if ($LogCsvPath) {
    $log | Export-Csv -Path $LogCsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Log written to: $LogCsvPath" -ForegroundColor Green
  } else {
    $log | Select-Object Phase, UserPrincipalName, Action, Error | Format-Table -AutoSize
  }
}

Main
