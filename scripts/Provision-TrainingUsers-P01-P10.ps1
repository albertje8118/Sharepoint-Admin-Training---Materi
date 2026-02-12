<#
.SYNOPSIS
  Provisions training user accounts P01-P10 in Microsoft Entra ID (Azure AD) using Microsoft Graph PowerShell.

.DESCRIPTION
  - Creates users P01..P10 with UPNs like P01@<UpnSuffix>
  - Sets a default password (per request) and forces change at next sign-in
  - Skips users that already exist by default (optional reset)

PREREQUISITES
  - PowerShell 7+ or Windows PowerShell 5.1
  - Microsoft Graph PowerShell module (Microsoft.Graph)
  - Permissions: Directory.ReadWrite.All (or User.ReadWrite.All) in Microsoft Graph
  - Role: typically User Administrator or Global Administrator

SECURITY NOTE
  Hard-coding passwords in scripts is not recommended. This script supports overriding the password via parameter.

.EXAMPLE
  ./Provision-TrainingUsers-P01-P10.ps1 -UpnSuffix "contoso.onmicrosoft.com"

.EXAMPLE
  ./Provision-TrainingUsers-P01-P10.ps1 -UpnSuffix "contoso.onmicrosoft.com" -ResetIfExists
#>

[CmdletBinding(SupportsShouldProcess)]
param(
  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]$UpnSuffix,

  [Parameter()]
  [ValidateNotNullOrEmpty()]
  [string]$DefaultPasswordPlain = 'Pa55w.rd1234!',

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
  [string]$DisplayNamePrefix = 'Northwind Training Admin',

  [Parameter()]
  [ValidateNotNullOrEmpty()]
  [string]$UsageLocation = 'ID',

  [Parameter()]
  [switch]$ResetIfExists
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-GraphModule {
  if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    throw "Microsoft.Graph module not found. Install it with: Install-Module Microsoft.Graph -Scope CurrentUser"
  }
}

function Connect-Graph {
  # Directory.ReadWrite.All is sufficient to create users (admin consent required in many tenants)
  $scopes = @('Directory.ReadWrite.All')
  Write-Host "Connecting to Microsoft Graph with scopes: $($scopes -join ', ')" -ForegroundColor Cyan
  Connect-MgGraph -Scopes $scopes | Out-Null
}

function Format-Id([int]$i) {
  return '{0}{1:00}' -f $UserPrefix, $i
}

function Get-ExistingUser([string]$upn) {
  try {
    return Get-MgUser -UserId $upn -ErrorAction Stop
  } catch {
    return $null
  }
}

function New-TrainingUser([string]$id, [string]$upn) {
  $mailNickname = $id
  $passwordProfile = @{ 
    password = $DefaultPasswordPlain
    forceChangePasswordNextSignIn = $true
  }

  $body = @{ 
    accountEnabled  = $true
    displayName     = "$DisplayNamePrefix $id"
    mailNickname    = $mailNickname
    userPrincipalName = $upn
    passwordProfile = $passwordProfile
    usageLocation   = $UsageLocation
  }

  if ($PSCmdlet.ShouldProcess($upn, 'Create user')) {
    New-MgUser -BodyParameter $body | Out-Null
  }
}

function Reset-TrainingUserPassword([string]$upn) {
  $passwordProfile = @{ 
    password = $DefaultPasswordPlain
    forceChangePasswordNextSignIn = $true
  }

  if ($PSCmdlet.ShouldProcess($upn, 'Reset password + require change')) {
    Update-MgUser -UserId $upn -PasswordProfile $passwordProfile | Out-Null
  }
}

function Main {
  if ($End -lt $Start) {
    throw "End ($End) must be >= Start ($Start)."
  }

  Ensure-GraphModule
  Connect-Graph

  Write-Host "Provisioning users ${UserPrefix}{Start:00}..${UserPrefix}{End:00} in UPN suffix '$UpnSuffix'" -ForegroundColor Green

  $results = @()

  for ($i = $Start; $i -le $End; $i++) {
    $id = Format-Id -i $i
    $upn = "$id@$UpnSuffix"

    $existing = Get-ExistingUser -upn $upn
    if ($null -ne $existing) {
      if ($ResetIfExists) {
        Reset-TrainingUserPassword -upn $upn
        $results += [pscustomobject]@{ Id = $id; UPN = $upn; Action = 'ResetPassword' }
        Write-Host "[$id] Exists -> password reset" -ForegroundColor Yellow
      } else {
        $results += [pscustomobject]@{ Id = $id; UPN = $upn; Action = 'SkippedExists' }
        Write-Host "[$id] Exists -> skipped" -ForegroundColor DarkYellow
      }
      continue
    }

    New-TrainingUser -id $id -upn $upn
    $results += [pscustomobject]@{ Id = $id; UPN = $upn; Action = 'Created' }
    Write-Host "[$id] Created" -ForegroundColor Cyan
  }

  Write-Host "Done." -ForegroundColor Green
  $results | Format-Table -AutoSize
}

Main
