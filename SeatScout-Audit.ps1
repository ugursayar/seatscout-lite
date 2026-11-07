<#
.SYNOPSIS
    SeatScout - Microsoft 365 license waste audit. Finds unused/inactive/over-provisioned
    licenses and quantifies how much money you can recover.

.DESCRIPTION
    Runs entirely inside YOUR tenant using the Microsoft Graph PowerShell SDK (read-only).
    No data ever leaves your environment - SeatScout has no servers and collects nothing.

    Produces:
      * CSV exports for each finding (for remediation)
      * A console summary

    Findings:
      * Unassigned purchased seats   (paying for seats nobody has)
      * Disabled-but-licensed users  (blocked accounts still consuming licenses)
      * Never-signed-in licensed users
      * Inactive licensed users      (no sign-in in N days)

.PARAMETER InactiveDays
    A licensed, enabled user with no sign-in in this many days is flagged inactive. Default 30.

.PARAMETER OutputPath
    Folder for the report + CSVs. Default: .\SeatScout-Report-<timestamp>

.PARAMETER PriceTablePath
    CSV mapping skuPartNumber -> friendly name + monthly price. Default: .\config\PriceTable.csv

.PARAMETER Lite
    Free mode: exports the inactive/unused-license CSV + a summary only (no $ totals, no HTML report).

.PARAMETER MockDataPath
    Internal/testing: load tenant data from a JSON file instead of calling Microsoft Graph.

.EXAMPLE
    .\SeatScout-Audit.ps1 -InactiveDays 60

.NOTES
    SeatScout  |  seatscout.dev  |  read-only, runs in your tenant.
#>

[CmdletBinding()]
param(
    [int]    $InactiveDays   = 30,
    [string] $OutputPath,
    [string] $PriceTablePath = (Join-Path (Join-Path $PSScriptRoot 'config') 'PriceTable.csv'),
    [switch] $Lite,
    [string] $MockDataPath
)

$ErrorActionPreference = 'Stop'
$script:Version = '1.0.1'
$script:Edition = 'Lite'   # build edition (set per package): Lite | Solo | Pro
$nowUtc = (Get-Date).ToUniversalTime()

# ----------------------------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------------------------
function Write-Step { param($Msg) Write-Host "[SeatScout] $Msg" -ForegroundColor Cyan }
function Write-Warn { param($Msg) Write-Host "[SeatScout] $Msg" -ForegroundColor Yellow }

function Get-PriceTable {
    param([string]$Path)
    if (-not (Test-Path $Path)) { throw "Price table not found: $Path" }
    $rows = Import-Csv -Path $Path | Where-Object { $_.SkuPartNumber -and -not $_.SkuPartNumber.StartsWith('#') }
    $map = @{}
    foreach ($r in $rows) {
        $price = $null
        if ($r.MonthlyPriceUSD -and ($r.MonthlyPriceUSD -as [double]) -ne $null) { $price = [double]$r.MonthlyPriceUSD }
        $map[$r.SkuPartNumber] = [pscustomobject]@{
            PartNumber   = $r.SkuPartNumber
            FriendlyName = if ($r.FriendlyName) { $r.FriendlyName } else { $r.SkuPartNumber }
            Price        = $price
        }
    }
    return $map
}


# ----------------------------------------------------------------------------------------------
# Data acquisition (Graph OR mock)
# ----------------------------------------------------------------------------------------------
function Get-TenantData {
    param([string]$MockPath)

    if ($MockPath) {
        Write-Step "Loading MOCK data from $MockPath"
        $raw = Get-Content -Path $MockPath -Raw | ConvertFrom-Json
        # Mock dates are absolute, so age/inactivity math runs against the mock's own asOf instead of
        # the clock; otherwise the demo output (and the published sample figures) drift day by day.
        if ($raw.asOf) { $script:nowUtc = ([datetime]$raw.asOf).ToUniversalTime() }
        return $raw
    }

    Write-Step "Connecting to Microsoft Graph (read-only)..."
    $scopes = @('Organization.Read.All','User.Read.All','Directory.Read.All','AuditLog.Read.All')
    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction SilentlyContinue
        Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue
    } catch {
        throw "Microsoft Graph PowerShell SDK not installed. Run:  Install-Module Microsoft.Graph -Scope CurrentUser"
    }
    Connect-MgGraph -Scopes $scopes -NoWelcome

    Write-Step "Reading organization + subscribed SKUs..."
    $org  = Get-MgOrganization
    $skus = Get-MgSubscribedSku -All

    Write-Step "Reading users (license + sign-in activity)..."
    # signInActivity requires Entra ID P1; on non-premium tenants Graph 403s the whole query,
    # so try with it, then fall back to a query without it (disabled + unassigned-seat analysis only).
    $baseProps = 'id,displayName,userPrincipalName,accountEnabled,assignedLicenses,createdDateTime'
    $signInRetrievable = $true
    try {
        $users = Get-MgUser -All -Property ($baseProps + ',signInActivity') -ErrorAction Stop
    } catch {
        Write-Warn "Sign-in activity needs Microsoft Entra ID P1 - retrying without it (disabled + unassigned-seat analysis only)."
        $signInRetrievable = $false
        $users = Get-MgUser -All -Property $baseProps -ErrorAction Stop
    }

    # Normalize Graph objects into the same shape as the mock JSON
    $data = [pscustomobject]@{
        organization      = [pscustomobject]@{ displayName = $org.DisplayName; id = $org.Id }
        signInRetrievable = $signInRetrievable
        subscribedSkus = $skus | ForEach-Object {
            [pscustomobject]@{
                skuId         = $_.SkuId
                skuPartNumber = $_.SkuPartNumber
                prepaidUnits  = [pscustomobject]@{ enabled = $_.PrepaidUnits.Enabled }
                consumedUnits = $_.ConsumedUnits
            }
        }
        users = $users | ForEach-Object {
            $sia = $_.SignInActivity
            [pscustomobject]@{
                id                = $_.Id
                displayName       = $_.DisplayName
                userPrincipalName = $_.UserPrincipalName
                accountEnabled    = $_.AccountEnabled
                createdDateTime   = $_.CreatedDateTime
                assignedLicenses  = @($_.AssignedLicenses | ForEach-Object { [pscustomobject]@{ skuId = $_.SkuId } })
                signInActivity    = if ($sia) { [pscustomobject]@{ lastSignInDateTime = $sia.LastSignInDateTime } } else { $null }
            }
        }
    }
    return $data
}

# ----------------------------------------------------------------------------------------------
# Analysis
# ----------------------------------------------------------------------------------------------
function Invoke-Analysis {
    param($Data, $PriceMap, [int]$InactiveDays)

    # skuId -> details (partNumber, friendly, price)
    $skuById = @{}
    $skuMeta = @()
    $signInDataAvailable = $false

    foreach ($s in $Data.subscribedSkus) {
        $meta = $PriceMap[$s.skuPartNumber]
        $friendly = if ($meta) { $meta.FriendlyName } else { $s.skuPartNumber }
        $price    = if ($meta) { $meta.Price } else { $null }
        $purchased = [int]$s.prepaidUnits.enabled
        $consumed  = [int]$s.consumedUnits
        $isUnlimited = $purchased -gt 100000   # Microsoft sentinel for free / unlimited self-service SKUs
        $unassigned = if ($isUnlimited) { 0 } else { [math]::Max(0, $purchased - $consumed) }
        $info = [pscustomobject]@{
            SkuId=$s.skuId; PartNumber=$s.skuPartNumber; Friendly=$friendly; Price=$price
            Purchased=$purchased; Consumed=$consumed; Unassigned=$unassigned; IsUnlimited=$isUnlimited
        }
        $skuById[$s.skuId] = $info
        $skuMeta += $info
    }

    $cutoff = $nowUtc.AddDays(-$InactiveDays)
    $signInRetrievable = if ($null -ne $Data.signInRetrievable) { [bool]$Data.signInRetrievable } else { $true }

    $disabled=@(); $never=@(); $inactive=@()

    foreach ($u in $Data.users) {
        $lic = @($u.assignedLicenses | ForEach-Object { $skuById[$_.skuId] } | Where-Object { $_ })
        if ($lic.Count -eq 0) { continue }   # unlicensed users are irrelevant to license waste

        $userMonthly = ($lic | Where-Object { $_.Price } | Measure-Object -Property Price -Sum).Sum
        if (-not $userMonthly) { $userMonthly = 0 }
        $licNames = ($lic | ForEach-Object { $_.Friendly }) -join '; '

        $lastSignIn = $null
        if ($u.signInActivity -and $u.signInActivity.lastSignInDateTime) {
            $signInDataAvailable = $true
            try { $lastSignIn = [datetime]$u.signInActivity.lastSignInDateTime } catch { $lastSignIn = $null }
        }
        $created = $null
        if ($u.createdDateTime) { try { $created = [datetime]$u.createdDateTime } catch {} }

        $row = [pscustomobject]@{
            DisplayName=$u.displayName; UserPrincipalName=$u.userPrincipalName
            AccountEnabled=$u.accountEnabled; Licenses=$licNames
            LastSignIn = if ($lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd') } else { 'never' }
            MonthlyCost=[math]::Round($userMonthly,2)
        }

        if (-not $u.accountEnabled) {
            $disabled += $row
        }
        elseif ($signInRetrievable) {
            if (-not $lastSignIn) {
                # never signed in - only flag if the account is old enough to be meaningful
                if (-not $created -or $created -lt $cutoff) { $never += $row }
            }
            elseif ($lastSignIn -lt $cutoff) {
                $inactive += $row
            }
        }

    }

    $result = [ordered]@{
        Org=$Data.organization
        Skus=$skuMeta
        Disabled=$disabled; Never=$never; Inactive=$inactive
        SignInDataAvailable=$signInRetrievable
    }
    return [pscustomobject]$result
}

# ----------------------------------------------------------------------------------------------
# Reporting
# ----------------------------------------------------------------------------------------------

# ----------------------------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------------------------
Write-Host ""
Write-Host "  SeatScout v$script:Version - Microsoft 365 license waste audit" -ForegroundColor Green
Write-Host "  Read-only. Runs in your tenant. Nothing leaves your environment." -ForegroundColor DarkGray
Write-Host ""

if (-not $OutputPath) { $OutputPath = Join-Path (Get-Location) ("SeatScout-Report-" + $nowUtc.ToString('yyyyMMdd-HHmmss')) }
New-Item -ItemType Directory -Force -Path $OutputPath | Out-Null

$priceMap = Get-PriceTable -Path $PriceTablePath
$data     = Get-TenantData -MockPath $MockDataPath
$R        = Invoke-Analysis -Data $data -PriceMap $priceMap -InactiveDays $InactiveDays

# CSV exports (always)
$R.Disabled  | Export-Csv (Join-Path $OutputPath 'disabled-but-licensed.csv') -NoTypeInformation -Encoding utf8
$R.Never     | Export-Csv (Join-Path $OutputPath 'never-signed-in.csv') -NoTypeInformation -Encoding utf8
$R.Inactive  | Export-Csv (Join-Path $OutputPath 'inactive-users.csv') -NoTypeInformation -Encoding utf8

if ($Lite -or $script:Edition -eq 'Lite') {
    Write-Host ""
    Write-Step "LITE mode - free edition"
    Write-Host ("  Disabled but licensed : {0}" -f $R.Disabled.Count)
    Write-Host ("  Never signed in       : {0}" -f $R.Never.Count)
    Write-Host ("  Inactive > $InactiveDays days     : {0}" -f $R.Inactive.Count)
    Write-Host ("  Unassigned seats      : {0}" -f (($R.Skus | Measure-Object -Property Unassigned -Sum).Sum))
    Write-Host ""
    Write-Host "  Solo/Pro add the HTML report — seatscout.dev" -ForegroundColor Yellow
    Write-Host ""
    Write-Step "CSV saved to: $OutputPath"
    return
}

