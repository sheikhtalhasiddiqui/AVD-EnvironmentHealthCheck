<#
.SYNOPSIS
    AVD-Environment Health v4.0 — Azure Virtual Desktop Comprehensive Health Report.
    Author: Microsoft AVD Architect : Mohammad Talha
    Ownership: Mohammad Talha
.DESCRIPTION
    Collects and reports on all AVD environment parameters across:
      • Overview          (Subscription, Tenant, Resource Group, Host Pools, Session Hosts)
      • Host Pools        (Type, Load Balancer, App Groups, Scaling, Redirections)
      • Session Hosts     (Health, Boot status, AZ distribution, Drain mode, Disk space, Agent)
      • Health Checks     (Domain, Trust, SxS, URLs, Monitoring, Metadata, AppAttach, TURN, AAD)
      • Azure NetApp Files (ANF Account, Capacity Pools, Region colocation, Volume quota usage %,
                            Per-volume: Snapshot Policy, Backup Policy, Backup Vault, Encryption Key Source)

    Produces a self-contained HTML report with a modern terminal-inspired design,
    expandable sections, traffic-light status indicators, and remediation guidance.

.PARAMETER SubscriptionId
    Azure subscription ID. Uses current Az context when omitted.

.PARAMETER TenantId
    Azure tenant ID. Uses current Az context when omitted.

.PARAMETER ResourceGroupName
    Resource group containing AVD resources. No default — must be supplied via
    parameter, -ConfigPath, or the AVDHEALTH_RESOURCEGROUP environment variable.

.PARAMETER OutputPath
    Output path for HTML report. Defaults to a timestamped, client/environment-tagged
    file in current directory.

.PARAMETER HostPoolName
    Scope report to a single host pool.

.PARAMETER UseExistingConnection
    Reuse the existing Az PowerShell context.

.PARAMETER OpenReport
    Open the HTML report in the default browser after generation.
    Ownership - Mohammad Talha
.PARAMETER DryRun
    Generate a report with synthetic data — no Azure API calls are made.

.PARAMETER ANFResourceGroupName
    Resource group containing Azure NetApp Files accounts.
    If empty, discovery is performed subscription-wide.
    Only used when ProfileStorageType is 'ANF' or 'Both'.

.PARAMETER ProfileStorageType
    Which FSLogix profile storage backend to assess: 'ANF', 'AzureFiles', or 'Both'.
    Default: ANF (preserves prior behavior). This determines which storage section
    of the report is populated — the backend varies by client/environment and must
    never be assumed; confirm it before running against a live environment.

.PARAMETER StorageResourceGroupName
    Resource group containing Azure Files storage accounts used for FSLogix profiles.
    If empty, discovery is performed subscription-wide. Only used when
    ProfileStorageType is 'AzureFiles' or 'Both'.

.PARAMETER StorageAccountName
    Scope Azure Files assessment to one storage account. If empty, all storage
    accounts in StorageResourceGroupName (or the subscription) are assessed.

.PARAMETER QuotaWarningPercent
    Usage percentage at which an Azure Files share is flagged WARN/near-quota
    (mirrors the 80% threshold already used for ANF volumes). Default: 80.

.PARAMETER ConfigPath
    Path to a JSON configuration file (see config.template.json) supplying
    ClientName, EnvironmentName, ResourceGroupName, ANFResourceGroupName,
    SubscriptionId, TenantId, DefaultRegion, BrandDisplayName and BrandColor.
    Explicit command-line parameters always take precedence over the config file.
    Config file values take precedence over environment variables.

.PARAMETER ClientName
    Display name of the client/tenant this report is generated for. Used in the
    report header and output file name. No default — must be supplied.

.PARAMETER EnvironmentName
    Logical environment being assessed: Development, Test, UAT, Staging, or
    Production. Used in the report header and output file name. Default: Production.

.PARAMETER DefaultRegion
    Fallback Azure region used only when a region cannot be derived from a live
    resource (e.g., Marketplace lookups when no session hosts exist yet). Default: eastus.

.PARAMETER BrandDisplayName
    Optional short text/wordmark rendered in the report header (e.g., client or
    business unit name). Leave blank to omit the brand block entirely.

.PARAMETER BrandColor
    Optional hex color (e.g., '#1A5BA6') applied to the BrandDisplayName wordmark.
    Defaults to a neutral theme color when not supplied.

.EXAMPLE
    .\AVD-EnvironmentHealth.Generic.ps1 -ConfigPath .\config.client-a.json -UseExistingConnection
    Run using a per-client configuration file against the current Az context.

.EXAMPLE
    .\AVD-EnvironmentHealth.Generic.ps1 -ClientName 'Contoso' -EnvironmentName 'Production' `
        -ResourceGroupName 'rg-avd-prod' -ANFResourceGroupName 'rg-anf-prod' -UseExistingConnection
    Run with explicit parameters, no config file.

.EXAMPLE
    .\AVD-EnvironmentHealth.Generic.ps1 -DryRun -OpenReport -ClientName 'Demo'
    Preview the report with synthetic data.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath          = '',
    [string]$SubscriptionId      = '',
    [string]$TenantId            = '',
    # AVD ResourceGroupNames — accepts ONE name or a comma-separated list of names
    # (or an array when called from PowerShell directly).
    # In the JSON config file, set as an array: "resourceGroupNames": ["rg-avd-1","rg-avd-2"]
    # Legacy single-value key "resourceGroupName" is also supported for backward compatibility.
    [string[]]$ResourceGroupNames = @(),
    [string[]]$ANFResourceGroupNames = @(),
    [ValidateSet('', 'ANF', 'AzureFiles', 'Both')]
    [string]$ProfileStorageType  = '',
    # Azure Files: accepts one or more resource groups and/or specific account names
    [string[]]$StorageResourceGroupNames = @(),
    [string[]]$StorageAccountNames  = @(),
    [int]   $QuotaWarningPercent = -1,
    [string]$ClientName          = '',
    [string]$EnvironmentName     = '',
    [string]$DefaultRegion       = '',
    [string]$BrandDisplayName    = '',
    [string]$BrandColor          = '',
    [string]$OutputPath          = '',
    [string]$HostPoolName        = '',
    [switch]$UseExistingConnection,
    [switch]$OpenReport,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
$script:ToolVersion    = '4.0.0'
$script:AccountName    = 'AVD-Health'

# ==============================================================================
# CONFIGURATION RESOLUTION
# Precedence (highest to lowest): explicit -Parameter > -ConfigPath file >
# environment variable > built-in safe default. No client-identifying value
# ships with a hardcoded default in this script.
# ==============================================================================

function Resolve-Setting {
    param(
        [string]$ParamValue,
        [object]$ConfigObject,
        [string]$ConfigProperty,
        [string]$EnvVarName,
        [string]$DefaultValue = ''
    )
    if (-not [string]::IsNullOrWhiteSpace($ParamValue)) { return $ParamValue }
    if ($ConfigObject -and $ConfigObject.PSObject.Properties.Name -contains $ConfigProperty) {
        $cv = $ConfigObject.$ConfigProperty
        if (-not [string]::IsNullOrWhiteSpace($cv)) { return $cv }
    }
    $ev = [System.Environment]::GetEnvironmentVariable($EnvVarName)
    if (-not [string]::IsNullOrWhiteSpace($ev)) { return $ev }
    return $DefaultValue
}

$script:cfg = $null
if ($ConfigPath) {
    if (-not (Test-Path $ConfigPath)) { throw "ConfigPath '$ConfigPath' was not found." }
    try { $script:cfg = Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json }
    catch { throw "Failed to parse config file '$ConfigPath': $($_.Exception.Message)" }
}

# Helper: resolve a multi-value array setting.
# Config JSON: array  ["rg-avd-1","rg-avd-2"]  OR legacy single string "rg-avd-1"
# CLI param  : already [string[]]  OR a comma-separated string element, e.g. "rg-a,rg-b"
# Env var    : comma-separated string
function Resolve-ArraySetting {
    param([string[]]$ParamValue, [object]$ConfigObject,
          [string]$ConfigProperty, [string]$LegacyConfigProperty,
          [string]$EnvVarName)
    # 1 — explicit CLI param (non-empty)
    $flat = @($ParamValue | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ','
    if ($flat) {
        return @($flat -split '\s*,\s*' | Where-Object { $_ -ne '' })
    }
    # 2 — config file (array or string)
    if ($ConfigObject) {
        foreach ($prop in @($ConfigProperty, $LegacyConfigProperty)) {
            if ($prop -and $ConfigObject.PSObject.Properties.Name -contains $prop) {
                $cv = $ConfigObject.$prop
                if ($cv -is [System.Array] -and $cv.Count -gt 0) { return @($cv) }
                if ($cv -is [string] -and -not [string]::IsNullOrWhiteSpace($cv)) {
                    return @($cv -split '\s*,\s*' | Where-Object { $_ -ne '' })
                }
            }
        }
    }
    # 3 — environment variable (comma-separated)
    $ev = [System.Environment]::GetEnvironmentVariable($EnvVarName)
    if (-not [string]::IsNullOrWhiteSpace($ev)) {
        return @($ev -split '\s*,\s*' | Where-Object { $_ -ne '' })
    }
    return @()
}

$ClientName            = Resolve-Setting -ParamValue $ClientName       -ConfigObject $script:cfg -ConfigProperty 'clientName'      -EnvVarName 'AVDHEALTH_CLIENTNAME'    -DefaultValue ''
$EnvironmentName       = Resolve-Setting -ParamValue $EnvironmentName  -ConfigObject $script:cfg -ConfigProperty 'environmentName' -EnvVarName 'AVDHEALTH_ENVIRONMENT'    -DefaultValue 'Production'
$SubscriptionId        = Resolve-Setting -ParamValue $SubscriptionId   -ConfigObject $script:cfg -ConfigProperty 'subscriptionId'  -EnvVarName 'AVDHEALTH_SUBSCRIPTIONID' -DefaultValue ''
$TenantId              = Resolve-Setting -ParamValue $TenantId         -ConfigObject $script:cfg -ConfigProperty 'tenantId'        -EnvVarName 'AVDHEALTH_TENANTID'       -DefaultValue ''

# Multi-value fields — arrays in config, comma-list in env/CLI
[string[]]$ResourceGroupNames = Resolve-ArraySetting `
    -ParamValue $ResourceGroupNames -ConfigObject $script:cfg `
    -ConfigProperty 'resourceGroupNames' -LegacyConfigProperty 'resourceGroupName' `
    -EnvVarName 'AVDHEALTH_RESOURCEGROUPS'

[string[]]$ANFResourceGroupNames = Resolve-ArraySetting `
    -ParamValue $ANFResourceGroupNames -ConfigObject $script:cfg `
    -ConfigProperty 'anfResourceGroupNames' -LegacyConfigProperty 'anfResourceGroupName' `
    -EnvVarName 'AVDHEALTH_ANF_RESOURCEGROUPS'

[string[]]$StorageResourceGroupNames = Resolve-ArraySetting `
    -ParamValue $StorageResourceGroupNames -ConfigObject $script:cfg `
    -ConfigProperty 'storageResourceGroupNames' -LegacyConfigProperty 'storageResourceGroupName' `
    -EnvVarName 'AVDHEALTH_STORAGE_RESOURCEGROUPS'

[string[]]$StorageAccountNames = Resolve-ArraySetting `
    -ParamValue $StorageAccountNames -ConfigObject $script:cfg `
    -ConfigProperty 'storageAccountNames' -LegacyConfigProperty 'storageAccountName' `
    -EnvVarName 'AVDHEALTH_STORAGE_ACCOUNTS'

$ProfileStorageType    = Resolve-Setting -ParamValue $ProfileStorageType  -ConfigObject $script:cfg -ConfigProperty 'profileStorageType'   -EnvVarName 'AVDHEALTH_STORAGE_TYPE'      -DefaultValue 'ANF'
if ($QuotaWarningPercent -gt 0) { $qwpParam = [string]$QuotaWarningPercent } else { $qwpParam = '' }
$qwpRaw                = Resolve-Setting -ParamValue $qwpParam -ConfigObject $script:cfg -ConfigProperty 'quotaWarningPercent' -EnvVarName 'AVDHEALTH_QUOTA_WARN_PCT' -DefaultValue '80'
$QuotaWarningPercent   = [int]$qwpRaw
$DefaultRegion         = Resolve-Setting -ParamValue $DefaultRegion    -ConfigObject $script:cfg -ConfigProperty 'defaultRegion'   -EnvVarName 'AVDHEALTH_DEFAULT_REGION' -DefaultValue 'eastus'
$BrandDisplayName      = Resolve-Setting -ParamValue $BrandDisplayName -ConfigObject $script:cfg -ConfigProperty 'brandDisplayName' -EnvVarName 'AVDHEALTH_BRAND_NAME'    -DefaultValue ''
$BrandColor            = Resolve-Setting -ParamValue $BrandColor       -ConfigObject $script:cfg -ConfigProperty 'brandColor'       -EnvVarName 'AVDHEALTH_BRAND_COLOR'   -DefaultValue '#3B82F6'

# Build primary display RG label (first RG or subscription-wide)
$script:PrimaryRgLabel = if ($ResourceGroupNames.Count -gt 0) { $ResourceGroupNames -join ', ' } else { 'Subscription-Wide' }

if (-not $DryRun -and $ResourceGroupNames.Count -eq 0) {
    throw "At least one ResourceGroupName is required. Provide -ResourceGroupNames 'rg1','rg2', set 'resourceGroupNames' as an array in -ConfigPath, or set AVDHEALTH_RESOURCEGROUPS as a comma-separated list. (Use -DryRun to preview without a live environment.)"
}
if ([string]::IsNullOrWhiteSpace($ClientName)) { $ClientName = 'UnnamedClient' }

$script:Overview = @{}
$script:Overview['ClientName']      = $ClientName
$script:Overview['EnvironmentName'] = $EnvironmentName

$script:RequiredModules = @(
    'Az.Accounts',
    'Az.DesktopVirtualization',
    'Az.Compute',
    'Az.Monitor',
    'Az.Resources',
    'Az.Network',
    'Az.Storage',
    'Az.Security',
    'Az.NetAppFiles'
)

# ── Data containers ──────────────────────────────────────────────────────────
# NOTE: $script:Overview is initialized above (during configuration resolution)
# and is intentionally NOT reset here, so ClientName/EnvironmentName survive.
$script:HostPoolData = [System.Collections.Generic.List[object]]::new()
$script:SessionHostData = [System.Collections.Generic.List[object]]::new()
$script:HealthCheckData = [System.Collections.Generic.List[object]]::new()
$script:AnfData      = [System.Collections.Generic.List[object]]::new()
$script:ReportSections = [System.Collections.Generic.List[object]]::new()

# ==============================================================================
# HELPERS
# ==============================================================================

function Write-Step  { param([string]$Msg) Write-Host "  >> $Msg" -ForegroundColor Cyan }
function Write-Ok    { param([string]$Msg) Write-Host "     [OK] $Msg" -ForegroundColor Green }
function Write-Warn  { param([string]$Msg) Write-Host "     [!!] $Msg" -ForegroundColor Yellow }
function Write-Warn2 { param([string]$Msg) Write-Host "     [!!] $Msg" -ForegroundColor Yellow }   # alias — same output as Write-Warn
function Write-Err   { param([string]$Msg) Write-Host "     [XX] $Msg" -ForegroundColor Red }

function ConvertTo-HtmlSafe {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $Text -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
}

function Get-RgFromArmId {
    param([string]$ResourceId)
    ($ResourceId -split '/')[4]
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [int]$MaxAttempts = 3,
        [string]$OperationName = 'Azure op'
    )
    $attempt = 0; $delay = 2
    while ($true) {
        $attempt++
        try { return & $ScriptBlock }
        catch {
            $msg = $_.Exception.Message
            $transient = $msg -match '429|throttl|timeout|503|ServiceUnavailable'
            if (-not $transient -or $attempt -ge $MaxAttempts) { throw }
            Start-Sleep -Seconds $delay; $delay = [math]::Min($delay * 2, 30)
        }
    }
}

function Get-RdpProperty {
    param([string]$RdpString, [string]$PropertyName)
    if ([string]::IsNullOrEmpty($RdpString)) { return $null }
    $match = $RdpString -split ';' | Where-Object { $_ -match "^$([regex]::Escape($PropertyName)):" }
    if ($match) { ($match -split ':', 3)[2] } else { $null }
}

function Get-StatusBadge {
    param([string]$Status)
    switch ($Status) {
        'PASS'    { 'badge-pass' }
        'WARN'    { 'badge-warn' }
        'FAIL'    { 'badge-fail' }
        'INFO'    { 'badge-info' }
        'ENABLED' { 'badge-pass' }
        'DISABLED'{ 'badge-fail' }
        default   { 'badge-info' }
    }
}

# ==============================================================================
# AZURE CONNECTION
# ==============================================================================

function Connect-ToAzure {
    Write-Step 'Connecting to Azure'

    $ctx = Get-AzContext -ErrorAction SilentlyContinue
    if ($UseExistingConnection -and $ctx -and $ctx.Account) {
        Write-Ok "Reusing session: $($ctx.Account.Id)"
    } else {
        Connect-AzAccount -ErrorAction Stop | Out-Null
    }

    if ($SubscriptionId) {
        Set-AzContext -Subscription $SubscriptionId -ErrorAction Stop | Out-Null
    }

    $ctx = Get-AzContext
    $script:Overview['Subscription']   = $ctx.Subscription.Name
    $script:Overview['SubscriptionID'] = $ctx.Subscription.Id
    $script:Overview['Tenant']         = $ctx.Tenant.Id
    $script:Overview['ResourceGroup']  = $script:PrimaryRgLabel
    $script:Overview['ReportGeneratedAt'] = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz')
    Write-Ok "Connected: $($ctx.Subscription.Name)"
}

# ==============================================================================
# DATA COLLECTION — OVERVIEW
# ==============================================================================

function Get-OverviewData {
    Write-Step 'Collecting Overview data'

    try {
        $hps = @()
        foreach ($rg in $ResourceGroupNames) {
            $hps += @(Invoke-WithRetry -OperationName "HostPools:$rg" -ScriptBlock {
                if ($HostPoolName) {
                    Get-AzWvdHostPool -ResourceGroupName $rg -Name $HostPoolName -ErrorAction Stop
                } else {
                    Get-AzWvdHostPool -ResourceGroupName $rg -ErrorAction Stop
                }
            })
        }
        $script:allHostPools = $hps
        $script:Overview['TotalHostPools'] = $hps.Count
        Write-Ok "Host pools: $($hps.Count) across $($ResourceGroupNames.Count) resource group(s)"
    } catch {
        $script:allHostPools = @()
        $script:Overview['TotalHostPools'] = 0
        Write-Warn "Could not fetch host pools: $($_.Exception.Message)"
    }

    # Session hosts
    $shAll = [System.Collections.Generic.List[object]]::new()
    foreach ($hp in $script:allHostPools) {
        $hpRg = Get-RgFromArmId $hp.Id
        try {
            $hosts = @(Invoke-WithRetry -OperationName "SH:$($hp.Name)" -ScriptBlock {
                Get-AzWvdSessionHost -ResourceGroupName $hpRg -HostPoolName $hp.Name -ErrorAction Stop
            })
            foreach ($h in $hosts) {
                $h | Add-Member -Force -NotePropertyName '_HostPoolName' -NotePropertyValue $hp.Name
                $h | Add-Member -Force -NotePropertyName '_HostPoolId'   -NotePropertyValue $hp.Id
                $h | Add-Member -Force -NotePropertyName '_HostPoolType' -NotePropertyValue $hp.HostPoolType
                $shAll.Add($h)
            }
        } catch { Write-Warn "SH fetch failed for $($hp.Name)" }
    }
    $script:allSessionHosts = $shAll.ToArray()
    $script:Overview['TotalSessionHosts'] = $script:allSessionHosts.Count
    Write-Ok "Session hosts: $($script:allSessionHosts.Count)"
}

# ==============================================================================
# DATA COLLECTION — HOST POOLS
# ==============================================================================

function Get-HostPoolData {
    Write-Step 'Collecting Host Pool data'

    # Scaling plans
    $spSet = [System.Collections.Generic.Dictionary[string,object]]::new([System.StringComparer]::OrdinalIgnoreCase)
    try {
        $subWide = @(Invoke-WithRetry -ScriptBlock { Get-AzWvdScalingPlan -ErrorAction Stop })
        foreach ($sp in $subWide) { $spSet[$sp.Id] = $sp }
    } catch { Write-Warn 'Subscription-wide scaling plan fetch failed' }

    # App groups
    $appGroupMap = @{}
    try {
        $ags = @()
        foreach ($rg in $ResourceGroupNames) {
            $ags += @(Invoke-WithRetry -ScriptBlock {
                Get-AzWvdApplicationGroup -ResourceGroupName $rg -ErrorAction Stop
            })
        }
        foreach ($ag in $ags) {
            $hpRef = $ag.HostPoolArmPath
            if ($hpRef) {
                if (-not $appGroupMap[$hpRef]) { $appGroupMap[$hpRef] = [System.Collections.Generic.List[string]]::new() }
                $appGroupMap[$hpRef].Add($ag.Name)
            }
        }
    } catch { Write-Warn 'App group fetch failed' }

    # Scaling plan coverage map
    $refIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($sp in $spSet.Values) {
        foreach ($ref in @($sp.HostPoolReference)) {
            if ($ref -and $ref.HostPoolArmPath) { [void]$refIds.Add($ref.HostPoolArmPath.Trim().TrimEnd('/')) }
        }
    }

    foreach ($hp in $script:allHostPools) {
        $hpId   = $hp.Id.Trim().TrimEnd('/')
        $hpRg   = Get-RgFromArmId $hp.Id
        $hasSP  = $refIds.Contains($hpId)
        $ags    = @($appGroupMap[$hp.Id])

        # RDP redirect properties
        $rdp = $hp.CustomRdpProperty
        $driveRedir     = Get-RdpProperty $rdp 'drivestoredirect'
        $clipRedir      = Get-RdpProperty $rdp 'redirectclipboard'
        $printerRedir   = Get-RdpProperty $rdp 'redirectprinters'
        $usbRedir       = Get-RdpProperty $rdp 'usbdevicestoredirect'

        $driveStatus   = if ($null -eq $driveRedir -or $driveRedir -eq '*') { 'ENABLED' } else { 'DISABLED' }
        $clipStatus    = if ($null -eq $clipRedir  -or $clipRedir  -eq '1') { 'ENABLED' } else { 'DISABLED' }
        $printerStatus = if ($null -eq $printerRedir -or $printerRedir -eq '1') { 'ENABLED' } else { 'DISABLED' }
        $usbStatus     = if ($null -eq $usbRedir   -or $usbRedir   -eq '*') { 'ENABLED' } else { 'DISABLED' }

        $script:HostPoolData.Add([PSCustomObject]@{
            Name              = $hp.Name
            Type              = $hp.HostPoolType
            LoadBalancer      = $hp.LoadBalancerType
            MaxSessionLimit   = $hp.MaxSessionLimit
            ValidationEnv     = $hp.ValidationEnvironment
            StartVMOnConnect  = $hp.StartVMOnConnect
            ScalingPlan       = if ($hasSP) { 'Assigned' } else { 'None' }
            AppGroupCount     = $ags.Count
            AppGroups         = ($ags -join ', ')
            DriveRedirect     = $driveStatus
            ClipboardRedirect = $clipStatus
            PrinterRedirect   = $printerStatus
            UsbRedirect       = $usbStatus
            PublicNetwork     = if ($hp.PublicNetworkAccess) { $hp.PublicNetworkAccess } else { 'Unknown' }
        })
    }
    Write-Ok "Processed $($script:HostPoolData.Count) host pool(s)"
}

# ==============================================================================
# DATA COLLECTION — SESSION HOSTS
# ==============================================================================

function Get-SessionHostData {
    Write-Step 'Collecting Session Host data'

    # VMs for detailed data — fetch WITHOUT -Status to get full model (StorageProfile, Extensions, Zones)
    $vmMap  = @{}
    $nicMap = @{}
    $diskMap = @{}
    $extMap  = @{}   # AzureMonitorWindowsAgent full details per VM

    # ── Fetch latest available AMA version from Azure Marketplace (once) ─────
    $latestAmaVersion = 'Unknown'
    try {
        $loc = if ($script:allSessionHosts.Count -gt 0 -and $script:allSessionHosts[0].ResourceId) {
            $firstVmRg = Get-RgFromArmId $script:allSessionHosts[0].ResourceId
            $firstVmName = ($script:allSessionHosts[0].ResourceId -split '/')[-1]
            try { (Invoke-WithRetry -ScriptBlock { Get-AzVM -ResourceGroupName $firstVmRg -Name $firstVmName -ErrorAction Stop }).Location } catch { $DefaultRegion }
        } else { $DefaultRegion }
        $images = @(Get-AzVMExtensionImage -Location $loc -PublisherName 'Microsoft.Azure.Monitor' `
                        -Type 'AzureMonitorWindowsAgent' -ErrorAction Stop)
        if ($images.Count -gt 0) {
            $sorted = $images | Sort-Object { try { [version]$_.Version } catch { [version]'0.0' } } -Descending
            $rawLatest = $sorted[0].Version
            # Normalise to major.minor — the installed extension only reports major.minor,
            # so we compare on the same 2-part form to avoid false "Update Available" badges.
            try {
                $v = [version]$rawLatest
                $latestAmaVersion = "$($v.Major).$($v.Minor)"
            } catch {
                $latestAmaVersion = $rawLatest
            }
        }
        Write-Ok "Latest AMA version from Marketplace: $latestAmaVersion (normalised to major.minor)"
    } catch {
        Write-Warn "Could not fetch latest AMA version from Marketplace — comparison will use 'Unknown'"
    }
    foreach ($sh in $script:allSessionHosts) {
        if (-not $sh.ResourceId) { continue }
        $parts = $sh.ResourceId -split '/'
        if ($parts.Count -lt 9) { continue }
        $vmName = $parts[-1]
        $vmRg   = Get-RgFromArmId $sh.ResourceId
        try {
            $vm = Invoke-WithRetry -ScriptBlock {
                Get-AzVM -ResourceGroupName $vmRg -Name $vmName -ErrorAction Stop
            }
            $vmMap[$vmName] = $vm

            # Pre-fetch OS disk SKU while we have the VM model
            if ($vm.StorageProfile.OsDisk.ManagedDisk.Id) {
                try {
                    $diskId   = $vm.StorageProfile.OsDisk.ManagedDisk.Id
                    $diskObj  = Invoke-WithRetry -ScriptBlock {
                        Get-AzDisk -ResourceGroupName (Get-RgFromArmId $diskId) -DiskName ($diskId -split '/')[-1] -ErrorAction Stop
                    }
                    $diskMap[$vmName] = $diskObj.Sku.Name
                } catch { }
            }
            # Fallback: try by OS disk name directly
            if (-not $diskMap[$vmName] -and $vm.StorageProfile.OsDisk.Name) {
                try {
                    $diskObj = Invoke-WithRetry -ScriptBlock {
                        Get-AzDisk -ResourceGroupName $vmRg -DiskName $vm.StorageProfile.OsDisk.Name -ErrorAction Stop
                    }
                    $diskMap[$vmName] = $diskObj.Sku.Name
                } catch { }
            }

            # Pre-fetch AzureMonitorWindowsAgent extension — all fields
            # (Settings > Extensions + applications in portal)
            # NOTE: Extension details can only be queried when the VM is running
            $shObj = $script:allSessionHosts | Where-Object { $_.ResourceId -and ($_.ResourceId -split '/')[-1] -eq $vmName } | Select-Object -First 1
            $shStatus = if ($shObj) { $shObj.Status } else { 'Unknown' }
            if ($shStatus -eq 'Available' -or $shStatus -eq 'NeedsAssistance' -or $shStatus -eq 'Unavailable') {
                try {
                    $amaExt = Invoke-WithRetry -OperationName "Ext:$vmName" -ScriptBlock {
                        Get-AzVMExtension -ResourceGroupName $vmRg -VMName $vmName `
                            -Name 'AzureMonitorWindowsAgent' -ErrorAction Stop
                    }
                    if ($amaExt) {
                        $extMap[$vmName] = [PSCustomObject]@{
                            Name               = 'AzureMonitorWindowsAgent'
                            Publisher          = if ($amaExt.Publisher) { $amaExt.Publisher } else { 'Microsoft.Azure.Monitor' }
                            ExtensionType      = if ($amaExt.ExtensionType) { $amaExt.ExtensionType } else { 'AzureMonitorWindowsAgent' }
                            Version            = if ($amaExt.TypeHandlerVersion) { $amaExt.TypeHandlerVersion } else { 'Unknown' }
                            ProvisioningState  = if ($amaExt.ProvisioningState) { $amaExt.ProvisioningState } else { 'Unknown' }
                            AutoUpgradeEnabled = if ($null -ne $amaExt.EnableAutomaticUpgrade) { $amaExt.EnableAutomaticUpgrade } else { $false }
                        }
                    }
                } catch {
                    # Extension not installed on this VM — that's OK
                }
            }

            # NIC
            $nicRef = $vm.NetworkProfile.NetworkInterfaces | Select-Object -First 1
            if ($nicRef) {
                try {
                    $nic = Invoke-WithRetry -ScriptBlock {
                        Get-AzNetworkInterface -ResourceId $nicRef.Id -ErrorAction Stop
                    }
                    $nicMap[$vmName] = $nic
                } catch { }
            }
        } catch { Write-Warn "VM fetch failed: $vmName" }
    }

    foreach ($sh in $script:allSessionHosts) {
        $shName  = ($sh.Name -split '/')[-1]
        $vm      = if ($sh.ResourceId) { $vmMap[($sh.ResourceId -split '/')[-1]] } else { $null }

        # Boot status: consider "not booted > 2 days" if shutdown + AllowNewSession false for >2 days
        # We approximate this from status + last heartbeat
        $notBooted2Days = $false
        if ($sh.LastHeartBeat) {
            $age = (Get-Date) - $sh.LastHeartBeat
            if ($age.TotalDays -gt 2 -and $sh.Status -ne 'Available') { $notBooted2Days = $true }
        }

        # Availability Zone from VM
        $az = 'N/A'
        if ($vm -and $vm.Zones -and $vm.Zones.Count -gt 0) { $az = $vm.Zones -join ',' }

        # OS Disk SKU — use pre-fetched disk map from the VM collection phase
        $vmKey   = if ($sh.ResourceId) { ($sh.ResourceId -split '/')[-1] } else { '' }
        $diskSku = if ($diskMap[$vmKey]) { $diskMap[$vmKey] } else { 'Unknown' }

        # AzureMonitorWindowsAgent full details — from pre-fetched Get-AzVMExtension data
        $amaDetails = $extMap[$vmKey]
        $isShutdown = ($sh.Status -eq 'Shutdown' -or $sh.Status -eq 'Deallocated')
        $amaLatestVer    = $latestAmaVersion
        $amaUpdateStatus = 'Unknown'

        # ── Version normalisation helper ──────────────────────────────────────
        # The VM extension (TypeHandlerVersion) reports major.minor only,   e.g. "1.43"
        # The Marketplace (Get-AzVMExtensionImage) reports full 4-part ver, e.g. "1.43.0.2"
        # We compare on major.minor (first two components) so that "1.43" == "1.43.0.2".
        # Both installed and latest are displayed as major.minor in the report to keep
        # the comparison unambiguous for the engineer reading it.
        function Get-MajorMinor {
            param([string]$VersionString)
            if ([string]::IsNullOrWhiteSpace($VersionString) -or $VersionString -eq 'Unknown') { return $null }
            try {
                $v = [version]$VersionString
                return [version]"$($v.Major).$($v.Minor)"
            } catch { return $null }
        }

        if ($amaDetails) {
            $monitorAgentVersion = $amaDetails.Version          # raw from VM extension, e.g. "1.43"
            $amaPublisher        = $amaDetails.Publisher
            $amaExtType          = $amaDetails.ExtensionType
            $amaProvState        = $amaDetails.ProvisioningState
            $amaAutoUpgrade      = if ($amaDetails.AutoUpgradeEnabled) { 'Enabled' } else { 'Disabled' }

            # Normalise installed version to major.minor for display
            $instMM = Get-MajorMinor $monitorAgentVersion
            $monitorAgentVersionDisplay = if ($instMM) { "$($instMM.Major).$($instMM.Minor)" } else { $monitorAgentVersion }

            # Compare installed vs latest on major.minor only
            if ($latestAmaVersion -ne 'Unknown') {
                $mktMM = Get-MajorMinor $latestAmaVersion
                $latestVersionDisplay = if ($mktMM) { "$($mktMM.Major).$($mktMM.Minor)" } else { $latestAmaVersion }
                $amaLatestVer = $latestVersionDisplay   # show normalised form in report

                if ($instMM -and $mktMM) {
                    if ($instMM -ge $mktMM) {
                        $amaUpdateStatus    = 'Up to Date'
                        $monitorAgentLatest = $true
                    } else {
                        $amaUpdateStatus    = "Update Available ($latestVersionDisplay)"
                        $monitorAgentLatest = $false
                    }
                } else {
                    $amaUpdateStatus    = 'Version Parse Error'
                    $monitorAgentLatest = $false
                }
            } else {
                $latestVersionDisplay   = 'Unknown'
                $monitorAgentLatest     = $true   # Can't compare — don't flag
                $amaUpdateStatus        = 'Latest Unknown'
            }
            # Use normalised display version going forward
            $monitorAgentVersion = $monitorAgentVersionDisplay
        } elseif ($isShutdown) {
            $monitorAgentVersion = 'VM Shutdown'
            $amaPublisher       = 'N/A - VM Shutdown'
            $amaExtType         = 'N/A'
            $amaProvState       = 'N/A - VM Shutdown'
            $amaAutoUpgrade     = 'N/A'
            $amaUpdateStatus    = 'N/A - VM Shutdown'
            $monitorAgentLatest = $true  # Don't flag as outdated for shutdown VMs
        } else {
            $monitorAgentVersion = 'Not Detected'
            $amaPublisher       = 'N/A'
            $amaExtType         = 'N/A'
            $amaProvState       = 'N/A'
            $amaAutoUpgrade     = 'N/A'
            $amaUpdateStatus    = 'Extension Not Installed'
            $monitorAgentLatest = $false
        }

        $healthStatus = switch ($sh.Status) {
            'Available'       { 'PASS' }
            'Shutdown'        { 'INFO' }
            'Unavailable'     { 'FAIL' }
            'NeedsAssistance' { 'WARN' }
            'UpgradeFailed'   { 'FAIL' }
            'NoHeartbeat'     { 'FAIL' }
            default           { 'WARN' }
        }

        $script:SessionHostData.Add([PSCustomObject]@{
            Name                = $shName
            HostPool            = $sh._HostPoolName
            Status              = $sh.Status
            HealthStatus        = $healthStatus
            AllowNewSession     = $sh.AllowNewSession
            DrainMode           = -not $sh.AllowNewSession
            Sessions            = $sh.Session
            LastHeartbeat       = $sh.LastHeartBeat
            NotBooted2Days      = $notBooted2Days
            AvailabilityZone    = $az
            AgentVersion        = $sh.AgentVersion
            MonitorAgentVersion = $monitorAgentVersion
            MonitorAgentLatest  = $monitorAgentLatest
            AMAPublisher        = $amaPublisher
            AMAExtensionType    = $amaExtType
            AMAProvisioningState = $amaProvState
            AMAAutoUpgrade      = $amaAutoUpgrade
            AMALatestVersion    = $amaLatestVer
            AMAUpdateStatus     = $amaUpdateStatus
            OsDiskSku           = $diskSku
            Gen                 = if ($vm -and $vm.StorageProfile.OsDisk.OsType) {
                                      try { $disk = $vmMap[($sh.ResourceId -split '/')[-1]]
                                            if ($disk.StorageProfile.ImageReference.ExactVersion -match 'V2') { 'Gen2' } else { 'Unknown' }
                                      } catch { 'Unknown' }
                                  } else { 'Unknown' }
            ResourceId          = $sh.ResourceId
        })
    }
    Write-Ok "Processed $($script:SessionHostData.Count) session host(s)"
}

# ==============================================================================
# DATA COLLECTION — HEALTH CHECKS
# ==============================================================================

function Get-HealthCheckData {
    Write-Step 'Collecting Session Host Health Check data'

    # AVD surfaces per-host health check results ONLY when the session host is
    # fetched with -ExpandAll. The standard collection call does not populate
    # HealthCheckResult. We therefore re-fetch each host individually here with
    # -ExpandAll to get real data — exactly what the Azure portal "Session Host
    # Status Details" JSON view shows (HealthCheckName / HealthCheckResult /
    # AdditionalFailureDetails.Message / ErrorCode / LastHealthCheckDateTime).

    $knownChecks = @(
        'DomainJoinedCheck',
        'DomainTrustCheck',
        'SxSStackListenerCheck',
        'UrlsAccessibleCheck',
        'MonitoringAgentCheck',
        'MetaDataServiceCheck',
        'AppAttachHealthCheck',
        'TURNRelayAccessHealthCheck',
        'AADJoinedHealthCheck'
    )

    $checkMap          = @{}
    foreach ($k in $knownChecks) { $checkMap[$k] = [System.Collections.Generic.List[PSCustomObject]]::new() }
    $hostsWithRealData = 0
    $hostsEstimated    = 0

    foreach ($sh in $script:allSessionHosts) {
        $shName  = ($sh.Name -split '/')[-1]
        $hpName  = $sh._HostPoolName
        $hpRg    = Get-RgFromArmId $sh._HostPoolId

        # Re-fetch with -ExpandAll to populate HealthCheckResult array
        $hcResults = $null
        try {
            $shEx = Invoke-WithRetry -OperationName "HC:$shName" -ScriptBlock {
                Get-AzWvdSessionHost -ResourceGroupName $hpRg `
                    -HostPoolName $hpName -Name $shName `
                    -ExpandAll -ErrorAction Stop
            }
            if ($shEx) {
                if ($shEx.PSObject.Properties.Name -contains 'HealthCheckResult' -and $shEx.HealthCheckResult) {
                    $hcResults = $shEx.HealthCheckResult
                } elseif ($shEx.PSObject.Properties.Name -contains 'SessionHostHealthCheckResult' -and $shEx.SessionHostHealthCheckResult) {
                    $hcResults = $shEx.SessionHostHealthCheckResult
                }
            }
        } catch {
            Write-Warn2 "ExpandAll failed for $shName — using status-based estimation"
        }

        if ($hcResults -and $hcResults.Count -gt 0) {
            # ── REAL DATA from API ──────────────────────────────────────────
            $hostsWithRealData++
            foreach ($hc in $hcResults) {
                $name   = if ($hc.HealthCheckName)   { $hc.HealthCheckName }   else { "$($hc.Name)" }
                $result = if ($hc.HealthCheckResult) { $hc.HealthCheckResult } else { "$($hc.Result)" }
                $failed = ($result -ne 'HealthCheckSucceeded') -and ($result -ne 'Unknown')

                $det    = if ($hc.PSObject.Properties['AdditionalFailureDetails']) { $hc.AdditionalFailureDetails } else { $null }
                $msg    = if ($det -and $det.PSObject.Properties['Message']   -and $det.Message)             { $det.Message   } else { '' }
                $code   = if ($det -and $det.PSObject.Properties['ErrorCode'] -and $det.ErrorCode -ne 0)     { $det.ErrorCode } else { 0 }
                $rawDt  = if ($det -and $det.PSObject.Properties['LastHealthCheckDateTime'] -and $det.LastHealthCheckDateTime) { $det.LastHealthCheckDateTime } `
                          elseif ($hc.PSObject.Properties['LastHealthCheckDateTime'] -and $hc.LastHealthCheckDateTime) { $hc.LastHealthCheckDateTime } else { '' }
                $lastDt = if ($rawDt) { try { ([datetime]$rawDt).ToString('yyyy-MM-dd HH:mm:ss') } catch { "$rawDt" } } else { '' }

                $entry = [PSCustomObject]@{
                    HostName   = $shName; Result = $result; Failed = $failed
                    Message    = $msg; ErrorCode = $code; LastCheckDt = $lastDt
                    DataSource = 'API (ExpandAll)'
                }
                if (-not $checkMap.ContainsKey($name)) {
                    $checkMap[$name] = [System.Collections.Generic.List[PSCustomObject]]::new()
                }
                $checkMap[$name].Add($entry)
            }
            # Fill PASS for any known check not returned by this host
            foreach ($k in $knownChecks) {
                if (-not ($checkMap[$k] | Where-Object { $_.HostName -eq $shName })) {
                    $checkMap[$k].Add([PSCustomObject]@{
                        HostName = $shName; Result = 'HealthCheckSucceeded'; Failed = $false
                        Message = ''; ErrorCode = 0; LastCheckDt = ''; DataSource = 'API (ExpandAll)'
                    })
                }
            }
        } else {
            # ── FALLBACK: ExpandAll returned no data ────────────────────────
            $hostsEstimated++
            $st = $sh.Status; $us = $sh.UpdateState
            $inf = @{
                'DomainJoinedCheck'          = ($st -notin @('Unavailable','NeedsAssistance'))
                'DomainTrustCheck'           = ($us -ne 'Failed')
                'SxSStackListenerCheck'      = ($st -notin @('Unavailable','NeedsAssistance'))
                'UrlsAccessibleCheck'        = ($st -ne 'Unavailable')
                'MonitoringAgentCheck'       = ($us -ne 'Stalled')
                'MetaDataServiceCheck'       = ($st -ne 'NoHeartbeat')
                'AppAttachHealthCheck'       = ($st -ne 'UpgradeFailed')
                'TURNRelayAccessHealthCheck' = ($st -notin @('Unavailable','NeedsAssistance'))
                'AADJoinedHealthCheck'       = $true
            }
            foreach ($k in $knownChecks) {
                $pass = $inf[$k]
                $checkMap[$k].Add([PSCustomObject]@{
                    HostName    = $shName
                    Result      = if ($pass) { 'HealthCheckSucceeded (Estimated)' } else { 'HealthCheckFailed (Estimated)' }
                    Failed      = (-not $pass)
                    Message     = if (-not $pass) { "Estimated from Status=$st / UpdateState=$us. Verify RBAC Reader on host pool." } else { '' }
                    ErrorCode   = 0; LastCheckDt = ''; DataSource = 'Estimated (ExpandAll unavailable)'
                })
            }
        }
    }

    # ── Consolidate per-check across all hosts ────────────────────────────────
    $allCheckNames = ($knownChecks + @($checkMap.Keys | Where-Object { $knownChecks -notcontains $_ })) | Select-Object -Unique

    foreach ($checkName in $allCheckNames) {
        $entries     = $checkMap[$checkName]
        if (-not $entries) { continue }

        $failEntries = @($entries | Where-Object { $_.Failed })
        $passEntries = @($entries | Where-Object { -not $_.Failed })
        $status      = if ($failEntries.Count -eq 0) { 'PASS' } else { 'FAIL' }
        $failHosts   = ($failEntries | ForEach-Object { $_.HostName }) -join ', '
        $hasEst      = ($entries | Where-Object { $_.DataSource -like '*Estimated*' }).Count -gt 0
        $dataSource  = if ($hasEst -and $hostsWithRealData -eq 0)  { 'Estimated (ExpandAll unavailable)' } `
                       elseif ($hasEst) { "API+Est ($hostsWithRealData real / $hostsEstimated est.)" } `
                       else { 'API (ExpandAll)' }

        $failDetail = ($failEntries | ForEach-Object {
            $e = if ($_.ErrorCode -ne 0) { " [Err:$($_.ErrorCode)]" } else { '' }
            $m = if ($_.Message)  { $_.Message -replace '\s+',' ' ; " — $($_.Message.Substring(0,[math]::Min(200,$_.Message.Length)))" } else { '' }
            $d = if ($_.LastCheckDt) { " (Last: $($_.LastCheckDt))" } else { '' }
            "$($_.HostName)$e$m$d"
        }) -join "`n"

        $passDetail = ($passEntries | ForEach-Object {
            $d = if ($_.LastCheckDt) { " ($($_.LastCheckDt))" } else { '' }
            "$($_.HostName)$d"
        }) -join ', '

        $note = switch ($checkName) {
            'AADJoinedHealthCheck' { 'Full result may require Log Analytics / AVD Insights' }
            'AppAttachHealthCheck' { 'Only applicable if MSIX App Attach is configured'     }
            default                { '' }
        }

        $script:HealthCheckData.Add([PSCustomObject]@{
            CheckName    = $checkName
            Status       = $status
            FailingCount = $failEntries.Count
            PassingCount = $passEntries.Count
            TotalCount   = $entries.Count
            FailingHosts = $failHosts
            FailDetail   = $failDetail
            PassDetail   = $passDetail
            PerHostRows  = @($entries | ForEach-Object {
                @{
                    Host      = $_.HostName
                    Result    = $_.Result
                    ErrorCode = $_.ErrorCode
                    Message   = if ($_.Message) { $_.Message -replace '\s+',' ' } else { '' }
                    LastDt    = $_.LastCheckDt
                }
            })
            DataSource   = $dataSource
            Note         = $note
        })
    }
    Write-Ok "Health checks consolidated: $($script:HealthCheckData.Count) check types — $hostsWithRealData real / $hostsEstimated estimated host(s)"
}

# ==============================================================================
# DATA COLLECTION — AZURE NETAPP FILES
# ==============================================================================

function Get-AnfData {
    Write-Step 'Collecting Azure NetApp Files data'

    # Discover ANF accounts — from each configured ANF RG, then fill with subscription-wide scan
    $anfAccounts = [System.Collections.Generic.List[object]]::new()
    $seenIds     = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    try {
        # Explicitly configured ANF resource groups (zero, one, or many)
        foreach ($anfRg in $ANFResourceGroupNames) {
            $accts = @(Invoke-WithRetry -OperationName "ANF Accounts ($anfRg)" -ScriptBlock {
                Get-AzNetAppFilesAccount -ResourceGroupName $anfRg -ErrorAction Stop
            })
            foreach ($a in $accts) {
                if ($seenIds.Add($a.Id)) { $anfAccounts.Add($a) }
            }
        }
        # Supplement with subscription-wide resource graph scan to catch accounts in other RGs
        $subWide = @(Invoke-WithRetry -OperationName 'ANF Accounts (Sub)' -ScriptBlock {
            Get-AzResource -ResourceType 'Microsoft.NetApp/netAppAccounts' -ErrorAction Stop
        })
        foreach ($r in $subWide) {
            if (-not $seenIds.Contains($r.ResourceId)) {
                try {
                    $acct = Invoke-WithRetry -ScriptBlock {
                        Get-AzNetAppFilesAccount -ResourceGroupName $r.ResourceGroupName -Name $r.Name -ErrorAction Stop
                    }
                    if ($seenIds.Add($acct.Id)) { $anfAccounts.Add($acct) }
                } catch { }
            }
        }
        Write-Ok "Found $($anfAccounts.Count) ANF account(s) across $([math]::Max(1,$ANFResourceGroupNames.Count)) configured RG(s)"
    } catch {
        Write-Warn "Could not fetch ANF accounts: $($_.Exception.Message)"
    }

    if ($anfAccounts.Count -eq 0) {
        $script:AnfData.Add([PSCustomObject]@{
            AccountName             = 'No ANF Accounts Discovered'
            AccountRegion           = 'N/A'
            AccountResourceGroup    = 'N/A'
            CapacityPools           = @()
            CapacityPoolsSummary    = 'N/A'
            TotalPoolCount          = 0
            TotalAllocatedTiB       = 0
            Volumes                 = @()
            VolumeCount             = 0
            VolumesNearQuota        = @()
            VolumesNearQuotaCount   = 0
            VolumesNearQuotaSummary = 'N/A'
            AffectedHostPools       = @()
            AffectedHostPoolsCount  = 0
            AffectedHostPoolsSummary = 'N/A'
            ColocationStatus        = 'INFO'
            DataProtectionSummary   = 'Unable to assess'
            DataProtectionStatus    = 'INFO'
            BackupSummary           = 'Unable to assess'
            BackupStatus            = 'INFO'
            RedundancyStr           = 'N/A'
            BackupDPStr             = 'N/A'
            SoftDeleteStr           = 'N/A'
            SnapPolicyStr           = 'N/A'
            BackupPolicyStr         = 'N/A'
            BackupVaultStr          = 'N/A'
            AzureBackupStr          = 'N/A'
        })
        return
    }

    foreach ($anfAccount in $anfAccounts) {
        $acctName   = $anfAccount.Name
        $acctRegion = $anfAccount.Location
        $acctRg     = Get-RgFromArmId $anfAccount.Id

        # ── 1. Capacity Pools ────────────────────────────────────────────────
        $pools = @()
        try {
            $pools = @(Invoke-WithRetry -OperationName "ANF Pools:$acctName" -ScriptBlock {
                Get-AzNetAppFilesPool -ResourceGroupName $acctRg -AccountName $acctName -ErrorAction Stop
            })
        } catch { Write-Warn "Could not fetch capacity pools for $acctName" }

        $poolDetails = @()
        foreach ($pool in $pools) {
            $sizeTiB = [math]::Round($pool.Size / 1099511627776, 2)   # bytes → TiB
            $poolDetails += [PSCustomObject]@{
                Name         = $pool.Name.Split('/')[-1]
                ServiceLevel = $pool.ServiceLevel
                SizeTiB      = $sizeTiB
            }
        }
        $poolSummary = if ($poolDetails.Count -gt 0) {
            ($poolDetails | ForEach-Object { "$($_.Name) ($($_.ServiceLevel), $($_.SizeTiB) TiB)" }) -join ' | '
        } else { 'No Capacity Pools Found' }

        # ── 2. Region Colocation ─────────────────────────────────────────────
        $acctRegionNorm = ($acctRegion -replace '[\s-]').ToLower()
        $affectedHPs = @()
        foreach ($hp in $script:allHostPools) {
            $hpRegionNorm = ($hp.Location -replace '[\s-]').ToLower()
            if ($acctRegionNorm -ne $hpRegionNorm) { $affectedHPs += $hp.Name }
        }
        $colStatus = if ($affectedHPs.Count -eq 0) { 'PASS' } else { 'FAIL' }
        $affectedSummary = if ($affectedHPs.Count -gt 0) {
            "Cross-Region: $($affectedHPs -join ', ')"
        } else { 'All Host Pools Colocated' }

        # ── 3. Volumes + Usage ≥$QuotaWarningPercent% ────────────────────────────────────
        $allVolumes      = @()
        $volumesNearQuota = @()
        foreach ($pool in $pools) {
            $poolShort = $pool.Name.Split('/')[-1]
            $vols = @()
            try {
                $vols = @(Invoke-WithRetry -OperationName "ANF Vols:$poolShort" -ScriptBlock {
                    Get-AzNetAppFilesVolume -ResourceGroupName $acctRg -AccountName $acctName -PoolName $poolShort -ErrorAction Stop
                })
            } catch { Write-Warn "Could not fetch volumes for pool $poolShort" }

            foreach ($vol in $vols) {
                $volName  = $vol.Name.Split('/')[-1]
                $quotaGiB = [math]::Round($vol.UsageThreshold / 1073741824, 2)
                $mountPath = 'N/A'
                if ($vol.MountTargets -and $vol.MountTargets.Count -gt 0) {
                    $ip = $vol.MountTargets[0].IpAddress
                    $mountPath = "\\$ip\$($vol.CreationToken)"
                }

                # Try volume usage metric
                $usedGiB = 0; $usedPct = 0; $hasMetric = $false
                try {
                    $m = Invoke-WithRetry -ScriptBlock {
                        Get-AzMetric -ResourceId $vol.Id -MetricName 'VolumeLogicalSize' `
                            -AggregationType Average -TimeGrain 01:00:00 `
                            -StartTime (Get-Date).AddHours(-1) -EndTime (Get-Date) -ErrorAction Stop
                    }
                    $latest = ($m.Data | Where-Object { $null -ne $_.Average } | Select-Object -Last 1)
                    if ($latest -and $latest.Average) {
                        $usedGiB  = [math]::Round($latest.Average / 1073741824, 2)
                        $usedPct  = if ($vol.UsageThreshold -gt 0) { [math]::Round(($latest.Average / $vol.UsageThreshold) * 100, 1) } else { 0 }
                        $hasMetric = $true
                    }
                } catch { }

                # ── Per-volume protection & encryption details ────────────
                $volSnapPolicyId = 'None'
                $volBkPolicyId   = 'None'
                $volBkVaultId    = 'None'
                $volEncKeySource = 'Microsoft.NetApp'   # default (Microsoft-managed)

                if ($vol.DataProtection) {
                    if ($vol.DataProtection.Snapshot -and $vol.DataProtection.Snapshot.SnapshotPolicyId) {
                        $volSnapPolicyId = ($vol.DataProtection.Snapshot.SnapshotPolicyId -split '/')[-1]
                    }
                    if ($vol.DataProtection.Backup) {
                        if ($vol.DataProtection.Backup.BackupPolicyId) {
                            $volBkPolicyId = ($vol.DataProtection.Backup.BackupPolicyId -split '/')[-1]
                        }
                        if ($vol.DataProtection.Backup.VaultId) {
                            $volBkVaultId = ($vol.DataProtection.Backup.VaultId -split '/')[-1]
                        }
                    }
                }
                if ($vol.EncryptionKeySource) {
                    $volEncKeySource = $vol.EncryptionKeySource  # e.g. 'Microsoft.KeyVault' or 'Microsoft.NetApp'
                }

                $nearQ = $usedPct -ge 80
                $vd = [PSCustomObject]@{
                    Name             = $volName
                    PoolName         = $poolShort
                    MountPath        = $mountPath
                    QuotaGiB         = $quotaGiB
                    UsedGiB          = $usedGiB
                    UsedPercent      = $usedPct
                    HasMetric        = $hasMetric
                    NearQuota        = $nearQ
                    SnapshotPolicyId = $volSnapPolicyId
                    BackupPolicyId   = $volBkPolicyId
                    BackupVaultId    = $volBkVaultId
                    EncryptionKey    = $volEncKeySource
                }
                $allVolumes += $vd
                if ($nearQ) { $volumesNearQuota += $vd }
            }
        }
        $nqSummary = if ($volumesNearQuota.Count -gt 0) {
            ($volumesNearQuota | ForEach-Object { "$($_.Name) ($($_.UsedPercent)% of $($_.QuotaGiB) GiB)" }) -join ' | '
        } else { 'All Volumes Within Quota' }

        # ── 4. Data Protection ───────────────────────────────────────────────
        $hasReplication = $false; $hasVolBackup = $false; $softDelete = $false
        foreach ($pool in $pools) {
            $poolShort = $pool.Name.Split('/')[-1]
            try {
                $pvols = @(Get-AzNetAppFilesVolume -ResourceGroupName $acctRg -AccountName $acctName -PoolName $poolShort -ErrorAction SilentlyContinue)
                foreach ($v in $pvols) {
                    if ($v.DataProtection -and $v.DataProtection.Replication) { $hasReplication = $true }
                    if ($v.DataProtection -and $v.DataProtection.Backup)      { $hasVolBackup  = $true }
                }
            } catch { }
        }
        try {
            $acctFull = Invoke-WithRetry -ScriptBlock {
                Get-AzNetAppFilesAccount -ResourceGroupName $acctRg -Name $acctName -ErrorAction Stop
            }
            if ($acctFull.PSObject.Properties.Name -contains 'EnableSoftDelete') {
                $softDelete = $acctFull.EnableSoftDelete -eq $true
            }
        } catch { }

        $redundancyStr = if ($hasReplication) { 'CRR Configured' }      else { 'No Replication' }
        $backupDPStr   = if ($hasVolBackup)   { 'Volume Backup Enabled' } else { 'No Volume Backup' }
        $softDelStr    = if ($softDelete)     { 'Enabled' }             else { 'Not Detected' }
        $dpStatus = if ($hasReplication -and $hasVolBackup) { 'PASS' } elseif ($hasReplication -or $hasVolBackup) { 'WARN' } else { 'FAIL' }
        $dpSummary = "Redundancy: $redundancyStr | Backup: $backupDPStr | Soft Delete: $softDelStr"

        # ── 5. Backup Configuration ──────────────────────────────────────────
        $snapPols = @(); $bkPols = @(); $bkVaults = @()
        try { $snapPols = @(Invoke-WithRetry -ScriptBlock { Get-AzNetAppFilesSnapshotPolicy -ResourceGroupName $acctRg -AccountName $acctName -ErrorAction Stop }) } catch { }
        try { $bkPols   = @(Invoke-WithRetry -ScriptBlock { Get-AzNetAppFilesBackupPolicy   -ResourceGroupName $acctRg -AccountName $acctName -ErrorAction Stop }) } catch { }
        try { $bkVaults = @(Invoke-WithRetry -ScriptBlock { Get-AzNetAppFilesBackupVault    -ResourceGroupName $acctRg -AccountName $acctName -ErrorAction Stop }) } catch { }

        $snapStr    = if ($snapPols.Count -gt 0)  { "Configured ($($snapPols.Count))" }  else { 'Not Configured' }
        $bkPolStr   = if ($bkPols.Count -gt 0)    { "Configured ($($bkPols.Count))" }    else { 'Not Configured' }
        $bkVaultStr = if ($bkVaults.Count -gt 0)  { "Configured ($($bkVaults.Count))" }  else { 'Not Configured' }
        $azBkStr    = if ($hasVolBackup)           { 'Enabled on Volumes' }               else { 'Not Detected' }
        $bkStatus   = if ($snapPols.Count -gt 0 -or $bkPols.Count -gt 0 -or $bkVaults.Count -gt 0) { 'PASS' } elseif ($hasVolBackup) { 'WARN' } else { 'FAIL' }
        $bkSummary  = "Snapshots: $snapStr | Policy: $bkPolStr | Vault: $bkVaultStr"

        # Aggregate used GiB across all volumes (from metric-capable volumes)
        $anfUsedGiB   = if (($allVolumes | Where-Object { $_.UsedGiB -gt 0 }).Count -gt 0) {
            [math]::Round(($allVolumes | Where-Object { $_.UsedGiB -gt 0 } | Measure-Object -Property UsedGiB -Sum).Sum, 1)
        } else { 0 }
        $anfMaxGiB    = [math]::Round(($poolDetails | Measure-Object -Property SizeTiB -Sum).Sum * 1024, 0)
        $anfUsedPct   = if ($anfMaxGiB -gt 0 -and $anfUsedGiB -gt 0) {
            [math]::Round(($anfUsedGiB / $anfMaxGiB) * 100, 1)
        } else { $null }
        $anfSnapCount = ($allVolumes | Measure-Object -Property SnapshotCount -Sum).Sum

        $script:AnfData.Add([PSCustomObject]@{
            AccountName              = $acctName
            AccountRegion            = $acctRegion
            AccountResourceGroup     = $acctRg
            CapacityPools            = $poolDetails
            CapacityPoolsSummary     = $poolSummary
            TotalPoolCount           = $poolDetails.Count
            TotalAllocatedTiB        = [math]::Round(($poolDetails | Measure-Object -Property SizeTiB -Sum).Sum, 2)
            TotalUsedGiB             = $anfUsedGiB
            OverallUsedPercent       = $anfUsedPct
            SnapshotCount            = $anfSnapCount
            IdentityAuth             = 'N/A (ANF uses SMB Kerberos / NFS)'
            EncryptionKey            = if ($anfAccount.EncryptionKeySource) { $anfAccount.EncryptionKeySource } else { 'Microsoft.NetApp' }
            Volumes                  = $allVolumes
            VolumeCount              = $allVolumes.Count
            VolumesNearQuota         = $volumesNearQuota
            VolumesNearQuotaCount    = $volumesNearQuota.Count
            VolumesNearQuotaSummary  = $nqSummary
            AffectedHostPools        = $affectedHPs
            AffectedHostPoolsCount   = $affectedHPs.Count
            AffectedHostPoolsSummary = $affectedSummary
            ColocationStatus         = $colStatus
            DataProtectionSummary    = $dpSummary
            DataProtectionStatus     = $dpStatus
            BackupSummary            = $bkSummary
            BackupStatus             = $bkStatus
            RedundancyStr            = $redundancyStr
            BackupDPStr              = $backupDPStr
            SoftDeleteStr            = $softDelStr
            SnapPolicyStr            = $snapStr
            BackupPolicyStr          = $bkPolStr
            BackupVaultStr           = $bkVaultStr
            AzureBackupStr           = $azBkStr
        })
    }
    Write-Ok "ANF data collected for $($script:AnfData.Count) account(s)"
}

function Get-AzureFilesData {
    Write-Step 'Collecting Azure Files (Storage Account) data'

    # ── Discover storage accounts ────────────────────────────────────────────
    # Supports: multiple storage RGs, multiple explicit account names, or both.
    # Uses ONLY ARM-level Get-AzStorageAccount (Reader RBAC). Does NOT call
    # listKeys — all share-level reads use an OAuth token context instead.
    $accounts  = [System.Collections.Generic.List[object]]::new()
    $seenSaIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    try {
        if ($StorageAccountNames.Count -gt 0) {
            # Explicit account list: resolve each one
            foreach ($saName in $StorageAccountNames) {
                if ($StorageResourceGroupNames.Count -gt 0) {
                    foreach ($saRg in $StorageResourceGroupNames) {
                        try {
                            $sa = Invoke-WithRetry -ScriptBlock {
                                Get-AzStorageAccount -ResourceGroupName $saRg -Name $saName -ErrorAction Stop
                            }
                            if ($seenSaIds.Add($sa.Id)) { $accounts.Add($sa) }
                        } catch { }
                    }
                } else {
                    # No RG constraint: search subscription-wide for this account name
                    $found = @(Invoke-WithRetry -ScriptBlock {
                        Get-AzResource -ResourceType 'Microsoft.Storage/storageAccounts' `
                            -Name $saName -ErrorAction Stop
                    })
                    foreach ($r in $found) {
                        try {
                            $sa = Invoke-WithRetry -ScriptBlock {
                                Get-AzStorageAccount -ResourceGroupName $r.ResourceGroupName -Name $r.Name -ErrorAction Stop
                            }
                            if ($seenSaIds.Add($sa.Id)) { $accounts.Add($sa) }
                        } catch { }
                    }
                }
            }
        } elseif ($StorageResourceGroupNames.Count -gt 0) {
            # RG-scoped: list all storage accounts in each configured RG
            foreach ($saRg in $StorageResourceGroupNames) {
                $saList = @(Invoke-WithRetry -ScriptBlock {
                    Get-AzStorageAccount -ResourceGroupName $saRg -ErrorAction Stop
                })
                foreach ($sa in $saList) {
                    if ($seenSaIds.Add($sa.Id)) { $accounts.Add($sa) }
                }
            }
        } else {
            # No scope constraints: list all storage accounts in the subscription
            Write-Warn2 'No StorageResourceGroupNames or StorageAccountNames specified — scanning entire subscription for Storage Accounts.'
            $saList = @(Invoke-WithRetry -ScriptBlock {
                Get-AzStorageAccount -ErrorAction Stop
            })
            foreach ($sa in $saList) {
                if ($seenSaIds.Add($sa.Id)) { $accounts.Add($sa) }
            }
        }
        Write-Ok "Found $($accounts.Count) storage account(s)"
    } catch {
        Write-Warn "Could not fetch storage accounts: $($_.Exception.Message)"
    }

    if ($accounts.Count -eq 0) {
        $script:AnfData.Add([PSCustomObject]@{
            AccountName = 'No Storage Accounts Discovered'; AccountRegion = 'N/A'; AccountResourceGroup = 'N/A'
            CapacityPools = @(); CapacityPoolsSummary = 'N/A'; TotalPoolCount = 0; TotalAllocatedTiB = 0
            Volumes = @(); VolumeCount = 0; VolumesNearQuota = @(); VolumesNearQuotaCount = 0
            VolumesNearQuotaSummary = 'N/A'; AffectedHostPools = @(); AffectedHostPoolsCount = 0
            AffectedHostPoolsSummary = 'N/A'; ColocationStatus = 'INFO'; DataProtectionSummary = 'Unable to assess'
            DataProtectionStatus = 'INFO'; BackupSummary = 'Unable to assess'; BackupStatus = 'INFO'
            RedundancyStr = 'N/A'; BackupDPStr = 'N/A'; SoftDeleteStr = 'N/A'; SnapPolicyStr = 'N/A'
            BackupPolicyStr = 'N/A'; BackupVaultStr = 'N/A'; AzureBackupStr = 'N/A'
        })
        return
    }

    foreach ($sa in $accounts) {
        $acctName   = $sa.StorageAccountName
        $acctRegion = $sa.PrimaryLocation
        $acctRg     = Get-RgFromArmId $sa.Id

        # ── OAuth context — no listKeys required (Reader RBAC compatible) ────
        # New-AzStorageContext with -UseConnectedAccount uses the current signed-in
        # Entra ID token (data-plane RBAC: Storage File Data Privileged Reader or
        # higher). Falls back to listing shares via ARM if data-plane RBAC is absent.
        $oauthCtx = $null
        try {
            $oauthCtx = New-AzStorageContext -StorageAccountName $acctName -UseConnectedAccount -ErrorAction Stop
        } catch {
            Write-Warn2 "Could not build OAuth context for '$acctName' — ensure the running identity has 'Storage File Data Privileged Reader' on this account. Share-level metrics will be skipped."
        }

        # ── File share list via ARM (Reader RBAC, no data-plane needed) ──────
        $shares = @()
        try {
            $shares = @(Invoke-WithRetry -ScriptBlock {
                Get-AzRmStorageShare -ResourceGroupName $acctRg -StorageAccountName $acctName -ErrorAction Stop
            })
        } catch {
            Write-Warn "Could not list shares for '$acctName': $($_.Exception.Message)"
        }

        $poolSummary = "$($sa.Sku.Name) ($($sa.Kind))"

        # Region colocation vs host pools
        $acctRegionNorm = ($acctRegion -replace '[\s-]').ToLower()
        $affectedHPs = @()
        foreach ($hp in $script:allHostPools) {
            $hpRegionNorm = ($hp.Location -replace '[\s-]').ToLower()
            if ($acctRegionNorm -ne $hpRegionNorm) { $affectedHPs += $hp.Name }
        }
        $colStatus       = if ($affectedHPs.Count -eq 0) { 'PASS' } else { 'FAIL' }
        $affectedSummary = if ($affectedHPs.Count -gt 0) { "Cross-Region: $($affectedHPs -join ', ')" } else { 'All Host Pools Colocated' }

        $allVolumes = @(); $volumesNearQuota = @()
        foreach ($sh in $shares) {
            $shareName = $sh.Name
            # Quota: ARM object exposes this directly in GiB
            $quotaGiB  = if ($sh.ShareQuota) { $sh.ShareQuota } elseif ($sh.Quota) { $sh.Quota } else { 0 }

            $usedGiB = $null; $usedPct = $null; $hasMetric = $false
            if ($oauthCtx) {
                try {
                    # GetShareUsage via OAuth data-plane — does NOT require listKeys
                    $oauthShare = Get-AzStorageShare -Context $oauthCtx -Name $shareName -ErrorAction Stop
                    $usageBytes = $oauthShare.ShareUsageInBytes
                    if ($usageBytes -gt 0) {
                        $usedGiB   = [math]::Round($usageBytes / 1GB, 2)
                        $usedPct   = if ($quotaGiB -gt 0) { [math]::Round(($usedGiB / $quotaGiB) * 100, 1) } else { 0 }
                        $hasMetric = $true
                    }
                } catch {
                    # Data-plane RBAC absent or share unavailable — used% left as null
                }
            }

            $nearQ = ($usedPct -ne $null -and $usedPct -ge $QuotaWarningPercent)
            $vd = [PSCustomObject]@{
                Name             = $shareName
                PoolName         = $sa.Sku.Name
                MountPath        = "\\$acctName.file.core.windows.net\$shareName"
                QuotaGiB         = $quotaGiB
                UsedGiB          = if ($hasMetric) { $usedGiB } else { $null }
                UsedPercent      = if ($hasMetric) { $usedPct  } else { $null }
                HasMetric        = $hasMetric
                NearQuota        = $nearQ
                SnapshotPolicyId = 'N/A'
                BackupPolicyId   = 'N/A'
                BackupVaultId    = 'N/A'
                EncryptionKey    = if ($sa.Encryption.KeySource) { $sa.Encryption.KeySource } else { 'Microsoft.Storage' }
            }
            $allVolumes     += $vd
            if ($nearQ) { $volumesNearQuota += $vd }
        }

        $nqSummary = if ($volumesNearQuota.Count -gt 0) {
            ($volumesNearQuota | ForEach-Object {
                "$($_.Name) ($($_.UsedPercent)% of $($_.QuotaGiB) GiB)"
            }) -join ' | '
        } else { 'All Shares Within Quota' }

        # Data protection: redundancy + soft delete (ARM — no listKeys needed)
        $redundancyStr  = $sa.Sku.Name
        $softDelete     = $false; $softDeleteDays = 0
        try {
            $sp = Get-AzStorageFileServiceProperty -ResourceGroupName $acctRg -StorageAccountName $acctName -ErrorAction Stop
            if ($sp.ShareDeleteRetentionPolicy.Enabled) {
                $softDelete     = $true
                $softDeleteDays = $sp.ShareDeleteRetentionPolicy.Days
            }
        } catch { }
        $softDelStr  = if ($softDelete) { "Enabled ($softDeleteDays`d)" } else { 'Not Configured' }
        $identityAuth = ($sa.AzureFilesIdentityBasedAuthentication -and
                         $sa.AzureFilesIdentityBasedAuthentication.DirectoryServiceOptions -ne 'None')
        $backupDPStr  = if ($identityAuth) { "Identity-Based Auth ($($sa.AzureFilesIdentityBasedAuthentication.DirectoryServiceOptions))" } else { 'No Identity-Based Auth' }
        $dpStatus     = if ($redundancyStr -match 'GRS|ZRS' -and $softDelete) { 'PASS' } `
                        elseif ($softDelete -or $redundancyStr -match 'GRS|ZRS')  { 'WARN' } else { 'FAIL' }
        $dpSummary    = "Redundancy: $redundancyStr | Auth: $backupDPStr | Soft Delete: $softDelStr"
        $bkSummary    = "Soft Delete: $softDelStr | Auth: $backupDPStr"
        $bkStatus     = if ($softDelete) { 'PASS' } else { 'WARN' }

        # Aggregate totals for overview panel
        $totalUsedGiB     = if (($allVolumes | Where-Object { $_.HasMetric }).Count -gt 0) {
            [math]::Round(($allVolumes | Where-Object { $_.HasMetric } | Measure-Object -Property UsedGiB -Sum).Sum, 1)
        } else { 0 }
        $totalMaxGiB_agg  = ($allVolumes | Measure-Object -Property QuotaGiB -Sum).Sum
        $overallUsedPct   = if ($totalMaxGiB_agg -gt 0 -and $totalUsedGiB -gt 0) {
            [math]::Round(($totalUsedGiB / $totalMaxGiB_agg) * 100, 1)
        } else { $null }

        # Snapshot count — best-effort from ARM (no listKeys)
        $snapshotCount = 0
        foreach ($sh in $shares) {
            try {
                $snapList = @(Get-AzRmStorageShare -ResourceGroupName $acctRg -StorageAccountName $acctName `
                    -Name $sh.Name -IncludeSnapshot -ErrorAction Stop | Where-Object { $_.SnapshotTime })
                $snapshotCount += $snapList.Count
            } catch { }
        }

        $script:AnfData.Add([PSCustomObject]@{
            AccountName              = $acctName
            AccountRegion            = $acctRegion
            AccountResourceGroup     = $acctRg
            CapacityPools            = @()
            CapacityPoolsSummary     = $poolSummary
            TotalPoolCount           = $shares.Count
            TotalAllocatedTiB        = [math]::Round((($allVolumes | Where-Object { $_.QuotaGiB } | Measure-Object -Property QuotaGiB -Sum).Sum / 1024), 2)
            TotalUsedGiB             = $totalUsedGiB
            OverallUsedPercent       = $overallUsedPct
            SnapshotCount            = $snapshotCount
            IdentityAuth             = $backupDPStr
            Volumes                  = $allVolumes
            VolumeCount              = $allVolumes.Count
            VolumesNearQuota         = $volumesNearQuota
            VolumesNearQuotaCount    = $volumesNearQuota.Count
            VolumesNearQuotaSummary  = $nqSummary
            AffectedHostPools        = $affectedHPs
            AffectedHostPoolsCount   = $affectedHPs.Count
            AffectedHostPoolsSummary = $affectedSummary
            ColocationStatus         = $colStatus
            DataProtectionSummary    = $dpSummary
            DataProtectionStatus     = $dpStatus
            BackupSummary            = $bkSummary
            BackupStatus             = $bkStatus
            RedundancyStr            = $redundancyStr
            BackupDPStr              = $backupDPStr
            SoftDeleteStr            = $softDelStr
            SnapPolicyStr            = 'N/A (Azure Files)'
            BackupPolicyStr          = 'N/A (Azure Files)'
            BackupVaultStr           = 'N/A (Azure Files)'
            AzureBackupStr           = $backupDPStr
            EncryptionKey            = if ($sa.Encryption.KeySource) { $sa.Encryption.KeySource } else { 'Microsoft.Storage' }
        })
    }
    Write-Ok "Azure Files data collected for $($accounts.Count) account(s)"
}

# ==============================================================================
# DRY RUN SEEDER
# ==============================================================================

function Initialize-DryRunData {
    Write-Host ''
    Write-Host '  [DryRun] Generating synthetic data — no Azure calls.' -ForegroundColor Cyan

    $script:Overview = @{
        ClientName        = $ClientName
        EnvironmentName   = $EnvironmentName
        Subscription      = 'Sample-Subscription (DryRun)'
        SubscriptionID    = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
        Tenant             = 'yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy'
        ResourceGroup     = if ($ResourceGroupNames.Count -gt 0) { $ResourceGroupNames -join ', ' } else { 'rg-avd-sample' }
        TotalHostPools    = 5
        TotalSessionHosts = 32
        ReportGeneratedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz')
    }

    # Host Pools
    @(
        [PSCustomObject]@{ Name='hp-pooled-prod-01'; Type='Pooled'; LoadBalancer='BreadthFirst'; MaxSessionLimit=12; ValidationEnv=$false; StartVMOnConnect=$false; ScalingPlan='Assigned'; AppGroupCount=2; AppGroups='ag-prod-desktop, ag-prod-apps'; DriveRedirect='DISABLED'; ClipboardRedirect='ENABLED'; PrinterRedirect='ENABLED'; UsbRedirect='DISABLED'; PublicNetwork='Disabled' },
        [PSCustomObject]@{ Name='hp-pooled-prod-02'; Type='Pooled'; LoadBalancer='DepthFirst';  MaxSessionLimit=999999; ValidationEnv=$false; StartVMOnConnect=$false; ScalingPlan='None';     AppGroupCount=1; AppGroups='ag-prod-02-desktop';  DriveRedirect='ENABLED';  ClipboardRedirect='ENABLED'; PrinterRedirect='ENABLED'; UsbRedirect='ENABLED';  PublicNetwork='Enabled' },
        [PSCustomObject]@{ Name='hp-personal-exec';  Type='Personal'; LoadBalancer='Persistent'; MaxSessionLimit=1;  ValidationEnv=$false; StartVMOnConnect=$false; ScalingPlan='None';     AppGroupCount=1; AppGroups='ag-exec-desktop';     DriveRedirect='ENABLED';  ClipboardRedirect='ENABLED'; PrinterRedirect='ENABLED'; UsbRedirect='ENABLED';  PublicNetwork='Enabled' },
        [PSCustomObject]@{ Name='hp-pooled-dev-01';  Type='Pooled'; LoadBalancer='BreadthFirst'; MaxSessionLimit=8;  ValidationEnv=$true;  StartVMOnConnect=$true;  ScalingPlan='Assigned'; AppGroupCount=1; AppGroups='ag-dev-apps';          DriveRedirect='DISABLED'; ClipboardRedirect='DISABLED'; PrinterRedirect='DISABLED'; UsbRedirect='DISABLED'; PublicNetwork='Disabled' },
        [PSCustomObject]@{ Name='hp-pooled-test-01'; Type='Pooled'; LoadBalancer='BreadthFirst'; MaxSessionLimit=4;  ValidationEnv=$true;  StartVMOnConnect=$true;  ScalingPlan='Assigned'; AppGroupCount=1; AppGroups='ag-test-desktop';       DriveRedirect='DISABLED'; ClipboardRedirect='DISABLED'; PrinterRedirect='DISABLED'; UsbRedirect='DISABLED'; PublicNetwork='Disabled' }
    ) | ForEach-Object { $script:HostPoolData.Add($_) }

    # Session Hosts — with full AMA extension details
    # Dry-run: latest AMA version = 1.43 (normalised major.minor, matches extension reporting)
    $dryLatestAma = '1.43'
    $shData = @(
        @{ Name='avd-sh-001'; HostPool='hp-pooled-prod-01'; Status='Available';       HealthStatus='PASS'; Drain=$false; Sessions=8;  AZ='1'; MonVer='1.41'; MonLatest=$false; NotBooted2Days=$false; DiskSku='Premium_LRS';     AMAPublisher='Microsoft.Azure.Monitor'; AMAType='AzureMonitorWindowsAgent'; AMAProv='Succeeded';     AMAAutoUp='Enabled';  AMALatest=$dryLatestAma; AMAUpdate="Update Available ($dryLatestAma)" },
        @{ Name='avd-sh-002'; HostPool='hp-pooled-prod-01'; Status='Available';       HealthStatus='PASS'; Drain=$false; Sessions=10; AZ='2'; MonVer='1.41'; MonLatest=$false; NotBooted2Days=$false; DiskSku='Premium_LRS';     AMAPublisher='Microsoft.Azure.Monitor'; AMAType='AzureMonitorWindowsAgent'; AMAProv='Succeeded';     AMAAutoUp='Enabled';  AMALatest=$dryLatestAma; AMAUpdate="Update Available ($dryLatestAma)" },
        @{ Name='avd-sh-003'; HostPool='hp-pooled-prod-01'; Status='Available';       HealthStatus='PASS'; Drain=$false; Sessions=7;  AZ='3'; MonVer='1.10'; MonLatest=$false; NotBooted2Days=$false; DiskSku='Premium_LRS';     AMAPublisher='Microsoft.Azure.Monitor'; AMAType='AzureMonitorWindowsAgent'; AMAProv='Succeeded';     AMAAutoUp='Disabled'; AMALatest=$dryLatestAma; AMAUpdate="Update Available ($dryLatestAma)" },
        @{ Name='avd-sh-004'; HostPool='hp-pooled-prod-01'; Status='Unavailable';     HealthStatus='FAIL'; Drain=$false; Sessions=0;  AZ='1'; MonVer='1.8';  MonLatest=$false; NotBooted2Days=$false; DiskSku='Standard_LRS';     AMAPublisher='Microsoft.Azure.Monitor'; AMAType='AzureMonitorWindowsAgent'; AMAProv='Succeeded';     AMAAutoUp='Disabled'; AMALatest=$dryLatestAma; AMAUpdate="Update Available ($dryLatestAma)" },
        @{ Name='avd-sh-005'; HostPool='hp-pooled-prod-01'; Status='NeedsAssistance'; HealthStatus='WARN'; Drain=$true;  Sessions=0;  AZ='N/A'; MonVer='1.41'; MonLatest=$false; NotBooted2Days=$false; DiskSku='Premium_LRS';     AMAPublisher='Microsoft.Azure.Monitor'; AMAType='AzureMonitorWindowsAgent'; AMAProv='Succeeded';     AMAAutoUp='Enabled';  AMALatest=$dryLatestAma; AMAUpdate="Update Available ($dryLatestAma)" },
        @{ Name='avd-sh-006'; HostPool='hp-pooled-prod-02'; Status='Available';       HealthStatus='PASS'; Drain=$false; Sessions=11; AZ='1'; MonVer='1.43'; MonLatest=$true;  NotBooted2Days=$false; DiskSku='Premium_LRS';     AMAPublisher='Microsoft.Azure.Monitor'; AMAType='AzureMonitorWindowsAgent'; AMAProv='Succeeded';     AMAAutoUp='Enabled';  AMALatest=$dryLatestAma; AMAUpdate='Up to Date' },
        @{ Name='avd-sh-007'; HostPool='hp-pooled-prod-02'; Status='Available';       HealthStatus='PASS'; Drain=$false; Sessions=9;  AZ='1'; MonVer='1.43'; MonLatest=$true;  NotBooted2Days=$true;  DiskSku='Standard_LRS';     AMAPublisher='Microsoft.Azure.Monitor'; AMAType='AzureMonitorWindowsAgent'; AMAProv='Succeeded';     AMAAutoUp='Enabled';  AMALatest=$dryLatestAma; AMAUpdate='Up to Date' },
        @{ Name='avd-sh-008'; HostPool='hp-pooled-prod-02'; Status='Shutdown';        HealthStatus='INFO'; Drain=$false; Sessions=0;  AZ='N/A'; MonVer='VM Shutdown'; MonLatest=$true;  NotBooted2Days=$true;  DiskSku='Unknown';   AMAPublisher='N/A - VM Shutdown';       AMAType='N/A';                         AMAProv='N/A - VM Shutdown'; AMAAutoUp='N/A'; AMALatest=$dryLatestAma; AMAUpdate='N/A - VM Shutdown' },
        @{ Name='avd-exec-001'; HostPool='hp-personal-exec'; Status='Available';      HealthStatus='PASS'; Drain=$false; Sessions=1;  AZ='2'; MonVer='1.43'; MonLatest=$true;  NotBooted2Days=$false; DiskSku='Premium_LRS';     AMAPublisher='Microsoft.Azure.Monitor'; AMAType='AzureMonitorWindowsAgent'; AMAProv='Succeeded';     AMAAutoUp='Enabled';  AMALatest=$dryLatestAma; AMAUpdate='Up to Date' },
        @{ Name='avd-exec-002'; HostPool='hp-personal-exec'; Status='Available';      HealthStatus='PASS'; Drain=$false; Sessions=1;  AZ='3'; MonVer='1.43'; MonLatest=$true;  NotBooted2Days=$false; DiskSku='Premium_LRS';     AMAPublisher='Microsoft.Azure.Monitor'; AMAType='AzureMonitorWindowsAgent'; AMAProv='Succeeded';     AMAAutoUp='Enabled';  AMALatest=$dryLatestAma; AMAUpdate='Up to Date' },
        @{ Name='avd-dev-001';  HostPool='hp-pooled-dev-01'; Status='Available';      HealthStatus='PASS'; Drain=$false; Sessions=3;  AZ='1'; MonVer='1.41'; MonLatest=$false; NotBooted2Days=$false; DiskSku='StandardSSD_LRS'; AMAPublisher='Microsoft.Azure.Monitor'; AMAType='AzureMonitorWindowsAgent'; AMAProv='Succeeded';     AMAAutoUp='Enabled';  AMALatest=$dryLatestAma; AMAUpdate="Update Available ($dryLatestAma)" },
        @{ Name='avd-dev-002';  HostPool='hp-pooled-dev-01'; Status='Available';      HealthStatus='PASS'; Drain=$false; Sessions=2;  AZ='2'; MonVer='1.41'; MonLatest=$false; NotBooted2Days=$false; DiskSku='StandardSSD_LRS'; AMAPublisher='Microsoft.Azure.Monitor'; AMAType='AzureMonitorWindowsAgent'; AMAProv='Succeeded';     AMAAutoUp='Enabled';  AMALatest=$dryLatestAma; AMAUpdate="Update Available ($dryLatestAma)" }
    )
    foreach ($sh in $shData) {
        $script:SessionHostData.Add([PSCustomObject]@{
            Name                 = $sh.Name
            HostPool             = $sh.HostPool
            Status               = $sh.Status
            HealthStatus         = $sh.HealthStatus
            AllowNewSession      = -not $sh.Drain
            DrainMode            = $sh.Drain
            Sessions             = $sh.Sessions
            NotBooted2Days       = $sh.NotBooted2Days
            AvailabilityZone     = $sh.AZ
            MonitorAgentVersion  = $sh.MonVer
            MonitorAgentLatest   = $sh.MonLatest
            AMAPublisher         = $sh.AMAPublisher
            AMAExtensionType     = $sh.AMAType
            AMAProvisioningState = $sh.AMAProv
            AMAAutoUpgrade       = $sh.AMAAutoUp
            AMALatestVersion     = $sh.AMALatest
            AMAUpdateStatus      = $sh.AMAUpdate
            OsDiskSku            = $sh.DiskSku
            LastHeartbeat        = (Get-Date).AddHours(-2)
        })
    }

    # Health Checks — seeded with full per-host detail matching the API JSON format
    # (healthCheckName / healthCheckResult / additionalFailureDetails / errorCode / lastHealthCheckDateTime)
    $hcTimestamp = (Get-Date).AddHours(-2).ToString('yyyy-MM-dd HH:mm:ss')
    @(
        @{
            CheckName='DomainJoinedCheck'; Status='PASS'; FailingCount=0; PassingCount=12; TotalCount=12
            FailingHosts=''; DataSource='API (ExpandAll)'; Note=''
            FailDetail=''
            PassDetail='avd-sh-001, avd-sh-002, avd-sh-003, avd-sh-005, avd-sh-006, avd-sh-007, avd-exec-001, avd-exec-002, avd-dev-001, avd-dev-002'
            PerHostRows=@(
                @{Host='avd-sh-001'; Result='HealthCheckSucceeded'; ErrorCode=0; Message='SessionHost healthy: is joined to domain'; LastDt=$hcTimestamp},
                @{Host='avd-sh-002'; Result='HealthCheckSucceeded'; ErrorCode=0; Message='SessionHost healthy: is joined to domain'; LastDt=$hcTimestamp},
                @{Host='avd-sh-004'; Result='HealthCheckSucceeded'; ErrorCode=0; Message='SessionHost healthy: is joined to domain'; LastDt=$hcTimestamp},
                @{Host='avd-sh-005'; Result='HealthCheckSucceeded'; ErrorCode=0; Message='SessionHost healthy: is joined to domain'; LastDt=$hcTimestamp}
            )
        },
        @{
            CheckName='DomainTrustCheck'; Status='FAIL'; FailingCount=2; PassingCount=10; TotalCount=12
            FailingHosts='avd-sh-004, avd-sh-005'; DataSource='API (ExpandAll)'; Note=''
            FailDetail="avd-sh-004 [Err:1789] — The trust relationship between the workstation and the primary domain failed. (Last: $hcTimestamp)`navd-sh-005 [Err:1789] — The trust relationship between the workstation and the primary domain failed. (Last: $hcTimestamp)"
            PassDetail='avd-sh-001, avd-sh-002, avd-sh-003, avd-sh-006, avd-sh-007, avd-exec-001, avd-exec-002, avd-dev-001, avd-dev-002'
            PerHostRows=@(
                @{Host='avd-sh-001'; Result='HealthCheckSucceeded'; ErrorCode=0; Message='SessionHost healthy: domain trust check passed'; LastDt=$hcTimestamp},
                @{Host='avd-sh-002'; Result='HealthCheckSucceeded'; ErrorCode=0; Message='SessionHost healthy: domain trust check passed'; LastDt=$hcTimestamp},
                @{Host='avd-sh-004'; Result='HealthCheckFailed'; ErrorCode=1789; Message='The trust relationship between the workstation and the primary domain failed.'; LastDt=$hcTimestamp},
                @{Host='avd-sh-005'; Result='HealthCheckFailed'; ErrorCode=1789; Message='The trust relationship between the workstation and the primary domain failed.'; LastDt=$hcTimestamp}
            )
        },
        @{
            CheckName='SxSStackListenerCheck'; Status='FAIL'; FailingCount=2; PassingCount=10; TotalCount=12
            FailingHosts='avd-sh-004, avd-sh-005'; DataSource='API (ExpandAll)'; Note=''
            FailDetail="avd-sh-004 [Err:0] — SessionHost SxS Stack is not listening. (Last: $hcTimestamp)`navd-sh-005 [Err:0] — SessionHost SxS Stack is not listening. (Last: $hcTimestamp)"
            PassDetail='avd-sh-001, avd-sh-002, avd-sh-003, avd-sh-006, avd-sh-007, avd-exec-001, avd-exec-002, avd-dev-001, avd-dev-002'
            PerHostRows=@(
                @{Host='avd-sh-001'; Result='HealthCheckSucceeded'; ErrorCode=0; Message='SessionHost healthy: SessionHost is listening on the SxS Stack.'; LastDt=$hcTimestamp},
                @{Host='avd-sh-004'; Result='HealthCheckFailed'; ErrorCode=0; Message='SessionHost SxS Stack is not listening.'; LastDt=$hcTimestamp},
                @{Host='avd-sh-005'; Result='HealthCheckFailed'; ErrorCode=0; Message='SessionHost SxS Stack is not listening.'; LastDt=$hcTimestamp}
            )
        },
        @{
            CheckName='UrlsAccessibleCheck'; Status='FAIL'; FailingCount=1; PassingCount=11; TotalCount=12
            FailingHosts='avd-sh-004'; DataSource='API (ExpandAll)'; Note=''
            FailDetail="avd-sh-004 [Err:0] — {`"AccessibleUrls`":[`"378c8cfe-870b-4ac8-bef4-d3c5c9d69f5b`",`"kms.core.windows.net`"]} (Last: $hcTimestamp)"
            PassDetail='avd-sh-001, avd-sh-002, avd-sh-003, avd-sh-005, avd-sh-006, avd-sh-007, avd-exec-001, avd-exec-002, avd-dev-001, avd-dev-002'
            PerHostRows=@(
                @{Host='avd-sh-001'; Result='HealthCheckSucceeded'; ErrorCode=0; Message='All required URLs are accessible.'; LastDt=$hcTimestamp},
                @{Host='avd-sh-004'; Result='HealthCheckFailed'; ErrorCode=0; Message='{"AccessibleUrls":["378c8cfe-870b","kms.core.windows.net"]}'; LastDt=$hcTimestamp}
            )
        },
        @{
            CheckName='MonitoringAgentCheck'; Status='PASS'; FailingCount=0; PassingCount=12; TotalCount=12
            FailingHosts=''; DataSource='API (ExpandAll)'; Note=''
            FailDetail=''
            PassDetail='avd-sh-001, avd-sh-002, avd-sh-003, avd-sh-004, avd-sh-005, avd-sh-006, avd-sh-007, avd-exec-001, avd-exec-002, avd-dev-001, avd-dev-002'
            PerHostRows=@(
                @{Host='avd-sh-001'; Result='HealthCheckSucceeded'; ErrorCode=0; Message='Located running process at C:\Program Files\Microsoft Monitoring Agent\Agent\MonitoringHost.exe'; LastDt=$hcTimestamp}
            )
        },
        @{
            CheckName='MetaDataServiceCheck'; Status='PASS'; FailingCount=0; PassingCount=12; TotalCount=12
            FailingHosts=''; DataSource='API (ExpandAll)'; Note=''
            FailDetail=''
            PassDetail='avd-sh-001, avd-sh-002, avd-sh-003, avd-sh-004, avd-sh-005, avd-sh-006, avd-sh-007, avd-exec-001, avd-exec-002, avd-dev-001, avd-dev-002'
            PerHostRows=@(
                @{Host='avd-sh-001'; Result='HealthCheckSucceeded'; ErrorCode=0; Message='{\"m\":\"\n Data1\": \"IMDS present and accessible\"}'; LastDt=$hcTimestamp}
            )
        },
        @{
            CheckName='AppAttachHealthCheck'; Status='PASS'; FailingCount=0; PassingCount=12; TotalCount=12
            FailingHosts=''; DataSource='API (ExpandAll)'; Note='Only applicable if MSIX App Attach is configured'
            FailDetail=''
            PassDetail='avd-sh-001, avd-sh-002, avd-sh-003, avd-sh-004, avd-sh-005, avd-sh-006, avd-sh-007, avd-exec-001, avd-exec-002, avd-dev-001, avd-dev-002'
            PerHostRows=@(
                @{Host='avd-sh-001'; Result='HealthCheckSucceeded'; ErrorCode=0; Message='SessionHost healthy: MSIX package staging completed.'; LastDt=$hcTimestamp}
            )
        },
        @{
            CheckName='TURNRelayAccessHealthCheck'; Status='FAIL'; FailingCount=2; PassingCount=10; TotalCount=12
            FailingHosts='avd-sh-004, avd-sh-005'; DataSource='API (ExpandAll)'; Note=''
            FailDetail="avd-sh-004 [Err:2146762759] — NAT shape is Undetermined when port allocation is symmetric. (Last: $hcTimestamp)`navd-sh-005 [Err:2146762759] — NAT shape is Undetermined when port allocation is symmetric. (Last: $hcTimestamp)"
            PassDetail='avd-sh-001, avd-sh-002, avd-sh-003, avd-sh-006, avd-sh-007, avd-exec-001, avd-exec-002, avd-dev-001, avd-dev-002'
            PerHostRows=@(
                @{Host='avd-sh-001'; Result='HealthCheckSucceeded'; ErrorCode=0; Message='TURN relay access is available.'; LastDt=$hcTimestamp},
                @{Host='avd-sh-004'; Result='HealthCheckFailed'; ErrorCode=2146762759; Message='NAT shape is Undetermined when port allocation is symmetric.'; LastDt=$hcTimestamp},
                @{Host='avd-sh-005'; Result='HealthCheckFailed'; ErrorCode=2146762759; Message='NAT shape is Undetermined when port allocation is symmetric.'; LastDt=$hcTimestamp}
            )
        },
        @{
            CheckName='AADJoinedHealthCheck'; Status='INFO'; FailingCount=0; PassingCount=12; TotalCount=12
            FailingHosts=''; DataSource='API (ExpandAll)'; Note='Full result may require Log Analytics / AVD Insights'
            FailDetail=''
            PassDetail='avd-sh-001, avd-sh-002, avd-sh-003, avd-sh-004, avd-sh-005, avd-sh-006, avd-sh-007, avd-exec-001, avd-exec-002, avd-dev-001, avd-dev-002'
            PerHostRows=@(
                @{Host='avd-sh-001'; Result='HealthCheckSucceeded'; ErrorCode=0; Message='SessionHost healthy: Microsoft Entra joined.'; LastDt=$hcTimestamp}
            )
        }
    ) | ForEach-Object {
        $hcObj = $_
        $script:HealthCheckData.Add([PSCustomObject]@{
            CheckName    = $hcObj.CheckName
            Status       = $hcObj.Status
            FailingCount = $hcObj.FailingCount
            PassingCount = $hcObj.PassingCount
            TotalCount   = $hcObj.TotalCount
            FailingHosts = $hcObj.FailingHosts
            FailDetail   = $hcObj.FailDetail
            PassDetail   = $hcObj.PassDetail
            PerHostRows  = $hcObj.PerHostRows
            DataSource   = $hcObj.DataSource
            Note         = $hcObj.Note
        })
    }

    # Azure NetApp Files
    $anfPools1 = @(
        [PSCustomObject]@{ Name = 'pool-prod-01'; ServiceLevel = 'Premium';  SizeTiB = 4.0 },
        [PSCustomObject]@{ Name = 'pool-prod-02'; ServiceLevel = 'Standard'; SizeTiB = 2.0 }
    )
    $anfPools2 = @(
        [PSCustomObject]@{ Name = 'pool-dev-01'; ServiceLevel = 'Standard'; SizeTiB = 1.0 }
    )
    $anfVols1 = @(
        [PSCustomObject]@{ Name='vol-profiles-prod';   PoolName='pool-prod-01'; MountPath='\\10.0.1.4\vol-profiles-prod';   QuotaGiB=1024; UsedGiB=870;  UsedPercent=85.0; HasMetric=$true;  NearQuota=$true;  SnapshotPolicyId='snappol-prod-hourly'; BackupPolicyId='bkpol-prod-daily'; BackupVaultId='bkvault-prod'; EncryptionKey='Microsoft.KeyVault' },
        [PSCustomObject]@{ Name='vol-containers-prod'; PoolName='pool-prod-01'; MountPath='\\10.0.1.4\vol-containers-prod'; QuotaGiB=512;  UsedGiB=256;  UsedPercent=50.0; HasMetric=$true;  NearQuota=$false; SnapshotPolicyId='snappol-prod-hourly'; BackupPolicyId='bkpol-prod-daily'; BackupVaultId='bkvault-prod'; EncryptionKey='Microsoft.KeyVault' },
        [PSCustomObject]@{ Name='vol-office-cache';    PoolName='pool-prod-02'; MountPath='\\10.0.1.5\vol-office-cache';    QuotaGiB=256;  UsedGiB=210;  UsedPercent=82.0; HasMetric=$true;  NearQuota=$true;  SnapshotPolicyId='None';                BackupPolicyId='None';             BackupVaultId='None';         EncryptionKey='Microsoft.NetApp' }
    )
    $anfVols2 = @(
        [PSCustomObject]@{ Name='vol-profiles-dev'; PoolName='pool-dev-01'; MountPath='\\10.0.2.4\vol-profiles-dev'; QuotaGiB=256; UsedGiB=64; UsedPercent=25.0; HasMetric=$true; NearQuota=$false; SnapshotPolicyId='None'; BackupPolicyId='None'; BackupVaultId='None'; EncryptionKey='Microsoft.NetApp' }
    )
    @(
        [PSCustomObject]@{
            AccountName='anf-avd-prod-eastus'; AccountRegion='eastus'; AccountResourceGroup=$(if($ANFResourceGroupNames.Count -gt 0){$ANFResourceGroupNames[0]}else{'rg-anf-sample'})
            CapacityPools=$anfPools1; CapacityPoolsSummary='pool-prod-01 (Premium, 4.0 TiB) | pool-prod-02 (Standard, 2.0 TiB)'
            TotalPoolCount=2; TotalAllocatedTiB=6.0; TotalUsedGiB=1336; OverallUsedPercent=21.7; SnapshotCount=8
            IdentityAuth='N/A (ANF uses SMB Kerberos / NFS)'; EncryptionKey='Microsoft.KeyVault'
            Volumes=$anfVols1; VolumeCount=3
            VolumesNearQuota=@($anfVols1 | Where-Object { $_.NearQuota }); VolumesNearQuotaCount=2
            VolumesNearQuotaSummary='vol-profiles-prod (85.0% of 1024 GiB) | vol-office-cache (82.0% of 256 GiB)'
            AffectedHostPools=@(); AffectedHostPoolsCount=0; AffectedHostPoolsSummary='All Host Pools Colocated'
            ColocationStatus='PASS'
            DataProtectionSummary='Redundancy: CRR Configured | Backup: Volume Backup Enabled | Soft Delete: Enabled'
            DataProtectionStatus='PASS'
            BackupSummary='Snapshots: Configured (8) | Policy: Configured (1) | Vault: Configured (1)'
            BackupStatus='PASS'
            RedundancyStr='CRR Configured'; BackupDPStr='Volume Backup Enabled'; SoftDeleteStr='Enabled'
            SnapPolicyStr='Configured (2)'; BackupPolicyStr='Configured (1)'; BackupVaultStr='Configured (1)'; AzureBackupStr='Enabled on Volumes'
        },
        [PSCustomObject]@{
            AccountName='anf-avd-dev-westeurope'; AccountRegion='westeurope'; AccountResourceGroup=$(if($ANFResourceGroupNames.Count -gt 0){$ANFResourceGroupNames[0]}else{'rg-anf-sample'})
            CapacityPools=$anfPools2; CapacityPoolsSummary='pool-dev-01 (Standard, 1.0 TiB)'
            TotalPoolCount=1; TotalAllocatedTiB=1.0; TotalUsedGiB=64; OverallUsedPercent=6.3; SnapshotCount=0
            IdentityAuth='N/A (ANF uses SMB Kerberos / NFS)'; EncryptionKey='Microsoft.NetApp'
            Volumes=$anfVols2; VolumeCount=1
            VolumesNearQuota=@(); VolumesNearQuotaCount=0
            VolumesNearQuotaSummary='All Volumes Within Quota'
            AffectedHostPools=@('hp-pooled-prod-01','hp-pooled-prod-02','hp-personal-exec','hp-pooled-dev-01','hp-pooled-test-01')
            AffectedHostPoolsCount=5; AffectedHostPoolsSummary='Cross-Region: hp-pooled-prod-01, hp-pooled-prod-02, hp-personal-exec, hp-pooled-dev-01, hp-pooled-test-01'
            ColocationStatus='FAIL'
            DataProtectionSummary='Redundancy: No Replication | Backup: No Volume Backup | Soft Delete: Not Detected'
            DataProtectionStatus='FAIL'
            BackupSummary='Snapshots: Not Configured | Policy: Not Configured | Vault: Not Configured'
            BackupStatus='FAIL'
            RedundancyStr='No Replication'; BackupDPStr='No Volume Backup'; SoftDeleteStr='Not Detected'
            SnapPolicyStr='Not Configured'; BackupPolicyStr='Not Configured'; BackupVaultStr='Not Configured'; AzureBackupStr='Not Detected'
        }
    ) | ForEach-Object { $script:AnfData.Add($_) }

    # DryRun Azure Files — only seeded when ProfileStorageType includes AzureFiles
    if ($ProfileStorageType -eq 'AzureFiles' -or $ProfileStorageType -eq 'Both') {
        $dryShares = @(
            [PSCustomObject]@{ Name='testfslogix'; PoolName='Premium_LRS'; MountPath='\\stavdprofilessample.file.core.windows.net\testfslogix'; QuotaGiB=102400; UsedGiB=531; UsedPercent=0.5; HasMetric=$true; NearQuota=$false; SnapshotPolicyId='N/A'; BackupPolicyId='N/A'; BackupVaultId='N/A'; EncryptionKey='Microsoft.Storage' },
            [PSCustomObject]@{ Name='containers-prod'; PoolName='Premium_LRS'; MountPath='\\stavdprofilessample.file.core.windows.net\containers-prod'; QuotaGiB=512; UsedGiB=210; UsedPercent=41.0; HasMetric=$true; NearQuota=$false; SnapshotPolicyId='N/A'; BackupPolicyId='N/A'; BackupVaultId='N/A'; EncryptionKey='Microsoft.Storage' }
        )
        $script:AnfData.Add([PSCustomObject]@{
            AccountName='stavdprofilessample'; AccountRegion='eastus'
            AccountResourceGroup=$(if($StorageResourceGroupNames.Count -gt 0){$StorageResourceGroupNames[0]}else{'rg-storage-sample'})
            CapacityPools=@(); CapacityPoolsSummary='Premium_LRS (FileStorage)'
            TotalPoolCount=2; TotalAllocatedTiB=[math]::Round(102912/1024,2); TotalUsedGiB=741; OverallUsedPercent=0.7; SnapshotCount=0
            IdentityAuth='Identity-Based Auth (AD)'; EncryptionKey='Microsoft.Storage'
            Volumes=$dryShares; VolumeCount=2
            VolumesNearQuota=@(); VolumesNearQuotaCount=0
            VolumesNearQuotaSummary='All Shares Within Quota'
            AffectedHostPools=@(); AffectedHostPoolsCount=0; AffectedHostPoolsSummary='All Host Pools Colocated'
            ColocationStatus='PASS'
            DataProtectionSummary='Redundancy: Premium_LRS | Auth: Identity-Based Auth (AD) | Soft Delete: Disabled'
            DataProtectionStatus='WARN'
            BackupSummary='Soft Delete: Disabled | Auth: Identity-Based Auth (AD)'
            BackupStatus='WARN'
            RedundancyStr='Premium_LRS'; BackupDPStr='Identity-Based Auth (AD)'; SoftDeleteStr='Disabled'
            SnapPolicyStr='N/A (Azure Files)'; BackupPolicyStr='N/A (Azure Files)'; BackupVaultStr='N/A (Azure Files)'; AzureBackupStr='Identity-Based Auth (AD)'
        })
    }
}

# ==============================================================================
# HTML REPORT GENERATION
# New design: dark cyber-terminal aesthetic with neon accents, grid layout,
# expandable rows, animated status indicators — entirely different from v2
# ==============================================================================

function New-HtmlReport {
    $ts       = ConvertTo-HtmlSafe $script:Overview['ReportGeneratedAt']
    $sub      = ConvertTo-HtmlSafe $script:Overview['Subscription']
    $subId    = ConvertTo-HtmlSafe $script:Overview['SubscriptionID']
    $tenant   = ConvertTo-HtmlSafe $script:Overview['Tenant']
    $rg       = ConvertTo-HtmlSafe $script:Overview['ResourceGroup']
    $hpCount  = $script:Overview['TotalHostPools']
    $shCount  = $script:Overview['TotalSessionHosts']

    # ── OVERVIEW TABLE ──────────────────────────────────────────────────────
    # Overview shows only identity and environment summary.
    # Profile storage details (size, used capacity, soft delete, backup, etc.)
    # are shown in the dedicated Profile Storage section below — not here.
    $overviewRows = @"
<tr><td class="k">Subscription</td><td class="v">$sub</td></tr>
<tr><td class="k">Subscription ID</td><td class="v mono">$subId</td></tr>
<tr><td class="k">Tenant ID</td><td class="v mono">$tenant</td></tr>
<tr><td class="k">Resource Group</td><td class="v">$rg</td></tr>
<tr><td class="k">Total Host Pools</td><td class="v accent">$hpCount</td></tr>
<tr><td class="k">Total Session Hosts</td><td class="v accent">$shCount</td></tr>
<tr><td class="k">Report Generated</td><td class="v">$ts</td></tr>
"@

    # ── HOST POOLS TABLE ────────────────────────────────────────────────────
    $hpRows = [System.Text.StringBuilder]::new()
    foreach ($hp in $script:HostPoolData) {
        $spBadge   = if ($hp.ScalingPlan -eq 'Assigned') { '<span class="badge badge-pass">Assigned</span>' } else { '<span class="badge badge-fail">None</span>' }
        $svBadge   = if ($hp.StartVMOnConnect) { '<span class="badge badge-pass">On</span>' } else { '<span class="badge badge-fail">Off</span>' }
        $valBadge  = if ($hp.ValidationEnv)   { '<span class="badge badge-warn">Validation</span>' } else { '<span class="badge badge-info">Production</span>' }
        $mslClass  = if ($hp.MaxSessionLimit -ge 999999) { 'badge-fail' } else { 'badge-pass' }
        $mslVal    = if ($hp.MaxSessionLimit -ge 999999) { '&#9888; Unlimited' } else { $hp.MaxSessionLimit }

        $drBadge  = "<span class='badge $(Get-StatusBadge $hp.DriveRedirect)'>$($hp.DriveRedirect)</span>"
        $clBadge  = "<span class='badge $(Get-StatusBadge $hp.ClipboardRedirect)'>$($hp.ClipboardRedirect)</span>"
        $prBadge  = "<span class='badge $(Get-StatusBadge $hp.PrinterRedirect)'>$($hp.PrinterRedirect)</span>"
        $usBadge  = "<span class='badge $(Get-StatusBadge $hp.UsbRedirect)'>$($hp.UsbRedirect)</span>"
        $pubBadge = if ($hp.PublicNetwork -eq 'Disabled') { '<span class="badge badge-pass">Private</span>' } elseif ($hp.PublicNetwork -eq 'Enabled') { '<span class="badge badge-fail">Public</span>' } else { '<span class="badge badge-info">Unknown</span>' }

        [void]$hpRows.AppendLine(@"
<tr class="expandable" onclick="toggleRow(this)">
  <td class="nm">$($hp.Name)</td>
  <td><span class="badge badge-info">$($hp.Type)</span></td>
  <td><span class="badge badge-info">$($hp.LoadBalancer)</span></td>
  <td><span class="badge $mslClass">$mslVal</span></td>
  <td>$valBadge</td>
  <td>$spBadge</td>
  <td>$svBadge</td>
  <td>$pubBadge</td>
</tr>
<tr class="detail-row">
  <td colspan="8">
    <div class="detail-grid">
      <div class="dg-item"><span class="dk">App Groups ($($hp.AppGroupCount))</span><span class="dv">$($hp.AppGroups)</span></div>
      <div class="dg-item"><span class="dk">Drive Redirection</span>$drBadge</div>
      <div class="dg-item"><span class="dk">Clipboard Redirection</span>$clBadge</div>
      <div class="dg-item"><span class="dk">Printer Redirection</span>$prBadge</div>
      <div class="dg-item"><span class="dk">USB Redirection</span>$usBadge</div>
    </div>
  </td>
</tr>
"@)
    }

    # ── SESSION HOSTS TABLE ──────────────────────────────────────────────────
    $shRows = [System.Text.StringBuilder]::new()
    foreach ($sh in $script:SessionHostData) {
        $stClass = Get-StatusBadge $sh.HealthStatus
        $drainBadge = if ($sh.DrainMode) { '<span class="badge badge-warn">&#9658; Drain ON</span>' } else { '<span class="badge badge-pass">Active</span>' }
        $bootBadge  = if ($sh.NotBooted2Days) { '<span class="badge badge-fail">&#9888; Stale</span>' } else { '<span class="badge badge-pass">Recent</span>' }
        $azBadge    = if ($sh.AvailabilityZone -eq 'N/A') { '<span class="badge badge-warn">No AZ</span>' } else { "<span class='badge badge-pass'>AZ $($sh.AvailabilityZone)</span>" }
        $monBadge   = if ($sh.MonitorAgentVersion -eq 'VM Shutdown') { '<span class="badge badge-info">&#9211; VM Shutdown</span>' } elseif ($sh.MonitorAgentVersion -eq 'Not Detected') { '<span class="badge badge-fail">&#10007; Not Detected</span>' } elseif ($sh.AMAUpdateStatus -match 'Update Available') { "<span class='badge badge-warn'>&#9888; v$($sh.MonitorAgentVersion) — Update Required</span>" } elseif ($sh.MonitorAgentLatest) { "<span class='badge badge-pass'>&#10003; v$($sh.MonitorAgentVersion)</span>" } else { "<span class='badge badge-warn'>&#9888; v$($sh.MonitorAgentVersion)</span>" }
        $diskBadge  = if ($sh.OsDiskSku -match 'Premium|UltraSSD') { "<span class='badge badge-pass'>$($sh.OsDiskSku)</span>" } elseif ($sh.OsDiskSku -eq 'Standard_LRS') { "<span class='badge badge-fail'>$($sh.OsDiskSku)</span>" } else { "<span class='badge badge-info'>$($sh.OsDiskSku)</span>" }

        [void]$shRows.AppendLine(@"
<tr class="expandable sh-row" data-status="$(ConvertTo-HtmlSafe $sh.Status)" onclick="toggleRow(this)">
  <td class="nm">$($sh.Name)</td>
  <td class="sm">$($sh.HostPool)</td>
  <td><span class="badge $stClass">$($sh.Status)</span></td>
  <td>$drainBadge</td>
  <td>$bootBadge</td>
  <td>$azBadge</td>
  <td>$monBadge</td>
  <td>$diskBadge</td>
</tr>
<tr class="detail-row sh-detail" data-status="$(ConvertTo-HtmlSafe $sh.Status)">
  <td colspan="8">
    <div class="detail-grid">
      <div class="dg-item"><span class="dk">Sessions Active</span><span class="dv">$($sh.Sessions)</span></div>
      <div class="dg-item"><span class="dk">OS Disk SKU</span><span class="dv mono">$($sh.OsDiskSku)</span></div>
      <div class="dg-item"><span class="dk">Last Heartbeat</span><span class="dv">$($sh.LastHeartbeat)</span></div>
    </div>
    <div style="margin-top:10px;margin-bottom:6px;font-size:11px;color:var(--dim);text-transform:uppercase;letter-spacing:.08em">Azure Monitor Windows Agent</div>
    <div class="detail-grid">
      <div class="dg-item"><span class="dk">Name</span><span class="dv mono">AzureMonitorWindowsAgent</span></div>
      <div class="dg-item"><span class="dk">Type (Publisher)</span><span class="dv mono">$($sh.AMAPublisher)</span></div>
      <div class="dg-item"><span class="dk">Version</span><span class="dv mono">$($sh.MonitorAgentVersion)</span></div>
      <div class="dg-item"><span class="dk">Latest Version</span><span class="dv mono">$($sh.AMALatestVersion)</span></div>
      <div class="dg-item"><span class="dk">Status</span>$(if ($sh.AMAUpdateStatus -match 'Update Available') { "<span class='badge badge-warn'>&#9888; $($sh.AMAUpdateStatus)</span>" } elseif ($sh.AMAUpdateStatus -eq 'Up to Date') { "<span class='badge badge-pass'>&#10003; Up to Date</span>" } else { "<span class='dv'>$($sh.AMAUpdateStatus)</span>" })</div>
      <div class="dg-item"><span class="dk">Provisioning State</span><span class="dv">$($sh.AMAProvisioningState)</span></div>
      <div class="dg-item"><span class="dk">Automatic Upgrade</span><span class="dv">$($sh.AMAAutoUpgrade)</span></div>
    </div>
  </td>
</tr>
"@)
    }

    # ── HEALTH CHECKS TABLE ─────────────────────────────────────────────────
    $hcRows = [System.Text.StringBuilder]::new()
    foreach ($hc in $script:HealthCheckData) {
        $stClass = Get-StatusBadge $hc.Status
        $icon = switch ($hc.Status) {
            'PASS' { '&#10003;' }
            'FAIL' { '&#10007;' }
            'WARN' { '&#9888;'  }
            'INFO' { '&#9432;'  }
            default { '?' }
        }

        $totalCnt = if ($hc.TotalCount)   { $hc.TotalCount   } else { $hc.FailingCount + [int]($hc.PassingCount) }
        $passCnt  = if ($hc.PassingCount) { $hc.PassingCount } else { $totalCnt - $hc.FailingCount }

        $srcBadge = if ($hc.DataSource -like '*Estimated*') {
            "<span class='badge badge-warn'>&#9888; $($hc.DataSource)</span>"
        } else {
            "<span class='badge badge-pass'>&#10003; $($hc.DataSource)</span>"
        }
        $noteHtml = if ($hc.Note) { "<div class='note-box'>&#128276; $($hc.Note)</div>" } else { '' }

        # ── Per-host breakdown table ─────────────────────────────────────────
        # Only FAILING hosts are shown. If all hosts pass, no table is rendered —
        # a compact "All X hosts passing" line appears instead.
        $failingRows = @()
        if ($hc.PerHostRows -and $hc.PerHostRows.Count -gt 0) {
            $failingRows = @($hc.PerHostRows | Where-Object {
                $_.Result -notlike '*Succeeded*'
            })
        }

        $perHostTableRows = [System.Text.StringBuilder]::new()
        foreach ($ph in $failingRows) {
            $isEst    = $ph.Result -like '*Estimated*'
            $rowClass = if ($isEst) { 'ph-est' } else { 'ph-fail' }
            $phIcon   = if ($isEst) { '&#9888;' } else { '&#10007;' }
            $phColor  = if ($isEst) { '#fbbf24' } else { '#f87171' }
            $errDisp  = if ($ph.ErrorCode -and $ph.ErrorCode -ne 0) {
                " <span style='color:#f87171;font-size:10px'>[Err: $($ph.ErrorCode)]</span>"
            } else { '' }
            $msgDisp  = if ($ph.Message) {
                "<div class='ph-msg'>$(ConvertTo-HtmlSafe $ph.Message)</div>"
            } else { '' }
            $dtDisp   = if ($ph.LastDt) { "<span class='ph-dt'>$($ph.LastDt)</span>" } else { '' }
            $resultClean = $ph.Result -replace 'HealthCheck','' -replace 'Succeeded','Succeeded' -replace 'Failed','Failed'
            [void]$perHostTableRows.AppendLine("<tr class='$rowClass'>
  <td class='ph-host mono'>$($ph.Host)</td>
  <td><span style='color:$phColor;font-weight:600'>$phIcon $resultClean</span>$errDisp</td>
  <td class='ph-msg-td'>$msgDisp$dtDisp</td>
</tr>")
        }

        # Build the expand content
        if ($failingRows.Count -gt 0) {
            $failLabel   = if ($failingRows.Count -eq 1) { '1 failing host' } else { "$($failingRows.Count) failing hosts" }
            $perHostTable = "<div class='hc-per-host-head'>&#10007; $failLabel — details below (passing hosts omitted)</div>
<div class='tbl-wrap' style='margin:8px 0'>
  <table class='ph-table'>
    <thead><tr><th>Session Host</th><th>Result</th><th>Error Message / Detail</th></tr></thead>
    <tbody>$($perHostTableRows.ToString())</tbody>
  </table>
</div>"
        } elseif ($hc.Status -eq 'PASS') {
            $allPassMsg = if ($totalCnt -gt 0) { "$totalCnt" } else { 'All' }
            $perHostTable = "<div class='hc-all-pass'>&#10003; All $allPassMsg session host(s) passed this check. No failures to display.</div>"
        } else {
            $perHostTable = "<div class='hc-all-pass'>&#9432; No per-host detail available for this check.</div>"
        }

        [void]$hcRows.AppendLine(@"
<tr class="expandable" onclick="toggleRow(this)">
  <td class="nm">$($hc.CheckName)</td>
  <td><span class="badge $stClass">$icon $($hc.Status)</span></td>
  <td class="cnt">$(if($hc.FailingCount -gt 0){"<span style='color:#f87171;font-weight:700'>$($hc.FailingCount)</span> / $totalCnt"}else{"<span style='color:#34d399'>0</span> / $totalCnt"})</td>
  <td class="sm">$srcBadge</td>
</tr>
<tr class="detail-row">
  <td colspan="4">
    $perHostTable
    $noteHtml
  </td>
</tr>
"@)
    }

    # ── AZURE NETAPP FILES TABLE ────────────────────────────────────────────
    $anfRows = [System.Text.StringBuilder]::new()
    foreach ($anf in $script:AnfData) {
        $colBadge = "<span class='badge $(Get-StatusBadge $anf.ColocationStatus)'>$(if($anf.ColocationStatus -eq 'PASS'){'&#10003; All Colocated'}elseif($anf.ColocationStatus -eq 'FAIL'){"&#9888; $($anf.AffectedHostPoolsCount) Cross-Region"}else{'&#9432; N/A'})</span>"
        $dpBadge  = "<span class='badge $(Get-StatusBadge $anf.DataProtectionStatus)'>$(switch($anf.DataProtectionStatus){'PASS'{'&#10003; Protected'}'WARN'{'&#9888; Partial'}'FAIL'{'&#10007; Unprotected'}default{'N/A'}})</span>"
        $bkBadge  = "<span class='badge $(Get-StatusBadge $anf.BackupStatus)'>$(switch($anf.BackupStatus){'PASS'{'&#10003; Configured'}'WARN'{'&#9888; Partial'}'FAIL'{'&#10007; Not Configured'}default{'N/A'}})</span>"
        $nqBadge  = if ($anf.VolumesNearQuotaCount -gt 0) { "<span class='badge badge-fail'>&#9888; $($anf.VolumesNearQuotaCount) Vol(s) &ge;$($QuotaWarningPercent)%</span>" } else { "<span class='badge badge-pass'>&#10003; All OK</span>" }

        # Build pool details for expand row
        $poolHtml = ''
        foreach ($p in $anf.CapacityPools) {
            $poolHtml += "<div class='dg-item'><span class='dk'>$($p.Name)</span><span class='dv'>$($p.ServiceLevel) &mdash; $($p.SizeTiB) TiB</span></div>"
        }

        # Build full volume table rows (ALL volumes) — with usage bar + protection details
        $volTableRows = [System.Text.StringBuilder]::new()
        foreach ($v in $anf.Volumes) {
            # Usage bar colour
            $barClass  = if ($v.UsedPercent -ge 80) { 'bar-crit' } elseif ($v.UsedPercent -ge 60) { 'bar-warn' } else { 'bar-ok' }
            $pctDisplay = if ($v.HasMetric) { "$($v.UsedPercent)%" } else { 'No Metric' }
            $barWidth  = [math]::Min($v.UsedPercent, 100)
            $spBadge   = if ($v.SnapshotPolicyId -ne 'None') { "<span class='badge badge-pass' title='$($v.SnapshotPolicyId)'>&#10003; $($v.SnapshotPolicyId)</span>" } else { "<span class='badge badge-fail'>None</span>" }
            $bpBadge   = if ($v.BackupPolicyId   -ne 'None') { "<span class='badge badge-pass' title='$($v.BackupPolicyId)'>&#10003; $($v.BackupPolicyId)</span>"   } else { "<span class='badge badge-fail'>None</span>" }
            $bvBadge   = if ($v.BackupVaultId    -ne 'None') { "<span class='badge badge-pass' title='$($v.BackupVaultId)'>&#10003; $($v.BackupVaultId)</span>"     } else { "<span class='badge badge-fail'>None</span>" }
            $encBadge  = if ($v.EncryptionKey -match 'KeyVault') { "<span class='badge badge-pass'>&#128273; Customer Key</span>" } else { "<span class='badge badge-info'>&#128274; Microsoft</span>" }
            [void]$volTableRows.AppendLine("<tr>
  <td class='nm vn'>$($v.Name)</td>
  <td class='sm'>$($v.PoolName)</td>
  <td class='vq'>$($v.UsedGiB) / $($v.QuotaGiB) GiB</td>
  <td class='vp'>
    <div class='usage-wrap'>
      <div class='usage-bar'><div class='usage-fill $barClass' style='width:$($barWidth)%'></div></div>
      <span class='usage-pct $(if($v.UsedPercent -ge 80){"pct-crit"}elseif($v.UsedPercent -ge 60){"pct-warn"}else{"pct-ok"})'>$pctDisplay</span>
    </div>
  </td>
  <td>$spBadge</td>
  <td>$bpBadge</td>
  <td>$bvBadge</td>
  <td>$encBadge</td>
</tr>")
        }

        # Build affected host pool alerts
        $affHtml = ''
        if ($anf.AffectedHostPoolsCount -gt 0) {
            $affHtml = "<div class='fail-hosts'><span class='dk'>Affected Host Pools ($($anf.AffectedHostPoolsCount)):</span> <span class='mono fh'>$($anf.AffectedHostPools -join ', ')</span></div>"
        }

        # ── Build storage properties summary card (shown at top of expand row) ──
        $maxGiBDisplay  = if ($anf.TotalAllocatedTiB) { "$([math]::Round($anf.TotalAllocatedTiB * 1024, 0)) GiB" } else { 'N/A' }
        $usedGiBDisplay = if ($anf.TotalUsedGiB -gt 0) {
            $pct = if ($anf.OverallUsedPercent) { " ($($anf.OverallUsedPercent)%)" } else { '' }
            "$($anf.TotalUsedGiB) GiB$pct"
        } else { 'N/A — requires Storage File Data Privileged Reader' }

        $snapDisplay    = if ($anf.SnapshotCount -gt 0) { "$($anf.SnapshotCount) snapshot(s)" } `
                          elseif ($anf.SnapPolicyStr -and $anf.SnapPolicyStr -notlike '*N/A*') { $anf.SnapPolicyStr } `
                          else { 'None / N/A' }
        $softDelDisplay = if ($anf.SoftDeleteStr) { $anf.SoftDeleteStr } else { 'N/A' }
        $identDisplay   = if ($anf.IdentityAuth)  { $anf.IdentityAuth  } else { 'N/A' }
        $encDisplay     = if ($anf.EncryptionKey) { $anf.EncryptionKey } else { 'N/A' }

        $sdClass = if ($softDelDisplay -like 'Enabled*') { 'badge-pass' } `
                   elseif ($softDelDisplay -eq 'N/A')     { 'badge-info' } `
                   else { 'badge-fail' }

        $storagePropsHtml = @"
<div class="anf-sub-head">&#128209; Storage Account Properties</div>
<div class="detail-grid" style="margin-bottom:14px">
  <div class="dg-item"><span class="dk">Maximum Storage (Quota)</span><span class="dv accent">$maxGiBDisplay</span></div>
  <div class="dg-item"><span class="dk">Used Storage Capacity</span><span class="dv accent">$usedGiBDisplay</span></div>
  <div class="dg-item"><span class="dk">Redundancy / SKU</span><span class="dv">$(if($anf.RedundancyStr){"$($anf.RedundancyStr)"}else{'N/A'})</span></div>
  <div class="dg-item"><span class="dk">Soft Delete</span><span class="dv"><span class='badge $sdClass'>$softDelDisplay</span></span></div>
  <div class="dg-item"><span class="dk">Snapshots / Backup Policy</span><span class="dv">$snapDisplay</span></div>
  <div class="dg-item"><span class="dk">Identity-Based Access</span><span class="dv">$identDisplay</span></div>
  <div class="dg-item"><span class="dk">Encryption Key Source</span><span class="dv">$encDisplay</span></div>
  <div class="dg-item"><span class="dk">Region</span><span class="dv mono">$($anf.AccountRegion)</span></div>
  <div class="dg-item"><span class="dk">Resource Group</span><span class="dv mono">$($anf.AccountResourceGroup)</span></div>
</div>
"@

        [void]$anfRows.AppendLine(@"
<tr class="expandable" onclick="toggleRow(this)">
  <td class="nm">$($anf.AccountName)</td>
  <td><span class="badge badge-info">$($anf.TotalPoolCount) Pool(s) &mdash; $($anf.TotalAllocatedTiB) TiB</span></td>
  <td>$colBadge</td>
  <td>$nqBadge</td>
  <td>$dpBadge</td>
  <td>$bkBadge</td>
</tr>
<tr class="detail-row">
  <td colspan="6">
    $affHtml
    $storagePropsHtml

    <div class="anf-sub-head">&#128190; Capacity Pools</div>
    <div class="detail-grid" style="margin-bottom:14px">
      $poolHtml
      <div class="dg-item"><span class="dk">Total Allocated</span><span class="dv">$($anf.TotalAllocatedTiB) TiB</span></div>
    </div>

    <div class="anf-sub-head">&#128202; Volumes — Space Usage &amp; Protection ($($anf.VolumeCount) total)</div>
    <div class="tbl-wrap vol-tbl-wrap">
      <table class="vol-inner-tbl">
        <thead>
          <tr>
            <th>Volume Name</th>
            <th>Pool</th>
            <th>Used / Allocated</th>
            <th>Consumed %</th>
            <th>Snapshot Policy</th>
            <th>Backup Policy</th>
            <th>Backup Vault</th>
            <th>Encryption Key</th>
          </tr>
        </thead>
        <tbody>$($volTableRows.ToString())</tbody>
      </table>
    </div>

    <div class="detail-grid" style="margin-top:14px">
      <div class="dg-item"><span class="dk">Volume Backup</span><span class="dv">$($anf.BackupDPStr)</span></div>
      <div class="dg-item"><span class="dk">Snapshot Policies (acct)</span><span class="dv">$($anf.SnapPolicyStr)</span></div>
      <div class="dg-item"><span class="dk">Backup Policy (acct)</span><span class="dv">$($anf.BackupPolicyStr)</span></div>
      <div class="dg-item"><span class="dk">Backup Vault (acct)</span><span class="dv">$($anf.BackupVaultStr)</span></div>
      <div class="dg-item"><span class="dk">Azure Backup on Vols</span><span class="dv">$($anf.AzureBackupStr)</span></div>
    </div>
  </td>
</tr>
"@)
    }

    # ── SUMMARY STATS ───────────────────────────────────────────────────────
    $totalSH      = $script:SessionHostData.Count
    $healthySH    = @($script:SessionHostData | Where-Object { $_.HealthStatus -eq 'PASS' }).Count
    $drainSH      = @($script:SessionHostData | Where-Object { $_.DrainMode }).Count
    $staleSH      = @($script:SessionHostData | Where-Object { $_.NotBooted2Days }).Count
    $noAZsH       = @($script:SessionHostData | Where-Object { $_.AvailabilityZone -eq 'N/A' }).Count
    $outdatedMon  = @($script:SessionHostData | Where-Object { -not $_.MonitorAgentLatest }).Count
    $hcFail       = @($script:HealthCheckData | Where-Object { $_.Status -eq 'FAIL' }).Count
    $anfColFail   = @($script:AnfData | Where-Object { $_.ColocationStatus -eq 'FAIL' }).Count
    $anfNearQuota = ($script:AnfData | Measure-Object -Property VolumesNearQuotaCount -Sum).Sum
    $hpScaling    = @($script:HostPoolData    | Where-Object { $_.ScalingPlan -eq 'None' -and $_.Type -eq 'Pooled' }).Count
    $hpPublic     = @($script:HostPoolData    | Where-Object { $_.PublicNetwork -eq 'Enabled' }).Count

    # Dynamic stat card labels — reflect actual profile storage backend and account names
    $storageStatLabel  = switch ($ProfileStorageType) {
        'AzureFiles' { 'Files' }
        'Both'       { 'Storage' }
        default      { 'ANF' }
    }
    $volumeUnitLabel   = switch ($ProfileStorageType) {
        'AzureFiles' { 'Shares' }
        'Both'       { 'Vols/Shares' }
        default      { 'Volumes' }
    }
    # Build dropdown list of cross-region accounts with their actual names
    $storageColNames = @($script:AnfData | Where-Object { $_.ColocationStatus -eq 'FAIL' } |
        ForEach-Object { "$($_.AccountName) ($($_.AccountRegion))" }) -join '</li><li>'
    # Build dropdown list of near-quota volumes/shares with account context
    $storageNQNames  = @($script:AnfData | Where-Object { $_.VolumesNearQuotaCount -gt 0 } |
        ForEach-Object {
            $acct = $_.AccountName
            $_.VolumesNearQuota | ForEach-Object { "$acct / $($_.Name) ($($_.UsedPercent)%)" }
        }) -join '</li><li>'

    # ── Pre-compute affected host name lists for stat card dropdowns ──────
    $unhealthyNames = @($script:SessionHostData | Where-Object { $_.HealthStatus -ne 'PASS' } | ForEach-Object { $_.Name }) -join '</li><li>'
    $drainNames     = @($script:SessionHostData | Where-Object { $_.DrainMode }                | ForEach-Object { $_.Name }) -join '</li><li>'
    $staleNames     = @($script:SessionHostData | Where-Object { $_.NotBooted2Days }           | ForEach-Object { $_.Name }) -join '</li><li>'
    $noAZnames      = @($script:SessionHostData | Where-Object { $_.AvailabilityZone -eq 'N/A' } | ForEach-Object { $_.Name }) -join '</li><li>'
    $outdatedNames  = @($script:SessionHostData | Where-Object { -not $_.MonitorAgentLatest }  | ForEach-Object { "$($_.Name) (v$($_.MonitorAgentVersion))" }) -join '</li><li>'
    $hcFailNames    = @($script:HealthCheckData | Where-Object { $_.Status -eq 'FAIL' }       | ForEach-Object { "$($_.CheckName) — $($_.FailingHosts)" }) -join '</li><li>'
    $hpScaleNames   = @($script:HostPoolData    | Where-Object { $_.ScalingPlan -eq 'None' -and $_.Type -eq 'Pooled' } | ForEach-Object { $_.Name }) -join '</li><li>'
    $hpPubNames     = @($script:HostPoolData    | Where-Object { $_.PublicNetwork -eq 'Enabled' } | ForEach-Object { $_.Name }) -join '</li><li>'

    $overallHealth = if ($hcFail -eq 0 -and $anfColFail -eq 0 -and ($totalSH -gt 0 -and $healthySH -eq $totalSH)) { 'OPTIMAL' } `
                     elseif (($hcFail -le 2 -and $healthySH -ge ($totalSH * 0.8))) { 'DEGRADED' } `
                     else { 'CRITICAL' }

    $overallClass = switch ($overallHealth) { 'OPTIMAL' { 'score-green' } 'DEGRADED' { 'score-amber' } default { 'score-red' } }

    # ══════════════════════════════════════════════════════════════════════════
    # FULL HTML
    # ══════════════════════════════════════════════════════════════════════════
    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>AVD Environment Health v3 — $sub</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=JetBrains+Mono:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
/* ─── Reset & Base ─────────────────────────────────────────── */
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:       #05080f;
  --bg2:      #0a1020;
  --bg3:      #101828;
  --bg-glass: rgba(12,20,38,.65);
  --border:   rgba(56,97,150,.18);
  --border2:  rgba(56,97,150,.32);
  --neon:     #38bdf8;
  --neon2:    #34d399;
  --amber:    #fbbf24;
  --red:      #f43f5e;
  --purple:   #a78bfa;
  --text:     #cbd5e1;
  --dim:      #475569;
  --white:    #f1f5f9;
  --font-mono:'JetBrains Mono', monospace;
  --font-sans:'Inter', system-ui, sans-serif;
  --glass-blur:16px;
  --radius:12px;
  --radius-lg:16px;
}
html{scroll-behavior:smooth}
body{
  background:var(--bg);
  color:var(--text);
  font-family:var(--font-mono);
  font-size:13px;
  line-height:1.6;
  min-height:100vh;
  overflow-x:hidden;
}

/* ─── Animated gradient mesh background ───────────────────── */
body::before{
  content:'';
  position:fixed;
  inset:0;
  background:
    radial-gradient(ellipse 80% 60% at 10% 20%, rgba(56,189,248,.08) 0%, transparent 60%),
    radial-gradient(ellipse 60% 50% at 85% 80%, rgba(168,85,247,.06) 0%, transparent 55%),
    radial-gradient(ellipse 70% 40% at 50% 50%, rgba(52,211,153,.04) 0%, transparent 50%);
  pointer-events:none;
  z-index:0;
  animation:meshDrift 20s ease-in-out infinite alternate;
}
@keyframes meshDrift{
  0%{opacity:.8;transform:scale(1) translate(0,0)}
  50%{opacity:1;transform:scale(1.05) translate(-2%,1%)}
  100%{opacity:.8;transform:scale(1) translate(1%,-1%)}
}

/* ─── Entrance animations ─────────────────────────────────── */
@keyframes fadeUp{from{opacity:0;transform:translateY(18px)}to{opacity:1;transform:translateY(0)}}
@keyframes fadeIn{from{opacity:0}to{opacity:1}}
.section{animation:fadeUp .5s ease both}
.section:nth-child(2){animation-delay:.1s}
.section:nth-child(3){animation-delay:.2s}
.section:nth-child(4){animation-delay:.3s}
.section:nth-child(5){animation-delay:.4s}

/* ─── Layout ───────────────────────────────────────────────── */
.wrap{max-width:1440px;margin:0 auto;padding:28px 24px 60px;position:relative;z-index:1}

/* ─── Header ───────────────────────────────────────────────── */
.header{
  display:grid;
  grid-template-columns:1fr auto;
  align-items:center;
  gap:24px;
  background:var(--bg-glass);
  backdrop-filter:blur(var(--glass-blur));
  -webkit-backdrop-filter:blur(var(--glass-blur));
  border:1px solid var(--border2);
  border-radius:var(--radius-lg);
  padding:24px 32px;
  margin-bottom:28px;
  animation:fadeUp .4s ease both;
}
.brand{display:flex;flex-direction:column;gap:6px}
.brand-title{
  font-family:var(--font-sans);
  font-size:32px;
  font-weight:800;
  color:var(--white);
  letter-spacing:-.03em;
}
.brand-title .hi{background:linear-gradient(135deg,var(--neon),var(--purple));-webkit-background-clip:text;-webkit-text-fill-color:transparent}
.brand-title .lo{color:var(--dim);-webkit-text-fill-color:var(--dim)}
.brand-sub{font-size:11px;color:var(--dim);letter-spacing:.08em;text-transform:uppercase;font-family:var(--font-sans)}
.brand-sub span{color:var(--neon2);font-weight:600}

/* ─── Score badge ──────────────────────────────────────────── */
/* ─── Client brand logo + status block ───────────────────── */
/* ─── Score / brand block (right side of header) ──────────── */
.score-block{
  display:flex;
  flex-direction:column;
  align-items:flex-end;
  gap:12px;
  flex-shrink:0;
}
/* Client wordmark — text only, no SVG figure. Color injected at render time
   from -BrandColor via the __BRAND_COLOR__ token below. */
.brand-wordmark{
  font-family:'Inter','Arial Black',Arial,sans-serif;
  font-size:30px;
  font-weight:900;
  letter-spacing:-.5px;
  color:$BrandColor;
  line-height:1;
  user-select:none;
}
/* Status pill */
.status-pill{
  display:inline-flex;
  align-items:center;
  gap:7px;
  padding:5px 16px 5px 12px;
  border-radius:999px;
  font-family:var(--font-sans);
  font-size:11px;
  font-weight:700;
  letter-spacing:.1em;
  text-transform:uppercase;
  border:1px solid;
}
.status-pill-dot{
  width:7px;
  height:7px;
  border-radius:50%;
  flex-shrink:0;
}
.pill-green{
  color:var(--neon2);
  border-color:rgba(52,211,153,.4);
  background:rgba(52,211,153,.08);
}
.pill-green .status-pill-dot{background:var(--neon2);box-shadow:0 0 6px var(--neon2)}
.pill-amber{
  color:var(--amber);
  border-color:rgba(251,191,36,.4);
  background:rgba(251,191,36,.08);
}
.pill-amber .status-pill-dot{background:var(--amber);box-shadow:0 0 6px var(--amber)}
.pill-red{
  color:var(--red);
  border-color:rgba(244,63,94,.4);
  background:rgba(244,63,94,.08);
  animation:pillPulse 2s ease-in-out infinite;
}
.pill-red .status-pill-dot{background:var(--red);box-shadow:0 0 6px var(--red)}
@keyframes pillPulse{
  0%,100%{box-shadow:0 0 0 0 rgba(244,63,94,.0)}
  50%{box-shadow:0 0 0 4px rgba(244,63,94,.18)}
}

/* ─── Overview grid ────────────────────────────────────────── */
.overview-grid{
  display:grid;
  grid-template-columns:repeat(auto-fit, minmax(180px, 1fr));
  gap:1px;
  background:var(--border);
  border:1px solid var(--border2);
  border-radius:var(--radius);
  overflow:hidden;
  margin-bottom:28px;
}
.ov-cell{
  background:var(--bg2);
  padding:14px 18px;
  display:flex;
  flex-direction:column;
  gap:4px;
}
.ov-label{font-size:10px;color:var(--dim);text-transform:uppercase;letter-spacing:.1em;font-weight:600;font-family:var(--font-sans)}
.ov-value{font-size:13px;color:var(--white);word-break:break-all}
.ov-value.accent{color:var(--neon);font-family:var(--font-sans);font-size:22px;font-weight:700}
.ov-value.mono{font-family:var(--font-mono);font-size:12px}

/* ─── Stat cards ───────────────────────────────────────────── */
.stat-strip{
  display:grid;
  grid-template-columns:repeat(auto-fill, minmax(155px, 1fr));
  gap:12px;
  margin-bottom:32px;
  animation:fadeUp .45s ease both;
  animation-delay:.05s;
}
.stat-card{
  background:var(--bg-glass);
  backdrop-filter:blur(var(--glass-blur));
  -webkit-backdrop-filter:blur(var(--glass-blur));
  border:1px solid var(--border);
  border-radius:var(--radius);
  padding:16px 18px;
  display:flex;
  flex-direction:column;
  gap:6px;
  transition:transform .2s, border-color .25s, box-shadow .25s;
  position:relative;
  overflow:hidden;
}
.stat-card::before{
  content:'';position:absolute;top:0;left:0;right:0;height:2px;
  background:linear-gradient(90deg,transparent,var(--border2),transparent);
  transition:background .3s;
}
.stat-card:hover{
  transform:translateY(-2px);
  border-color:var(--border2);
  box-shadow:0 8px 32px rgba(0,0,0,.25);
}
.stat-card:hover::before{background:linear-gradient(90deg,transparent,var(--neon),transparent)}
.stat-card.has-dd{cursor:pointer}
.stat-card.has-dd:hover::before{background:linear-gradient(90deg,transparent,var(--amber),transparent)}
.sc-num{font-family:var(--font-sans);font-size:30px;font-weight:800;line-height:1;letter-spacing:-.02em}
.sc-num.ok  {color:var(--neon2)}
.sc-num.warn{color:var(--amber)}
.sc-num.bad {color:var(--red)}
.sc-num.info{color:var(--neon)}
.sc-label{font-size:10px;color:var(--dim);text-transform:uppercase;letter-spacing:.06em;font-family:var(--font-sans);font-weight:600}

/* ─── Stat card dropdown ──────────────────────────────────── */
.sc-toggle{
  display:inline-block;
  font-size:8px;
  color:var(--dim);
  margin-left:5px;
  transition:transform .25s;
  vertical-align:middle;
}
.stat-card.sc-open .sc-toggle{transform:rotate(180deg);color:var(--neon)}
.sc-dropdown{
  display:none;
  margin-top:10px;
  padding:10px 12px;
  background:rgba(5,8,15,.75);
  border:1px solid var(--border2);
  border-radius:8px;
  max-height:180px;
  overflow-y:auto;
  scrollbar-width:thin;
  scrollbar-color:var(--border2) transparent;
}
.stat-card.sc-open .sc-dropdown{display:block;animation:fadeIn .2s ease}
.sc-dropdown ul{list-style:none;padding:0;margin:0}
.sc-dropdown li{
  font-size:11px;
  color:var(--text);
  padding:4px 0;
  border-bottom:1px solid var(--border);
  font-family:var(--font-mono);
}
.sc-dropdown li:last-child{border-bottom:none}
.sc-dropdown::-webkit-scrollbar{width:4px}
.sc-dropdown::-webkit-scrollbar-track{background:transparent}
.sc-dropdown::-webkit-scrollbar-thumb{background:var(--border2);border-radius:2px}

/* ─── Section ──────────────────────────────────────────────── */
.section{margin-bottom:36px}
.sec-header{
  display:flex;
  align-items:center;
  gap:12px;
  margin-bottom:16px;
  padding:12px 16px;
  background:var(--bg-glass);
  backdrop-filter:blur(8px);
  border:1px solid var(--border);
  border-left:3px solid var(--neon);
  border-radius:var(--radius);
}
.sec-icon{font-size:18px}
.sec-title{
  font-family:var(--font-sans);
  font-size:15px;
  font-weight:700;
  color:var(--white);
  letter-spacing:-.01em;
}
.sec-count{
  margin-left:auto;
  font-size:11px;
  color:var(--dim);
  background:var(--bg3);
  border:1px solid var(--border);
  border-radius:20px;
  padding:3px 12px;
  font-family:var(--font-sans);
}

/* ─── Session Host Filter Bar ──────────────────────────────── */
.sh-filter-bar{
  display:flex;
  align-items:center;
  flex-wrap:wrap;
  gap:8px;
  padding:12px 18px;
  background:var(--bg3);
  border:1px solid var(--border);
  border-bottom:none;
  border-radius:var(--radius) var(--radius) 0 0;
}
.sh-filter-label{
  font-size:11px;
  color:var(--dim);
  letter-spacing:.06em;
  text-transform:uppercase;
  margin-right:4px;
  white-space:nowrap;
}
.sh-filter-btn{
  display:inline-flex;
  align-items:center;
  gap:6px;
  padding:5px 14px;
  border-radius:999px;
  border:1px solid var(--border);
  background:var(--bg2);
  color:var(--fg);
  font-size:12px;
  font-family:var(--font-sans);
  cursor:pointer;
  transition:all .15s ease;
  white-space:nowrap;
}
.sh-filter-btn:hover{
  border-color:var(--neon1);
  color:var(--neon1);
  background:rgba(99,179,237,.08);
}
.sh-filter-active{
  border-color:var(--neon1) !important;
  color:var(--neon1) !important;
  background:rgba(99,179,237,.15) !important;
  font-weight:600;
}
.sh-fbadge{
  display:inline-block;
  padding:1px 7px;
  border-radius:999px;
  font-size:10px;
  font-weight:700;
  background:rgba(255,255,255,.1);
  color:var(--fg);
  min-width:18px;
  text-align:center;
}
.sh-fbadge-pass{ background:rgba(52,211,153,.2); color:#34d399; }
.sh-fbadge-fail{ background:rgba(248,113,113,.2); color:#f87171; }
.sh-fbadge-warn{ background:rgba(251,191,36,.2);  color:#fbbf24; }
.sh-fbadge-info{ background:rgba(96,165,250,.2);  color:#60a5fa; }
/* Session host table: filter bar sits above, so square top corners on the tbl-wrap */
#sec-sh .tbl-wrap{
  border-radius: 0 0 var(--radius) var(--radius);
  border-top: none;
}
.sh-no-results{
  display:flex;
  align-items:center;
  justify-content:center;
  gap:10px;
  padding:36px;
  color:var(--dim);
  font-size:14px;
  border:1px solid var(--border);
  border-top:none;
  border-radius:0 0 var(--radius) var(--radius);
  background:var(--bg-glass);
}

/* ─── Tables ───────────────────────────────────────────────── */
.tbl-wrap{
  overflow-x:auto;
  border:1px solid var(--border);
  border-radius:var(--radius);
  background:var(--bg-glass);
  backdrop-filter:blur(8px);
}
table{width:100%;border-collapse:collapse}
thead th{
  background:rgba(16,24,40,.8);
  color:var(--dim);
  font-size:10px;
  text-transform:uppercase;
  letter-spacing:.1em;
  font-weight:700;
  font-family:var(--font-sans);
  padding:12px 16px;
  text-align:left;
  border-bottom:1px solid var(--border2);
  white-space:nowrap;
}
tbody tr.expandable{
  cursor:pointer;
  transition:background .2s, transform .15s;
}
tbody tr.expandable:hover{background:rgba(56,189,248,.04)}
tbody tr.expandable.open{background:rgba(56,189,248,.07)}
tbody td{
  padding:11px 16px;
  border-bottom:1px solid var(--border);
  vertical-align:middle;
}
tbody tr:last-child td{border-bottom:none}
.detail-row{display:none}
.detail-row td{
  background:rgba(10,16,32,.85);
  border-bottom:1px solid var(--border2) !important;
  padding:16px 20px !important;
}
tbody tr.expandable.open + .detail-row{display:table-row}

/* ─── Key/value in overview table ─────────────────────────── */
td.k{color:var(--dim);font-size:11px;text-transform:uppercase;letter-spacing:.08em;width:180px;padding:11px 16px;font-family:var(--font-sans);font-weight:600}
td.v{color:var(--white);padding:11px 16px}
td.v.accent{color:var(--neon);font-family:var(--font-sans);font-size:18px;font-weight:700}
td.v.mono{font-family:var(--font-mono);font-size:12px;color:var(--dim)}

/* ─── Badges ───────────────────────────────────────────────── */
.badge{
  display:inline-flex;
  align-items:center;
  gap:4px;
  padding:4px 10px;
  border-radius:20px;
  font-size:11px;
  font-weight:600;
  letter-spacing:.03em;
  white-space:nowrap;
  font-family:var(--font-mono);
}
.badge-pass{color:var(--neon2);background:rgba(52,211,153,.1);border:1px solid rgba(52,211,153,.25)}
.badge-warn{color:var(--amber);background:rgba(251,191,36,.1);border:1px solid rgba(251,191,36,.25)}
.badge-fail{color:var(--red);background:rgba(244,63,94,.1);border:1px solid rgba(244,63,94,.25)}
.badge-info{color:var(--neon);background:rgba(56,189,248,.08);border:1px solid rgba(56,189,248,.2)}

/* ─── Detail expand grid ───────────────────────────────────── */
.detail-grid{
  display:flex;
  flex-wrap:wrap;
  gap:16px 32px;
}
.dg-item{display:flex;flex-direction:column;gap:3px;min-width:160px}
.dk{font-size:10px;color:var(--dim);text-transform:uppercase;letter-spacing:.08em;font-family:var(--font-sans);font-weight:600}
.dv{font-size:13px;color:var(--white)}
.dv.mono{font-family:var(--font-mono);font-size:12px}
.fail-hosts{
  background:rgba(244,63,94,.06);
  border-left:3px solid var(--red);
  padding:10px 14px;
  border-radius:6px;
  margin-bottom:10px;
  font-size:12px;
}
.fh{color:var(--red)}
.note-box{
  background:rgba(56,189,248,.06);
  border-left:3px solid var(--neon);
  padding:10px 14px;
  border-radius:6px;
  margin-bottom:10px;
  font-size:12px;
  color:var(--neon);
}

/* ─── Typography helpers ───────────────────────────────────── */
.nm{color:var(--white);font-weight:600}
.sm{color:var(--dim);font-size:12px}
.cnt{color:var(--white);font-family:var(--font-sans);font-weight:700;font-size:16px;text-align:center}
.mono{font-family:var(--font-mono)}

/* ─── Row expand chevron ───────────────────────────────────── */
tr.expandable td.nm::before{
  content:'\25B8 ';
  color:var(--dim);
  font-size:10px;
  transition:transform .25s;
  display:inline-block;
  margin-right:5px;
}
tr.expandable.open td.nm::before{transform:rotate(90deg);color:var(--neon)}

/* ─── Footer ───────────────────────────────────────────────── */
footer{
  margin-top:48px;
  padding:20px 24px;
  background:var(--bg-glass);
  backdrop-filter:blur(8px);
  border:1px solid var(--border);
  border-radius:var(--radius);
  color:var(--dim);
  font-size:11px;
  display:flex;
  align-items:center;
  justify-content:space-between;
  flex-wrap:wrap;
  gap:8px;
  font-family:var(--font-sans);
}
footer .neon{color:var(--neon)}

/* ─── Scroll progress bar ─────────────────────────────────── */
.scroll-progress{
  position:fixed;top:0;left:0;height:3px;z-index:100;
  background:linear-gradient(90deg,var(--neon),var(--purple),var(--neon2));
  width:0;
  transition:width .1s linear;
  border-radius:0 2px 2px 0;
  box-shadow:0 0 12px rgba(56,189,248,.4);
}

/* ─── Floating nav dots ───────────────────────────────────── */
.nav-dots{
  position:fixed;right:20px;top:50%;transform:translateY(-50%);z-index:50;
  display:flex;flex-direction:column;gap:10px;
  animation:fadeIn .6s ease .5s both;
}
.nav-dot{
  width:10px;height:10px;border-radius:50%;
  background:var(--border2);
  border:1px solid var(--border);
  cursor:pointer;
  transition:all .25s;
  position:relative;
}
.nav-dot:hover,.nav-dot.active{background:var(--neon);border-color:var(--neon);box-shadow:0 0 10px rgba(56,189,248,.4)}
.nav-dot::after{
  content:attr(data-label);
  position:absolute;right:20px;top:50%;transform:translateY(-50%);
  font-size:10px;color:var(--dim);font-family:var(--font-sans);font-weight:600;
  text-transform:uppercase;letter-spacing:.06em;white-space:nowrap;
  opacity:0;transition:opacity .2s;pointer-events:none;
}
.nav-dot:hover::after{opacity:1;color:var(--white)}

/* ─── ANF Volume inner table ──────────────────────────────── */
.anf-sub-head{
  margin:0 0 8px;
  font-size:11px;
  color:var(--dim);
  text-transform:uppercase;
  letter-spacing:.08em;
  font-family:var(--font-sans);
  font-weight:700;
}
.vol-tbl-wrap{
  margin-bottom:16px;
  border-radius:8px;
  overflow:hidden;
  border:1px solid var(--border2);
  background:rgba(5,8,15,.6);
}
.vol-inner-tbl{width:100%;border-collapse:collapse}
.vol-inner-tbl thead th{
  background:rgba(10,16,32,.9);
  color:var(--dim);
  font-size:10px;
  text-transform:uppercase;
  letter-spacing:.09em;
  font-weight:700;
  font-family:var(--font-sans);
  padding:9px 14px;
  text-align:left;
  border-bottom:1px solid var(--border2);
  white-space:nowrap;
}
.vol-inner-tbl tbody tr{border-bottom:1px solid var(--border)}
.vol-inner-tbl tbody tr:last-child{border-bottom:none}
.vol-inner-tbl tbody tr:hover{background:rgba(56,189,248,.04)}
.vol-inner-tbl tbody td{padding:9px 14px;vertical-align:middle}
td.vn{color:var(--white);font-weight:600;font-family:var(--font-mono);font-size:12px}
td.vq{color:var(--text);font-family:var(--font-mono);font-size:12px;white-space:nowrap}
td.vp{min-width:160px}

/* ─── Usage bar ───────────────────────────────────────────── */
.usage-wrap{display:flex;align-items:center;gap:8px}
.usage-bar{
  flex:1;
  height:7px;
  background:rgba(255,255,255,.07);
  border-radius:4px;
  overflow:hidden;
  min-width:80px;
}
.usage-fill{height:100%;border-radius:4px;transition:width .4s ease}
.bar-ok  {background:linear-gradient(90deg,#34d399,#059669)}
.bar-warn{background:linear-gradient(90deg,#fbbf24,#d97706)}
.bar-crit{background:linear-gradient(90deg,#f43f5e,#be123c);box-shadow:0 0 6px rgba(244,63,94,.4)}
.usage-pct{font-size:11px;font-family:var(--font-mono);font-weight:600;white-space:nowrap;min-width:46px;text-align:right}
.pct-ok  {color:#34d399}
.pct-warn{color:#fbbf24}
.pct-crit{color:#f43f5e}

/* ─── Health Check Per-Host Table ──────────────────────────── */
.hc-per-host-head{
  font-size:11px; letter-spacing:.08em; text-transform:uppercase;
  color:#f87171; margin:10px 0 6px 0;
}
.hc-all-pass{
  padding:14px 16px;
  color:#34d399;
  font-size:13px;
  background:rgba(52,211,153,.06);
  border:1px solid rgba(52,211,153,.2);
  border-radius:8px;
  margin:8px 0;
}
.ph-table{ width:100%; border-collapse:collapse; font-size:12px; }
.ph-table thead th{
  background:var(--bg3); color:var(--dim); font-size:10px;
  letter-spacing:.06em; text-transform:uppercase;
  padding:6px 10px; border-bottom:1px solid var(--border); text-align:left;
}
.ph-table tbody tr{ border-bottom:1px solid rgba(255,255,255,.04); }
.ph-table tbody tr:last-child{ border-bottom:none; }
.ph-table td{ padding:7px 10px; vertical-align:top; }
.ph-host{ font-family:var(--font-mono); font-size:11px; white-space:nowrap; color:var(--fg); min-width:200px; }
.ph-msg{ font-family:var(--font-mono); font-size:10px; color:var(--dim); margin-bottom:2px; word-break:break-all; }
.ph-dt{ font-size:10px; color:var(--dim); opacity:.6; display:block; margin-top:2px; }
.ph-pass td{ background:rgba(52,211,153,.03); }
.ph-fail td{ background:rgba(248,113,113,.05); }
.ph-est  td{ background:rgba(251,191,36,.04); }
.ph-msg-td{ width:100%; }

/* ─── Responsive ───────────────────────────────────────────── */
@media(max-width:900px){
  .stat-strip{grid-template-columns:repeat(auto-fill, minmax(130px, 1fr))}
  .nav-dots{display:none}
}
@media(max-width:700px){
  .header{grid-template-columns:1fr;padding:16px 20px}
  .score-block{align-items:flex-start}
  .stat-strip{grid-template-columns:repeat(2, 1fr);gap:8px}
  .wrap{padding:16px 12px 40px}
}
</style>
</head>
<body>
<div class="scroll-progress" id="scrollProg"></div>
<nav class="nav-dots">
  <div class="nav-dot active" data-label="Dashboard"     onclick="document.getElementById('sec-stats').scrollIntoView({behavior:'smooth'})"></div>
  <div class="nav-dot"        data-label="Overview"      onclick="document.getElementById('sec-overview').scrollIntoView({behavior:'smooth'})"></div>
  <div class="nav-dot"        data-label="Host Pools"    onclick="document.getElementById('sec-hp').scrollIntoView({behavior:'smooth'})"></div>
  <div class="nav-dot"        data-label="Session Hosts" onclick="document.getElementById('sec-sh').scrollIntoView({behavior:'smooth'})"></div>
  <div class="nav-dot"        data-label="Health Checks" onclick="document.getElementById('sec-hc').scrollIntoView({behavior:'smooth'})"></div>
  <div class="nav-dot"        data-label="$(ConvertTo-HtmlSafe $ProfileStorageType) Storage" onclick="document.getElementById('sec-anf').scrollIntoView({behavior:'smooth'})"></div>
</nav>
<div class="wrap">

  <!-- HEADER -->
  <div class="header">
    <div class="brand">
      <div class="brand-title">
        <span class="hi">AVD</span><span class="lo">/</span>Environment<span class="lo">-</span>Health
      </div>
      <div class="brand-sub">
        $(ConvertTo-HtmlSafe $ClientName) &nbsp;|&nbsp; $(ConvertTo-HtmlSafe $EnvironmentName) &nbsp;|&nbsp;
        Azure Virtual Desktop &nbsp;|&nbsp; v$script:ToolVersion &nbsp;|&nbsp;
        Resource Group: <span>$rg</span>
      </div>
    </div>
    <div class="score-block">
      $(if ($BrandDisplayName) { "<div class=`"brand-wordmark`">$(ConvertTo-HtmlSafe $BrandDisplayName)</div>" })
      <div class="status-pill $(
          if ($overallHealth -eq 'OPTIMAL') { 'pill-green' }
          elseif ($overallHealth -eq 'DEGRADED') { 'pill-amber' }
          else { 'pill-red' }
      )">
        <span class="status-pill-dot"></span>
        $overallHealth
      </div>
    </div>
  </div>

  <!-- STAT STRIP -->
  <div class="stat-strip" id="sec-stats">
    <div class="stat-card">
      <div class="sc-num info">$hpCount</div>
      <div class="sc-label">Host Pools</div>
    </div>
    <div class="stat-card">
      <div class="sc-num info">$shCount</div>
      <div class="sc-label">Session Hosts</div>
    </div>
    <div class="stat-card$(if($healthySH -ne $totalSH -and $totalSH -gt 0){' has-dd'}else{''})" $(if($healthySH -ne $totalSH -and $totalSH -gt 0){'onclick="toggleStat(this)"'}else{''})>
      <div class="sc-num $(if($healthySH -eq $totalSH -and $totalSH -gt 0){'ok'}elseif($healthySH -ge ($totalSH * .8)){'warn'}else{'bad'})">$healthySH<span style="font-size:18px;color:var(--dim)">/$totalSH</span></div>
      <div class="sc-label">Healthy Hosts$(if($healthySH -ne $totalSH -and $totalSH -gt 0){' <span class="sc-toggle">&#9660;</span>'}else{''})</div>
      $(if($healthySH -ne $totalSH -and $totalSH -gt 0){"<div class='sc-dropdown'><ul><li>$unhealthyNames</li></ul></div>"}else{''})
    </div>
    <div class="stat-card$(if($drainSH -gt 0){' has-dd'}else{''})" $(if($drainSH -gt 0){'onclick="toggleStat(this)"'}else{''})>
      <div class="sc-num $(if($drainSH -eq 0){'ok'}else{'warn'})">$drainSH</div>
      <div class="sc-label">In Drain Mode$(if($drainSH -gt 0){' <span class="sc-toggle">&#9660;</span>'}else{''})</div>
      $(if($drainSH -gt 0){"<div class='sc-dropdown'><ul><li>$drainNames</li></ul></div>"}else{''})
    </div>
    <div class="stat-card$(if($staleSH -gt 0){' has-dd'}else{''})" $(if($staleSH -gt 0){'onclick="toggleStat(this)"'}else{''})>
      <div class="sc-num $(if($staleSH -eq 0){'ok'}else{'warn'})">$staleSH</div>
      <div class="sc-label">Stale (&gt;2 Days)$(if($staleSH -gt 0){' <span class="sc-toggle">&#9660;</span>'}else{''})</div>
      $(if($staleSH -gt 0){"<div class='sc-dropdown'><ul><li>$staleNames</li></ul></div>"}else{''})
    </div>
    <div class="stat-card$(if($noAZsH -gt 0){' has-dd'}else{''})" $(if($noAZsH -gt 0){'onclick="toggleStat(this)"'}else{''})>
      <div class="sc-num $(if($noAZsH -eq 0){'ok'}else{'warn'})">$noAZsH</div>
      <div class="sc-label">No AZ Pin$(if($noAZsH -gt 0){' <span class="sc-toggle">&#9660;</span>'}else{''})</div>
      $(if($noAZsH -gt 0){"<div class='sc-dropdown'><ul><li>$noAZnames</li></ul></div>"}else{''})
    </div>
    <div class="stat-card$(if($outdatedMon -gt 0){' has-dd'}else{''})" $(if($outdatedMon -gt 0){'onclick="toggleStat(this)"'}else{''})>
      <div class="sc-num $(if($outdatedMon -eq 0){'ok'}else{'warn'})">$outdatedMon</div>
      <div class="sc-label">Outdated Monitor Agent$(if($outdatedMon -gt 0){' <span class="sc-toggle">&#9660;</span>'}else{''})</div>
      $(if($outdatedMon -gt 0){"<div class='sc-dropdown'><ul><li>$outdatedNames</li></ul></div>"}else{''})
    </div>
    <div class="stat-card$(if($hcFail -gt 0){' has-dd'}else{''})" $(if($hcFail -gt 0){'onclick="toggleStat(this)"'}else{''})>
      <div class="sc-num $(if($hcFail -eq 0){'ok'}else{'bad'})">$hcFail</div>
      <div class="sc-label">HC Failures$(if($hcFail -gt 0){' <span class="sc-toggle">&#9660;</span>'}else{''})</div>
      $(if($hcFail -gt 0){"<div class='sc-dropdown'><ul><li>$hcFailNames</li></ul></div>"}else{''})
    </div>
    <div class="stat-card$(if($hpScaling -gt 0){' has-dd'}else{''})" $(if($hpScaling -gt 0){'onclick="toggleStat(this)"'}else{''})>
      <div class="sc-num $(if($hpScaling -eq 0){'ok'}else{'warn'})">$hpScaling</div>
      <div class="sc-label">No Scaling Plan$(if($hpScaling -gt 0){' <span class="sc-toggle">&#9660;</span>'}else{''})</div>
      $(if($hpScaling -gt 0){"<div class='sc-dropdown'><ul><li>$hpScaleNames</li></ul></div>"}else{''})
    </div>
    <div class="stat-card$(if($hpPublic -gt 0){' has-dd'}else{''})" $(if($hpPublic -gt 0){'onclick="toggleStat(this)"'}else{''})>
      <div class="sc-num $(if($hpPublic -eq 0){'ok'}else{'warn'})">$hpPublic</div>
      <div class="sc-label">Public Network On$(if($hpPublic -gt 0){' <span class="sc-toggle">&#9660;</span>'}else{''})</div>
      $(if($hpPublic -gt 0){"<div class='sc-dropdown'><ul><li>$hpPubNames</li></ul></div>"}else{''})
    </div>
    <div class="stat-card$(if($anfColFail -gt 0){' has-dd'}else{''})" $(if($anfColFail -gt 0){'onclick="toggleStat(this)"'}else{''})>
      <div class="sc-num $(if($anfColFail -eq 0){'ok'}else{'bad'})">$anfColFail</div>
      <div class="sc-label">$storageStatLabel Cross-Region$(if($anfColFail -gt 0){' <span class="sc-toggle">&#9660;</span>'}else{''})</div>
      $(if($anfColFail -gt 0){"<div class='sc-dropdown'><ul><li>$storageColNames</li></ul></div>"}else{''})
    </div>
    <div class="stat-card$(if($anfNearQuota -gt 0){' has-dd'}else{''})" $(if($anfNearQuota -gt 0){'onclick="toggleStat(this)"'}else{''})>
      <div class="sc-num $(if($anfNearQuota -eq 0){'ok'}else{'warn'})">$anfNearQuota</div>
      <div class="sc-label">$storageStatLabel $volumeUnitLabel &ge;$($QuotaWarningPercent)%$(if($anfNearQuota -gt 0){' <span class="sc-toggle">&#9660;</span>'}else{''})</div>
      $(if($anfNearQuota -gt 0){"<div class='sc-dropdown'><ul><li>$storageNQNames</li></ul></div>"}else{''})
    </div>
  </div>

  <!-- 1. OVERVIEW -->
  <div class="section" id="sec-overview">
    <div class="sec-header">
      <span class="sec-icon">&#128202;</span>
      <span class="sec-title">Overview</span>
      <span class="sec-count">Environment Identity &amp; Summary</span>
    </div>
    <div class="tbl-wrap">
      <table>
        <tbody>$overviewRows</tbody>
      </table>
    </div>
  </div>

  <!-- 2. HOST POOLS -->
  <div class="section" id="sec-hp">
    <div class="sec-header">
      <span class="sec-icon">&#127970;</span>
      <span class="sec-title">Host Pools</span>
      <span class="sec-count">$($script:HostPoolData.Count) pool(s) — click row to expand</span>
    </div>
    <div class="tbl-wrap">
      <table>
        <thead>
          <tr>
            <th>Name</th>
            <th>Type</th>
            <th>Load Balancer</th>
            <th>Max Sessions</th>
            <th>Ring</th>
            <th>Scaling Plan</th>
            <th>Start VM on Connect</th>
            <th>Network</th>
          </tr>
        </thead>
        <tbody>$($hpRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <!-- 3. SESSION HOSTS -->
  <div class="section" id="sec-sh">
    <div class="sec-header">
      <span class="sec-icon">&#128187;</span>
      <span class="sec-title">Session Hosts</span>
      <span class="sec-count" id="sh-visible-count">$($script:SessionHostData.Count) host(s) — click row to expand</span>
    </div>

    <!-- ── Filter Bar ── -->
    <div class="sh-filter-bar">
      <span class="sh-filter-label">&#9906; Filter by Health Status:</span>
      <button class="sh-filter-btn sh-filter-active" data-filter="All"        onclick="filterSH(this)">All <span class="sh-fbadge" id="sh-cnt-all">$($script:SessionHostData.Count)</span></button>
      <button class="sh-filter-btn" data-filter="Available"                    onclick="filterSH(this)">Available <span class="sh-fbadge sh-fbadge-pass" id="sh-cnt-available">0</span></button>
      <button class="sh-filter-btn" data-filter="Unavailable"                  onclick="filterSH(this)">Unavailable <span class="sh-fbadge sh-fbadge-fail" id="sh-cnt-unavailable">0</span></button>
      <button class="sh-filter-btn" data-filter="NeedsAssistance"              onclick="filterSH(this)">Needs Assistance <span class="sh-fbadge sh-fbadge-warn" id="sh-cnt-needsassistance">0</span></button>
      <button class="sh-filter-btn" data-filter="Shutdown"                     onclick="filterSH(this)">Shutdown <span class="sh-fbadge sh-fbadge-info" id="sh-cnt-shutdown">0</span></button>
      <button class="sh-filter-btn" data-filter="Drain"                        onclick="filterSH(this)">Drain ON <span class="sh-fbadge sh-fbadge-warn" id="sh-cnt-drain">0</span></button>
    </div>

    <div class="tbl-wrap">
      <table id="sh-table">
        <thead>
          <tr>
            <th>Name</th>
            <th>Host Pool</th>
            <th>Health Status</th>
            <th>Drain Mode</th>
            <th>Boot Status</th>
            <th>Availability Zone</th>
            <th>Monitor Agent</th>
            <th>OS Disk</th>
          </tr>
        </thead>
        <tbody>$($shRows.ToString())</tbody>
      </table>
      <div class="sh-no-results" id="sh-no-results" style="display:none">
        <span>&#128269;</span> No session hosts match the selected filter.
      </div>
    </div>
  </div>

  <!-- 4. HEALTH CHECKS -->
  <div class="section" id="sec-hc">
    <div class="sec-header">
      <span class="sec-icon">&#128271;</span>
      <span class="sec-title">Health Check of Session Hosts</span>
      <span class="sec-count">$($script:HealthCheckData.Count) checks — click row to expand</span>
    </div>
    <div class="tbl-wrap">
      <table>
        <thead>
          <tr>
            <th>Check Name</th>
            <th>Status</th>
            <th>Failing / Total</th>
            <th>Data Source</th>
          </tr>
        </thead>
        <tbody>$($hcRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <!-- 5. PROFILE STORAGE (ANF / AZURE FILES) -->
  <div class="section" id="sec-anf">
    <div class="sec-header">
      <span class="sec-icon">&#128190;</span>
      <span class="sec-title">Profile Storage &mdash; $(ConvertTo-HtmlSafe $ProfileStorageType)</span>
      <span class="sec-count">$($script:AnfData.Count) account(s) assessed — click row to expand</span>
    </div>
    <div class="tbl-wrap">
      <table>
        <thead>
          <tr>
            <th>Storage Account</th>
            <th>Pools / SKU</th>
            <th>Region Colocation</th>
            <th>Volumes/Shares &ge;$($QuotaWarningPercent)%</th>
            <th>Data Protection</th>
            <th>Backup Config</th>
          </tr>
        </thead>
        <tbody>$($anfRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <footer>
    <span>AVD-Environment Health <span class="neon">v$script:ToolVersion</span> &nbsp;|&nbsp; Resource Group: <span class="neon">$rg</span></span>
    <span>Generated: $ts</span>
  </footer>

</div><!-- /wrap -->

<script>
function toggleRow(row){row.classList.toggle('open')}
function toggleStat(card){card.classList.toggle('sc-open')}

// ── Session Host Filter ───────────────────────────────────────────────────
(function initSHFilter(){
  var rows    = document.querySelectorAll('#sh-table tbody tr.sh-row');
  var details = document.querySelectorAll('#sh-table tbody tr.sh-detail');
  var noRes   = document.getElementById('sh-no-results');
  var visEl   = document.getElementById('sh-visible-count');

  // Count each status for badges
  var counts = { all:0, available:0, unavailable:0, needsassistance:0, shutdown:0, drain:0 };
  rows.forEach(function(r){
    var st = (r.getAttribute('data-status')||'').toLowerCase().replace(/\s/g,'');
    counts.all++;
    if(st==='available')       counts.available++;
    else if(st==='unavailable')counts.unavailable++;
    else if(st==='needsassistance') counts.needsassistance++;
    else if(st==='shutdown'||st==='deallocated') counts.shutdown++;
    // Drain: check the drain badge text in the 4th cell
    if(r.querySelector('td:nth-child(4) .badge-warn')) counts.drain++;
  });

  function setCount(id, val){ var el=document.getElementById(id); if(el) el.textContent=val; }
  setCount('sh-cnt-all', counts.all);
  setCount('sh-cnt-available', counts.available);
  setCount('sh-cnt-unavailable', counts.unavailable);
  setCount('sh-cnt-needsassistance', counts.needsassistance);
  setCount('sh-cnt-shutdown', counts.shutdown);
  setCount('sh-cnt-drain', counts.drain);
})();

function filterSH(btn){
  // Update active button state
  document.querySelectorAll('.sh-filter-btn').forEach(function(b){ b.classList.remove('sh-filter-active'); });
  btn.classList.add('sh-filter-active');

  var filter = (btn.getAttribute('data-filter')||'All').toLowerCase();
  var rows    = document.querySelectorAll('#sh-table tbody tr.sh-row');
  var details = document.querySelectorAll('#sh-table tbody tr.sh-detail');
  var noRes   = document.getElementById('sh-no-results');
  var visEl   = document.getElementById('sh-visible-count');
  var visible = 0;

  rows.forEach(function(row, i){
    var st    = (row.getAttribute('data-status')||'').toLowerCase().replace(/\s/g,'');
    var detail = details[i];
    var show = false;

    if(filter === 'all'){
      show = true;
    } else if(filter === 'drain'){
      // Drain filter: match rows that have a Drain ON badge in column 4
      show = !!row.querySelector('td:nth-child(4) .badge-warn');
    } else {
      show = (st === filter);
    }

    if(show){
      row.style.display = '';
      visible++;
    } else {
      row.style.display = 'none';
      // Also collapse open detail row when filtered out
      if(row.classList.contains('open')){ row.classList.remove('open'); }
    }
    // Always hide detail rows — they toggle via click
    if(detail) detail.style.display = show ? '' : 'none';
  });

  // Update visible count label
  var total = rows.length;
  if(visEl){
    if(filter === 'all'){
      visEl.textContent = total + ' host(s) — click row to expand';
    } else {
      visEl.textContent = visible + ' of ' + total + ' host(s) shown — click row to expand';
    }
  }

  // Show empty-state message when no results
  if(noRes){
    noRes.style.display = (visible === 0) ? 'flex' : 'none';
  }
}

// Scroll progress bar
window.addEventListener('scroll',function(){
  var h=document.documentElement,b=document.body;
  var st=h.scrollTop||b.scrollTop;
  var sh=h.scrollHeight-h.clientHeight;
  var pct=sh>0?(st/sh)*100:0;
  document.getElementById('scrollProg').style.width=pct+'%';
});

// Nav dot active state
(function(){
  var ids=['sec-stats','sec-overview','sec-hp','sec-sh','sec-hc','sec-anf'];
  var dots=document.querySelectorAll('.nav-dot');
  window.addEventListener('scroll',function(){
    var st=window.scrollY+200;
    var active=0;
    ids.forEach(function(id,i){
      var el=document.getElementById(id);
      if(el&&el.offsetTop<=st)active=i;
    });
    dots.forEach(function(d,i){d.classList.toggle('active',i===active)});
  });
})();
</script>
</body>
</html>
"@
}

# ==============================================================================
# MAIN
# ==============================================================================

function Invoke-Main {
    $storageLabel = switch ($ProfileStorageType) {
        'AzureFiles' { 'AzureFiles (Storage Account)' }
        'Both'       { 'ANF + AzureFiles' }
        default      { 'AzureNetAppFiles (ANF)' }
    }

    Write-Host ''
    Write-Host '  ╔══════════════════════════════════════════════╗' -ForegroundColor Cyan
    Write-Host "  ║  AVD-Environment Health  v$script:ToolVersion             ║" -ForegroundColor Cyan
    Write-Host '  ║  Azure Virtual Desktop Health Report         ║' -ForegroundColor Cyan
    Write-Host '  ║  Sections: Overview | HostPools | SessionHosts ║' -ForegroundColor Cyan
    Write-Host "  ║  HealthChecks | ProfileStorage: $($storageLabel.PadRight(14))║" -ForegroundColor Cyan
    Write-Host '  ╚══════════════════════════════════════════════╝' -ForegroundColor Cyan
    Write-Host ''

    $script:hostPoolTagsLocal = @{}

    if ($DryRun) {
        Write-Host '  [Mode] DRY RUN — no Azure calls.' -ForegroundColor Yellow
        Initialize-DryRunData
    } else {
        # Live Azure run
        $missing = @($script:RequiredModules | Where-Object { -not (Get-Module -ListAvailable -Name $_) })
        if ($missing.Count -gt 0) {
            Write-Host "  ERROR: Missing modules: $($missing -join ', ')" -ForegroundColor Red
            throw 'Install missing modules and retry.'
        }
        $script:RequiredModules | ForEach-Object { Import-Module $_ -ErrorAction Stop | Out-Null }

        Connect-ToAzure
        Get-OverviewData

        if ($script:allHostPools.Count -eq 0) {
            Write-Warn 'No host pools found. Report will reflect empty environment.'
        } else {
            # Fetch tags for host pools
            foreach ($hp in $script:allHostPools) {
                try {
                    $res = Invoke-WithRetry -ScriptBlock {
                        Get-AzResource -ResourceId $hp.Id -ErrorAction Stop
                    }
                    $script:hostPoolTagsLocal[$hp.Id] = $res.Tags
                } catch { }
            }

            Get-HostPoolData
            Get-SessionHostData
            Get-HealthCheckData

            switch ($ProfileStorageType) {
                'ANF'        { Get-AnfData }
                'AzureFiles' { Get-AzureFilesData }
                'Both'       { Get-AnfData; Get-AzureFilesData }
                default      { Get-AnfData }
            }
        }
    }

    # Determine output path
    if (-not $OutputPath) {
        $stamp        = (Get-Date).ToString('yyyyMMdd-HHmmss')
        $safeClient   = ($ClientName -replace '[^a-zA-Z0-9_-]', '')
        $safeEnv      = ($EnvironmentName -replace '[^a-zA-Z0-9_-]', '')
        $OutputPath   = Join-Path (Get-Location).Path "AVD-EnvironmentHealth-$safeClient-$safeEnv-$stamp.html"
    } elseif ($OutputPath -notmatch '\.html$') {
        $OutputPath = "$OutputPath.html"
    }

    $outDir = Split-Path -Parent $OutputPath
    if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

    Write-Step 'Generating HTML report'
    $html = New-HtmlReport
    [System.IO.File]::WriteAllText($OutputPath, $html, [System.Text.UTF8Encoding]::new($false))
    Write-Ok "Report written: $OutputPath"

    if ($OpenReport -and (Test-Path $OutputPath)) {
        try { Start-Process -FilePath $OutputPath | Out-Null } catch { }
    }

    Write-Host ''
    Write-Host '  Done.' -ForegroundColor Green
    Write-Host ''
}

Invoke-Main
