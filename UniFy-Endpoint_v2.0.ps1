<#PSScriptInfo

.VERSION: 2.0

.AUTHOR: UniFy-Endpoint

.MAINTAINER: Yoennis Olmo

.DESCRIPTION: Enterprise-grade backup and restore solution for Microsoft Intune configurations with advanced platform filtering, pagination support, and comprehensive policy coverage including Autopilot, ESP, Scripts, Remediations, WUfB, Assignment Filters, and App Protection policies.

.RELEASENOTES

Version 2.0:
- NEW: Windows Autopilot Profile support
- NEW: Autopilot Device Preparation Policies
- NEW: Enrollment Status Page (ESP) support
- NEW: Remediation Scripts support
- NEW: Windows Update for Business (WUfB) Policies
- NEW: Assignment Filter support
- NEW: App Configuration Policies (Managed Devices & Managed Apps)
- NEW: Platform filtering system (-Platform parameter)
- NEW: Nested assignment exporting for all policies
- ENHANCED: Clean JSON protocol for import-ready exports
- ENHANCED: Robust pagination for large tenants (1000+ policies)
- ENHANCED: Base64 script handling for all script types
- ENHANCED: Error handling with detailed diagnostics
- ENHANCED: Interactive menu system

#>

param(
    [Parameter(Mandatory = $false, HelpMessage = "Perform backup operation")]
    [switch]$Backup,

    [Parameter(Mandatory = $false, HelpMessage = "Perform restore operation")]
    [switch]$Restore,

    [Parameter(Mandatory = $false, HelpMessage = "Compare two backups")]
    [switch]$Compare,

    [Parameter(Mandatory = $false, HelpMessage = "List available backups")]
    [switch]$List,

    [Parameter(Mandatory = $false, HelpMessage = "Detect configuration drift")]
    [switch]$DriftDetection,

    [Parameter(Mandatory = $false, HelpMessage = "Export backup to different formats")]
    [switch]$Export,

    [Parameter(Mandatory = $false, HelpMessage = "Import configuration from file")]
    [switch]$Import,

    [Parameter(Mandatory = $false, HelpMessage = "Platform filter (Windows, iOS, Android, macOS, or All)")]
    [ValidateSet("All", "Windows", "iOS", "Android", "macOS")]
    [string]$Platform = "All",

    [Parameter(Mandatory = $false, HelpMessage = "Custom backup path")]
    [string]$BackupPath,

    [Parameter(Mandatory = $false, HelpMessage = "Export format (JSON, CSV, HTML)")]
    [ValidateSet("JSON", "CSV", "HTML", "All")]
    [string]$ExportFormat = "JSON",

    [Parameter(Mandatory = $false, HelpMessage = "Import file path")]
    [string]$ImportFile,

    [Parameter(Mandatory = $false, HelpMessage = "Disable logging (enabled by default)")]
    [switch]$DisableLogging,

    [Parameter(Mandatory = $false, HelpMessage = "Log file path")]
    [string]$LogPath,

    [Parameter(Mandatory = $false, HelpMessage = "Log level (Verbose, Info, Warning, Error)")]
    [ValidateSet("Verbose", "Info", "Warning", "Error")]
    [string]$LogLevel = "Info",

    [Parameter(Mandatory = $false, HelpMessage = "Retention days for old backups")]
    [int]$RetentionDays = 30,

    [Parameter(Mandatory = $false, HelpMessage = "Environment (Global, USGov, USGovDoD)")]
    [ValidateSet("Global", "USGov", "USGovDoD")]
    [string]$Environment = "Global",

    [Parameter(Mandatory = $false, HelpMessage = "App ID for authentication")]
    [string]$AppId,

    [Parameter(Mandatory = $false, HelpMessage = "Tenant ID for authentication")]
    [string]$TenantId,

    [Parameter(Mandatory = $false, HelpMessage = "Certificate Thumbprint for authentication")]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $false, HelpMessage = "Client Secret for app registration authentication")]
    [string]$ClientSecret,

    [Parameter(Mandatory = $false, HelpMessage = "Components to include in backup/restore")]
    [ValidateSet("All", "DeviceConfigurations", "CompliancePolicies", "SettingsCatalogPolicies",
                 "AppProtectionPolicies", "PowerShellScripts", "AutopilotProfiles",
                 "AutopilotDevicePrep", "EnrollmentStatusPage", "RemediationScripts",
                 "AssignmentFilters", "AppConfigManagedDevices", "AppConfigManagedApps",
                 "MacOSScripts", "MacOSCustomAttributes")]
    [string[]]$Components = @("All")
)

# Script configuration
$script:Version = "2.0.0"
$script:DefaultBackupPath = Join-Path $env:USERPROFILE "Documents\UniFy-Endpoint\backups"
$script:BackupLocation = if ($BackupPath) { $BackupPath } else { $script:DefaultBackupPath }
$script:GraphEnvironment = $Environment
$script:DefaultLogPath = Join-Path $env:USERPROFILE "Documents\UniFy-Endpoint\logs"
$script:LogLocation = if ($LogPath) { $LogPath } else { $script:DefaultLogPath }
$script:DefaultReportsPath = Join-Path $env:USERPROFILE "Documents\UniFy-Endpoint\reports"
$script:ReportsLocation = $script:DefaultReportsPath
$script:DefaultExportsPath = Join-Path $env:USERPROFILE "Documents\UniFy-Endpoint\exports"
$script:ExportsLocation = $script:DefaultExportsPath
$script:IsLiveReport = $false
$script:LoggingEnabled = -not $DisableLogging
$script:CurrentLogLevel = $LogLevel
$script:LogFile = $null
$script:SelectedPlatform = $Platform

# Color configuration
$script:Colors = @{
    Success = "Green"
    Error   = "Red"
    Warning = "Yellow"
    Info    = "Cyan"
    Header  = "Magenta"
    Default = "White"
}

# Platform mapping for filtering (per SKILL.md specification)
# Includes all platform enum values used by different policy types
$script:PlatformMapping = @{
    "Windows" = @("windows10AndLater", "windows", "windows10", "windows10X", "Windows")
    "iOS"     = @("iOS", "ios")
    "Android" = @("android", "androidEnterprise", "Android")
    "macOS"   = @("macOS", "macos")
}

#region Logging Functions

function Initialize-Logging {
    if ($script:LoggingEnabled) {
        if (-not (Test-Path $script:LogLocation)) {
            New-Item -ItemType Directory -Path $script:LogLocation -Force | Out-Null
        }

        $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
        $script:LogFile = Join-Path $script:LogLocation "UniFy-Endpoint_v2_$timestamp.log"

        Write-Log "========================================" -Level Info
        Write-Log "UniFy-Endpoint v$($script:Version) - Session Started" -Level Info
        Write-Log "========================================" -Level Info
        Write-Log "Log Level: $($script:CurrentLogLevel)" -Level Info
        Write-Log "Backup Location: $($script:BackupLocation)" -Level Info
        Write-Log "Log Location: $($script:LogLocation)" -Level Info
        Write-Log "Platform Filter: $($script:SelectedPlatform)" -Level Info
        Write-Log "Tenant: Connecting..." -Level Info
    }
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("Verbose", "Info", "Warning", "Error")]
        [string]$Level = "Info",
        [switch]$NoConsole
    )

    if (-not $script:LoggingEnabled) {
        return
    }

    # Check log level
    $levels = @{
        "Verbose" = 0
        "Info"    = 1
        "Warning" = 2
        "Error"   = 3
    }

    if ($levels[$Level] -lt $levels[$script:CurrentLogLevel]) {
        return
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $logEntry -Encoding UTF8
    }

    # Also write to console if not suppressed
    if (-not $NoConsole) {
        switch ($Level) {
            "Error" { Write-Host $Message -ForegroundColor Red }
            "Warning" { Write-Host $Message -ForegroundColor Yellow }
            "Info" { Write-Host $Message -ForegroundColor Cyan }
            "Verbose" { Write-Host $Message -ForegroundColor Gray }
        }
    }
}

function Close-Logging {
    if ($script:LoggingEnabled -and $script:LogFile) {
        Write-Log "UniFy-Endpoint v2.0 Session Ended" -Level Info
        Write-Log "================================================" -Level Info
    }
}

#endregion

#region Helper Functions

function Write-ColoredOutput {
    param(
        [string]$Message,
        [string]$Color = "White",
        [switch]$NoNewline
    )

    $params = @{
        Object          = $Message
        ForegroundColor = $Color
    }

    if ($NoNewline) {
        $params.NoNewline = $true
    }

    Write-Host @params

    # Also log if logging is enabled
    if ($script:LoggingEnabled) {
        $logLevel = switch ($Color) {
            "Red" { "Error" }
            "Yellow" { "Warning" }
            "Cyan" { "Info" }
            "Green" { "Info" }
            default { "Verbose" }
        }
        Write-Log $Message -Level $logLevel -NoConsole
    }
}

function Show-Banner {
    Clear-Host

    Write-Host ""
    Write-Host ""

    Write-Host "  Version: v$($script:Version)" -ForegroundColor White
    Write-Host "  Intelligent Intune Configuration Backup & Restore" -ForegroundColor Gray
    Write-Host "  By: UniFy-Endpoint" -ForegroundColor Gray

    $context = Get-MgContext
    if ($context) {
        Write-Host ""
        Write-Host "  Connected to Tenant: " -ForegroundColor Green -NoNewline
        Write-Host "$($context.TenantId)" -ForegroundColor White
        Write-Host "  Account: " -ForegroundColor Green -NoNewline
        Write-Host "$($context.Account)" -ForegroundColor White
    }
    Write-Host ""
}

function Get-SafeFileName {
    param (
        [string]$FileName
    )

    # Remove invalid characters for file system
    $invalid = [System.IO.Path]::GetInvalidFileNameChars() -join ''
    $regex = "[{0}]" -f [RegEx]::Escape($invalid)
    $safeName = $FileName -replace $regex, "_"

    # Replace other problematic characters including brackets
    $safeName = $safeName -replace '[:/\\*?"<>|\[\]]', '_'
    $safeName = $safeName -replace '\s+', ' '
    $safeName = $safeName.Trim()

    # Ensure filename is not too long (max 200 chars)
    if ($safeName.Length -gt 200) {
        $safeName = $safeName.Substring(0, 200)
    }

    return $safeName
}

function Get-PolicyTypeFromJson {
    <#
    .SYNOPSIS
    Detects the policy type from JSON content and returns the correct folder name.

    .DESCRIPTION
    Analyzes the @odata.type property and other indicators to determine the
    policy type and return the matching UniFy-Endpoint folder name for restore.

    .PARAMETER Policy
    The policy object (parsed from JSON)

    .EXAMPLE
    $folderName = Get-PolicyTypeFromJson -Policy $policyObject
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Policy
    )

    $odataType = $Policy.'@odata.type'

    # Legacy Intune Security Baselines (deviceManagement/intents API - old format)
    if ($odataType -eq "#microsoft.graph.deviceManagementIntent") {
        return "LegacyIntents"
    }

    # Settings Catalog Policies (including Autopilot Device Prep)
    if ($odataType -eq "#microsoft.graph.deviceManagementConfigurationPolicy" -or
        $Policy.settings -or
        $Policy.templateReference) {
        return "SettingsCatalogPolicies"
    }

    # Compliance Policies
    if ($odataType -like "*CompliancePolicy*" -or
        $odataType -like "*compliancePolicy*") {
        return "CompliancePolicies"
    }

    # Administrative Templates (Group Policy) — must be checked BEFORE Device Configurations
    # because '#microsoft.graph.groupPolicyConfiguration' also matches the "*Configuration*"
    # pattern below and would be mis-routed to DeviceConfigurations without this ordering.
    if ($odataType -eq "#microsoft.graph.groupPolicyConfiguration" -or
        $Policy.definitionValues) {
        return "AdministrativeTemplates"
    }

    # Device Configurations
    if ($odataType -like "*Configuration*" -and $odataType -notlike "*ConfigurationPolicy*") {
        return "DeviceConfigurations"
    }

    # App Protection Policies
    if ($odataType -like "*ProtectionPolicy*" -or
        $odataType -like "*ManagedAppProtection*") {
        return "AppProtectionPolicies"
    }

    # PowerShell Scripts
    if ($odataType -eq "#microsoft.graph.deviceManagementScript" -or
        $Policy.scriptContent) {
        return "PowerShellScripts"
    }

    # Remediation Scripts (Proactive Remediations)
    if ($odataType -eq "#microsoft.graph.deviceHealthScript" -or
        ($Policy.detectionScriptContent -and $Policy.remediationScriptContent)) {
        return "RemediationScripts"
    }

    # macOS Scripts
    if ($odataType -eq "#microsoft.graph.deviceShellScript") {
        return "MacOSScripts"
    }

    # macOS Custom Attributes
    if ($odataType -eq "#microsoft.graph.deviceCustomAttributeShellScript") {
        return "MacOSCustomAttributes"
    }

    # Autopilot Profiles
    if ($odataType -eq "#microsoft.graph.windowsAutopilotDeploymentProfile" -or
        $odataType -eq "#microsoft.graph.azureADWindowsAutopilotDeploymentProfile" -or
        $odataType -eq "#microsoft.graph.activeDirectoryWindowsAutopilotDeploymentProfile") {
        return "AutopilotProfiles"
    }

    # Enrollment Status Page
    if ($odataType -like "*EnrollmentCompletionPageConfiguration*") {
        return "EnrollmentStatusPage"
    }

    # Windows Update for Business
    if ($odataType -like "*windowsUpdateForBusiness*" -or
        $odataType -like "*WindowsQualityUpdateProfile*" -or
        $odataType -like "*WindowsFeatureUpdateProfile*" -or
        $odataType -like "*WindowsDriverUpdateProfile*") {
        return "WUfBPolicies"
    }

    # Assignment Filters
    if ($odataType -eq "#microsoft.graph.deviceAndAppManagementAssignmentFilter" -or
        $Policy.rule) {
        return "AssignmentFilters"
    }

    # App Configuration - Managed Devices
    if ($odataType -like "*managedDeviceMobileAppConfiguration*") {
        return "AppConfigManagedDevices"
    }

    # App Configuration - Managed Apps
    if ($odataType -like "*targetedManagedAppConfiguration*") {
        return "AppConfigManagedApps"
    }

    # Default to Settings Catalog if has 'name' property (common for Settings Catalog)
    # or if it has typical Settings Catalog structure
    if ($Policy.name -and -not $Policy.displayName) {
        return "SettingsCatalogPolicies"
    }

    # Fallback - try Settings Catalog as it's the most common modern policy type
    return "SettingsCatalogPolicies"
}

function Ensure-BackupDirectory {
    if (-not (Test-Path $script:BackupLocation)) {
        Write-ColoredOutput "Creating backup directory: $script:BackupLocation" -Color $script:Colors.Info
        New-Item -ItemType Directory -Path $script:BackupLocation -Force | Out-Null
    }
}

function Get-SelectedComponents {
    param (
        [string[]]$ComponentList
    )

    # NOTE: AutopilotDevicePrep is excluded - it's already exported via Settings Catalog
    # AutopilotDevicePrep is only available in component-specific backup (option 12 -> 1 -> 5)
    $allComponents = @(
        "DeviceConfigurations",
        "CompliancePolicies",
        "SettingsCatalogPolicies",
        "AppProtectionPolicies",
        "PowerShellScripts",
        "AutopilotProfiles",
        "EnrollmentStatusPage",
        "RemediationScripts",
        "AssignmentFilters",
        "AppConfigManagedDevices",
        "AppConfigManagedApps",
        "MacOSScripts",
        "MacOSCustomAttributes",
        "LegacyIntents"
    )

    if ($ComponentList -contains "All") {
        return $allComponents
    }
    else {
        return $ComponentList
    }
}

function Test-PlatformMatch {
    <#
    .SYNOPSIS
    Tests if a policy matches the selected platform filter.

    .DESCRIPTION
    Implements platform filtering logic per SKILL.md specification.
    Checks platform or appType properties against the platform mapping.

    .PARAMETER Policy
    The policy object to test

    .EXAMPLE
    Test-PlatformMatch -Policy $policy
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Policy
    )

    # If platform is set to "All", include everything
    if ($script:SelectedPlatform -eq "All") {
        return $true
    }

    # Get the platform values to match
    $platformValues = $script:PlatformMapping[$script:SelectedPlatform]

    if (-not $platformValues) {
        return $true
    }

    # Check the platform property
    if ($Policy.platform) {
        if ($platformValues -contains $Policy.platform) {
            return $true
        }
    }

    # Check the appType property (for app-related policies)
    if ($Policy.appType) {
        if ($platformValues -contains $Policy.appType) {
            return $true
        }
    }

    # Check platforms array (for some policy types)
    if ($Policy.platforms) {
        foreach ($plat in $Policy.platforms) {
            if ($platformValues -contains $plat) {
                return $true
            }
        }
    }

    # Check @odata.type for platform hints
    if ($Policy.'@odata.type') {
        $odataType = $Policy.'@odata.type'

        # Windows-specific types
        if ($script:SelectedPlatform -eq "Windows" -and
            ($odataType -like "*windows*" -or $odataType -like "*win10*")) {
            return $true
        }

        # iOS-specific types
        if ($script:SelectedPlatform -eq "iOS" -and
            ($odataType -like "*ios*" -or $odataType -like "*iPhone*" -or $odataType -like "*iPad*")) {
            return $true
        }

        # Android-specific types
        if ($script:SelectedPlatform -eq "Android" -and
            ($odataType -like "*android*")) {
            return $true
        }

        # macOS-specific types
        if ($script:SelectedPlatform -eq "macOS" -and
            ($odataType -like "*macOS*" -or $odataType -like "*mac*")) {
            return $true
        }
    }

    # If no platform info found, include it (could be cross-platform)
    return $false
}

function ConvertTo-PolicyObject {
    <#
    .SYNOPSIS
    Converts an API response (Hashtable/Dictionary) to a PSCustomObject for proper serialization.

    .DESCRIPTION
    Invoke-MgGraphRequest returns a Hashtable/Dictionary. When using Add-Member on these,
    NoteProperties are not serialized by ConvertTo-Json. This function converts the response
    to a PSCustomObject to ensure all properties (including id, createdDateTime, lastModifiedDateTime)
    are properly preserved when saving to JSON.

    .PARAMETER Response
    The API response object (typically a Hashtable from Invoke-MgGraphRequest)

    .EXAMPLE
    $policy = ConvertTo-PolicyObject -Response $apiResponse
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Response
    )

    if ($null -eq $Response) {
        return $null
    }

    # If it's already a PSCustomObject, return as-is
    if ($Response -is [PSCustomObject]) {
        return $Response
    }

    # Convert Hashtable/Dictionary to PSCustomObject
    $obj = [PSCustomObject]@{}
    foreach ($key in $Response.Keys) {
        # Skip OData metadata
        if ($key -eq '@odata.context') { continue }

        $value = $Response[$key]

        # Recursively convert nested hashtables
        if ($value -is [System.Collections.IDictionary]) {
            $value = ConvertTo-PolicyObject -Response $value
        }
        elseif ($value -is [System.Collections.IList]) {
            $convertedList = @()
            foreach ($item in $value) {
                if ($item -is [System.Collections.IDictionary]) {
                    $convertedList += ConvertTo-PolicyObject -Response $item
                }
                else {
                    $convertedList += $item
                }
            }
            $value = $convertedList
        }

        $obj | Add-Member -NotePropertyName $key -NotePropertyValue $value -Force
    }

    return $obj
}

# Comprehensive list of properties to ignore when comparing policies (metadata that changes or is computed)
$script:ComparisonIgnoreProperties = @(
    # Standard metadata
    'id',
    'createdDateTime',
    'lastModifiedDateTime',
    'modifiedDateTime',
    'version',

    # OData metadata
    '@odata.context',
    '@odata.type',
    '@odata.id',
    '@odata.etag',

    # Scope and role tags
    'roleScopeTagIds',
    'supportsScopeTags',
    'deviceManagementApplicabilityRuleOsEdition',
    'deviceManagementApplicabilityRuleOsVersion',
    'deviceManagementApplicabilityRuleDeviceMode',

    # Creation/modification metadata
    'creationSource',
    'isAssigned',
    'priority',
    'createdBy',
    'lastModifiedBy',

    # Assignment-related
    'assignments',
    'assignmentFilterEvaluationStatusDetails',

    # Template references (can include GUIDs)
    'templateReference',
    'templateId',
    'templateDisplayName',
    'templateDisplayVersion',
    'templateFamily',

    # Settings Catalog specific
    'settingCount',
    'settingDefinitionId',
    'settingInstanceTemplateId',
    'settingValueTemplateId',
    'settingValueTemplateReference',
    'settingInstanceTemplateReference',
    'odataType',

    # Additional GUID-based fields
    'secretReferenceValueId',
    'deviceConfigurationId',
    'groupId',
    'sourceId',
    'payloadId',

    # Autopilot and enrollment specific
    'deviceNameTemplate',
    'azureAdJoinType',
    'managementServiceAppId',

    # App protection and configuration specific
    'targetedAppManagementLevels',
    'appGroupType',
    'deployedAppCount',
    'apps',
    'deploymentSummary',
    'minimumRequiredPatchVersion',
    'minimumRequiredSdkVersion',
    'minimumWipeSdkVersion',
    'minimumWipePatchVersion',
    'minimumRequiredAppVersion',
    'minimumWipeAppVersion',
    'minimumRequiredOsVersion',
    'minimumWipeOsVersion',

    # Script specific (hashes change based on encoding)
    'scriptContentMd5Hash',
    'scriptContentSha256Hash',
    'detectionScriptContentMd5Hash',
    'detectionScriptContentSha256Hash',
    'remediationScriptContentMd5Hash',
    'remediationScriptContentSha256Hash',

    # Compliance policy specific
    'scheduledActionsForRule',
    'validOperatingSystemBuildRanges',

    # Remediation script specific
    'runSummary',
    'deviceRunStates',
    'isGlobalScript',
    'deviceHealthScriptType',
    'runSchedule',

    # PowerShell script runtime fields
    'fileName',
    'runAsAccount',
    'enforceSignatureCheck',
    'runAs32Bit',

    # App Configuration (Managed Apps) specific
    'customSettings',
    'encodedSettingXml',
    'targetedManagedAppGroupType',
    'configurationDeploymentSummaryPerApp',
    'customBrowserProtocol',
    'customBrowserPackageId',
    'customBrowserDisplayName',
    'customDialerAppProtocol',
    'customDialerAppPackageId',
    'customDialerAppDisplayName',
    'allowedAndroidDeviceManufacturers',
    'appActionIfAndroidDeviceManufacturerNotAllowed',
    'exemptedAppProtocols',
    'exemptedAppPackages',

    # Navigation properties (populated by $expand, not actual config data)
    'settingDefinitions',

    # Additional runtime/computed fields
    'technologies',
    'platformType',
    'settingsCount',
    'priorityMetaData',
    'assignedAppsCount'
)

function Remove-IgnoredProperties {
    param ($obj, $ignoreList)

    if ($null -eq $obj) { return $null }

    if ($obj -is [array]) {
        $cleanArray = @()
        foreach ($item in $obj) {
            $cleanArray += Remove-IgnoredProperties -obj $item -ignoreList $ignoreList
        }
        return $cleanArray
    }
    elseif ($obj -is [PSCustomObject] -or $obj -is [System.Collections.IDictionary]) {
        $cleanObj = [PSCustomObject]@{}
        $props = if ($obj -is [System.Collections.IDictionary]) { $obj.Keys } else { $obj.PSObject.Properties.Name }

        foreach ($propName in $props) {
            # Skip properties in the ignore list, OData annotations, and OData bound actions/functions
            if ($propName -notin $ignoreList -and $propName -notlike '*@odata*' -and $propName -notlike '#*') {
                $value = if ($obj -is [System.Collections.IDictionary]) { $obj[$propName] } else { $obj.$propName }
                $cleanValue = Remove-IgnoredProperties -obj $value -ignoreList $ignoreList
                $cleanObj | Add-Member -NotePropertyName $propName -NotePropertyValue $cleanValue -Force
            }
        }
        return $cleanObj
    }
    else {
        return $obj
    }
}

function Compare-PolicyContent {
    <#
    .SYNOPSIS
    Compares two policy objects to detect content differences.

    .DESCRIPTION
    Compares the settings/content of a backup policy with a tenant policy,
    ignoring metadata fields like id, createdDateTime, lastModifiedDateTime, version.

    .PARAMETER BackupPolicy
    The policy object from the backup file

    .PARAMETER TenantPolicy
    The policy object from the tenant (API response)

    .PARAMETER IgnoreProperties
    Additional properties to ignore during comparison

    .RETURNS
    PSCustomObject with IsEqual (bool) and Differences (array of strings)
    #>
    param (
        [Parameter(Mandatory = $true)]
        $BackupPolicy,

        [Parameter(Mandatory = $true)]
        $TenantPolicy,

        [Parameter(Mandatory = $false)]
        [string[]]$IgnoreProperties = @()
    )

    $allIgnore = $script:ComparisonIgnoreProperties + $IgnoreProperties

    $differences = @()

    # Convert both to JSON and back (normalizes the structure)
    $backupJson = $BackupPolicy | ConvertTo-Json -Depth 50 -Compress | ConvertFrom-Json
    $tenantJson = $TenantPolicy | ConvertTo-Json -Depth 50 -Compress | ConvertFrom-Json

    # Recursively remove ignored properties from both objects
    $cleanBackup = Remove-IgnoredProperties -obj $backupJson -ignoreList $allIgnore
    $cleanTenant = Remove-IgnoredProperties -obj $tenantJson -ignoreList $allIgnore

    # Normalize Base64 script content fields
    # Backup stores decoded text (human-readable), API returns Base64-encoded.
    # Decode any Base64 values so both sides are in plain text for comparison.
    $base64Fields = @('scriptContent', 'detectionScriptContent', 'remediationScriptContent')
    foreach ($field in $base64Fields) {
        foreach ($obj in @($cleanBackup, $cleanTenant)) {
            if ($null -ne $obj -and $obj.PSObject.Properties[$field]) {
                $val = $obj.$field
                if ($val -and $val -is [string]) {
                    try {
                        $decoded = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($val))
                        $obj | Add-Member -NotePropertyName $field -NotePropertyValue $decoded -Force
                    } catch {
                        # Not valid Base64 — already decoded text, leave as-is
                    }
                }
            }
        }
    }

    # Get all top-level properties for reporting differences
    $backupProps = @{}
    $tenantProps = @{}

    if ($cleanBackup.PSObject.Properties) {
        $cleanBackup.PSObject.Properties | ForEach-Object {
            $backupProps[$_.Name] = $_.Value
        }
    }

    if ($cleanTenant.PSObject.Properties) {
        $cleanTenant.PSObject.Properties | ForEach-Object {
            $tenantProps[$_.Name] = $_.Value
        }
    }

    # Compare properties
    $allKeys = @($backupProps.Keys) + @($tenantProps.Keys) | Select-Object -Unique

    foreach ($key in $allKeys) {
        $backupValue = $backupProps[$key]
        $tenantValue = $tenantProps[$key]

        # Convert to JSON for deep comparison
        $backupValueJson = if ($null -ne $backupValue) { $backupValue | ConvertTo-Json -Depth 50 -Compress } else { "null" }
        $tenantValueJson = if ($null -ne $tenantValue) { $tenantValue | ConvertTo-Json -Depth 50 -Compress } else { "null" }

        if ($backupValueJson -ne $tenantValueJson) {
            $differences += $key
        }
    }

    return [PSCustomObject]@{
        IsEqual = ($differences.Count -eq 0)
        Differences = $differences
        DifferenceCount = $differences.Count
    }
}

function Get-PolicyHash {
    <#
    .SYNOPSIS
    Computes SHA256 hash of a cleaned policy object for fast equality comparison.

    .DESCRIPTION
    Strips metadata/read-only properties (using the same ignore list as Compare-PolicyContent),
    normalizes Base64 script content, then computes SHA256 of the canonical JSON.
    Two policies with matching hashes are functionally identical.

    .PARAMETER Policy
    The policy object to hash.

    .RETURNS
    String - uppercase hex SHA256 hash.
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Policy
    )

    # Same ignore list as Compare-PolicyContent for consistency
    $ignoreList = @(
        'id', 'createdDateTime', 'lastModifiedDateTime', 'modifiedDateTime', 'version',
        '@odata.context', '@odata.type', '@odata.id', '@odata.etag',
        'roleScopeTagIds', 'supportsScopeTags',
        'deviceManagementApplicabilityRuleOsEdition', 'deviceManagementApplicabilityRuleOsVersion',
        'deviceManagementApplicabilityRuleDeviceMode',
        'creationSource', 'isAssigned', 'priority', 'createdBy', 'lastModifiedBy',
        'assignments', 'assignmentFilterEvaluationStatusDetails',
        'templateReference', 'templateId', 'templateDisplayName', 'templateDisplayVersion', 'templateFamily',
        'settingCount', 'settingDefinitionId', 'settingInstanceTemplateId',
        'settingValueTemplateId', 'settingValueTemplateReference', 'settingInstanceTemplateReference', 'odataType',
        'secretReferenceValueId', 'deviceConfigurationId', 'groupId', 'sourceId', 'payloadId',
        'deviceNameTemplate', 'azureAdJoinType', 'managementServiceAppId',
        'targetedAppManagementLevels', 'appGroupType', 'deployedAppCount', 'apps', 'deploymentSummary',
        'minimumRequiredPatchVersion', 'minimumRequiredSdkVersion', 'minimumWipeSdkVersion', 'minimumWipePatchVersion',
        'minimumRequiredAppVersion', 'minimumWipeAppVersion', 'minimumRequiredOsVersion', 'minimumWipeOsVersion',
        'scriptContentMd5Hash', 'scriptContentSha256Hash',
        'detectionScriptContentMd5Hash', 'detectionScriptContentSha256Hash',
        'remediationScriptContentMd5Hash', 'remediationScriptContentSha256Hash',
        'scheduledActionsForRule', 'validOperatingSystemBuildRanges',
        'runSummary', 'deviceRunStates', 'isGlobalScript', 'deviceHealthScriptType', 'runSchedule',
        'fileName', 'runAsAccount', 'enforceSignatureCheck', 'runAs32Bit',
        'customSettings', 'encodedSettingXml', 'targetedManagedAppGroupType',
        'configurationDeploymentSummaryPerApp',
        'customBrowserProtocol', 'customBrowserPackageId', 'customBrowserDisplayName',
        'customDialerAppProtocol', 'customDialerAppPackageId', 'customDialerAppDisplayName',
        'allowedAndroidDeviceManufacturers', 'appActionIfAndroidDeviceManufacturerNotAllowed',
        'exemptedAppProtocols', 'exemptedAppPackages',
        'settingDefinitions',
        'technologies', 'platformType', 'settingsCount', 'priorityMetaData', 'assignedAppsCount'
    )

    # Recursive cleaner (same logic as Compare-PolicyContent's Remove-IgnoredProperties)
    function Remove-ForHash {
        param ($obj, $ignore)
        if ($null -eq $obj) { return $null }
        if ($obj -is [array]) {
            $arr = @()
            foreach ($item in $obj) { $arr += Remove-ForHash -obj $item -ignore $ignore }
            return $arr
        }
        elseif ($obj -is [PSCustomObject] -or $obj -is [System.Collections.IDictionary]) {
            $clean = [PSCustomObject]@{}
            $props = if ($obj -is [System.Collections.IDictionary]) { $obj.Keys } else { $obj.PSObject.Properties.Name }
            foreach ($p in $props) {
                if ($p -notin $ignore -and $p -notlike '*@odata*' -and $p -notlike '#*') {
                    $val = if ($obj -is [System.Collections.IDictionary]) { $obj[$p] } else { $obj.$p }
                    $clean | Add-Member -NotePropertyName $p -NotePropertyValue (Remove-ForHash -obj $val -ignore $ignore) -Force
                }
            }
            return $clean
        }
        return $obj
    }

    $normalized = $Policy | ConvertTo-Json -Depth 50 -Compress | ConvertFrom-Json
    $cleaned = Remove-ForHash -obj $normalized -ignore $ignoreList

    # Normalize Base64 script content fields
    $base64Fields = @('scriptContent', 'detectionScriptContent', 'remediationScriptContent')
    foreach ($field in $base64Fields) {
        if ($null -ne $cleaned -and $cleaned.PSObject.Properties[$field]) {
            $val = $cleaned.$field
            if ($val -and $val -is [string]) {
                try {
                    $decoded = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($val))
                    $cleaned | Add-Member -NotePropertyName $field -NotePropertyValue $decoded -Force
                } catch { }
            }
        }
    }

    $json = $cleaned | ConvertTo-Json -Depth 100 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
    return ([System.BitConverter]::ToString($hash) -replace '-', '')
}

function Get-PolicyExistenceStatus {
    <#
    .SYNOPSIS
    Checks policy existence on tenant by ID and name, returning detailed status.

    .DESCRIPTION
    Implements hybrid checking:
    1. First checks if the original policy ID still exists on tenant
    2. If ID exists, compares content to detect modifications
    3. If ID doesn't exist, checks if a policy with the same name exists
    4. Returns detailed status for restore decision

    .PARAMETER BackupPolicy
    The policy object from the backup file

    .PARAMETER ApiEndpoint
    The Graph API endpoint to query (e.g., "deviceManagement/deviceConfigurations")

    .PARAMETER ExistingPolicies
    Optional: Pre-fetched list of existing policies (for endpoints that don't support $filter)

    .PARAMETER RestoredName
    The name the policy will have after restore (with [Restored] prefix)

    .PARAMETER OriginalName
    The original policy name from the backup

    .PARAMETER NameField
    The field name used for the policy name (default: 'displayName', use 'name' for Settings Catalog)

    .RETURNS
    PSCustomObject with:
    - Status: 'Unchanged', 'Modified', 'Renamed', 'NameConflict', 'Deleted', 'New'
    - ExistingPolicy: The matching policy if found
    - Differences: Array of changed properties (if Modified)
    - Message: Human-readable status message
    #>
    param (
        [Parameter(Mandatory = $true)]
        $BackupPolicy,

        [Parameter(Mandatory = $true)]
        [string]$ApiEndpoint,

        [Parameter(Mandatory = $false)]
        $ExistingPolicies = $null,

        [Parameter(Mandatory = $true)]
        [string]$RestoredName,

        [Parameter(Mandatory = $true)]
        [string]$OriginalName,

        [Parameter(Mandatory = $false)]
        [string]$NameField = 'displayName',

        [Parameter(Mandatory = $false)]
        [switch]$UseV1Api,

        [Parameter(Mandatory = $false)]
        [string]$ExpandQuery = ''
    )

    $result = [PSCustomObject]@{
        Status = 'New'
        ExistingPolicy = $null
        Differences = @()
        DifferenceCount = 0
        Message = 'Policy not found on tenant'
        FoundById = $false
        FoundByName = $false
        IdChanged = $false
        NameChanged = $false
        SettingsChanged = $false
        TenantName = ''
        TenantId = ''
        BackupHash = ''
        TenantHash = ''
    }

    $apiVersion = if ($UseV1Api) { "v1.0" } else { "beta" }
    $backupId = $BackupPolicy.id

    # Step 1: Check by ID (if backup has an ID)
    if ($backupId) {
        try {
            $idCheckUri = "https://graph.microsoft.com/$apiVersion/$ApiEndpoint/$backupId"
            if ($ExpandQuery) { $idCheckUri += "?`$expand=$ExpandQuery" }
            $existingById = Invoke-MgGraphRequest -Uri $idCheckUri -Method GET -ErrorAction SilentlyContinue

            if ($existingById) {
                $result.FoundById = $true
                $result.ExistingPolicy = $existingById
                $result.TenantId = $existingById.id
                $tenantName = $existingById.$NameField
                $result.TenantName = $tenantName

                # Hash-based comparison (fast path)
                $bHash = Get-PolicyHash -Policy $BackupPolicy
                $tHash = Get-PolicyHash -Policy $existingById
                $result.BackupHash = $bHash
                $result.TenantHash = $tHash

                if ($bHash -eq $tHash) {
                    # Hashes match - policy is clean
                    if ($tenantName -eq $OriginalName) {
                        $result.Status = 'Unchanged'
                        $result.Message = "Policy exists with same ID and content (hash match)"
                    } else {
                        $result.Status = 'Renamed'
                        $result.NameChanged = $true
                        $result.Message = "Policy renamed on tenant: '$tenantName' (hash match)"
                    }
                } else {
                    # Hashes differ - run detailed comparison to confirm actual differences
                    # (Hash can differ due to property ordering while content is identical)
                    $comparison = Compare-PolicyContent -BackupPolicy $BackupPolicy -TenantPolicy $existingById

                    if ($comparison.IsEqual) {
                        # Hash differed but no actual content differences - treat as clean
                        if ($tenantName -eq $OriginalName) {
                            $result.Status = 'Unchanged'
                            $result.Message = "Policy exists with same ID and content (hash match)"
                        } else {
                            $result.Status = 'Renamed'
                            $result.NameChanged = $true
                            $result.Message = "Policy renamed on tenant: '$tenantName' (hash match)"
                        }
                    } else {
                        # Filter out name fields - a name change is a rename, not a settings change
                        $settingsDiffs = @($comparison.Differences | Where-Object { $_ -ne 'name' -and $_ -ne 'displayName' })
                        $nameChanged = ($tenantName -ne $OriginalName)

                        if ($settingsDiffs.Count -eq 0 -and $nameChanged) {
                            # Only the name field changed - treat as rename
                            $result.Status = 'Renamed'
                            $result.NameChanged = $true
                            $result.Message = "Policy renamed on tenant: '$tenantName'"
                        } elseif ($settingsDiffs.Count -gt 0) {
                            # Genuine settings differences found
                            $result.SettingsChanged = $true
                            $result.Differences = $settingsDiffs
                            $result.DifferenceCount = $settingsDiffs.Count

                            if ($nameChanged) {
                                $result.Status = 'Modified'
                                $result.NameChanged = $true
                                $result.Message = "Policy renamed to '$tenantName' + settings modified ($($settingsDiffs.Count) changes)"
                            } else {
                                $result.Status = 'Modified'
                                $result.Message = "Policy modified on tenant ($($settingsDiffs.Count) changes)"
                            }
                        } else {
                            # Only name fields in diff but names match - treat as unchanged
                            $result.Status = 'Unchanged'
                            $result.Message = "Policy exists with same ID and content"
                        }
                    }
                }

                return $result
            }
        }
        catch {
            # ID not found - policy was deleted or ID doesn't exist
            Write-Log "Policy ID $backupId not found on tenant: $($_.Exception.Message)" -Level Verbose
        }
    }

    # Step 2: Check by name (original name and restored name)
    $foundByName = $null

    if ($ExistingPolicies) {
        # Use pre-fetched list (for APIs that don't support $filter)
        $foundByName = $ExistingPolicies | Where-Object { $_.$NameField -eq $OriginalName }
        if (-not $foundByName) {
            $foundByName = $ExistingPolicies | Where-Object { $_.$NameField -eq $RestoredName }
        }
    } else {
        # Query API with $filter
        try {
            $escapedOriginalName = Get-ODataFilterValue -Value $OriginalName
            $nameCheckUri = "https://graph.microsoft.com/$apiVersion/$ApiEndpoint`?`$filter=$NameField eq '$escapedOriginalName'"
            $foundByName = Get-GraphData -Uri $nameCheckUri -SuppressErrors

            if (-not $foundByName -and $OriginalName -ne $RestoredName) {
                $escapedRestoredName = Get-ODataFilterValue -Value $RestoredName
                $nameCheckUri = "https://graph.microsoft.com/$apiVersion/$ApiEndpoint`?`$filter=$NameField eq '$escapedRestoredName'"
                $foundByName = Get-GraphData -Uri $nameCheckUri -SuppressErrors
            }
        }
        catch {
            Write-Log "Error checking policy by name: $($_.Exception.Message)" -Level Verbose
        }
    }

    if ($foundByName) {
        # Re-fetch with $expand if needed (e.g., Settings Catalog needs $expand=settings)
        if ($ExpandQuery -and $foundByName.id) {
            try {
                $expandUri = "https://graph.microsoft.com/$apiVersion/$ApiEndpoint/$($foundByName.id)?`$expand=$ExpandQuery"
                $foundByName = Invoke-MgGraphRequest -Uri $expandUri -Method GET
            } catch {
                Write-Log "Error expanding policy details: $($_.Exception.Message)" -Level Verbose
            }
        }

        $result.FoundByName = $true
        $result.ExistingPolicy = $foundByName

        # Hash-based comparison (fast path)
        $bHash = Get-PolicyHash -Policy $BackupPolicy
        $tHash = Get-PolicyHash -Policy $foundByName
        $result.BackupHash = $bHash
        $result.TenantHash = $tHash

        # Policy exists with same name but different ID (or backup had no ID)
        if ($backupId -and $foundByName.id -ne $backupId) {
            $result.IdChanged = $true
            $result.TenantId = $foundByName.id
            $result.TenantName = $foundByName.$NameField

            if ($bHash -eq $tHash) {
                $result.Status = 'NameConflict'
                $result.SettingsChanged = $false
                $result.Message = "Policy ID changed (settings unchanged - hash match) - Tenant ID: $($foundByName.id)"
            } else {
                $comparison = Compare-PolicyContent -BackupPolicy $BackupPolicy -TenantPolicy $foundByName
                $result.Status = 'NameConflict'
                if ($comparison.IsEqual) {
                    $result.SettingsChanged = $false
                    $result.Message = "Policy ID changed (settings unchanged) - Tenant ID: $($foundByName.id)"
                } else {
                    $settingsDiffs = @($comparison.Differences | Where-Object { $_ -ne 'name' -and $_ -ne 'displayName' })
                    if ($settingsDiffs.Count -gt 0) {
                        $result.SettingsChanged = $true
                        $result.Differences = $settingsDiffs
                        $result.DifferenceCount = $settingsDiffs.Count
                        $result.Message = "Policy ID changed + settings modified ($($settingsDiffs.Count) changes)"
                    } else {
                        $result.SettingsChanged = $false
                        $result.Message = "Policy ID changed (settings unchanged) - Tenant ID: $($foundByName.id)"
                    }
                }
            }
        } else {
            if ($bHash -eq $tHash) {
                $result.Status = 'Unchanged'
                $result.Message = "Policy exists with same name and content (hash match)"
            } else {
                $comparison = Compare-PolicyContent -BackupPolicy $BackupPolicy -TenantPolicy $foundByName
                $settingsDiffs = @($comparison.Differences | Where-Object { $_ -ne 'name' -and $_ -ne 'displayName' })
                if ($comparison.IsEqual -or $settingsDiffs.Count -eq 0) {
                    # Hash differed but no actual settings differences - treat as clean
                    $result.Status = 'Unchanged'
                    $result.Message = "Policy exists with same name and content"
                } else {
                    $result.Status = 'Modified'
                    $result.SettingsChanged = $true
                    $result.Differences = $settingsDiffs
                    $result.DifferenceCount = $settingsDiffs.Count
                    $result.TenantId = $foundByName.id
                    $result.TenantName = $foundByName.$NameField
                    $result.Message = "Policy modified on tenant ($($settingsDiffs.Count) changes)"
                }
            }
        }

        return $result
    }

    # Step 3: Policy not found by ID or name - it was deleted or is new
    if ($backupId) {
        $result.Status = 'Deleted'
        $result.Message = 'Policy was deleted from tenant'
    } else {
        $result.Status = 'New'
        $result.Message = 'New policy (not found on tenant)'
    }

    return $result
}

function Write-RestoreStatus {
    <#
    .SYNOPSIS
    Writes the restore status message and updates statistics based on policy existence check.

    .DESCRIPTION
    Helper function that displays appropriate status messages for restore operations
    based on the existence status check, and updates the restore statistics.

    .PARAMETER Status
    The status object from Get-PolicyExistenceStatus

    .PARAMETER Preview
    Whether this is a preview/dry-run operation

    .PARAMETER RestoreStats
    Reference to the restore statistics hashtable

    .PARAMETER IsImport
    Whether this is an import operation (affects display labels)

    .RETURNS
    Boolean - $true if restore should proceed, $false if should skip
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Status,

        [Parameter(Mandatory = $true)]
        [bool]$Preview,

        [Parameter(Mandatory = $true)]
        [hashtable]$RestoreStats,

        [Parameter(Mandatory = $false)]
        [bool]$IsImport = $false
    )

    $shouldRestore = $false

    switch ($Status.Status) {
        'Unchanged' {
            if ($IsImport) {
                Write-ColoredOutput " [SKIPPED - Already Exists]" -Color $script:Colors.Warning
            } elseif ($Preview) {
                Write-ColoredOutput " [UNCHANGED]" -Color $script:Colors.Success
            } else {
                Write-ColoredOutput " [UNCHANGED]" -Color $script:Colors.Success
            }
            $RestoreStats.Unchanged++
            $RestoreStats.Skipped++
        }
        'Modified' {
            if ($IsImport) {
                Write-ColoredOutput " [SKIPPED - Already Exists]" -Color $script:Colors.Warning
            } elseif ($Preview) {
                # Settings changed - [CHANGED]
                $settingWord = if ($Status.DifferenceCount -eq 1) { 'Setting' } else { 'Settings' }
                $changeMsg = " [CHANGED] - [$($Status.DifferenceCount) $settingWord Differ]"
                if ($Status.NameChanged) { $changeMsg += " [RENAMED] -> [$($Status.TenantName)]" }
                Write-ColoredOutput $changeMsg -Color $script:Colors.Warning
            } else {
                Write-ColoredOutput " [SKIPPED - CHANGED]" -Color $script:Colors.Warning
            }
            $RestoreStats.Modified++
            $RestoreStats.Skipped++
        }
        'Renamed' {
            if ($IsImport) {
                Write-ColoredOutput " [SKIPPED - Already Exists]" -Color $script:Colors.Warning
            } else {
                $tenantDisplayName = if ($Status.TenantName) { $Status.TenantName } else { $Status.ExistingPolicy.displayName }
                if ($Preview) {
                    Write-ColoredOutput " [RENAMED] -> [$tenantDisplayName]" -Color $script:Colors.Info
                } else {
                    Write-ColoredOutput " [SKIPPED - RENAMED]" -Color $script:Colors.Info
                }
            }
            $RestoreStats.Renamed++
            $RestoreStats.Skipped++
        }
        'NameConflict' {
            if ($IsImport) {
                Write-ColoredOutput " [SKIPPED - Already Exists]" -Color $script:Colors.Warning
            } else {
                # Build detailed message based on whether settings also changed
                if ($Status.SettingsChanged) {
                    $msg = "New ID + Settings Changed ($($Status.DifferenceCount))"
                    $color = $script:Colors.Error
                } else {
                    $msg = "New ID - Settings Unchanged"
                    $color = $script:Colors.Warning
                }

                if ($Preview) {
                    Write-ColoredOutput " [RESTORED] [$msg]" -Color $color
                    if ($Status.SettingsChanged -and $Status.Differences.Count -gt 0) {
                        $diffList = ($Status.Differences | Select-Object -First 5) -join ', '
                        if ($Status.Differences.Count -gt 5) { $diffList += ", ... +$($Status.Differences.Count - 5) more" }
                        Write-ColoredOutput "      Changed: $diffList" -Color $color
                    }
                } else {
                    Write-ColoredOutput " [RESTORED] [SKIPPED - $msg]" -Color $color
                }
            }
            $RestoreStats.NameConflict++
            $RestoreStats.Skipped++
        }
        'Deleted' {
            if ($Preview) {
                $previewLabel = if ($IsImport) { '[NEW] - [Would Create]' } else { '[DELETED] - [Would Restore]' }
                $previewColor = if ($IsImport) { $script:Colors.Info } else { $script:Colors.Error }
                Write-ColoredOutput " $previewLabel" -Color $previewColor
                $RestoreStats.Deleted++
                $RestoreStats.Success++
            } else {
                $restoreVerb = if ($IsImport) { 'Importing' } else { 'Restoring' }
                $label = if ($IsImport) { '[NEW]' } else { '[DELETED]' }
                $labelColor = if ($IsImport) { $script:Colors.Info } else { $script:Colors.Error }
                Write-ColoredOutput " $label - $restoreVerb..." -Color $labelColor -NoNewline
                $RestoreStats.Deleted++
                $shouldRestore = $true
            }
        }
        'New' {
            if ($Preview) {
                Write-ColoredOutput " [NEW - Would Create]" -Color $script:Colors.Info
                $RestoreStats.Success++
            } else {
                Write-ColoredOutput " [NEW] - Creating..." -Color $script:Colors.Info -NoNewline
                $shouldRestore = $true
            }
        }
    }

    return $shouldRestore
}

function Write-DifferenceReport {
    <#
    .SYNOPSIS
    Displays a formatted difference report for a dirty/modified policy.

    .DESCRIPTION
    Shows per-property backup vs tenant values with hash information.
    Used in the conflict resolution phase to give admins full visibility
    before deciding whether to restore.

    .PARAMETER PolicyName
    The display name of the policy.

    .PARAMETER PolicyType
    The component type (e.g., "DeviceConfigurations").

    .PARAMETER Status
    The status object from Get-PolicyExistenceStatus containing Differences, hash info, etc.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$PolicyName,

        [Parameter(Mandatory = $true)]
        [string]$PolicyType,

        [Parameter(Mandatory = $true)]
        $Status
    )

    Write-Host ""
    Write-Host "  ┌─────────────────────────────────────────────────────────────────" -ForegroundColor Yellow
    Write-Host "  │ DIFFERENCE REPORT: $PolicyName" -ForegroundColor Yellow
    Write-Host "  │ Type: $PolicyType" -ForegroundColor Gray

    # Show hash info if available
    if ($Status.BackupHash -and $Status.TenantHash) {
        $shortBackup = $Status.BackupHash.Substring(0, [Math]::Min(12, $Status.BackupHash.Length))
        $shortTenant = $Status.TenantHash.Substring(0, [Math]::Min(12, $Status.TenantHash.Length))
        Write-Host "  │ Hash: Backup=$shortBackup...  Tenant=$shortTenant..." -ForegroundColor Gray
    }

    Write-Host "  │ Changes: $($Status.DifferenceCount) setting(s) differ" -ForegroundColor Yellow
    Write-Host "  ├─────────────────────────────────────────────────────────────────" -ForegroundColor Yellow

    if ($Status.SettingsChanged -and $Status.Differences.Count -gt 0) {
        foreach ($diff in $Status.Differences) {
            Write-Host "  │" -ForegroundColor Yellow
            Write-Host "  │ Property: " -ForegroundColor Gray -NoNewline
            Write-Host "$diff" -ForegroundColor White

            # Try to get actual values from backup and tenant policies if available
            # (Differences array contains property names as strings from Compare-PolicyContent)
        }
    }

    Write-Host "  │" -ForegroundColor Yellow
    Write-Host "  └─────────────────────────────────────────────────────────────────" -ForegroundColor Yellow
}

function ConvertFrom-JsonDate {
    <#
    .SYNOPSIS
    Parses date strings from JSON backup files, handling multiple formats.

    .DESCRIPTION
    Handles both ISO 8601 format (from Graph API) and Microsoft JSON date format
    (/Date(milliseconds)/) that PowerShell's ConvertTo-Json creates.

    .PARAMETER DateString
    The date string to parse

    .PARAMETER Format
    Output format (default: "yyyy-MM-dd HH:mm")

    .EXAMPLE
    $formatted = ConvertFrom-JsonDate -DateString $policy.createdDateTime
    #>
    param (
        [Parameter(Mandatory = $false)]
        [string]$DateString,

        [Parameter(Mandatory = $false)]
        [string]$Format = "yyyy-MM-dd HH:mm"
    )

    if ([string]::IsNullOrWhiteSpace($DateString)) {
        return "N/A"
    }

    try {
        # Check for Microsoft JSON date format: /Date(1234567890123)/
        if ($DateString -match '\/Date\((-?\d+)\)\/') {
            $milliseconds = [long]$Matches[1]
            # Check for DateTime.MinValue (negative or zero milliseconds indicating year 0001)
            if ($milliseconds -le 0) {
                return "N/A"
            }
            $epoch = [DateTime]::new(1970, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)
            $dateTime = $epoch.AddMilliseconds($milliseconds).ToLocalTime()
            return $dateTime.ToString($Format)
        }

        # Try standard ISO 8601 parsing
        $dateTime = [DateTime]::Parse($DateString)

        # Check for DateTime.MinValue (year 0001) - indicates no valid date
        if ($dateTime.Year -lt 1970) {
            return "N/A"
        }

        return $dateTime.ToString($Format)
    }
    catch {
        return "N/A"
    }
}

function Test-ValidDateTime {
    <#
    .SYNOPSIS
    Checks if a date value from Graph API is valid (not null, not MinValue).

    .PARAMETER DateValue
    The date value to check (can be string or DateTime)

    .RETURNS
    $true if the date is valid, $false otherwise
    #>
    param (
        [Parameter(Mandatory = $false)]
        $DateValue
    )

    if ($null -eq $DateValue) {
        return $false
    }

    $dateString = $DateValue.ToString()

    if ([string]::IsNullOrWhiteSpace($dateString)) {
        return $false
    }

    # Check for Microsoft JSON date format with negative/zero value (DateTime.MinValue)
    if ($dateString -match '\/Date\((-?\d+)\)\/') {
        $milliseconds = [long]$Matches[1]
        if ($milliseconds -le 0) {
            return $false
        }
        return $true
    }

    # Try to parse as DateTime
    try {
        $dateTime = [DateTime]::Parse($dateString)
        if ($dateTime.Year -lt 1970) {
            return $false
        }
        return $true
    }
    catch {
        return $false
    }
}

function Copy-ValidDates {
    <#
    .SYNOPSIS
    Copies valid date properties from source to target object if target dates are invalid.

    .DESCRIPTION
    Used to preserve valid dates from list API calls when detail API calls return invalid dates.

    .PARAMETER Source
    The source object (e.g., from list API call)

    .PARAMETER Target
    The target object to update (e.g., from detail API call)
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Source,

        [Parameter(Mandatory = $true)]
        $Target
    )

    # Check and copy createdDateTime
    if (-not (Test-ValidDateTime -DateValue $Target.createdDateTime)) {
        if (Test-ValidDateTime -DateValue $Source.createdDateTime) {
            if ($Target.PSObject.Properties['createdDateTime']) {
                $Target.createdDateTime = $Source.createdDateTime
            } else {
                $Target | Add-Member -NotePropertyName 'createdDateTime' -NotePropertyValue $Source.createdDateTime -Force
            }
            Write-Log "Preserved createdDateTime from list API for $($Target.displayName)" -Level Verbose
        }
    }

    # Check and copy lastModifiedDateTime
    if (-not (Test-ValidDateTime -DateValue $Target.lastModifiedDateTime)) {
        if (Test-ValidDateTime -DateValue $Source.lastModifiedDateTime) {
            if ($Target.PSObject.Properties['lastModifiedDateTime']) {
                $Target.lastModifiedDateTime = $Source.lastModifiedDateTime
            } else {
                $Target | Add-Member -NotePropertyName 'lastModifiedDateTime' -NotePropertyValue $Source.lastModifiedDateTime -Force
            }
            Write-Log "Preserved lastModifiedDateTime from list API for $($Target.displayName)" -Level Verbose
        }
    }

    return $Target
}

function Get-AuditLogDates {
    <#
    .SYNOPSIS
    Attempts to retrieve creation/modification dates from audit logs for a resource.

    .DESCRIPTION
    Makes an API call to audit logs to find when a resource was created or modified.
    This is a fallback when the resource itself doesn't have valid date properties.

    .PARAMETER ResourceId
    The ID of the resource to look up

    .PARAMETER ResourceType
    The type of resource (for filtering audit logs)
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$ResourceId,

        [Parameter(Mandatory = $false)]
        [string]$ResourceType = ""
    )

    $result = @{
        CreatedDateTime = $null
        LastModifiedDateTime = $null
    }

    try {
        # Query audit logs for this resource
        $auditUri = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$filter=targetResources/any(t: t/id eq '$ResourceId')&`$top=50&`$orderby=activityDateTime desc"

        $auditLogs = Get-GraphData -Uri $auditUri -SuppressErrors

        if ($auditLogs -and $auditLogs.Count -gt 0) {
            # Find the earliest "Add" or "Create" activity for creation date
            $createActivity = $auditLogs | Where-Object {
                $_.activityDisplayName -like "*Add*" -or
                $_.activityDisplayName -like "*Create*"
            } | Sort-Object activityDateTime | Select-Object -First 1

            if ($createActivity) {
                $result.CreatedDateTime = $createActivity.activityDateTime
            }

            # The most recent activity is the last modified
            $latestActivity = $auditLogs | Sort-Object activityDateTime -Descending | Select-Object -First 1
            if ($latestActivity) {
                $result.LastModifiedDateTime = $latestActivity.activityDateTime
            }
        }
    }
    catch {
        Write-Log "Could not retrieve audit logs for resource $ResourceId : $($_.Exception.Message)" -Level Verbose
    }

    return $result
}

function Ensure-ValidDates {
    <#
    .SYNOPSIS
    Ensures a policy object has valid date properties, trying multiple sources.

    .DESCRIPTION
    Attempts to get valid dates from:
    1. The object itself
    2. A fallback source object (e.g., from list API)
    3. Audit logs as last resort

    .PARAMETER Policy
    The policy object to ensure has valid dates

    .PARAMETER FallbackSource
    Optional fallback object with potentially valid dates

    .PARAMETER TryAuditLogs
    Whether to try audit logs if other sources fail (default: $true)
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Policy,

        [Parameter(Mandatory = $false)]
        $FallbackSource = $null,

        [Parameter(Mandatory = $false)]
        [bool]$TryAuditLogs = $true
    )

    # First, try to copy from fallback source if provided
    if ($null -ne $FallbackSource) {
        $Policy = Copy-ValidDates -Source $FallbackSource -Target $Policy
    }

    # Check if we still need dates
    $needCreated = -not (Test-ValidDateTime -DateValue $Policy.createdDateTime)
    $needModified = -not (Test-ValidDateTime -DateValue $Policy.lastModifiedDateTime)

    # Try audit logs as last resort
    if ($TryAuditLogs -and ($needCreated -or $needModified) -and $Policy.id) {
        $auditDates = Get-AuditLogDates -ResourceId $Policy.id

        if ($needCreated -and $auditDates.CreatedDateTime) {
            if ($Policy.PSObject.Properties['createdDateTime']) {
                $Policy.createdDateTime = $auditDates.CreatedDateTime
            } else {
                $Policy | Add-Member -NotePropertyName 'createdDateTime' -NotePropertyValue $auditDates.CreatedDateTime -Force
            }
            Write-Log "Retrieved createdDateTime from audit logs for $($Policy.displayName)" -Level Verbose
        }

        if ($needModified -and $auditDates.LastModifiedDateTime) {
            if ($Policy.PSObject.Properties['lastModifiedDateTime']) {
                $Policy.lastModifiedDateTime = $auditDates.LastModifiedDateTime
            } else {
                $Policy | Add-Member -NotePropertyName 'lastModifiedDateTime' -NotePropertyValue $auditDates.LastModifiedDateTime -Force
            }
            Write-Log "Retrieved lastModifiedDateTime from audit logs for $($Policy.displayName)" -Level Verbose
        }
    }

    return $Policy
}

function Remove-ReadOnlyProperties {
    <#
    .SYNOPSIS
    Removes read-only and system-generated properties from a policy object.

    .DESCRIPTION
    Implements the "Clean JSON" protocol per SKILL.md specification.
    Strips properties that cause 400 Bad Request errors during import/POST operations.
    Also converts WCF date format (/Date(...)/) to ISO 8601 format.

    .PARAMETER Policy
    The policy object to clean

    .EXAMPLE
    $cleanPolicy = Remove-ReadOnlyProperties -Policy $policy
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Policy
    )

    # Properties to remove per SKILL.md Clean JSON Protocol
    # NOTE: @odata.type is intentionally PRESERVED - Graph API requires it for polymorphic
    # endpoints (DeviceConfigurations, CompliancePolicies, EnrollmentStatusPage, etc.)
    $readOnlyProps = @(
        'id',
        'createdDateTime',
        'lastModifiedDateTime',
        'modifiedDateTime',
        'version',
        'isAssigned',
        '@odata.context',
        '@odata.etag',
        'roleScopeTagIds',
        'supportsScopeTags',
        'assignments',
        'settingDefinitions',
        'creationSource',
        'priority'
    )

    # Create a copy to avoid modifying the original
    $cleanPolicy = $Policy | ConvertTo-Json -Depth 100 | ConvertFrom-Json

    # Remove read-only properties and any @odata annotation properties (except @odata.type)
    $propsToRemove = @($cleanPolicy.PSObject.Properties.Name | Where-Object {
        $_ -in $readOnlyProps -or ($_ -like '*@odata*' -and $_ -ne '@odata.type')
    })
    foreach ($prop in $propsToRemove) {
        $cleanPolicy.PSObject.Properties.Remove($prop)
    }

    return $cleanPolicy
}

function Convert-WcfDatesInJson {
    <#
    .SYNOPSIS
    Converts WCF date format /Date(...)/ to ISO 8601 in a JSON string.

    .DESCRIPTION
    The Graph API rejects WCF date format strings like /Date(-62135596800000)/
    (which represents DateTime.MinValue). This function converts them to ISO 8601
    format or removes them if they represent invalid/minimum dates.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$JsonString
    )

    # Replace /Date(milliseconds)/ patterns with ISO 8601
    $result = [regex]::Replace($JsonString, '"\\/Date\((-?\d+)\)\\/"', {
        param($match)
        $milliseconds = [long]$match.Groups[1].Value
        if ($milliseconds -le 0) {
            # DateTime.MinValue or before epoch - use null
            '"0001-01-01T00:00:00Z"'
        }
        else {
            $epoch = [DateTime]::new(1970, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)
            $dateTime = $epoch.AddMilliseconds($milliseconds)
            """$($dateTime.ToString('yyyy-MM-ddTHH:mm:ss.fffZ'))"""
        }
    })

    return $result
}

#endregion

#region Authentication Functions

function Connect-UniFy-Endpoint {
    Write-Host ""
    Write-Host "  Connecting to Microsoft Graph..." -ForegroundColor Cyan

    try {
        # Define required permissions for v2.0 (expanded set)
        $requiredPermissions = @(
            "DeviceManagementConfiguration.ReadWrite.All",
            "DeviceManagementServiceConfig.ReadWrite.All",
            "DeviceManagementApps.ReadWrite.All",
            "DeviceManagementManagedDevices.ReadWrite.All",
            "AuditLog.Read.All"
        )

        # Check authentication method: Certificate > Client Secret > Interactive
        if ($AppId -and $TenantId -and $CertificateThumbprint) {
            Write-ColoredOutput "Using certificate authentication..." -Color $script:Colors.Info
            $connectionResult = Connect-MgGraph -ClientId $AppId -TenantId $TenantId `
                -Environment $script:GraphEnvironment -CertificateThumbprint $CertificateThumbprint `
                -NoWelcome -ErrorAction Stop
        }
        elseif ($AppId -and $TenantId -and $ClientSecret) {
            Write-ColoredOutput "Using client secret authentication..." -Color $script:Colors.Info
            $secureSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
            $clientCredential = New-Object System.Management.Automation.PSCredential($AppId, $secureSecret)
            $connectionResult = Connect-MgGraph -TenantId $TenantId `
                -Environment $script:GraphEnvironment -ClientSecretCredential $clientCredential `
                -NoWelcome -ErrorAction Stop
        }
        else {
            Write-ColoredOutput "Using interactive authentication..." -Color $script:Colors.Info
            $connectionResult = Connect-MgGraph -Scopes $requiredPermissions `
                -Environment $script:GraphEnvironment -NoWelcome -ErrorAction Stop
        }

        # Verify connection
        $context = Get-MgContext
        if ($context) {
            Write-Host "  Successfully connected to tenant: " -ForegroundColor Green -NoNewline
            Write-Host "$($context.TenantId)" -ForegroundColor White
            Write-Host "  Account: " -ForegroundColor Cyan -NoNewline
            Write-Host "$($context.Account)" -ForegroundColor White
            return $true
        }
        else {
            throw "Failed to establish Graph connection"
        }
    }
    catch {
        $errorMessage = $_.Exception.Message

        # Check if user canceled the authentication
        if ($errorMessage -match "cancel" -or $errorMessage -match "cancelled" -or $errorMessage -match "user.*abort") {
            Write-ColoredOutput "`nAuthentication was cancelled by user." -Color $script:Colors.Warning
            Write-Log "Authentication cancelled by user" -Level Info
            return "cancelled"
        }
        else {
            Write-ColoredOutput "Failed to connect to Microsoft Graph: $_" -Color $script:Colors.Error
            Write-Log "Connection Error: $errorMessage" -Level Error
            return $false
        }
    }
}

#endregion

#region Platform Mapping Functions

function Get-ValidComponentsForPlatform {
    <#
    .SYNOPSIS
    Returns the list of valid component types for a specific platform.

    .DESCRIPTION
    Different policy types are only applicable to certain platforms.
    This function returns the list of components that should be backed up
    for the specified platform to avoid unnecessary API calls and errors.

    .PARAMETER Platform
    The platform name (Windows, iOS, Android, macOS, or All)

    .EXAMPLE
    $validComponents = Get-ValidComponentsForPlatform -Platform "iOS"
    #>
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet("Windows", "iOS", "Android", "macOS", "All")]
        [string]$Platform
    )

    # Define platform-specific component mappings based on Intune capabilities
    # NOTE: AutopilotDevicePrep is already exported via Settings Catalog - only available in component-specific backup
    # NOTE: There is only ONE Assignment Filters endpoint (deviceManagement/assignmentFilters) - filters by platform property
    $platformMap = @{
        "Windows" = @(
            "DeviceConfigurations",         # Includes Administrative Templates
            "CompliancePolicies",
            "SettingsCatalogPolicies",        # Settings Catalog (includes Autopilot Device Prep)
            "AppProtectionPolicies",        # Windows (Microsoft Edge)
            "PowerShellScripts",            # Windows only
            "AutopilotProfiles",            # Windows only
            "EnrollmentStatusPage",         # Windows only
            "RemediationScripts",           # Windows only
            "AssignmentFilters"             # Assignment Filters (filtered by platform)
        )
        "iOS" = @(
            "DeviceConfigurations",
            "CompliancePolicies",
            "SettingsCatalogPolicies",        # Settings Catalog
            "AppProtectionPolicies",        # iOS supported
            "AppConfigManagedDevices",      # iOS and Android only
            "AppConfigManagedApps",         # iOS and Android only
            "AssignmentFilters"             # Assignment Filters (filtered by platform)
        )
        "Android" = @(
            "DeviceConfigurations",
            "CompliancePolicies",
            "SettingsCatalogPolicies",        # Settings Catalog
            "AppProtectionPolicies",        # Android supported
            "AppConfigManagedDevices",      # iOS and Android only
            "AppConfigManagedApps",         # iOS and Android only
            "AssignmentFilters"             # Assignment Filters (filtered by platform)
        )
        "macOS" = @(
            "DeviceConfigurations",
            "CompliancePolicies",
            "SettingsCatalogPolicies",        # Settings Catalog
            # NOTE: macOS does NOT support AppProtectionPolicies
            "MacOSScripts",                 # macOS Shell Scripts
            "MacOSCustomAttributes",        # macOS Custom Attributes
            "AssignmentFilters"             # Assignment Filters (filtered by platform)
        )
        "All" = @(
            "DeviceConfigurations",
            "CompliancePolicies",
            "SettingsCatalogPolicies",
            "AppProtectionPolicies",
            "PowerShellScripts",
            "AutopilotProfiles",
            "EnrollmentStatusPage",
            "RemediationScripts",
            "AssignmentFilters",
            "AppConfigManagedDevices",
            "AppConfigManagedApps",
            "MacOSScripts",
            "MacOSCustomAttributes"
        )
    }

    return $platformMap[$Platform]
}

#endregion

#region UI Helper Functions

function Show-OpenFileDialog {
    <#
    .SYNOPSIS
    Shows a Windows file browser dialog for selecting a JSON file.

    .DESCRIPTION
    Opens a native Windows file selection dialog that allows users to browse
    and select a JSON file instead of manually typing the path.

    .PARAMETER InitialDirectory
    The starting directory for the file browser. Defaults to Documents folder.

    .EXAMPLE
    $filePath = Show-OpenFileDialog
    #>
    param (
        [string]$InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
    )

    Add-Type -AssemblyName System.Windows.Forms

    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = $InitialDirectory
    $openFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $openFileDialog.FilterIndex = 1
    $openFileDialog.Title = "Select Import File"
    $openFileDialog.Multiselect = $false

    $result = $openFileDialog.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $openFileDialog.FileName
    }
    else {
        return $null
    }
}

function Show-FolderBrowserDialog {
    <#
    .SYNOPSIS
    Shows a modern Windows folder browser dialog for selecting a directory.

    .DESCRIPTION
    Opens a modern Vista-style folder selection dialog (same style as file open dialog)
    that allows users to browse and select a folder using the IFileOpenDialog COM interface.

    .PARAMETER Description
    The title text shown in the dialog.

    .PARAMETER SelectedPath
    The initially selected path.

    .EXAMPLE
    $folderPath = Show-FolderBrowserDialog -Description "Select backup folder"
    #>
    param (
        [string]$Description = "Select folder",
        [string]$SelectedPath = [Environment]::GetFolderPath("MyDocuments")
    )

    # Use COM interop for modern Vista-style folder picker (same as file open dialog)
    $source = @"
using System;
using System.Runtime.InteropServices;

public class FolderPicker {
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHCreateItemFromParsingName(
        [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
        IntPtr pbc,
        [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
        out IntPtr ppv);

    [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
    private class FileOpenDialogClass { }

    [ComImport, Guid("42f85136-db7e-439c-85f1-e4075d135fc8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileOpenDialog {
        [PreserveSig] int Show(IntPtr hwndOwner);
        void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
        void SetFileTypeIndex(uint iFileType);
        void GetFileTypeIndex(out uint piFileType);
        void Advise(IntPtr pfde, out uint pdwCookie);
        void Unadvise(uint dwCookie);
        void SetOptions(uint fos);
        void GetOptions(out uint pfos);
        void SetDefaultFolder(IntPtr psi);
        void SetFolder(IntPtr psi);
        void GetFolder(out IntPtr ppsi);
        void GetCurrentSelection(out IntPtr ppsi);
        void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void GetResult(out IntPtr ppsi);
        void AddPlace(IntPtr psi, int fdap);
        void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
        void Close(int hr);
        void SetClientGuid(ref Guid guid);
        void ClearClientData();
        void SetFilter(IntPtr pFilter);
        void GetResults(out IntPtr ppenum);
        void GetSelectedItems(out IntPtr ppsai);
    }

    [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItem {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
        void GetParent(out IntPtr ppsi);
        void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IntPtr psi, uint hint, out int piOrder);
    }

    private const uint FOS_PICKFOLDERS = 0x00000020;
    private const uint FOS_FORCEFILESYSTEM = 0x00000040;
    private const uint SIGDN_FILESYSPATH = 0x80058000;
    private static readonly Guid IShellItemGuid = new Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE");

    public static string ShowDialog(string title, string initialPath) {
        IFileOpenDialog dialog = (IFileOpenDialog)new FileOpenDialogClass();
        try {
            dialog.SetOptions(FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);
            dialog.SetTitle(title);

            if (!string.IsNullOrEmpty(initialPath) && System.IO.Directory.Exists(initialPath)) {
                IntPtr folderPtr;
                if (SHCreateItemFromParsingName(initialPath, IntPtr.Zero, IShellItemGuid, out folderPtr) == 0) {
                    dialog.SetFolder(folderPtr);
                    Marshal.Release(folderPtr);
                }
            }

            if (dialog.Show(IntPtr.Zero) == 0) {
                IntPtr resultPtr;
                dialog.GetResult(out resultPtr);
                IShellItem result = (IShellItem)Marshal.GetObjectForIUnknown(resultPtr);
                string path;
                result.GetDisplayName(SIGDN_FILESYSPATH, out path);
                Marshal.Release(resultPtr);
                return path;
            }
        }
        finally {
            Marshal.ReleaseComObject(dialog);
        }
        return null;
    }
}
"@

    try {
        # Check if the type is already loaded
        $typeExists = [System.Type]::GetType("FolderPicker", $false, $true)
        if (-not $typeExists) {
            Add-Type -TypeDefinition $source -Language CSharp -ErrorAction Stop
        }

        $result = [FolderPicker]::ShowDialog($Description, $SelectedPath)
        return $result
    }
    catch {
        # Fallback to standard FolderBrowserDialog if COM interop fails
        Add-Type -AssemblyName System.Windows.Forms
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = $Description
        $folderBrowser.SelectedPath = $SelectedPath
        $folderBrowser.ShowNewFolderButton = $true

        $dialogResult = $folderBrowser.ShowDialog()
        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            return $folderBrowser.SelectedPath
        }
        return $null
    }
}

#endregion

#region Graph API Functions

function Get-ODataFilterValue {
    <#
    .SYNOPSIS
    Escapes a string value for use in OData $filter queries.

    .DESCRIPTION
    Properly escapes special characters in strings for OData filter syntax:
    - Single quotes are doubled (' becomes '')
    - Special characters are URL-encoded (& [ ] ( ) + # etc.)

    .PARAMETER Value
    The string value to escape

    .EXAMPLE
    $escapedName = Get-ODataFilterValue -Value "[Restored] Policy & Test's Name (24H2+)"
    # Returns: %5BRestored%5D Policy %26 Test''s Name %2824H2%2B%29
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    # First escape single quotes (OData requires '' for literal single quote)
    $escaped = $Value -replace "'", "''"

    # URL-encode special characters that cause issues in OData filter queries
    $escaped = $escaped -replace '&', '%26'      # Ampersand
    $escaped = $escaped -replace '\[', '%5B'     # Opening square bracket
    $escaped = $escaped -replace '\]', '%5D'     # Closing square bracket
    $escaped = $escaped -replace '\(', '%28'     # Opening parenthesis
    $escaped = $escaped -replace '\)', '%29'     # Closing parenthesis
    $escaped = $escaped -replace '\+', '%2B'     # Plus sign
    $escaped = $escaped -replace '#', '%23'      # Hash/pound sign

    return $escaped
}

function Get-GraphData {
    <#
    .SYNOPSIS
    Retrieves data from Microsoft Graph API with automatic pagination support.

    .DESCRIPTION
    Implements robust pagination per SKILL.md specification using $top=999 and @odata.nextLink.
    Ensures all data is retrieved in large tenants (1000+ policies).

    .PARAMETER Uri
    The Graph API URI to query

    .PARAMETER Method
    The HTTP method (default: GET)

    .PARAMETER SuppressErrors
    If true, suppresses error output to console (errors still logged). Use for optional API calls
    like audit logs where permission errors are expected and non-critical.

    .EXAMPLE
    $policies = Get-GraphData -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies"

    .EXAMPLE
    $auditLogs = Get-GraphData -Uri "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits" -SuppressErrors
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $false)]
        [string]$Method = "GET",

        [Parameter(Mandatory = $false)]
        [switch]$SuppressErrors
    )

    $allData = @()

    # Add pagination parameter only for collection endpoints (not single item fetches by ID)
    # Single item endpoints have a GUID at the end of the path and don't support $top
    # Some IDs are compound: guid_suffix (e.g., ESP: guid_Windows10EnrollmentCompletionPageConfiguration)
    $uriPath = ($Uri -split '\?')[0]  # Get the path without query string
    $isCollectionEndpoint = $uriPath -notmatch '[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}(_[A-Za-z0-9]+)?$'

    $addedTop = $false
    $originalUri = $Uri
    if ($isCollectionEndpoint -and $Uri -notlike "*`$top=*") {
        $separator = if ($Uri.Contains("?")) { "&" } else { "?" }
        $Uri = "$Uri$separator`$top=999"
        $addedTop = $true
    }

    $nextLink = $Uri
    $retried500 = $false

    while ($nextLink) {
        try {
            Write-Log "Fetching: $nextLink" -Level Verbose
            $response = Invoke-MgGraphRequest -Uri $nextLink -Method $Method

            if ($null -ne $response.value) {
                # Collection response - only add if there are items
                if ($response.value.Count -gt 0) {
                    $allData += $response.value
                    Write-Log "Retrieved $($response.value.Count) items" -Level Verbose
                } else {
                    Write-Log "Query returned empty collection" -Level Verbose
                }
            }
            elseif ($response -and $response.GetType().Name -ne "Hashtable") {
                # Single item response (non-hashtable)
                $allData += $response
            }
            elseif ($response -and -not $response.ContainsKey('value') -and -not $response.'@odata.nextLink') {
                # Single item in hashtable format (but NOT a collection response with empty value)
                $allData += $response
            }

            # Check for pagination
            $nextLink = $response.'@odata.nextLink'

            if ($nextLink) {
                Write-Log "More data available, following nextLink..." -Level Verbose
            }
        }
        catch {
            $errorDetails = $_.Exception.Message

            # Parse specific error codes per SKILL.md specification
            # Only display errors to console if SuppressErrors is not set
            if ($errorDetails -like "*403*" -or $errorDetails -like "*Forbidden*") {
                if (-not $SuppressErrors) {
                    Write-ColoredOutput "Permission Error (403): Insufficient permissions for $Uri" -Color $script:Colors.Error
                }
                Write-Log "403 Forbidden: $errorDetails" -Level Verbose
            }
            elseif ($errorDetails -like "*404*" -or $errorDetails -like "*Not Found*") {
                if (-not $SuppressErrors) {
                    Write-ColoredOutput "Resource Not Found (404): $Uri" -Color $script:Colors.Warning
                }
                Write-Log "404 Not Found: $errorDetails" -Level Verbose
            }
            elseif ($errorDetails -like "*400*" -or $errorDetails -like "*Bad Request*") {
                if ($addedTop) {
                    # Endpoint doesn't support $top, retry without pagination
                    Write-Log "Endpoint does not support `$top parameter, retrying without pagination: $originalUri" -Level Verbose
                    $nextLink = $originalUri
                    $addedTop = $false
                    continue
                }
                if (-not $SuppressErrors) {
                    Write-ColoredOutput "Bad Request (400): Invalid syntax for $Uri" -Color $script:Colors.Error
                }
                Write-Log "400 Bad Request: $errorDetails" -Level Verbose
            }
            elseif ($errorDetails -like "*InternalServerError*" -or $errorDetails -like "*Internal Server Error*" -or $errorDetails -like "*500*") {
                if (-not $retried500) {
                    Write-Log "Transient 500 error, retrying after 2s: $nextLink" -Level Verbose
                    $retried500 = $true
                    Start-Sleep -Seconds 2
                    continue
                }
                if (-not $SuppressErrors) {
                    Write-ColoredOutput "Server Error (500): $Uri - skipping, will treat as empty." -Color $script:Colors.Warning
                }
                Write-Log "500 Internal Server Error (after retry): $errorDetails" -Level Verbose
            }
            else {
                if (-not $SuppressErrors) {
                    Write-ColoredOutput "Error fetching data from $Uri : $errorDetails" -Color $script:Colors.Warning
                }
                Write-Log "Graph API Error: $errorDetails" -Level Verbose
            }

            break
        }
    }

    return $allData
}

function Get-PolicyAssignments {
    <#
    .SYNOPSIS
    Retrieves assignment information for a specific policy.

    .DESCRIPTION
    Queries the /assignments endpoint for a policy and returns the assigned groups.
    Per SKILL.md specification, assignments are nested within the policy JSON.

    .PARAMETER PolicyId
    The ID of the policy

    .PARAMETER PolicyType
    The type of policy (used to construct the URI)

    .EXAMPLE
    $assignments = Get-PolicyAssignments -PolicyId "abc123" -PolicyType "deviceConfigurations"
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$PolicyId,

        [Parameter(Mandatory = $true)]
        [string]$PolicyType
    )

    try {
        $uri = "https://graph.microsoft.com/beta/deviceManagement/$PolicyType/$PolicyId/assignments"
        $assignments = Get-GraphData -Uri $uri

        Write-Log "Retrieved $($assignments.Count) assignments for policy $PolicyId" -Level Verbose

        return $assignments
    }
    catch {
        Write-Log "Could not retrieve assignments for $PolicyId : $($_.Exception.Message)" -Level Warning
        return @()
    }
}

function Get-PolicyDetails {
    <#
    .SYNOPSIS
    Retrieves complete details for a specific policy including assignments.

    .DESCRIPTION
    Fetches full policy details and embeds assignment information.
    Converts response to PSCustomObject to ensure proper serialization.

    .PARAMETER PolicyId
    The ID of the policy

    .PARAMETER PolicyType
    The type of policy

    .PARAMETER IncludeAssignments
    Whether to include assignment information

    .EXAMPLE
    $policy = Get-PolicyDetails -PolicyId "abc123" -PolicyType "deviceConfigurations" -IncludeAssignments
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$PolicyId,

        [Parameter(Mandatory = $true)]
        [string]$PolicyType,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeAssignments
    )

    $uri = "https://graph.microsoft.com/beta/deviceManagement/$PolicyType/$PolicyId"

    try {
        $response = Invoke-MgGraphRequest -Uri $uri -Method GET

        # Convert Hashtable to PSCustomObject for proper serialization
        $policy = ConvertTo-PolicyObject -Response $response

        # Add assignments if requested
        if ($IncludeAssignments) {
            $assignments = Get-PolicyAssignments -PolicyId $PolicyId -PolicyType $PolicyType
            $policy | Add-Member -NotePropertyName "assignments" -NotePropertyValue $assignments -Force
            $policy | Add-Member -NotePropertyName "isAssigned" -NotePropertyValue ($assignments.Count -gt 0) -Force
        }

        return $policy
    }
    catch {
        Write-ColoredOutput "Failed to get details for policy $PolicyId : $_" -Color $script:Colors.Warning
        Write-Log "Policy Details Error: $($_.Exception.Message)" -Level Warning
        return $null
    }
}

#endregion

#region Export Functions - v2.0

function Export-AutopilotProfiles {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    Write-ColoredOutput "`n$(if ($script:IsLiveReport) { 'Fetching' } else { 'Exporting' }) Windows Autopilot Profiles..." -Color $script:Colors.Info

    try {
        $uri = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles"
        $profiles = Get-GraphData -Uri $uri

        if (-not $profiles -or $profiles.Count -eq 0) {
            Write-ColoredOutput "  No Autopilot Profiles found" -Color $script:Colors.Info
            return 0
        }

        $exportedCount = 0
        $profilePath = Join-Path $BackupPath "AutopilotProfiles"
        New-Item -ItemType Directory -Path $profilePath -Force | Out-Null

        foreach ($profile in $profiles) {
            if ($script:SelectedPlatform -ne "All" -and $script:SelectedPlatform -ne "Windows") {
                continue
            }

            Write-ColoredOutput "  - $($profile.displayName)" -Color $script:Colors.Default

            $fullProfile = Get-PolicyDetails -PolicyId $profile.id `
                -PolicyType "windowsAutopilotDeploymentProfiles" `
                -IncludeAssignments

            if ($fullProfile) {
                # Ensure valid dates - try list API data first, then audit logs
                $fullProfile = Ensure-ValidDates -Policy $fullProfile -FallbackSource $profile -TryAuditLogs $true

                $platformTag = "Windows"
                $fileName = Get-SafeFileName -FileName "$($profile.displayName)_${platformTag}_$($profile.id)"
                $filePath = Join-Path $profilePath "$fileName.json"

                $fullProfile | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding UTF8
                $exportedCount++
                Write-Log "Exported: $($profile.displayName)" -Level Verbose
            }
        }
        Write-ColoredOutput "  Total: $exportedCount profiles" -Color $script:Colors.Success

        return $exportedCount
    }
    catch {
        Write-ColoredOutput "  Error exporting Autopilot Profiles: $_" -Color $script:Colors.Error
        Write-Log "Autopilot Profiles export error: $($_.Exception.Message)" -Level Error
        return 0
    }
}

function Export-AutopilotDevicePrep {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    Write-ColoredOutput "`n$(if ($script:IsLiveReport) { 'Fetching' } else { 'Exporting' }) Autopilot Device Preparation Policies..." -Color $script:Colors.Info

    try {
        $templateId = "20f9c2d1-e508-4122-8692-7f284b3956f1"
        $uri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?`$filter=templateReference/templateId eq '$templateId'"
        $policies = Get-GraphData -Uri $uri

        if (-not $policies -or $policies.Count -eq 0) {
            Write-ColoredOutput "  No Device Preparation Policies found" -Color $script:Colors.Info
            return 0
        }

        $exportedCount = 0
        $prepPath = Join-Path $BackupPath "AutopilotDevicePrep"
        New-Item -ItemType Directory -Path $prepPath -Force | Out-Null

        foreach ($policy in $policies) {
            # Skip policies with null/empty IDs or names
            if ([string]::IsNullOrWhiteSpace($policy.id) -or [string]::IsNullOrWhiteSpace($policy.name)) {
                continue
            }

            if ($script:SelectedPlatform -ne "All" -and $script:SelectedPlatform -ne "Windows") {
                continue
            }

            Write-ColoredOutput "  - $($policy.name)" -Color $script:Colors.Default

            try {
                $uri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$($policy.id)?`$expand=settings"
                $response = Invoke-MgGraphRequest -Uri $uri -Method GET
                $fullPolicy = ConvertTo-PolicyObject -Response $response

                # Ensure valid dates - try list API data first, then audit logs
                $fullPolicy = Ensure-ValidDates -Policy $fullPolicy -FallbackSource $policy -TryAuditLogs $true

                $assignUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$($policy.id)/assignments"
                $assignments = Get-GraphData -Uri $assignUri
                $fullPolicy | Add-Member -NotePropertyName "assignments" -NotePropertyValue $assignments -Force
                $fullPolicy | Add-Member -NotePropertyName "isAssigned" -NotePropertyValue ($null -ne $assignments -and $assignments.Count -gt 0) -Force

                $platformTag = "Windows"
                $fileName = Get-SafeFileName -FileName "$($policy.name)_${platformTag}_$($policy.id)"
                $filePath = Join-Path $prepPath "$fileName.json"

                $fullPolicy | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding UTF8
                $exportedCount++
                Write-Log "Exported: $($policy.name)" -Level Verbose
            }
            catch {
                Write-Log "Error exporting Device Prep policy $($policy.name): $_" -Level Warning
            }
        }
        Write-ColoredOutput "  Total: $exportedCount policies" -Color $script:Colors.Success

        return $exportedCount
    }
    catch {
        Write-ColoredOutput "  Error exporting Device Prep Policies: $_" -Color $script:Colors.Error
        Write-Log "Device Prep export error: $($_.Exception.Message)" -Level Error
        return 0
    }
}

function Export-EnrollmentStatusPage {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    Write-ColoredOutput "`n$(if ($script:IsLiveReport) { 'Fetching' } else { 'Exporting' }) Enrollment Status Page (ESP)..." -Color $script:Colors.Info

    try {
        $uri = "https://graph.microsoft.com/v1.0/deviceManagement/deviceEnrollmentConfigurations"
        $configs = Get-GraphData -Uri $uri

        if (-not $configs -or $configs.Count -eq 0) {
            Write-ColoredOutput "  No ESP configurations found" -Color $script:Colors.Info
            return 0
        }

        $espConfigs = $configs | Where-Object {
            $_.'@odata.type' -eq '#microsoft.graph.windows10EnrollmentCompletionPageConfiguration'
        }

        if (-not $espConfigs -or $espConfigs.Count -eq 0) {
            Write-ColoredOutput "  No ESP configurations found" -Color $script:Colors.Info
            return 0
        }

        $exportedCount = 0
        $espPath = Join-Path $BackupPath "EnrollmentStatusPage"
        New-Item -ItemType Directory -Path $espPath -Force | Out-Null

        foreach ($config in $espConfigs) {
            if ($script:SelectedPlatform -ne "All" -and $script:SelectedPlatform -ne "Windows") {
                continue
            }

            Write-ColoredOutput "  - $($config.displayName)" -Color $script:Colors.Default

            $detailUri = "https://graph.microsoft.com/v1.0/deviceManagement/deviceEnrollmentConfigurations/$($config.id)"
            $response = Invoke-MgGraphRequest -Uri $detailUri -Method GET
            $fullConfig = ConvertTo-PolicyObject -Response $response

            # Ensure valid dates - try list API data first, then audit logs
            $fullConfig = Ensure-ValidDates -Policy $fullConfig -FallbackSource $config -TryAuditLogs $true

            $assignUri = "https://graph.microsoft.com/v1.0/deviceManagement/deviceEnrollmentConfigurations/$($config.id)/assignments"
            $assignments = Get-GraphData -Uri $assignUri
            $fullConfig | Add-Member -NotePropertyName "assignments" -NotePropertyValue $assignments -Force
            $fullConfig | Add-Member -NotePropertyName "isAssigned" -NotePropertyValue ($null -ne $assignments -and $assignments.Count -gt 0) -Force

            $platformTag = "Windows"
            $fileName = Get-SafeFileName -FileName "$($config.displayName)_${platformTag}_$($config.id)"
            $filePath = Join-Path $espPath "$fileName.json"

            $fullConfig | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding UTF8
            $exportedCount++
            Write-Log "Exported: $($config.displayName)" -Level Verbose
        }
        Write-ColoredOutput "  Total: $exportedCount configurations" -Color $script:Colors.Success

        return $exportedCount
    }
    catch {
        Write-ColoredOutput "  Error exporting ESP: $_" -Color $script:Colors.Error
        Write-Log "ESP export error: $($_.Exception.Message)" -Level Error
        return 0
    }
}

function Export-PowerShellScripts {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    Write-ColoredOutput "`n$(if ($script:IsLiveReport) { 'Fetching' } else { 'Exporting' }) PowerShell Scripts..." -Color $script:Colors.Info

    try {
        $uri = "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts"
        $scripts = Get-GraphData -Uri $uri

        if (-not $scripts -or $scripts.Count -eq 0) {
            Write-ColoredOutput "  No PowerShell Scripts found" -Color $script:Colors.Info
            return 0
        }

        $exportedCount = 0
        $scriptsPath = Join-Path $BackupPath "PowerShellScripts"
        New-Item -ItemType Directory -Path $scriptsPath -Force | Out-Null

        foreach ($scriptItem in $scripts) {
            if ($script:SelectedPlatform -ne "All" -and $script:SelectedPlatform -ne "Windows") {
                continue
            }

            Write-ColoredOutput "  - $($scriptItem.displayName)" -Color $script:Colors.Default

            $detailUri = "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts/$($scriptItem.id)"
            $response = Invoke-MgGraphRequest -Uri $detailUri -Method GET
            $fullScript = ConvertTo-PolicyObject -Response $response

            # Ensure valid dates - try list API data first, then audit logs
            $fullScript = Ensure-ValidDates -Policy $fullScript -FallbackSource $scriptItem -TryAuditLogs $true

            if ($fullScript.scriptContent) {
                try {
                    $decodedContent = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($fullScript.scriptContent))
                    $fullScript.scriptContent = $decodedContent
                    Write-Log "Decoded Base64 script content for $($scriptItem.displayName)" -Level Verbose
                }
                catch {
                    Write-Log "Could not decode script content for $($scriptItem.displayName), keeping original" -Level Warning
                }
            }

            $assignUri = "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts/$($scriptItem.id)/assignments"
            $assignments = Get-GraphData -Uri $assignUri
            $fullScript | Add-Member -NotePropertyName "assignments" -NotePropertyValue $assignments -Force
            $fullScript | Add-Member -NotePropertyName "isAssigned" -NotePropertyValue ($null -ne $assignments -and $assignments.Count -gt 0) -Force

            $platformTag = "Windows"
            $fileName = Get-SafeFileName -FileName "$($scriptItem.displayName)_${platformTag}_$($scriptItem.id)"
            $filePath = Join-Path $scriptsPath "$fileName.json"

            $fullScript | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding UTF8
            $exportedCount++
            Write-Log "Exported: $($scriptItem.displayName)" -Level Verbose
        }
        Write-ColoredOutput "  Total: $exportedCount scripts" -Color $script:Colors.Success

        return $exportedCount
    }
    catch {
        Write-ColoredOutput "  Error exporting PowerShell Scripts: $_" -Color $script:Colors.Error
        Write-Log "PowerShell Scripts export error: $($_.Exception.Message)" -Level Error
        return 0
    }
}

function Export-RemediationScripts {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    Write-ColoredOutput "`n$(if ($script:IsLiveReport) { 'Fetching' } else { 'Exporting' }) Remediation Scripts..." -Color $script:Colors.Info

    try {
        $uri = "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts"
        $remediations = Get-GraphData -Uri $uri

        if (-not $remediations -or $remediations.Count -eq 0) {
            Write-ColoredOutput "  No Remediation Scripts found" -Color $script:Colors.Info
            return 0
        }

        $exportedCount = 0
        $remediationsPath = Join-Path $BackupPath "RemediationScripts"
        New-Item -ItemType Directory -Path $remediationsPath -Force | Out-Null

        foreach ($remediation in $remediations) {
            if ($script:SelectedPlatform -ne "All" -and $script:SelectedPlatform -ne "Windows") {
                continue
            }

            Write-ColoredOutput "  - $($remediation.displayName)" -Color $script:Colors.Default

            $detailUri = "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts/$($remediation.id)"
            $response = Invoke-MgGraphRequest -Uri $detailUri -Method GET
            $fullRemediation = ConvertTo-PolicyObject -Response $response

            # Ensure valid dates - try list API data first, then audit logs
            $fullRemediation = Ensure-ValidDates -Policy $fullRemediation -FallbackSource $remediation -TryAuditLogs $true

            if ($fullRemediation.detectionScriptContent) {
                try {
                    $decodedDetection = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($fullRemediation.detectionScriptContent))
                    $fullRemediation.detectionScriptContent = $decodedDetection
                    Write-Log "Decoded detection script for $($remediation.displayName)" -Level Verbose
                }
                catch {
                    Write-Log "Could not decode detection script for $($remediation.displayName)" -Level Warning
                }
            }

            if ($fullRemediation.remediationScriptContent) {
                try {
                    $decodedRemediation = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($fullRemediation.remediationScriptContent))
                    $fullRemediation.remediationScriptContent = $decodedRemediation
                    Write-Log "Decoded remediation script for $($remediation.displayName)" -Level Verbose
                }
                catch {
                    Write-Log "Could not decode remediation script for $($remediation.displayName)" -Level Warning
                }
            }

            $assignUri = "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts/$($remediation.id)/assignments"
            $assignments = Get-GraphData -Uri $assignUri
            $fullRemediation | Add-Member -NotePropertyName "assignments" -NotePropertyValue $assignments -Force
            $fullRemediation | Add-Member -NotePropertyName "isAssigned" -NotePropertyValue ($null -ne $assignments -and $assignments.Count -gt 0) -Force

            $platformTag = "Windows"
            $fileName = Get-SafeFileName -FileName "$($remediation.displayName)_${platformTag}_$($remediation.id)"
            $filePath = Join-Path $remediationsPath "$fileName.json"

            $fullRemediation | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding UTF8
            $exportedCount++
            Write-Log "Exported: $($remediation.displayName)" -Level Verbose
        }
        Write-ColoredOutput "  Total: $exportedCount remediations" -Color $script:Colors.Success

        return $exportedCount
    }
    catch {
        Write-ColoredOutput "  Error exporting Remediation Scripts: $_" -Color $script:Colors.Error
        Write-Log "Remediation Scripts export error: $($_.Exception.Message)" -Level Error
        return 0
    }
}

function Export-WUfBPolicies {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    Write-ColoredOutput "`n$(if ($script:IsLiveReport) { 'Fetching' } else { 'Exporting' }) Windows Update for Business Policies..." -Color $script:Colors.Info

    try {
        $uri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations"
        $allConfigs = Get-GraphData -Uri $uri
        $policies = @($allConfigs | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.windowsUpdateForBusinessConfiguration' })

        if (-not $policies -or $policies.Count -eq 0) {
            Write-ColoredOutput "  No WUfB Policies found" -Color $script:Colors.Info
            return 0
        }

        $exportedCount = 0
        $wufbPath = Join-Path $BackupPath "WUfBPolicies"
        New-Item -ItemType Directory -Path $wufbPath -Force | Out-Null

        foreach ($policy in $policies) {
            if ($script:SelectedPlatform -ne "All" -and $script:SelectedPlatform -ne "Windows") {
                continue
            }

            Write-ColoredOutput "  - $($policy.displayName)" -Color $script:Colors.Default

            $detailUri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$($policy.id)"
            $response = Invoke-MgGraphRequest -Uri $detailUri -Method GET
            $fullPolicy = ConvertTo-PolicyObject -Response $response

            # Ensure valid dates - try list API data first, then audit logs
            $fullPolicy = Ensure-ValidDates -Policy $fullPolicy -FallbackSource $policy -TryAuditLogs $true

            $assignUri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$($policy.id)/assignments"
            $assignments = Get-GraphData -Uri $assignUri
            $fullPolicy | Add-Member -NotePropertyName "assignments" -NotePropertyValue $assignments -Force
            $fullPolicy | Add-Member -NotePropertyName "isAssigned" -NotePropertyValue ($null -ne $assignments -and $assignments.Count -gt 0) -Force

            $platformTag = "Windows"
            $fileName = Get-SafeFileName -FileName "$($policy.displayName)_${platformTag}_$($policy.id)"
            $filePath = Join-Path $wufbPath "$fileName.json"

            $fullPolicy | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding UTF8
            $exportedCount++
            Write-Log "Exported: $($policy.displayName)" -Level Verbose
        }
        Write-ColoredOutput "  Total: $exportedCount policies" -Color $script:Colors.Success

        return $exportedCount
    }
    catch {
        Write-ColoredOutput "  Error exporting WUfB Policies: $_" -Color $script:Colors.Error
        Write-Log "WUfB Policies export error: $($_.Exception.Message)" -Level Error
        return 0
    }
}

function Export-AssignmentFilters {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    Write-ColoredOutput "`n$(if ($script:IsLiveReport) { 'Fetching' } else { 'Exporting' }) Assignment Filters..." -Color $script:Colors.Info

    try {
        # Assignment Filters endpoint (filters by platform property)
        $uri = "https://graph.microsoft.com/beta/deviceManagement/assignmentFilters"
        $filters = Get-GraphData -Uri $uri

        if (-not $filters -or $filters.Count -eq 0) {
            Write-ColoredOutput "  No Assignment Filters found" -Color $script:Colors.Info
            return 0
        }

        $exportedCount = 0
        $filtersPath = Join-Path $BackupPath "AssignmentFilters"
        New-Item -ItemType Directory -Path $filtersPath -Force | Out-Null

        foreach ($filter in $filters) {
            if (-not (Test-PlatformMatch -Policy $filter)) {
                continue
            }

            Write-ColoredOutput "  - $($filter.displayName)" -Color $script:Colors.Default

            $detailUri = "https://graph.microsoft.com/beta/deviceManagement/assignmentFilters/$($filter.id)"
            $response = Invoke-MgGraphRequest -Uri $detailUri -Method GET
            $fullFilter = ConvertTo-PolicyObject -Response $response

            # Ensure valid dates - try list API data first, then audit logs
            $fullFilter = Ensure-ValidDates -Policy $fullFilter -FallbackSource $filter -TryAuditLogs $true

            $platformTag = if ($filter.platform) { $filter.platform } else { "All" }

            $fileName = Get-SafeFileName -FileName "$($filter.displayName)_${platformTag}_$($filter.id)"
            $filePath = Join-Path $filtersPath "$fileName.json"

            $fullFilter | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding UTF8
            $exportedCount++
            Write-Log "Exported: $($filter.displayName)" -Level Verbose
        }
        Write-ColoredOutput "  Total: $exportedCount filters" -Color $script:Colors.Success

        return $exportedCount
    }
    catch {
        Write-ColoredOutput "  Error exporting Assignment Filters: $_" -Color $script:Colors.Error
        Write-Log "Assignment Filters export error: $($_.Exception.Message)" -Level Error
        return 0
    }
}

function Export-AppProtectionPolicies {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    Write-ColoredOutput "`n$(if ($script:IsLiveReport) { 'Fetching' } else { 'Exporting' }) App Protection Policies (MAM)..." -Color $script:Colors.Info

    try {
        $exportedCount = 0
        $appProtectionPath = Join-Path $BackupPath "AppProtectionPolicies"
        New-Item -ItemType Directory -Path $appProtectionPath -Force | Out-Null

        # iOS App Protection Policies (only process if platform is All or iOS)
        if ($script:SelectedPlatform -eq "All" -or $script:SelectedPlatform -eq "iOS") {
            $iosUri = "https://graph.microsoft.com/beta/deviceAppManagement/iosManagedAppProtections"
            $iosPolicies = Get-GraphData -Uri $iosUri

            if ($iosPolicies -and $iosPolicies.Count -gt 0) {
                foreach ($policy in $iosPolicies) {
                    Write-ColoredOutput "  - [iOS] $($policy.displayName)" -Color $script:Colors.Default

                    $detailUri = "https://graph.microsoft.com/beta/deviceAppManagement/iosManagedAppProtections/$($policy.id)"
                    $response = Invoke-MgGraphRequest -Uri $detailUri -Method GET
                    $fullPolicy = ConvertTo-PolicyObject -Response $response

                    # Add @odata.type if not present (API detail endpoint doesn't always return it)
                    if (-not $fullPolicy.'@odata.type') {
                        $fullPolicy | Add-Member -NotePropertyName '@odata.type' -NotePropertyValue '#microsoft.graph.iosManagedAppProtection' -Force
                    }

                    # Ensure valid dates - try list API data first, then audit logs
                    $fullPolicy = Ensure-ValidDates -Policy $fullPolicy -FallbackSource $policy -TryAuditLogs $true

                    $assignUri = "https://graph.microsoft.com/beta/deviceAppManagement/iosManagedAppProtections/$($policy.id)/assignments"
                    $assignments = Get-GraphData -Uri $assignUri
                    $fullPolicy | Add-Member -NotePropertyName "assignments" -NotePropertyValue $assignments -Force
                    $fullPolicy | Add-Member -NotePropertyName "isAssigned" -NotePropertyValue ($null -ne $assignments -and $assignments.Count -gt 0) -Force

                    $fileName = Get-SafeFileName -FileName "$($policy.displayName)_iOS_$($policy.id)"
                    $filePath = Join-Path $appProtectionPath "$fileName.json"

                    $fullPolicy | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding UTF8
                    $exportedCount++
                    Write-Log "Exported iOS policy: $($policy.displayName)" -Level Verbose
                }
            }
        }

        # Android App Protection Policies (only process if platform is All or Android)
        if ($script:SelectedPlatform -eq "All" -or $script:SelectedPlatform -eq "Android") {
            $androidUri = "https://graph.microsoft.com/beta/deviceAppManagement/androidManagedAppProtections"
            $androidPolicies = Get-GraphData -Uri $androidUri

            if ($androidPolicies -and $androidPolicies.Count -gt 0) {
                foreach ($policy in $androidPolicies) {
                    Write-ColoredOutput "  - [Android] $($policy.displayName)" -Color $script:Colors.Default

                    $detailUri = "https://graph.microsoft.com/beta/deviceAppManagement/androidManagedAppProtections/$($policy.id)"
                    $response = Invoke-MgGraphRequest -Uri $detailUri -Method GET
                    $fullPolicy = ConvertTo-PolicyObject -Response $response

                    # Add @odata.type if not present (API detail endpoint doesn't always return it)
                    if (-not $fullPolicy.'@odata.type') {
                        $fullPolicy | Add-Member -NotePropertyName '@odata.type' -NotePropertyValue '#microsoft.graph.androidManagedAppProtection' -Force
                    }

                    # Ensure valid dates - try list API data first, then audit logs
                    $fullPolicy = Ensure-ValidDates -Policy $fullPolicy -FallbackSource $policy -TryAuditLogs $true

                    $assignUri = "https://graph.microsoft.com/beta/deviceAppManagement/androidManagedAppProtections/$($policy.id)/assignments"
                    $assignments = Get-GraphData -Uri $assignUri
                    $fullPolicy | Add-Member -NotePropertyName "assignments" -NotePropertyValue $assignments -Force
                    $fullPolicy | Add-Member -NotePropertyName "isAssigned" -NotePropertyValue ($null -ne $assignments -and $assignments.Count -gt 0) -Force

                    $fileName = Get-SafeFileName -FileName "$($policy.displayName)_Android_$($policy.id)"
                    $filePath = Join-Path $appProtectionPath "$fileName.json"

                    $fullPolicy | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding UTF8
                    $exportedCount++
                    Write-Log "Exported Android policy: $($policy.displayName)" -Level Verbose
                }
            }
        }

        # Windows App Protection Policies (only process if platform is All or Windows)
        if ($script:SelectedPlatform -eq "All" -or $script:SelectedPlatform -eq "Windows") {
            $windowsUri = "https://graph.microsoft.com/beta/deviceAppManagement/windowsManagedAppProtections"
            $windowsPolicies = Get-GraphData -Uri $windowsUri

            if ($windowsPolicies -and $windowsPolicies.Count -gt 0) {
                foreach ($policy in $windowsPolicies) {
                    Write-ColoredOutput "  - [Windows] $($policy.displayName)" -Color $script:Colors.Default

                    $detailUri = "https://graph.microsoft.com/beta/deviceAppManagement/windowsManagedAppProtections/$($policy.id)"
                    $response = Invoke-MgGraphRequest -Uri $detailUri -Method GET
                    $fullPolicy = ConvertTo-PolicyObject -Response $response

                    # Add @odata.type if not present (API detail endpoint doesn't always return it)
                    if (-not $fullPolicy.'@odata.type') {
                        $fullPolicy | Add-Member -NotePropertyName '@odata.type' -NotePropertyValue '#microsoft.graph.windowsManagedAppProtection' -Force
                    }

                    # Ensure valid dates - try list API data first, then audit logs
                    $fullPolicy = Ensure-ValidDates -Policy $fullPolicy -FallbackSource $policy -TryAuditLogs $true

                    $assignUri = "https://graph.microsoft.com/beta/deviceAppManagement/windowsManagedAppProtections/$($policy.id)/assignments"
                    $assignments = Get-GraphData -Uri $assignUri
                    $fullPolicy | Add-Member -NotePropertyName "assignments" -NotePropertyValue $assignments -Force
                    $fullPolicy | Add-Member -NotePropertyName "isAssigned" -NotePropertyValue ($null -ne $assignments -and $assignments.Count -gt 0) -Force

                    $fileName = Get-SafeFileName -FileName "$($policy.displayName)_Windows_$($policy.id)"
                    $filePath = Join-Path $appProtectionPath "$fileName.json"

                    $fullPolicy | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding UTF8
                    $exportedCount++
                    Write-Log "Exported Windows policy: $($policy.displayName)" -Level Verbose
                }
            }
        }

        Write-ColoredOutput "  Total: $exportedCount policies" -Color $script:Colors.Success

        return $exportedCount
    }
    catch {
        Write-ColoredOutput "  Error exporting App Protection Policies: $_" -Color $script:Colors.Error
        Write-Log "App Protection Policies export error: $($_.Exception.Message)" -Level Error
        return 0
    }
}

function Export-AppConfigManagedDevices {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    Write-ColoredOutput "`n$(if ($script:IsLiveReport) { 'Fetching' } else { 'Exporting' }) App Configuration Policies (Managed Devices)..." -Color $script:Colors.Info

    try {
        # Correct endpoint for Managed Devices app configuration
        $uri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations"
        $configs = Get-GraphData -Uri $uri

        if (-not $configs -or $configs.Count -eq 0) {
            Write-ColoredOutput "  No App Configuration Policies (Managed Devices) found" -Color $script:Colors.Info
            return 0
        }

        $exportedCount = 0
        $configPath = Join-Path $BackupPath "AppConfigManagedDevices"
        New-Item -ItemType Directory -Path $configPath -Force | Out-Null

        foreach ($config in $configs) {
            if (-not (Test-PlatformMatch -Policy $config)) {
                continue
            }

            Write-ColoredOutput "  - $($config.displayName)" -Color $script:Colors.Default

            $detailUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations/$($config.id)"
            $response = Invoke-MgGraphRequest -Uri $detailUri -Method GET
            $fullConfig = ConvertTo-PolicyObject -Response $response

            # Ensure valid dates - try list API data first, then audit logs
            $fullConfig = Ensure-ValidDates -Policy $fullConfig -FallbackSource $config -TryAuditLogs $true

            $assignUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations/$($config.id)/assignments"
            $assignments = Get-GraphData -Uri $assignUri
            $fullConfig | Add-Member -NotePropertyName "assignments" -NotePropertyValue $assignments -Force
            $fullConfig | Add-Member -NotePropertyName "isAssigned" -NotePropertyValue ($null -ne $assignments -and $assignments.Count -gt 0) -Force

            $platformTag = "All"
            if ($config.'@odata.type') {
                if ($config.'@odata.type' -like "*ios*") { $platformTag = "iOS" }
                elseif ($config.'@odata.type' -like "*android*") { $platformTag = "Android" }
            }

            $fileName = Get-SafeFileName -FileName "$($config.displayName)_${platformTag}_$($config.id)"
            $filePath = Join-Path $configPath "$fileName.json"

            $fullConfig | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding UTF8
            $exportedCount++
            Write-Log "Exported: $($config.displayName)" -Level Verbose
        }
        Write-ColoredOutput "  Total: $exportedCount configurations" -Color $script:Colors.Success

        return $exportedCount
    }
    catch {
        Write-ColoredOutput "  Error exporting App Config (Managed Devices): $_" -Color $script:Colors.Error
        Write-Log "App Config (Managed Devices) export error: $($_.Exception.Message)" -Level Error
        return 0
    }
}

function Export-AppConfigManagedApps {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    Write-ColoredOutput "`n$(if ($script:IsLiveReport) { 'Fetching' } else { 'Exporting' }) App Configuration Policies (Managed Apps)..." -Color $script:Colors.Info

    try {
        # Use $expand to get apps and assignments in the same call to avoid separate requests with potential null IDs
        $uri = "https://graph.microsoft.com/beta/deviceAppManagement/targetedManagedAppConfigurations?`$expand=apps,assignments"
        $configs = Get-GraphData -Uri $uri

        if (-not $configs -or $configs.Count -eq 0) {
            Write-ColoredOutput "  No App Configuration Policies (Managed Apps) found" -Color $script:Colors.Info
            return 0
        }

        $exportedCount = 0
        $configPath = Join-Path $BackupPath "AppConfigManagedApps"
        New-Item -ItemType Directory -Path $configPath -Force | Out-Null

        foreach ($config in $configs) {
            # Skip configs with null/empty IDs
            if ([string]::IsNullOrWhiteSpace($config.id)) {
                continue
            }

            if (-not (Test-PlatformMatch -Policy $config)) {
                continue
            }

            Write-ColoredOutput "  - $($config.displayName)" -Color $script:Colors.Default

            # Get full config details with settings
            $detailUri = "https://graph.microsoft.com/beta/deviceAppManagement/targetedManagedAppConfigurations/$($config.id)?`$expand=apps,assignments"
            $response = Invoke-MgGraphRequest -Uri $detailUri -Method GET
            $fullConfig = ConvertTo-PolicyObject -Response $response

            # Ensure valid dates - try list API data first, then audit logs
            $fullConfig = Ensure-ValidDates -Policy $fullConfig -FallbackSource $config -TryAuditLogs $true

            # Derive isAssigned from embedded assignments (fetched via $expand=apps,assignments)
            $fullConfig | Add-Member -NotePropertyName "isAssigned" -NotePropertyValue ($null -ne $fullConfig.assignments -and @($fullConfig.assignments).Count -gt 0) -Force

            $platformTag = "All"
            if ($config.apps -and $config.apps.Count -gt 0) {
                $firstApp = $config.apps[0]
                if ($firstApp.mobileAppIdentifier) {
                    $appId = $firstApp.mobileAppIdentifier.'@odata.type'
                    if ($appId -like "*ios*") { $platformTag = "iOS" }
                    elseif ($appId -like "*android*") { $platformTag = "Android" }
                }
            }

            $fileName = Get-SafeFileName -FileName "$($config.displayName)_${platformTag}_$($config.id)"
            $filePath = Join-Path $configPath "$fileName.json"

            $fullConfig | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding UTF8
            $exportedCount++
            Write-Log "Exported: $($config.displayName)" -Level Verbose
        }
        Write-ColoredOutput "  Total: $exportedCount configurations" -Color $script:Colors.Success

        return $exportedCount
    }
    catch {
        Write-ColoredOutput "  Error exporting App Config (Managed Apps): $_" -Color $script:Colors.Error
        Write-Log "App Config (Managed Apps) export error: $($_.Exception.Message)" -Level Error
        return 0
    }
}

function Export-MacOSScripts {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    Write-ColoredOutput "`n$(if ($script:IsLiveReport) { 'Fetching' } else { 'Exporting' }) macOS Shell Scripts..." -Color $script:Colors.Info

    try {
        # Use $expand=assignments to get assignments in the same call
        $uri = "https://graph.microsoft.com/beta/deviceManagement/deviceShellScripts?`$expand=assignments"
        $scripts = Get-GraphData -Uri $uri

        if (-not $scripts -or $scripts.Count -eq 0) {
            Write-ColoredOutput "  No macOS Shell Scripts found" -Color $script:Colors.Info
            return 0
        }

        $exportedCount = 0
        $scriptsPath = Join-Path $BackupPath "MacOSScripts"
        New-Item -ItemType Directory -Path $scriptsPath -Force | Out-Null

        foreach ($scriptItem in $scripts) {
            # Skip scripts with null/empty IDs
            if ([string]::IsNullOrWhiteSpace($scriptItem.id)) {
                continue
            }

            Write-ColoredOutput "  - $($scriptItem.displayName)" -Color $script:Colors.Default

            # Get full script details including content (the list doesn't include scriptContent)
            $detailUri = "https://graph.microsoft.com/beta/deviceManagement/deviceShellScripts/$($scriptItem.id)"
            $response = Invoke-MgGraphRequest -Uri $detailUri -Method GET
            $fullScript = ConvertTo-PolicyObject -Response $response

            # Ensure valid dates - try list API data first, then audit logs
            $fullScript = Ensure-ValidDates -Policy $fullScript -FallbackSource $scriptItem -TryAuditLogs $true

            # Add assignments from the expanded query
            $fullScript | Add-Member -NotePropertyName "isAssigned" -NotePropertyValue ($null -ne $scriptItem.assignments -and @($scriptItem.assignments).Count -gt 0) -Force
            if ($scriptItem.assignments) {
                $fullScript | Add-Member -NotePropertyName "assignments" -NotePropertyValue $scriptItem.assignments -Force
            }

            $fileName = Get-SafeFileName -FileName "$($scriptItem.displayName)_macOS_$($scriptItem.id)"
            $filePath = Join-Path $scriptsPath "$fileName.json"

            $fullScript | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding UTF8
            $exportedCount++
            Write-Log "Exported macOS Script: $($scriptItem.displayName)" -Level Verbose
        }

        Write-ColoredOutput "  Total: $exportedCount scripts" -Color $script:Colors.Success

        return $exportedCount
    }
    catch {
        Write-ColoredOutput "  Error exporting macOS Shell Scripts: $_" -Color $script:Colors.Error
        Write-Log "macOS Shell Scripts export error: $($_.Exception.Message)" -Level Error
        return 0
    }
}

function Export-MacOSCustomAttributes {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    Write-ColoredOutput "`n$(if ($script:IsLiveReport) { 'Fetching' } else { 'Exporting' }) macOS Custom Attributes..." -Color $script:Colors.Info

    try {
        # Use $expand=assignments to get assignments in the same call
        $uri = "https://graph.microsoft.com/beta/deviceManagement/deviceCustomAttributeShellScripts?`$expand=assignments"
        $attributes = Get-GraphData -Uri $uri

        if (-not $attributes -or $attributes.Count -eq 0) {
            Write-ColoredOutput "  No macOS Custom Attributes found" -Color $script:Colors.Info
            return 0
        }

        $exportedCount = 0
        $attrPath = Join-Path $BackupPath "MacOSCustomAttributes"
        New-Item -ItemType Directory -Path $attrPath -Force | Out-Null

        foreach ($attr in $attributes) {
            # Skip attributes with null/empty IDs
            if ([string]::IsNullOrWhiteSpace($attr.id)) {
                continue
            }

            Write-ColoredOutput "  - $($attr.displayName)" -Color $script:Colors.Default

            # Get full attribute details (the list doesn't include scriptContent)
            $detailUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCustomAttributeShellScripts/$($attr.id)"
            $response = Invoke-MgGraphRequest -Uri $detailUri -Method GET
            $fullAttr = ConvertTo-PolicyObject -Response $response

            # Ensure valid dates - try list API data first, then audit logs
            $fullAttr = Ensure-ValidDates -Policy $fullAttr -FallbackSource $attr -TryAuditLogs $true

            # Add assignments from the expanded query
            $fullAttr | Add-Member -NotePropertyName "isAssigned" -NotePropertyValue ($null -ne $attr.assignments -and @($attr.assignments).Count -gt 0) -Force
            if ($attr.assignments) {
                $fullAttr | Add-Member -NotePropertyName "assignments" -NotePropertyValue $attr.assignments -Force
            }

            $fileName = Get-SafeFileName -FileName "$($attr.displayName)_macOS_$($attr.id)"
            $filePath = Join-Path $attrPath "$fileName.json"

            $fullAttr | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding UTF8
            $exportedCount++
            Write-Log "Exported macOS Custom Attribute: $($attr.displayName)" -Level Verbose
        }

        Write-ColoredOutput "  Total: $exportedCount custom attributes" -Color $script:Colors.Success

        return $exportedCount
    }
    catch {
        Write-ColoredOutput "  Error exporting macOS Custom Attributes: $_" -Color $script:Colors.Error
        Write-Log "macOS Custom Attributes export error: $($_.Exception.Message)" -Level Error
        return 0
    }
}

#endregion

#region Legacy Export Functions (Updated for v2.0)

function Export-DeviceConfigurations {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    Write-ColoredOutput "`n$(if ($script:IsLiveReport) { 'Fetching' } else { 'Exporting' }) Device Configuration Policies..." -Color $script:Colors.Info

    try {
        $uri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations"
        $configs = Get-GraphData -Uri $uri

        if (-not $configs -or $configs.Count -eq 0) {
            Write-ColoredOutput "  No Device Configurations found" -Color $script:Colors.Info
            return 0
        }

        $exportedCount = 0
        $configPath = Join-Path $BackupPath "DeviceConfigurations"
        New-Item -ItemType Directory -Path $configPath -Force | Out-Null

        foreach ($config in $configs) {
            # Skip WUfB policies - exported separately by Export-WUfBPolicies
            if ($config.'@odata.type' -eq '#microsoft.graph.windowsUpdateForBusinessConfiguration') {
                continue
            }

            if (-not (Test-PlatformMatch -Policy $config)) {
                continue
            }

            Write-ColoredOutput "  - $($config.displayName)" -Color $script:Colors.Default

            $fullPolicy = Get-PolicyDetails -PolicyId $config.id -PolicyType "deviceConfigurations" -IncludeAssignments

            if ($fullPolicy) {
                # Ensure valid dates - try list API data first, then audit logs
                $fullPolicy = Ensure-ValidDates -Policy $fullPolicy -FallbackSource $config -TryAuditLogs $true

                $cleanPolicy = Remove-ReadOnlyProperties -Policy $fullPolicy

                $platformTag = if ($fullPolicy.'@odata.type') {
                    switch -Wildcard ($fullPolicy.'@odata.type') {
                        "*windows*" { "Windows" }
                        "*ios*" { "iOS" }
                        "*android*" { "Android" }
                        "*macOS*" { "macOS" }
                        default { if ($config.platform) { $config.platform } else { "All" } }
                    }
                } elseif ($config.platform) { $config.platform } else { "All" }
                $fileName = Get-SafeFileName -FileName "$($config.displayName)_${platformTag}_$($config.id)"
                $filePath = Join-Path $configPath "$fileName.json"

                $fullPolicy | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding UTF8
                $exportedCount++
            }
        }

        Write-ColoredOutput "  Total: $exportedCount policies" -Color $script:Colors.Success
        return $exportedCount
    }
    catch {
        Write-ColoredOutput "  Error: $_" -Color $script:Colors.Error
        return 0
    }
}

function Export-CompliancePolicies {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    Write-ColoredOutput "`n$(if ($script:IsLiveReport) { 'Fetching' } else { 'Exporting' }) Compliance Policies..." -Color $script:Colors.Info

    try {
        $uri = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies"
        $policies = Get-GraphData -Uri $uri

        if (-not $policies -or $policies.Count -eq 0) {
            Write-ColoredOutput "  No Compliance Policies found" -Color $script:Colors.Info
            return 0
        }

        $exportedCount = 0
        $compliancePath = Join-Path $BackupPath "CompliancePolicies"
        New-Item -ItemType Directory -Path $compliancePath -Force | Out-Null

        foreach ($policy in $policies) {
            if (-not (Test-PlatformMatch -Policy $policy)) {
                continue
            }

            Write-ColoredOutput "  - $($policy.displayName)" -Color $script:Colors.Default

            $fullPolicy = Get-PolicyDetails -PolicyId $policy.id -PolicyType "deviceCompliancePolicies" -IncludeAssignments

            if ($fullPolicy) {
                # Fetch scheduledActionsForRule via $expand on the policy GET.
                # The sub-resource navigation endpoint (/scheduledActionsForRule) is not
                # supported by Graph API. Use $expand on the policy itself instead.
                try {
                    $expandUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies/$($policy.id)?`$expand=scheduledActionsForRule(`$expand=scheduledActionConfigurations)"
                    $expandedResponse = Invoke-MgGraphRequest -Uri $expandUri -Method GET -ErrorAction Stop
                    if ($expandedResponse.scheduledActionsForRule) {
                        $scheduledActions = [array]$expandedResponse.scheduledActionsForRule
                        $fullPolicy | Add-Member -NotePropertyName "scheduledActionsForRule" -NotePropertyValue $scheduledActions -Force
                    }
                } catch {
                    Write-Log "Warning: Could not fetch scheduledActionsForRule for $($policy.displayName): $_" -Level Warning
                }

                # Ensure valid dates - try list API data first, then audit logs
                $fullPolicy = Ensure-ValidDates -Policy $fullPolicy -FallbackSource $policy -TryAuditLogs $true

                $cleanPolicy = Remove-ReadOnlyProperties -Policy $fullPolicy

                $platformTag = if ($policy.'@odata.type') {
                    switch -Wildcard ($policy.'@odata.type') {
                        "*windows*" { "Windows" }
                        "*ios*" { "iOS" }
                        "*android*" { "Android" }
                        "*macOS*" { "macOS" }
                        default { "All" }
                    }
                } else { "All" }

                $fileName = Get-SafeFileName -FileName "$($policy.displayName)_${platformTag}_$($policy.id)"
                $filePath = Join-Path $compliancePath "$fileName.json"

                $fullPolicy | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding UTF8
                $exportedCount++
            }
        }

        Write-ColoredOutput "  Total: $exportedCount policies" -Color $script:Colors.Success
        return $exportedCount
    }
    catch {
        Write-ColoredOutput "  Error: $_" -Color $script:Colors.Error
        return 0
    }
}

function Export-SettingsCatalogPolicies {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    Write-ColoredOutput "`n$(if ($script:IsLiveReport) { 'Fetching' } else { 'Exporting' }) Settings Catalog Policies..." -Color $script:Colors.Info

    try {
        $uri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
        $allPolicies = Get-GraphData -Uri $uri

        # Exclude Autopilot Device Prep
        $policies = $allPolicies | Where-Object {
            -not ($_.templateReference -and $_.templateReference.templateId -eq "20f9c2d1-e508-4122-8692-7f284b3956f1")
        }

        if (-not $policies -or $policies.Count -eq 0) {
            Write-ColoredOutput "  No Settings Catalog Policies found" -Color $script:Colors.Info
            return 0
        }

        $exportedCount = 0
        $settingsPath = Join-Path $BackupPath "SettingsCatalogPolicies"
        New-Item -ItemType Directory -Path $settingsPath -Force | Out-Null

        foreach ($policy in $policies) {
            if (-not (Test-PlatformMatch -Policy $policy)) {
                continue
            }

            Write-ColoredOutput "  - $($policy.name)" -Color $script:Colors.Default

            $uri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$($policy.id)?`$expand=settings"
            $response = Invoke-MgGraphRequest -Uri $uri -Method GET
            $fullPolicy = ConvertTo-PolicyObject -Response $response

            # Ensure valid dates - try list API data first, then audit logs
            $fullPolicy = Ensure-ValidDates -Policy $fullPolicy -FallbackSource $policy -TryAuditLogs $true

            $assignUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$($policy.id)/assignments"
            $assignments = Get-GraphData -Uri $assignUri
            $fullPolicy | Add-Member -NotePropertyName "assignments" -NotePropertyValue $assignments -Force
            $fullPolicy | Add-Member -NotePropertyName "isAssigned" -NotePropertyValue ($null -ne $assignments -and $assignments.Count -gt 0) -Force

            $platformTag = if ($policy.platforms -and $policy.platforms.Count -gt 0) { $policy.platforms[0] } else { "All" }
            $fileName = Get-SafeFileName -FileName "$($policy.name)_${platformTag}_$($policy.id)"
            $filePath = Join-Path $settingsPath "$fileName.json"

            $fullPolicy | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding UTF8
            $exportedCount++
        }

        Write-ColoredOutput "  Total: $exportedCount policies" -Color $script:Colors.Success
        return $exportedCount
    }
    catch {
        Write-ColoredOutput "  Error: $_" -Color $script:Colors.Error
        return 0
    }
}

function Export-AdministrativeTemplates {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    Write-ColoredOutput "`n$(if ($script:IsLiveReport) { 'Fetching' } else { 'Exporting' }) Administrative Templates..." -Color $script:Colors.Info

    try {
        if ($script:SelectedPlatform -ne "All" -and $script:SelectedPlatform -ne "Windows") {
            Write-ColoredOutput "  Skipped (Platform filter: $($script:SelectedPlatform))" -Color $script:Colors.Warning
            return 0
        }

        $uri = "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations"
        $templates = Get-GraphData -Uri $uri

        if (-not $templates -or $templates.Count -eq 0) {
            Write-ColoredOutput "  No Administrative Templates found" -Color $script:Colors.Info
            return 0
        }

        $exportedCount = 0
        $templatesPath = Join-Path $BackupPath "AdministrativeTemplates"
        New-Item -ItemType Directory -Path $templatesPath -Force | Out-Null

        foreach ($template in $templates) {
            Write-ColoredOutput "  - $($template.displayName)" -Color $script:Colors.Default

            $uri = "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations/$($template.id)/definitionValues?`$expand=definition"
            $definitionValues = Get-GraphData -Uri $uri

            $templateData = $template
            $templateData | Add-Member -NotePropertyName "definitionValues" -NotePropertyValue $definitionValues -Force

            $assignUri = "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations/$($template.id)/assignments"
            $assignments = Get-GraphData -Uri $assignUri
            $templateData | Add-Member -NotePropertyName "assignments" -NotePropertyValue $assignments -Force
            $templateData | Add-Member -NotePropertyName "isAssigned" -NotePropertyValue ($null -ne $assignments -and $assignments.Count -gt 0) -Force

            # Ensure valid dates - try audit logs if dates are invalid
            $templateData = Ensure-ValidDates -Policy $templateData -TryAuditLogs $true

            $fileName = Get-SafeFileName -FileName "$($template.displayName)_Windows_$($template.id)"
            $filePath = Join-Path $templatesPath "$fileName.json"

            $templateData | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding UTF8
            $exportedCount++
        }

        Write-ColoredOutput "  Total: $exportedCount templates" -Color $script:Colors.Success
        return $exportedCount
    }
    catch {
        Write-ColoredOutput "  Error: $_" -Color $script:Colors.Error
        return 0
    }
}

#endregion

#region Main Backup Function

function Backup-IntuneConfiguration {
    param (
        [string]$BackupName,
        [string[]]$SelectedComponents
    )

    Write-Host ""
    Write-Host "  STARTING BACKUP PROCESS - UniFy-Endpoint v2.0" -ForegroundColor Cyan
    Write-Host "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan

    if (-not $SelectedComponents) {
        $SelectedComponents = Get-SelectedComponents -ComponentList $script:Components
    }

    if ($SelectedComponents.Count -eq 0) {
        Write-ColoredOutput "No components selected for backup." -Color $script:Colors.Warning
        return
    }

    # Filter components based on platform compatibility
    if ($script:SelectedPlatform -and $script:SelectedPlatform -ne "All") {
        $validComponents = Get-ValidComponentsForPlatform -Platform $script:SelectedPlatform
        $originalCount = $SelectedComponents.Count
        $SelectedComponents = $SelectedComponents | Where-Object { $validComponents -contains $_ }

        $filteredCount = $originalCount - $SelectedComponents.Count
        if ($filteredCount -gt 0) {
            Write-ColoredOutput "`nFiltered out $filteredCount incompatible component(s) for $($script:SelectedPlatform)" -Color $script:Colors.Warning
            Write-Log "Filtered $filteredCount components incompatible with $($script:SelectedPlatform)" -Level Info
        }

        if ($SelectedComponents.Count -eq 0) {
            Write-ColoredOutput "No compatible components selected for $($script:SelectedPlatform) platform." -Color $script:Colors.Warning
            return
        }
    }

    # Component display name mapping for user-friendly output
    $componentDisplayNames = @{
        "DeviceConfigurations"    = "Device Configurations"
        "CompliancePolicies"      = "Compliance Policies"
        "SettingsCatalogPolicies"   = "Settings Catalog"
        "AppProtectionPolicies"   = "App Protection Policies"
        "PowerShellScripts"       = "PowerShell Scripts"
        "AutopilotProfiles"       = "Autopilot Profiles"
        "AutopilotDevicePrep"     = "Autopilot Device Preparation"
        "EnrollmentStatusPage"    = "Enrollment Status Page"
        "RemediationScripts"      = "Remediation Scripts"
        "AssignmentFilters"       = "Assignment Filters"
        "AppConfigManagedDevices" = "App Configuration (Managed Devices)"
        "AppConfigManagedApps"    = "App Configuration (Managed Apps)"
        "MacOSScripts"            = "macOS Shell Scripts"
        "MacOSCustomAttributes"   = "macOS Custom Attributes"
    }

    Write-ColoredOutput "`nPlatform Filter: $($script:SelectedPlatform)" -Color $script:Colors.Info
    Write-ColoredOutput "Components to backup:" -Color $script:Colors.Info
    foreach ($component in $SelectedComponents) {
        $displayName = if ($componentDisplayNames.ContainsKey($component)) { $componentDisplayNames[$component] } else { $component }
        Write-ColoredOutput "  - $displayName" -Color $script:Colors.Default
    }

    # Create backup folder
    $backupDate = Get-Date -Format 'yyyy-MM-dd-HHmmss'
    $backupFolderName = if ($BackupName) { "backup-$BackupName-$backupDate" } else { "backup-$backupDate" }
    $backupFolder = Join-Path $script:BackupLocation $backupFolderName

    Write-ColoredOutput "`nBackup folder: $backupFolder" -Color $script:Colors.Info
    New-Item -ItemType Directory -Path $backupFolder -Force | Out-Null

    # Initialize counters
    $backupStartTime = Get-Date
    $backupCounts = @{
        DeviceConfigurations          = 0
        CompliancePolicies            = 0
        SettingsCatalogPolicies         = 0
        AppProtectionPolicies         = 0
        PowerShellScripts             = 0
        AutopilotProfiles             = 0
        AutopilotDevicePrep           = 0
        EnrollmentStatusPage          = 0
        RemediationScripts            = 0
        AssignmentFilters             = 0
        AppConfigManagedDevices       = 0
        AppConfigManagedApps          = 0
        MacOSScripts                  = 0
        MacOSCustomAttributes         = 0
    }

    # Execute exports based on selected components

    if ($SelectedComponents -contains "DeviceConfigurations") {
        $backupCounts.DeviceConfigurations = Export-DeviceConfigurations -BackupPath $backupFolder
    }

    if ($SelectedComponents -contains "CompliancePolicies") {
        $backupCounts.CompliancePolicies = Export-CompliancePolicies -BackupPath $backupFolder
    }

    if ($SelectedComponents -contains "SettingsCatalogPolicies") {
        $backupCounts.SettingsCatalogPolicies = Export-SettingsCatalogPolicies -BackupPath $backupFolder
    }

    if ($SelectedComponents -contains "AppProtectionPolicies") {
        $backupCounts.AppProtectionPolicies = Export-AppProtectionPolicies -BackupPath $backupFolder
    }

    if ($SelectedComponents -contains "PowerShellScripts") {
        $backupCounts.PowerShellScripts = Export-PowerShellScripts -BackupPath $backupFolder
    }

    if ($SelectedComponents -contains "AutopilotProfiles") {
        $backupCounts.AutopilotProfiles = Export-AutopilotProfiles -BackupPath $backupFolder
    }

    if ($SelectedComponents -contains "AutopilotDevicePrep") {
        $backupCounts.AutopilotDevicePrep = Export-AutopilotDevicePrep -BackupPath $backupFolder
    }

    if ($SelectedComponents -contains "EnrollmentStatusPage") {
        $backupCounts.EnrollmentStatusPage = Export-EnrollmentStatusPage -BackupPath $backupFolder
    }

    if ($SelectedComponents -contains "RemediationScripts") {
        $backupCounts.RemediationScripts = Export-RemediationScripts -BackupPath $backupFolder
    }

    if ($SelectedComponents -contains "AssignmentFilters") {
        $backupCounts.AssignmentFilters = Export-AssignmentFilters -BackupPath $backupFolder
    }

    if ($SelectedComponents -contains "AppConfigManagedDevices") {
        $backupCounts.AppConfigManagedDevices = Export-AppConfigManagedDevices -BackupPath $backupFolder
    }

    if ($SelectedComponents -contains "AppConfigManagedApps") {
        $backupCounts.AppConfigManagedApps = Export-AppConfigManagedApps -BackupPath $backupFolder
    }

    if ($SelectedComponents -contains "MacOSScripts") {
        $backupCounts.MacOSScripts = Export-MacOSScripts -BackupPath $backupFolder
    }

    if ($SelectedComponents -contains "MacOSCustomAttributes") {
        $backupCounts.MacOSCustomAttributes = Export-MacOSCustomAttributes -BackupPath $backupFolder
    }

    # Calculate backup duration
    $backupEndTime = Get-Date
    $backupDuration = $backupEndTime - $backupStartTime
    $durationSeconds = [Math]::Round($backupDuration.TotalSeconds)

    # Format duration
    if ($durationSeconds -lt 60) {
        $durationFormatted = "${durationSeconds}s"
    }
    else {
        $minutes = [Math]::Floor($durationSeconds / 60)
        $seconds = $durationSeconds % 60
        $durationFormatted = "${minutes}m ${seconds}s"
    }

    # Create metadata
    $totalItems = ($backupCounts.Values | Measure-Object -Sum).Sum
    $metadata = @{
        BackupDate         = $backupDate
        BackupFolder       = $backupFolderName
        StartTime          = $backupStartTime.ToString("yyyy-MM-dd HH:mm:ss")
        EndTime            = $backupEndTime.ToString("yyyy-MM-dd HH:mm:ss")
        Duration           = $durationFormatted
        DurationSeconds    = $durationSeconds
        TenantId           = (Get-MgContext).TenantId
        ItemCounts         = $backupCounts
        TotalItems         = $totalItems
        Version            = $script:Version
        SelectedComponents = $SelectedComponents
        PlatformFilter     = $script:SelectedPlatform
    }

    $metadataPath = Join-Path $backupFolder "metadata.json"
    $metadata | ConvertTo-Json -Depth 10 | Out-File -FilePath $metadataPath -Encoding UTF8

    # Summary
    Write-Host ""
    Write-Host "  ════════════════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host "                BACKUP COMPLETED SUCCESSFULLY (v2.0)                " -ForegroundColor Green
    Write-Host "  ════════════════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Location: " -ForegroundColor Cyan -NoNewline
    Write-Host "$backupFolder" -ForegroundColor White
    Write-Host "  Duration: " -ForegroundColor Cyan -NoNewline
    Write-Host "$durationFormatted" -ForegroundColor White
    Write-Host "  Total items: " -ForegroundColor Cyan -NoNewline
    Write-Host "$totalItems" -ForegroundColor White
    Write-Host "  Platform: " -ForegroundColor Cyan -NoNewline
    Write-Host "$($script:SelectedPlatform)" -ForegroundColor White
    Write-Host ""

    # Display breakdown
    Write-Host "  Breakdown:" -ForegroundColor Cyan
    foreach ($key in $backupCounts.Keys | Sort-Object) {
        if ($backupCounts[$key] -gt 0) {
            Write-Host "     $key" -ForegroundColor Gray -NoNewline
            Write-Host ": " -ForegroundColor DarkGray -NoNewline
            Write-Host "$($backupCounts[$key])" -ForegroundColor White
        }
    }
    Write-Host ""

    Write-Log "Backup completed successfully: $totalItems items in $durationFormatted" -Level Info

    return $backupFolder
}

#endregion

#region Restore Functions

function Restore-ConflictPolicy {
    <#
    .SYNOPSIS
    Restores a single conflicting policy from backup as a new [Restored] copy.

    .DESCRIPTION
    Handles the creation logic for each policy type when the admin chooses to restore
    a conflicting policy during the conflict resolution phase.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$ConflictInfo
    )

    $backupFile = $ConflictInfo.BackupFile
    $restoredName = $ConflictInfo.RestoredName
    $policyType = $ConflictInfo.PolicyType
    $apiEndpoint = $ConflictInfo.ApiEndpoint
    $nameField = if ($ConflictInfo.NameField) { $ConflictInfo.NameField } else { 'displayName' }
    $useV1Api = if ($ConflictInfo.UseV1Api) { $true } else { $false }

    $policy = Get-Content $backupFile | ConvertFrom-Json

    try {
        switch ($policyType) {
            'SettingsCatalogPolicies' {
                # Clean settings: strip settingDefinitions and OData annotations
                $cleanSettings = @()
                if ($policy.settings) {
                    foreach ($setting in @($policy.settings)) {
                        $settingJson = $setting | ConvertTo-Json -Depth 50 -Compress | ConvertFrom-Json
                        $propsToRemove = @($settingJson.PSObject.Properties.Name | Where-Object {
                            $_ -eq 'settingDefinitions' -or $_ -like '*@odata*'
                        })
                        foreach ($prop in $propsToRemove) {
                            $settingJson.PSObject.Properties.Remove($prop)
                        }
                        $cleanSettings += $settingJson
                    }
                }
                $newPolicy = @{
                    name              = $restoredName
                    description       = $policy.description
                    platforms         = $policy.platforms
                    technologies      = $policy.technologies
                    settings          = $cleanSettings
                    templateReference = $policy.templateReference
                }
                $body = $newPolicy | ConvertTo-Json -Depth 50
                $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/$apiEndpoint" `
                    -Method POST -Body $body -ContentType "application/json"
            }
            'AutopilotDevicePrep' {
                $newPolicy = @{
                    name              = $restoredName
                    description       = $policy.description
                    platforms         = $policy.platforms
                    technologies      = $policy.technologies
                    settings          = $policy.settings
                    templateReference = $policy.templateReference
                }
                $body = $newPolicy | ConvertTo-Json -Depth 50
                $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/$apiEndpoint" `
                    -Method POST -Body $body -ContentType "application/json"
            }
            'AdministrativeTemplates' {
                $body = @{
                    displayName = $restoredName
                    description = $policy.description
                } | ConvertTo-Json -Depth 20
                $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/$apiEndpoint" `
                    -Method POST -Body $body -ContentType "application/json"
            }
            'PowerShellScripts' {
                $policy = Remove-ReadOnlyProperties -Policy $policy
                $policy.displayName = $restoredName
                $policy.PSObject.Properties.Remove('scriptContentMd5Hash')
                $policy.PSObject.Properties.Remove('scriptContentSha256Hash')
                # Re-encode script content to Base64
                if ($policy.scriptContent -and -not [string]::IsNullOrWhiteSpace($policy.scriptContent)) {
                    try {
                        $policy.scriptContent = [System.Convert]::ToBase64String(
                            [System.Text.Encoding]::UTF8.GetBytes($policy.scriptContent))
                    } catch { }
                }
                $body = $policy | ConvertTo-Json -Depth 20
                $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/$apiEndpoint" `
                    -Method POST -Body $body -ContentType "application/json"
            }
            'RemediationScripts' {
                $policy = Remove-ReadOnlyProperties -Policy $policy
                $policy.displayName = $restoredName
                $policy.PSObject.Properties.Remove('detectionScriptMd5Hash')
                $policy.PSObject.Properties.Remove('detectionScriptSha256Hash')
                $policy.PSObject.Properties.Remove('remediationScriptMd5Hash')
                $policy.PSObject.Properties.Remove('remediationScriptSha256Hash')
                # Re-encode script content to Base64
                foreach ($scriptProp in @('detectionScriptContent', 'remediationScriptContent')) {
                    if ($policy.$scriptProp -and -not [string]::IsNullOrWhiteSpace($policy.$scriptProp)) {
                        try {
                            $policy.$scriptProp = [System.Convert]::ToBase64String(
                                [System.Text.Encoding]::UTF8.GetBytes($policy.$scriptProp))
                        } catch { }
                    }
                }
                $body = $policy | ConvertTo-Json -Depth 20
                $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/$apiEndpoint" `
                    -Method POST -Body $body -ContentType "application/json"
            }
            'MacOSScripts' {
                $policy = Remove-ReadOnlyProperties -Policy $policy
                $policy.displayName = $restoredName
                $policy.PSObject.Properties.Remove('scriptContentMd5Hash')
                $policy.PSObject.Properties.Remove('scriptContentSha256Hash')
                if ($policy.scriptContent -and -not [string]::IsNullOrWhiteSpace($policy.scriptContent)) {
                    try {
                        $policy.scriptContent = [System.Convert]::ToBase64String(
                            [System.Text.Encoding]::UTF8.GetBytes($policy.scriptContent))
                    } catch { }
                }
                $body = $policy | ConvertTo-Json -Depth 20
                $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/$apiEndpoint" `
                    -Method POST -Body $body -ContentType "application/json"
            }
            'MacOSCustomAttributes' {
                $policy = Remove-ReadOnlyProperties -Policy $policy
                $policy.displayName = $restoredName
                $policy.PSObject.Properties.Remove('scriptContentMd5Hash')
                $policy.PSObject.Properties.Remove('scriptContentSha256Hash')
                if ($policy.customAttributeScript -and $policy.customAttributeScript.scriptContent) {
                    try {
                        $policy.customAttributeScript.scriptContent = [System.Convert]::ToBase64String(
                            [System.Text.Encoding]::UTF8.GetBytes($policy.customAttributeScript.scriptContent))
                    } catch { }
                }
                $body = $policy | ConvertTo-Json -Depth 20
                $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/$apiEndpoint" `
                    -Method POST -Body $body -ContentType "application/json"
            }
            default {
                # Generic handler for: DeviceConfigurations, CompliancePolicies, AppProtectionPolicies,
                # AppConfigManagedDevices, AutopilotProfiles, EnrollmentStatusPage, WUfBPolicies,
                # AssignmentFilters, AppConfigManagedApps
                $policy = Remove-ReadOnlyProperties -Policy $policy

                # Compliance policies require scheduledActionsForRule with a block action
                if ($policyType -eq 'CompliancePolicies') {
                    if ($policy.scheduledActionsForRule -and ($policy.scheduledActionsForRule | Measure-Object).Count -gt 0) {
                        foreach ($rule in $policy.scheduledActionsForRule) {
                            if ($rule.PSObject.Properties['id']) { $rule.PSObject.Properties.Remove('id') }
                            if ($rule.scheduledActionConfigurations) {
                                foreach ($config in $rule.scheduledActionConfigurations) {
                                    if ($config.PSObject.Properties['id']) { $config.PSObject.Properties.Remove('id') }
                                }
                            }
                        }
                        $hasBlock = $policy.scheduledActionsForRule | Where-Object {
                            $_.scheduledActionConfigurations | Where-Object { $_.actionType -eq "block" }
                        }
                        if (-not $hasBlock) {
                            $blockConfig = [PSCustomObject]@{ actionType = "block"; gracePeriodHours = 0; notificationTemplateId = ""; notificationMessageCCList = @() }
                            $policy.scheduledActionsForRule[0].scheduledActionConfigurations = @($blockConfig) + @($policy.scheduledActionsForRule[0].scheduledActionConfigurations)
                        }
                    } else {
                        $defaultAction = [PSCustomObject]@{
                            ruleName = "PasswordRequired"
                            scheduledActionConfigurations = @(
                                [PSCustomObject]@{ actionType = "block"; gracePeriodHours = 0; notificationTemplateId = ""; notificationMessageCCList = @() }
                            )
                        }
                        $policy | Add-Member -NotePropertyName "scheduledActionsForRule" -NotePropertyValue @($defaultAction) -Force
                    }
                }

                # App Protection policies: apps must be assigned separately
                if ($policyType -eq 'AppProtectionPolicies') {
                    $policy.PSObject.Properties.Remove('apps')
                }

                if ($nameField -eq 'name') {
                    $policy.name = $restoredName
                } else {
                    $policy.displayName = $restoredName
                }
                $body = $policy | ConvertTo-Json -Depth 20
                $body = Convert-WcfDatesInJson -JsonString $body
                $apiVersion = if ($useV1Api) { "v1.0" } else { "beta" }
                $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/$apiVersion/$apiEndpoint" `
                    -Method POST -Body $body -ContentType "application/json"
            }
        }

        return @{ Success = $true; Error = $null }
    }
    catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Restore-IntuneConfiguration {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BackupPath,

        [Parameter(Mandatory = $false)]
        [switch]$Preview,

        [Parameter(Mandatory = $false)]
        [string[]]$SelectedComponents = @("All"),

        [Parameter(Mandatory = $false)]
        [switch]$IsImport
    )

    if (-not (Test-Path $BackupPath)) {
        Write-ColoredOutput "Backup path not found: $BackupPath" -Color $script:Colors.Error
        return
    }

    $metadataPath = Join-Path $BackupPath "metadata.json"
    if (-not (Test-Path $metadataPath)) {
        Write-ColoredOutput "Invalid backup: metadata.json not found" -Color $script:Colors.Error
        return
    }

    $metadata = Get-Content $metadataPath | ConvertFrom-Json

    $configHeader = if ($IsImport) { 'Import Configuration:' } else { 'Restore Configuration:' }
    Write-ColoredOutput "`n$configHeader" -Color $script:Colors.Header
    Write-ColoredOutput "Backup Date: $($metadata.BackupDate)" -Color $script:Colors.Info
    Write-ColoredOutput "Total Items: $($metadata.TotalItems)" -Color $script:Colors.Info
    Write-ColoredOutput "Preview Mode: $Preview" -Color $script:Colors.Info
    $behaviorMsg = if ($IsImport) {
        "Policy Behavior: New policies get '- [New Policy]' suffix, existing-match policies get '- [Imported]' suffix"
    } else {
        "Policy Behavior: New policies get '- [New Policy]' suffix, dirty policies get '- [Restored]' suffix"
    }
    Write-ColoredOutput $behaviorMsg -Color $script:Colors.Info

    if (-not $Preview) {
        Write-ColoredOutput "`nWARNING: This will create NEW policies with naming suffixes in your Intune tenant!" -Color $script:Colors.Warning
        Write-ColoredOutput "Existing policies will NOT be modified or deleted - they will be skipped." -Color $script:Colors.Info
        $confirm = Read-Host "`nAre you sure you want to continue? (yes/no)"
        if ($confirm -ne "yes") {
            Write-ColoredOutput "Restore cancelled." -Color $script:Colors.Warning
            return
        }
    }

    # Get components to restore
    if ($SelectedComponents -contains "All") {
        $componentsToRestore = @("DeviceConfigurations", "CompliancePolicies", "SettingsCatalogPolicies",
                                  "AppProtectionPolicies", "PowerShellScripts", "AutopilotProfiles",
                                  "AutopilotDevicePrep", "EnrollmentStatusPage", "RemediationScripts",
                                  "AssignmentFilters", "AppConfigManagedDevices", "AppConfigManagedApps",
                                  "MacOSScripts", "MacOSCustomAttributes", "LegacyIntents")
    }
    else {
        $componentsToRestore = $SelectedComponents
    }

    # Filter components based on platform compatibility
    if ($script:SelectedPlatform -and $script:SelectedPlatform -ne "All") {
        $validComponents = Get-ValidComponentsForPlatform -Platform $script:SelectedPlatform
        $originalCount = $componentsToRestore.Count
        $componentsToRestore = $componentsToRestore | Where-Object { $validComponents -contains $_ }

        $filteredCount = $originalCount - $componentsToRestore.Count
        if ($filteredCount -gt 0) {
            Write-ColoredOutput "`nFiltered out $filteredCount incompatible component(s) for $($script:SelectedPlatform)" -Color $script:Colors.Warning
            Write-Log "Filtered $filteredCount components incompatible with $($script:SelectedPlatform)" -Level Info
        }

        if ($componentsToRestore.Count -eq 0) {
            Write-ColoredOutput "No compatible components selected for $($script:SelectedPlatform) platform." -Color $script:Colors.Warning
            return
        }
    }

    $conflictPolicies = [System.Collections.ArrayList]::new()

    $restoreStats = @{
        Success      = 0
        Skipped      = 0
        Failed       = 0
        Unchanged    = 0
        Modified     = 0
        Deleted      = 0
        NameConflict = 0
        Renamed      = 0
        ConflictResolved = 0
    }

    $actionVerb = if ($IsImport) { 'Importing' } else { 'Restoring' }

    # 1. Restore Device Configurations
    if ($componentsToRestore -contains "DeviceConfigurations") {
        $configPath = Join-Path $BackupPath "DeviceConfigurations"
        if (Test-Path $configPath) {
            Write-ColoredOutput "`n$actionVerb Device Configurations..." -Color $script:Colors.Info
            $configs = Get-ChildItem -Path $configPath -Filter "*.json"

            foreach ($configFile in $configs) {
                $config = Get-Content $configFile.FullName | ConvertFrom-Json
                $originalName = $config.displayName

                # Strip existing restore markers (backward compat: old prefix and new suffix)
                if ($originalName -like "[Restored] *") {
                    $baseName = $originalName -replace '^\[Restored\] ', ''
                } elseif ($originalName -match ' - \[(Restored|New Policy)\]$') {
                    $baseName = $originalName -replace ' - \[(Restored|New Policy)\]$', ''
                } else {
                    $baseName = $originalName
                }
                $restoredName = "$baseName - [Restored]"
                $newPolicyName = if ($IsImport) { "$baseName - [New Policy]" } else { "$baseName - [Restored]" }

                Write-ColoredOutput "  - $baseName" -Color $script:Colors.Default -NoNewline

                # Use hybrid checking (ID + Name + Content comparison)
                $status = Get-PolicyExistenceStatus -BackupPolicy $config `
                    -ApiEndpoint "deviceManagement/deviceConfigurations" `
                    -RestoredName $restoredName -OriginalName $originalName

                $shouldRestore = Write-RestoreStatus -Status $status -Preview $Preview -RestoreStats $restoreStats -IsImport $IsImport
                if (-not $shouldRestore) {
                    if (-not $Preview -and -not $IsImport -and $status.Status -in @('NameConflict', 'Modified', 'Renamed')) {
                        [void]$conflictPolicies.Add(@{
                            PolicyType = 'DeviceConfigurations'; PolicyName = $baseName; BackupFile = $configFile.FullName
                            Status = $status; ApiEndpoint = 'deviceManagement/deviceConfigurations'
                            NameField = 'displayName'; RestoredName = $restoredName; OriginalName = $originalName
                        })
                    }
                    continue
                }

                $config = Remove-ReadOnlyProperties -Policy $config
                $config.displayName = $newPolicyName

                try {
                    $body = $config | ConvertTo-Json -Depth 20
                    # Convert WCF date format /Date(...)/ to ISO 8601 (e.g., WUfB policies)
                    $body = Convert-WcfDatesInJson -JsonString $body
                    $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations" `
                        -Method POST -Body $body -ContentType "application/json"

                    Write-ColoredOutput " [CREATED]" -Color $script:Colors.Success
                    $restoreStats.Success++
                }
                catch {
                    Write-ColoredOutput " [FAILED: $_]" -Color $script:Colors.Error
                    $restoreStats.Failed++
                }
            }
        }
    }

    # 2. Restore Compliance Policies
    if ($componentsToRestore -contains "CompliancePolicies") {
        $compliancePath = Join-Path $BackupPath "CompliancePolicies"
        if (Test-Path $compliancePath) {
            Write-ColoredOutput "`n$actionVerb Compliance Policies..." -Color $script:Colors.Info
            $policies = Get-ChildItem -Path $compliancePath -Filter "*.json"

            foreach ($policyFile in $policies) {
                $policy = Get-Content $policyFile.FullName | ConvertFrom-Json
                $originalName = $policy.displayName

                # Strip existing restore markers (backward compat: old prefix and new suffix)
                if ($originalName -like "[Restored] *") {
                    $baseName = $originalName -replace '^\[Restored\] ', ''
                } elseif ($originalName -match ' - \[(Restored|New Policy)\]$') {
                    $baseName = $originalName -replace ' - \[(Restored|New Policy)\]$', ''
                } else {
                    $baseName = $originalName
                }
                $restoredName = "$baseName - [Restored]"
                $newPolicyName = if ($IsImport) { "$baseName - [New Policy]" } else { "$baseName - [Restored]" }

                Write-ColoredOutput "  - $baseName" -Color $script:Colors.Default -NoNewline

                # Use hybrid checking (ID + Name + Content comparison)
                $status = Get-PolicyExistenceStatus -BackupPolicy $policy `
                    -ApiEndpoint "deviceManagement/deviceCompliancePolicies" `
                    -RestoredName $restoredName -OriginalName $originalName

                $shouldRestore = Write-RestoreStatus -Status $status -Preview $Preview -RestoreStats $restoreStats -IsImport $IsImport
                if (-not $shouldRestore) {
                    if (-not $Preview -and -not $IsImport -and $status.Status -in @('NameConflict', 'Modified', 'Renamed')) {
                        [void]$conflictPolicies.Add(@{
                            PolicyType = 'CompliancePolicies'; PolicyName = $baseName; BackupFile = $policyFile.FullName
                            Status = $status; ApiEndpoint = 'deviceManagement/deviceCompliancePolicies'
                            NameField = 'displayName'; RestoredName = $restoredName; OriginalName = $originalName
                        })
                    }
                    continue
                }

                $policy = Remove-ReadOnlyProperties -Policy $policy

                # Handle scheduledActionsForRule: strip nested IDs and ensure a block action exists
                # (Graph API requires exactly one block action; field is a navigation property not in default GET)
                if ($policy.scheduledActionsForRule -and ($policy.scheduledActionsForRule | Measure-Object).Count -gt 0) {
                    # Strip read-only IDs from nested objects (IDs cause 400 errors on POST)
                    foreach ($rule in $policy.scheduledActionsForRule) {
                        if ($rule.PSObject.Properties['id']) { $rule.PSObject.Properties.Remove('id') }
                        if ($rule.scheduledActionConfigurations) {
                            foreach ($config in $rule.scheduledActionConfigurations) {
                                if ($config.PSObject.Properties['id']) { $config.PSObject.Properties.Remove('id') }
                            }
                        }
                    }
                    # Ensure at least one block action exists
                    $hasBlock = $policy.scheduledActionsForRule | Where-Object {
                        $_.scheduledActionConfigurations | Where-Object { $_.actionType -eq "block" }
                    }
                    if (-not $hasBlock) {
                        $blockConfig = [PSCustomObject]@{ actionType = "block"; gracePeriodHours = 0; notificationTemplateId = ""; notificationMessageCCList = @() }
                        $policy.scheduledActionsForRule[0].scheduledActionConfigurations = @($blockConfig) + @($policy.scheduledActionsForRule[0].scheduledActionConfigurations)
                    }
                } else {
                    # No scheduledActionsForRule in backup JSON — inject default (required by Graph API)
                    $defaultAction = [PSCustomObject]@{
                        ruleName = "PasswordRequired"
                        scheduledActionConfigurations = @(
                            [PSCustomObject]@{ actionType = "block"; gracePeriodHours = 0; notificationTemplateId = ""; notificationMessageCCList = @() }
                        )
                    }
                    $policy | Add-Member -NotePropertyName "scheduledActionsForRule" -NotePropertyValue @($defaultAction) -Force
                }

                $policy.displayName = $newPolicyName

                try {
                    $body = $policy | ConvertTo-Json -Depth 20
                    $body = Convert-WcfDatesInJson -JsonString $body
                    $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies" `
                        -Method POST -Body $body -ContentType "application/json"

                    Write-ColoredOutput " [CREATED]" -Color $script:Colors.Success
                    $restoreStats.Success++
                }
                catch {
                    Write-ColoredOutput " [FAILED: $_]" -Color $script:Colors.Error
                    $restoreStats.Failed++
                }
            }
        }
    }

    # 3. Restore Settings Catalog Policies
    if ($componentsToRestore -contains "SettingsCatalogPolicies") {
        $settingsPath = Join-Path $BackupPath "SettingsCatalogPolicies"
        if (Test-Path $settingsPath) {
            Write-ColoredOutput "`n$actionVerb Settings Catalog Policies..." -Color $script:Colors.Info
            $policies = Get-ChildItem -Path $settingsPath -Filter "*.json"

            foreach ($policyFile in $policies) {
                $policy = Get-Content $policyFile.FullName | ConvertFrom-Json
                # Settings Catalog uses 'name' but external JSON may use 'displayName'
                $originalName = if ($policy.name) { $policy.name } else { $policy.displayName }

                # Strip existing restore markers (backward compat: old prefix and new suffix)
                if ($originalName -like "[Restored] *") {
                    $baseName = $originalName -replace '^\[Restored\] ', ''
                } elseif ($originalName -match ' - \[(Restored|New Policy)\]$') {
                    $baseName = $originalName -replace ' - \[(Restored|New Policy)\]$', ''
                } else {
                    $baseName = $originalName
                }
                $restoredName = "$baseName - [Restored]"
                $newPolicyName = if ($IsImport) { "$baseName - [New Policy]" } else { "$baseName - [Restored]" }

                Write-ColoredOutput "  - $baseName" -Color $script:Colors.Default -NoNewline

                # Use hybrid checking (ID + Name + Content comparison) - Settings Catalog uses 'name' field
                $status = Get-PolicyExistenceStatus -BackupPolicy $policy `
                    -ApiEndpoint "deviceManagement/configurationPolicies" `
                    -RestoredName $restoredName -OriginalName $originalName -NameField 'name' `
                    -ExpandQuery 'settings'

                $shouldRestore = Write-RestoreStatus -Status $status -Preview $Preview -RestoreStats $restoreStats -IsImport $IsImport
                if (-not $shouldRestore) {
                    if (-not $Preview -and -not $IsImport -and $status.Status -in @('NameConflict', 'Modified', 'Renamed')) {
                        [void]$conflictPolicies.Add(@{
                            PolicyType = 'SettingsCatalogPolicies'; PolicyName = $baseName; BackupFile = $policyFile.FullName
                            Status = $status; ApiEndpoint = 'deviceManagement/configurationPolicies'
                            NameField = 'name'; RestoredName = $restoredName; OriginalName = $originalName
                            ExpandQuery = 'settings'
                        })
                    }
                    continue
                }

                # Clean settings: strip settingDefinitions and OData annotations from each setting item
                $cleanSettings = @()
                if ($policy.settings) {
                    foreach ($setting in @($policy.settings)) {
                        $settingJson = $setting | ConvertTo-Json -Depth 50 -Compress | ConvertFrom-Json
                        # Remove settingDefinitions and any @odata annotation properties recursively
                        $propsToRemove = @($settingJson.PSObject.Properties.Name | Where-Object {
                            $_ -eq 'settingDefinitions' -or $_ -like '*@odata*'
                        })
                        foreach ($prop in $propsToRemove) {
                            $settingJson.PSObject.Properties.Remove($prop)
                        }
                        $cleanSettings += $settingJson
                    }
                }

                # Fix secret-type settings: Graph API exports Secret values as plain strings,
                # but reimporting them requires the SecretSettingValue type. Recursively find
                # any simpleSettingValue whose parent settingDefinitionId implies a secret
                # (contains 'password', 'secret', 'credential', or 'apikey') and fix the type.
                $fixSecretSettings = $null
                $fixSecretSettings = {
                    param ($node)
                    if ($node -is [System.Management.Automation.PSCustomObject]) {
                        if ($node.settingDefinitionId -and
                            ($node.settingDefinitionId -imatch 'password|secret|credential|apikey') -and
                            $node.simpleSettingValue -and
                            ($node.simpleSettingValue.'@odata.type' -like '*StringSettingValue')) {
                            $node.simpleSettingValue.'@odata.type' = '#microsoft.graph.deviceManagementConfigurationSecretSettingValue'
                            if (-not ($node.simpleSettingValue.PSObject.Properties['valueState'])) {
                                $node.simpleSettingValue | Add-Member -NotePropertyName 'valueState' -NotePropertyValue 'notEncrypted' -Force
                            }
                        }
                        foreach ($prop in $node.PSObject.Properties) {
                            & $fixSecretSettings $prop.Value
                        }
                    } elseif ($node -is [System.Collections.IEnumerable] -and $node -isnot [string]) {
                        foreach ($item in $node) { & $fixSecretSettings $item }
                    }
                }
                foreach ($s in $cleanSettings) { & $fixSecretSettings $s }

                # Fallback: if a secret-keyword field still has a placeholder StringSettingValue
                # (e.g. <YOUR WIFI PASSWORD> from community baselines), convert to empty SecretSettingValue.
                # Graph API rejects StringSettingValue for Secret-type definitions; empty notEncrypted is accepted.
                $stripSecretPlaceholders = $null
                $stripSecretPlaceholders = {
                    param ($node)
                    if ($node -is [System.Management.Automation.PSCustomObject]) {
                        if ($node.settingDefinitionId -and
                            ($node.settingDefinitionId -imatch 'password|secret|credential|apikey') -and
                            $node.simpleSettingValue -and
                            ($node.simpleSettingValue.'@odata.type' -like '*StringSettingValue') -and
                            ($node.simpleSettingValue.value -match '^<.*>$')) {
                            $node.simpleSettingValue.'@odata.type' = '#microsoft.graph.deviceManagementConfigurationSecretSettingValue'
                            $node.simpleSettingValue.value = ''
                            if (-not ($node.simpleSettingValue.PSObject.Properties['valueState'])) {
                                $node.simpleSettingValue | Add-Member -NotePropertyName 'valueState' -NotePropertyValue 'notEncrypted' -Force
                            }
                        }
                        foreach ($prop in $node.PSObject.Properties) {
                            & $stripSecretPlaceholders $prop.Value
                        }
                    } elseif ($node -is [System.Collections.IEnumerable] -and $node -isnot [string]) {
                        foreach ($item in $node) { & $stripSecretPlaceholders $item }
                    }
                }
                foreach ($s in $cleanSettings) { & $stripSecretPlaceholders $s }

                $newPolicy = @{
                    name              = $newPolicyName
                    description       = $policy.description
                    platforms         = $policy.platforms
                    technologies      = $policy.technologies
                    settings          = $cleanSettings
                    templateReference = $policy.templateReference
                }

                try {
                    $body = $newPolicy | ConvertTo-Json -Depth 50
                    $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" `
                        -Method POST -Body $body -ContentType "application/json"

                    Write-ColoredOutput " [CREATED]" -Color $script:Colors.Success
                    $restoreStats.Success++
                }
                catch {
                    Write-ColoredOutput " [FAILED: $_]" -Color $script:Colors.Error
                    $restoreStats.Failed++
                }
            }
        }
    }

    # 4. Restore App Protection Policies
    if ($componentsToRestore -contains "AppProtectionPolicies") {
        $appProtectionPath = Join-Path $BackupPath "AppProtectionPolicies"
        if (Test-Path $appProtectionPath) {
            Write-ColoredOutput "`n$actionVerb App Protection Policies..." -Color $script:Colors.Info

            # Pre-fetch all existing App Protection Policies (API doesn't support $filter by displayName reliably)
            $existingAndroidPolicies = @(Get-GraphData -Uri "https://graph.microsoft.com/beta/deviceAppManagement/androidManagedAppProtections")
            $existingiOSPolicies = @(Get-GraphData -Uri "https://graph.microsoft.com/beta/deviceAppManagement/iosManagedAppProtections")
            $existingWindowsPolicies = @(Get-GraphData -Uri "https://graph.microsoft.com/beta/deviceAppManagement/windowsManagedAppProtections")


            $policies = Get-ChildItem -Path $appProtectionPath -Filter "*.json"

            foreach ($policyFile in $policies) {
                $policy = Get-Content $policyFile.FullName | ConvertFrom-Json
                $originalName = $policy.displayName

                # Strip existing restore markers (backward compat: old prefix and new suffix)
                if ($originalName -like "[Restored] *") {
                    $baseName = $originalName -replace '^\[Restored\] ', ''
                } elseif ($originalName -match ' - \[(Restored|New Policy)\]$') {
                    $baseName = $originalName -replace ' - \[(Restored|New Policy)\]$', ''
                } else {
                    $baseName = $originalName
                }
                $restoredName = "$baseName - [Restored]"
                $newPolicyName = if ($IsImport) { "$baseName - [New Policy]" } else { "$baseName - [Restored]" }

                $policyType = $policy.'@odata.type'
                $endpoint = $null
                $existingPolicies = @()

                # If @odata.type is missing, detect from filename (backup files include _iOS_, _Android_, _Windows_ suffix)
                if ([string]::IsNullOrWhiteSpace($policyType)) {
                    $fileName = $policyFile.Name
                    if ($fileName -match '_iOS_') {
                        $policyType = "#microsoft.graph.iosManagedAppProtection"
                    }
                    elseif ($fileName -match '_Android_') {
                        $policyType = "#microsoft.graph.androidManagedAppProtection"
                    }
                    elseif ($fileName -match '_Windows_') {
                        $policyType = "#microsoft.graph.windowsManagedAppProtection"
                    }
                }

                switch ($policyType) {
                    "#microsoft.graph.androidManagedAppProtection" {
                        $endpoint = "deviceAppManagement/androidManagedAppProtections"
                        $existingPolicies = $existingAndroidPolicies
                    }
                    "#microsoft.graph.iosManagedAppProtection" {
                        $endpoint = "deviceAppManagement/iosManagedAppProtections"
                        $existingPolicies = $existingiOSPolicies
                    }
                    "#microsoft.graph.windowsManagedAppProtection" {
                        $endpoint = "deviceAppManagement/windowsManagedAppProtections"
                        $existingPolicies = $existingWindowsPolicies
                    }
                }

                Write-ColoredOutput "  - $baseName" -Color $script:Colors.Default -NoNewline

                if (-not $endpoint) {
                    Write-ColoredOutput " [FAILED: Unknown policy type]" -Color $script:Colors.Error
                    $restoreStats.Failed++
                    continue
                }

                # Use hybrid checking with pre-fetched policies (API doesn't support $filter reliably)
                $status = Get-PolicyExistenceStatus -BackupPolicy $policy `
                    -ApiEndpoint $endpoint `
                    -ExistingPolicies $existingPolicies `
                    -RestoredName $restoredName -OriginalName $originalName

                $shouldRestore = Write-RestoreStatus -Status $status -Preview $Preview -RestoreStats $restoreStats -IsImport $IsImport
                if (-not $shouldRestore) {
                    if (-not $Preview -and -not $IsImport -and $status.Status -in @('NameConflict', 'Modified', 'Renamed')) {
                        [void]$conflictPolicies.Add(@{
                            PolicyType = 'AppProtectionPolicies'; PolicyName = $baseName; BackupFile = $policyFile.FullName
                            Status = $status; ApiEndpoint = $endpoint
                            NameField = 'displayName'; RestoredName = $restoredName; OriginalName = $originalName
                        })
                    }
                    continue
                }

                $policy.displayName = $newPolicyName
                $policy.PSObject.Properties.Remove('id')
                $policy.PSObject.Properties.Remove('createdDateTime')
                $policy.PSObject.Properties.Remove('lastModifiedDateTime')
                $policy.PSObject.Properties.Remove('version')
                $policy.PSObject.Properties.Remove('assignments')
                $policy.PSObject.Properties.Remove('apps')  # apps must be assigned separately via apps endpoint

                # Get the short endpoint for POST
                $shortEndpoint = $endpoint -replace 'deviceAppManagement/', ''

                try {
                    $body = $policy | ConvertTo-Json -Depth 10
                    $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/$shortEndpoint" `
                        -Method POST -Body $body -ContentType "application/json"

                    Write-ColoredOutput " [CREATED]" -Color $script:Colors.Success
                    $restoreStats.Success++
                }
                catch {
                    Write-ColoredOutput " [FAILED: $_]" -Color $script:Colors.Error
                    $restoreStats.Failed++
                }
            }
        }
    }

    # 5. Restore PowerShell Scripts
    if ($componentsToRestore -contains "PowerShellScripts") {
        $scriptsPath = Join-Path $BackupPath "PowerShellScripts"
        if (Test-Path $scriptsPath) {
            Write-ColoredOutput "`n$actionVerb PowerShell Scripts..." -Color $script:Colors.Info
            $scripts = Get-ChildItem -Path $scriptsPath -Filter "*.json"

            foreach ($scriptFile in $scripts) {
                $psScript = Get-Content $scriptFile.FullName | ConvertFrom-Json
                $originalName = $psScript.displayName

                # Strip existing restore markers (backward compat: old prefix and new suffix)
                if ($originalName -like "[Restored] *") {
                    $baseName = $originalName -replace '^\[Restored\] ', ''
                } elseif ($originalName -match ' - \[(Restored|New Policy)\]$') {
                    $baseName = $originalName -replace ' - \[(Restored|New Policy)\]$', ''
                } else {
                    $baseName = $originalName
                }
                $restoredName = "$baseName - [Restored]"
                $newPolicyName = if ($IsImport) { "$baseName - [New Policy]" } else { "$baseName - [Restored]" }

                Write-ColoredOutput "  - $baseName" -Color $script:Colors.Default -NoNewline

                # Use hybrid checking (ID + Name + Content comparison)
                $status = Get-PolicyExistenceStatus -BackupPolicy $psScript `
                    -ApiEndpoint "deviceManagement/deviceManagementScripts" `
                    -RestoredName $restoredName -OriginalName $originalName

                $shouldRestore = Write-RestoreStatus -Status $status -Preview $Preview -RestoreStats $restoreStats -IsImport $IsImport
                if (-not $shouldRestore) {
                    if (-not $Preview -and -not $IsImport -and $status.Status -in @('NameConflict', 'Modified', 'Renamed')) {
                        [void]$conflictPolicies.Add(@{
                            PolicyType = 'PowerShellScripts'; PolicyName = $baseName; BackupFile = $scriptFile.FullName
                            Status = $status; ApiEndpoint = 'deviceManagement/deviceManagementScripts'
                            NameField = 'displayName'; RestoredName = $restoredName; OriginalName = $originalName
                        })
                    }
                    continue
                }

                # Re-encode script content to Base64
                if ($psScript.scriptContent -and -not [string]::IsNullOrWhiteSpace($psScript.scriptContent)) {
                    try {
                        $psScript.scriptContent = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($psScript.scriptContent))
                    }
                    catch {
                        Write-Log "Script content might already be Base64 encoded" -Level Warning
                    }
                }

                $psScript.displayName = $newPolicyName
                $psScript.PSObject.Properties.Remove('id')
                $psScript.PSObject.Properties.Remove('createdDateTime')
                $psScript.PSObject.Properties.Remove('lastModifiedDateTime')
                $psScript.PSObject.Properties.Remove('assignments')
                $psScript.PSObject.Properties.Remove('scriptContentMd5Hash')
                $psScript.PSObject.Properties.Remove('scriptContentSha256Hash')

                try {
                    $body = $psScript | ConvertTo-Json -Depth 10
                    $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts" `
                        -Method POST -Body $body -ContentType "application/json"

                    Write-ColoredOutput " [CREATED]" -Color $script:Colors.Success
                    $restoreStats.Success++
                }
                catch {
                    Write-ColoredOutput " [FAILED: $_]" -Color $script:Colors.Error
                    $restoreStats.Failed++
                }
            }
        }
    }

    # 6. Restore Administrative Templates
    if ($componentsToRestore -contains "AdministrativeTemplates") {
        $templatesPath = Join-Path $BackupPath "AdministrativeTemplates"
        if (Test-Path $templatesPath) {
            Write-ColoredOutput "`n$actionVerb Administrative Templates..." -Color $script:Colors.Info
            $templates = Get-ChildItem -Path $templatesPath -Filter "*.json"

            foreach ($templateFile in $templates) {
                $template = Get-Content $templateFile.FullName | ConvertFrom-Json
                $originalName = $template.displayName

                # Strip existing restore markers (backward compat: old prefix and new suffix)
                if ($originalName -like "[Restored] *") {
                    $baseName = $originalName -replace '^\[Restored\] ', ''
                } elseif ($originalName -match ' - \[(Restored|New Policy)\]$') {
                    $baseName = $originalName -replace ' - \[(Restored|New Policy)\]$', ''
                } else {
                    $baseName = $originalName
                }
                $restoredName = "$baseName - [Restored]"
                $newPolicyName = if ($IsImport) { "$baseName - [New Policy]" } else { "$baseName - [Restored]" }

                Write-ColoredOutput "  - $baseName" -Color $script:Colors.Default -NoNewline

                # Use hybrid checking (ID + Name + Content comparison)
                $status = Get-PolicyExistenceStatus -BackupPolicy $template `
                    -ApiEndpoint "deviceManagement/groupPolicyConfigurations" `
                    -RestoredName $restoredName -OriginalName $originalName

                $shouldRestore = Write-RestoreStatus -Status $status -Preview $Preview -RestoreStats $restoreStats -IsImport $IsImport
                if (-not $shouldRestore) {
                    if (-not $Preview -and -not $IsImport -and $status.Status -in @('NameConflict', 'Modified', 'Renamed')) {
                        [void]$conflictPolicies.Add(@{
                            PolicyType = 'AdministrativeTemplates'; PolicyName = $baseName; BackupFile = $templateFile.FullName
                            Status = $status; ApiEndpoint = 'deviceManagement/groupPolicyConfigurations'
                            NameField = 'displayName'; RestoredName = $restoredName; OriginalName = $originalName
                        })
                    }
                    continue
                }

                $template.displayName = $newPolicyName
                $template.PSObject.Properties.Remove('id')
                $template.PSObject.Properties.Remove('createdDateTime')
                $template.PSObject.Properties.Remove('lastModifiedDateTime')
                $template.PSObject.Properties.Remove('version')
                $template.PSObject.Properties.Remove('assignments')

                try {
                    $templateBody = @{
                        displayName = $template.displayName
                        description = $template.description
                    } | ConvertTo-Json

                    $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations" `
                        -Method POST -Body $templateBody -ContentType "application/json"

                    Write-ColoredOutput " [CREATED]" -Color $script:Colors.Success
                    $restoreStats.Success++
                }
                catch {
                    Write-ColoredOutput " [FAILED: $_]" -Color $script:Colors.Error
                    $restoreStats.Failed++
                }
            }
        }
    }

    # 7. Restore Remediation Scripts
    if ($componentsToRestore -contains "RemediationScripts") {
        $remediationPath = Join-Path $BackupPath "RemediationScripts"
        if (Test-Path $remediationPath) {
            Write-ColoredOutput "`n$actionVerb Remediation Scripts..." -Color $script:Colors.Info
            $scripts = Get-ChildItem -Path $remediationPath -Filter "*.json"

            foreach ($scriptFile in $scripts) {
                $remediation = Get-Content $scriptFile.FullName | ConvertFrom-Json
                $originalName = $remediation.displayName

                # Strip existing restore markers (backward compat: old prefix and new suffix)
                if ($originalName -like "[Restored] *") {
                    $baseName = $originalName -replace '^\[Restored\] ', ''
                } elseif ($originalName -match ' - \[(Restored|New Policy)\]$') {
                    $baseName = $originalName -replace ' - \[(Restored|New Policy)\]$', ''
                } else {
                    $baseName = $originalName
                }
                $restoredName = "$baseName - [Restored]"
                $newPolicyName = if ($IsImport) { "$baseName - [New Policy]" } else { "$baseName - [Restored]" }

                Write-ColoredOutput "  - $baseName" -Color $script:Colors.Default -NoNewline

                # Use hybrid checking (ID + Name + Content comparison)
                $status = Get-PolicyExistenceStatus -BackupPolicy $remediation `
                    -ApiEndpoint "deviceManagement/deviceHealthScripts" `
                    -RestoredName $restoredName -OriginalName $originalName

                $shouldRestore = Write-RestoreStatus -Status $status -Preview $Preview -RestoreStats $restoreStats -IsImport $IsImport
                if (-not $shouldRestore) {
                    if (-not $Preview -and -not $IsImport -and $status.Status -in @('NameConflict', 'Modified', 'Renamed')) {
                        [void]$conflictPolicies.Add(@{
                            PolicyType = 'RemediationScripts'; PolicyName = $baseName; BackupFile = $scriptFile.FullName
                            Status = $status; ApiEndpoint = 'deviceManagement/deviceHealthScripts'
                            NameField = 'displayName'; RestoredName = $restoredName; OriginalName = $originalName
                        })
                    }
                    continue
                }

                # Re-encode script content to Base64 if needed
                if ($remediation.detectionScriptContent -and -not [string]::IsNullOrWhiteSpace($remediation.detectionScriptContent)) {
                    try {
                        $remediation.detectionScriptContent = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($remediation.detectionScriptContent))
                    }
                    catch {
                        Write-Log "Detection script might already be Base64 encoded" -Level Warning
                    }
                }
                if ($remediation.remediationScriptContent -and -not [string]::IsNullOrWhiteSpace($remediation.remediationScriptContent)) {
                    try {
                        $remediation.remediationScriptContent = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($remediation.remediationScriptContent))
                    }
                    catch {
                        Write-Log "Remediation script might already be Base64 encoded" -Level Warning
                    }
                }

                $remediation.displayName = $newPolicyName
                $remediation.PSObject.Properties.Remove('id')
                $remediation.PSObject.Properties.Remove('createdDateTime')
                $remediation.PSObject.Properties.Remove('lastModifiedDateTime')
                $remediation.PSObject.Properties.Remove('assignments')
                $remediation.PSObject.Properties.Remove('detectionScriptMd5Hash')
                $remediation.PSObject.Properties.Remove('detectionScriptSha256Hash')
                $remediation.PSObject.Properties.Remove('remediationScriptMd5Hash')
                $remediation.PSObject.Properties.Remove('remediationScriptSha256Hash')

                try {
                    $body = $remediation | ConvertTo-Json -Depth 10
                    $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts" `
                        -Method POST -Body $body -ContentType "application/json"

                    Write-ColoredOutput " [CREATED]" -Color $script:Colors.Success
                    $restoreStats.Success++
                }
                catch {
                    Write-ColoredOutput " [FAILED: $_]" -Color $script:Colors.Error
                    $restoreStats.Failed++
                }
            }
        }
    }

    # 8. Restore App Configuration - Managed Devices
    if ($componentsToRestore -contains "AppConfigManagedDevices") {
        $appConfigPath = Join-Path $BackupPath "AppConfigManagedDevices"
        if (Test-Path $appConfigPath) {
            Write-ColoredOutput "`n$actionVerb App Configuration (Managed Devices)..." -Color $script:Colors.Info
            $configs = Get-ChildItem -Path $appConfigPath -Filter "*.json"

            foreach ($configFile in $configs) {
                $config = Get-Content $configFile.FullName | ConvertFrom-Json
                $originalName = $config.displayName

                # Strip existing restore markers (backward compat: old prefix and new suffix)
                if ($originalName -like "[Restored] *") {
                    $baseName = $originalName -replace '^\[Restored\] ', ''
                } elseif ($originalName -match ' - \[(Restored|New Policy)\]$') {
                    $baseName = $originalName -replace ' - \[(Restored|New Policy)\]$', ''
                } else {
                    $baseName = $originalName
                }
                $restoredName = "$baseName - [Restored]"
                $newPolicyName = if ($IsImport) { "$baseName - [New Policy]" } else { "$baseName - [Restored]" }

                Write-ColoredOutput "  - $baseName" -Color $script:Colors.Default -NoNewline

                # Use hybrid checking (ID + Name + Content comparison)
                $status = Get-PolicyExistenceStatus -BackupPolicy $config `
                    -ApiEndpoint "deviceAppManagement/mobileAppConfigurations" `
                    -RestoredName $restoredName -OriginalName $originalName

                $shouldRestore = Write-RestoreStatus -Status $status -Preview $Preview -RestoreStats $restoreStats -IsImport $IsImport
                if (-not $shouldRestore) {
                    if (-not $Preview -and -not $IsImport -and $status.Status -in @('NameConflict', 'Modified', 'Renamed')) {
                        [void]$conflictPolicies.Add(@{
                            PolicyType = 'AppConfigManagedDevices'; PolicyName = $baseName; BackupFile = $configFile.FullName
                            Status = $status; ApiEndpoint = 'deviceAppManagement/mobileAppConfigurations'
                            NameField = 'displayName'; RestoredName = $restoredName; OriginalName = $originalName
                        })
                    }
                    continue
                }

                $config.displayName = $newPolicyName
                $config.PSObject.Properties.Remove('id')
                $config.PSObject.Properties.Remove('createdDateTime')
                $config.PSObject.Properties.Remove('lastModifiedDateTime')
                $config.PSObject.Properties.Remove('version')
                $config.PSObject.Properties.Remove('assignments')

                try {
                    $body = $config | ConvertTo-Json -Depth 10
                    $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations" `
                        -Method POST -Body $body -ContentType "application/json"

                    Write-ColoredOutput " [CREATED]" -Color $script:Colors.Success
                    $restoreStats.Success++
                }
                catch {
                    Write-ColoredOutput " [FAILED: $_]" -Color $script:Colors.Error
                    $restoreStats.Failed++
                }
            }
        }
    }

    # 9. Restore Autopilot Profiles
    if ($componentsToRestore -contains "AutopilotProfiles") {
        $autopilotPath = Join-Path $BackupPath "AutopilotProfiles"
        if (Test-Path $autopilotPath) {
            Write-ColoredOutput "`n$actionVerb Autopilot Profiles..." -Color $script:Colors.Info
            $profiles = Get-ChildItem -Path $autopilotPath -Filter "*.json"

            foreach ($profileFile in $profiles) {
                $profile = Get-Content $profileFile.FullName | ConvertFrom-Json
                $originalName = $profile.displayName

                # Strip existing restore markers (backward compat: old prefix and new suffix)
                if ($originalName -like "[Restored] *") {
                    $baseName = $originalName -replace '^\[Restored\] ', ''
                } elseif ($originalName -match ' - \[(Restored|New Policy)\]$') {
                    $baseName = $originalName -replace ' - \[(Restored|New Policy)\]$', ''
                } else {
                    $baseName = $originalName
                }
                $restoredName = "$baseName - [Restored]"
                $newPolicyName = if ($IsImport) { "$baseName - [New Policy]" } else { "$baseName - [Restored]" }

                Write-ColoredOutput "  - $baseName" -Color $script:Colors.Default -NoNewline

                # Use hybrid checking (ID + Name + Content comparison)
                $status = Get-PolicyExistenceStatus -BackupPolicy $profile `
                    -ApiEndpoint "deviceManagement/windowsAutopilotDeploymentProfiles" `
                    -RestoredName $restoredName -OriginalName $originalName

                $shouldRestore = Write-RestoreStatus -Status $status -Preview $Preview -RestoreStats $restoreStats -IsImport $IsImport
                if (-not $shouldRestore) {
                    if (-not $Preview -and -not $IsImport -and $status.Status -in @('NameConflict', 'Modified', 'Renamed')) {
                        [void]$conflictPolicies.Add(@{
                            PolicyType = 'AutopilotProfiles'; PolicyName = $baseName; BackupFile = $profileFile.FullName
                            Status = $status; ApiEndpoint = 'deviceManagement/windowsAutopilotDeploymentProfiles'
                            NameField = 'displayName'; RestoredName = $restoredName; OriginalName = $originalName
                        })
                    }
                    continue
                }

                $profile.displayName = $newPolicyName
                $profile.PSObject.Properties.Remove('id')
                $profile.PSObject.Properties.Remove('createdDateTime')
                $profile.PSObject.Properties.Remove('lastModifiedDateTime')
                $profile.PSObject.Properties.Remove('assignments')

                try {
                    $body = $profile | ConvertTo-Json -Depth 10
                    $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles" `
                        -Method POST -Body $body -ContentType "application/json"

                    Write-ColoredOutput " [CREATED]" -Color $script:Colors.Success
                    $restoreStats.Success++
                }
                catch {
                    Write-ColoredOutput " [FAILED: $_]" -Color $script:Colors.Error
                    $restoreStats.Failed++
                }
            }
        }
    }

    # 10. Restore Autopilot Device Prep
    if ($componentsToRestore -contains "AutopilotDevicePrep") {
        $devicePrepPath = Join-Path $BackupPath "AutopilotDevicePrep"
        if (Test-Path $devicePrepPath) {
            Write-ColoredOutput "`n$actionVerb Autopilot Device Preparation Policies..." -Color $script:Colors.Info
            $policies = Get-ChildItem -Path $devicePrepPath -Filter "*.json"

            foreach ($policyFile in $policies) {
                $policy = Get-Content $policyFile.FullName | ConvertFrom-Json
                $originalName = $policy.name

                # Strip existing restore markers (backward compat: old prefix and new suffix)
                if ($originalName -like "[Restored] *") {
                    $baseName = $originalName -replace '^\[Restored\] ', ''
                } elseif ($originalName -match ' - \[(Restored|New Policy)\]$') {
                    $baseName = $originalName -replace ' - \[(Restored|New Policy)\]$', ''
                } else {
                    $baseName = $originalName
                }
                $restoredName = "$baseName - [Restored]"
                $newPolicyName = if ($IsImport) { "$baseName - [New Policy]" } else { "$baseName - [Restored]" }

                Write-ColoredOutput "  - $baseName" -Color $script:Colors.Default -NoNewline

                $status = Get-PolicyExistenceStatus -BackupPolicy $policy `
                    -ApiEndpoint "deviceManagement/configurationPolicies" `
                    -RestoredName $restoredName -OriginalName $originalName `
                    -NameField 'name' -ExpandQuery 'settings'

                $shouldRestore = Write-RestoreStatus -Status $status -Preview $Preview -RestoreStats $restoreStats -IsImport $IsImport
                if (-not $shouldRestore) {
                    if (-not $Preview -and -not $IsImport -and $status.Status -in @('NameConflict', 'Modified', 'Renamed')) {
                        [void]$conflictPolicies.Add(@{
                            PolicyType = 'AutopilotDevicePrep'; PolicyName = $baseName; BackupFile = $policyFile.FullName
                            Status = $status; ApiEndpoint = 'deviceManagement/configurationPolicies'
                            NameField = 'name'; RestoredName = $restoredName; OriginalName = $originalName
                            ExpandQuery = 'settings'
                        })
                    }
                    continue
                }

                $newPolicy = @{
                    name                = $newPolicyName
                    description         = $policy.description
                    platforms           = $policy.platforms
                    technologies        = $policy.technologies
                    settings            = $policy.settings
                    templateReference   = $policy.templateReference
                }

                try {
                    $body = $newPolicy | ConvertTo-Json -Depth 10
                    $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" `
                        -Method POST -Body $body -ContentType "application/json"

                    Write-ColoredOutput " [SUCCESS]" -Color $script:Colors.Success
                    $restoreStats.Success++
                }
                catch {
                    Write-ColoredOutput " [FAILED: $_]" -Color $script:Colors.Error
                    $restoreStats.Failed++
                }
            }
        }
    }

    # 11. Restore Enrollment Status Page
    if ($componentsToRestore -contains "EnrollmentStatusPage") {
        $espPath = Join-Path $BackupPath "EnrollmentStatusPage"
        if (Test-Path $espPath) {
            Write-ColoredOutput "`n$actionVerb Enrollment Status Page configurations..." -Color $script:Colors.Info
            $configs = Get-ChildItem -Path $espPath -Filter "*.json"

            foreach ($configFile in $configs) {
                $config = Get-Content $configFile.FullName | ConvertFrom-Json
                $originalName = $config.displayName

                # Strip existing restore markers (backward compat: old prefix and new suffix)
                if ($originalName -like "[Restored] *") {
                    $baseName = $originalName -replace '^\[Restored\] ', ''
                } elseif ($originalName -match ' - \[(Restored|New Policy)\]$') {
                    $baseName = $originalName -replace ' - \[(Restored|New Policy)\]$', ''
                } else {
                    $baseName = $originalName
                }
                $restoredName = "$baseName - [Restored]"
                $newPolicyName = if ($IsImport) { "$baseName - [New Policy]" } else { "$baseName - [Restored]" }

                Write-ColoredOutput "  - $baseName" -Color $script:Colors.Default -NoNewline

                $status = Get-PolicyExistenceStatus -BackupPolicy $config `
                    -ApiEndpoint "deviceManagement/deviceEnrollmentConfigurations" `
                    -RestoredName $restoredName -OriginalName $originalName `
                    -UseV1Api

                $shouldRestore = Write-RestoreStatus -Status $status -Preview $Preview -RestoreStats $restoreStats -IsImport $IsImport
                if (-not $shouldRestore) {
                    if (-not $Preview -and -not $IsImport -and $status.Status -in @('NameConflict', 'Modified', 'Renamed')) {
                        [void]$conflictPolicies.Add(@{
                            PolicyType = 'EnrollmentStatusPage'; PolicyName = $baseName; BackupFile = $configFile.FullName
                            Status = $status; ApiEndpoint = 'deviceManagement/deviceEnrollmentConfigurations'
                            NameField = 'displayName'; RestoredName = $restoredName; OriginalName = $originalName
                            UseV1Api = $true
                        })
                    }
                    continue
                }

                $config.displayName = $newPolicyName
                $config.PSObject.Properties.Remove('id')
                $config.PSObject.Properties.Remove('createdDateTime')
                $config.PSObject.Properties.Remove('lastModifiedDateTime')
                $config.PSObject.Properties.Remove('version')
                $config.PSObject.Properties.Remove('assignments')

                try {
                    $body = $config | ConvertTo-Json -Depth 10
                    $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceEnrollmentConfigurations" `
                        -Method POST -Body $body -ContentType "application/json"

                    Write-ColoredOutput " [SUCCESS]" -Color $script:Colors.Success
                    $restoreStats.Success++
                }
                catch {
                    Write-ColoredOutput " [FAILED: $_]" -Color $script:Colors.Error
                    $restoreStats.Failed++
                }
            }
        }
    }

    # 12. Restore WUfB Policies
    if ($componentsToRestore -contains "WUfBPolicies") {
        $wufbPath = Join-Path $BackupPath "WUfBPolicies"
        if (Test-Path $wufbPath) {
            Write-ColoredOutput "`n$actionVerb Windows Update for Business Policies..." -Color $script:Colors.Info
            $policies = Get-ChildItem -Path $wufbPath -Filter "*.json"

            foreach ($policyFile in $policies) {
                $policy = Get-Content $policyFile.FullName | ConvertFrom-Json
                $originalName = $policy.displayName

                # Strip existing restore markers (backward compat: old prefix and new suffix)
                if ($originalName -like "[Restored] *") {
                    $baseName = $originalName -replace '^\[Restored\] ', ''
                } elseif ($originalName -match ' - \[(Restored|New Policy)\]$') {
                    $baseName = $originalName -replace ' - \[(Restored|New Policy)\]$', ''
                } else {
                    $baseName = $originalName
                }
                $restoredName = "$baseName - [Restored]"
                $newPolicyName = if ($IsImport) { "$baseName - [New Policy]" } else { "$baseName - [Restored]" }

                Write-ColoredOutput "  - $baseName" -Color $script:Colors.Default -NoNewline

                $status = Get-PolicyExistenceStatus -BackupPolicy $policy `
                    -ApiEndpoint "deviceManagement/deviceConfigurations" `
                    -RestoredName $restoredName -OriginalName $originalName

                $shouldRestore = Write-RestoreStatus -Status $status -Preview $Preview -RestoreStats $restoreStats -IsImport $IsImport
                if (-not $shouldRestore) {
                    if (-not $Preview -and -not $IsImport -and $status.Status -in @('NameConflict', 'Modified', 'Renamed')) {
                        [void]$conflictPolicies.Add(@{
                            PolicyType = 'WUfBPolicies'; PolicyName = $baseName; BackupFile = $policyFile.FullName
                            Status = $status; ApiEndpoint = 'deviceManagement/deviceConfigurations'
                            NameField = 'displayName'; RestoredName = $restoredName; OriginalName = $originalName
                        })
                    }
                    continue
                }

                $policy.displayName = $newPolicyName
                $policy.PSObject.Properties.Remove('id')
                $policy.PSObject.Properties.Remove('createdDateTime')
                $policy.PSObject.Properties.Remove('lastModifiedDateTime')
                $policy.PSObject.Properties.Remove('version')
                $policy.PSObject.Properties.Remove('assignments')

                try {
                    $body = $policy | ConvertTo-Json -Depth 10
                    $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations" `
                        -Method POST -Body $body -ContentType "application/json"

                    Write-ColoredOutput " [SUCCESS]" -Color $script:Colors.Success
                    $restoreStats.Success++
                }
                catch {
                    Write-ColoredOutput " [FAILED: $_]" -Color $script:Colors.Error
                    $restoreStats.Failed++
                }
            }
        }
    }

    # 13. Restore Assignment Filters
    if ($componentsToRestore -contains "AssignmentFilters") {
        $filterPath = Join-Path $BackupPath "AssignmentFilters"
        if (Test-Path $filterPath) {
            Write-ColoredOutput "`n$actionVerb Assignment Filters..." -Color $script:Colors.Info
            $filters = Get-ChildItem -Path $filterPath -Filter "*.json"

            # Fetch all existing assignment filters once (API doesn't support $filter parameter well)
            $existingFilters = @(Get-GraphData -Uri "https://graph.microsoft.com/beta/deviceManagement/assignmentFilters")

            foreach ($filterFile in $filters) {
                $filter = Get-Content $filterFile.FullName | ConvertFrom-Json
                $originalName = $filter.displayName

                # Strip existing restore markers (backward compat: old prefix and new suffix)
                if ($originalName -like "[Restored] *") {
                    $baseName = $originalName -replace '^\[Restored\] ', ''
                } elseif ($originalName -match ' - \[(Restored|New Policy)\]$') {
                    $baseName = $originalName -replace ' - \[(Restored|New Policy)\]$', ''
                } else {
                    $baseName = $originalName
                }
                $restoredName = "$baseName - [Restored]"
                $newPolicyName = if ($IsImport) { "$baseName - [New Policy]" } else { "$baseName - [Restored]" }

                Write-ColoredOutput "  - $baseName" -Color $script:Colors.Default -NoNewline

                $status = Get-PolicyExistenceStatus -BackupPolicy $filter `
                    -ApiEndpoint "deviceManagement/assignmentFilters" `
                    -ExistingPolicies $existingFilters `
                    -RestoredName $restoredName -OriginalName $originalName

                $shouldRestore = Write-RestoreStatus -Status $status -Preview $Preview -RestoreStats $restoreStats -IsImport $IsImport
                if (-not $shouldRestore) {
                    if (-not $Preview -and -not $IsImport -and $status.Status -in @('NameConflict', 'Modified', 'Renamed')) {
                        [void]$conflictPolicies.Add(@{
                            PolicyType = 'AssignmentFilters'; PolicyName = $baseName; BackupFile = $filterFile.FullName
                            Status = $status; ApiEndpoint = 'deviceManagement/assignmentFilters'
                            NameField = 'displayName'; RestoredName = $restoredName; OriginalName = $originalName
                        })
                    }
                    continue
                }

                $filter.displayName = $newPolicyName
                $filter.PSObject.Properties.Remove('id')
                $filter.PSObject.Properties.Remove('createdDateTime')
                $filter.PSObject.Properties.Remove('lastModifiedDateTime')

                try {
                    $body = $filter | ConvertTo-Json -Depth 10
                    $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/assignmentFilters" `
                        -Method POST -Body $body -ContentType "application/json"

                    Write-ColoredOutput " [SUCCESS]" -Color $script:Colors.Success
                    $restoreStats.Success++
                }
                catch {
                    Write-ColoredOutput " [FAILED: $_]" -Color $script:Colors.Error
                    $restoreStats.Failed++
                }
            }
        }
    }

    # 14. Restore App Configuration - Managed Apps
    if ($componentsToRestore -contains "AppConfigManagedApps") {
        $appConfigPath = Join-Path $BackupPath "AppConfigManagedApps"
        if (Test-Path $appConfigPath) {
            Write-ColoredOutput "`n$actionVerb App Configuration (Managed Apps)..." -Color $script:Colors.Info
            $configs = Get-ChildItem -Path $appConfigPath -Filter "*.json"

            foreach ($configFile in $configs) {
                $config = Get-Content $configFile.FullName | ConvertFrom-Json
                $originalName = $config.displayName

                # Strip existing restore markers (backward compat: old prefix and new suffix)
                if ($originalName -like "[Restored] *") {
                    $baseName = $originalName -replace '^\[Restored\] ', ''
                } elseif ($originalName -match ' - \[(Restored|New Policy)\]$') {
                    $baseName = $originalName -replace ' - \[(Restored|New Policy)\]$', ''
                } else {
                    $baseName = $originalName
                }
                $restoredName = "$baseName - [Restored]"
                $newPolicyName = if ($IsImport) { "$baseName - [New Policy]" } else { "$baseName - [Restored]" }

                Write-ColoredOutput "  - $baseName" -Color $script:Colors.Default -NoNewline

                $status = Get-PolicyExistenceStatus -BackupPolicy $config `
                    -ApiEndpoint "deviceAppManagement/targetedManagedAppConfigurations" `
                    -RestoredName $restoredName -OriginalName $originalName

                $shouldRestore = Write-RestoreStatus -Status $status -Preview $Preview -RestoreStats $restoreStats -IsImport $IsImport
                if (-not $shouldRestore) {
                    if (-not $Preview -and -not $IsImport -and $status.Status -in @('NameConflict', 'Modified', 'Renamed')) {
                        [void]$conflictPolicies.Add(@{
                            PolicyType = 'AppConfigManagedApps'; PolicyName = $baseName; BackupFile = $configFile.FullName
                            Status = $status; ApiEndpoint = 'deviceAppManagement/targetedManagedAppConfigurations'
                            NameField = 'displayName'; RestoredName = $restoredName; OriginalName = $originalName
                        })
                    }
                    continue
                }

                $config.displayName = $newPolicyName
                $config.PSObject.Properties.Remove('id')
                $config.PSObject.Properties.Remove('createdDateTime')
                $config.PSObject.Properties.Remove('lastModifiedDateTime')
                $config.PSObject.Properties.Remove('version')
                $config.PSObject.Properties.Remove('assignments')

                try {
                    $body = $config | ConvertTo-Json -Depth 10
                    # Correct endpoint for Managed Apps app configuration (MAM)
                    $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/targetedManagedAppConfigurations" `
                        -Method POST -Body $body -ContentType "application/json"

                    Write-ColoredOutput " [SUCCESS]" -Color $script:Colors.Success
                    $restoreStats.Success++
                }
                catch {
                    Write-ColoredOutput " [FAILED: $_]" -Color $script:Colors.Error
                    $restoreStats.Failed++
                }
            }
        }
    }

    # 15. Restore macOS Shell Scripts
    if ($componentsToRestore -contains "MacOSScripts") {
        $scriptsPath = Join-Path $BackupPath "MacOSScripts"
        if (Test-Path $scriptsPath) {
            Write-ColoredOutput "`n$actionVerb macOS Shell Scripts..." -Color $script:Colors.Info
            $scripts = Get-ChildItem -Path $scriptsPath -Filter "*.json"

            foreach ($scriptFile in $scripts) {
                $scriptData = Get-Content $scriptFile.FullName | ConvertFrom-Json
                $originalName = $scriptData.displayName

                # Strip existing restore markers (backward compat: old prefix and new suffix)
                if ($originalName -like "[Restored] *") {
                    $baseName = $originalName -replace '^\[Restored\] ', ''
                } elseif ($originalName -match ' - \[(Restored|New Policy)\]$') {
                    $baseName = $originalName -replace ' - \[(Restored|New Policy)\]$', ''
                } else {
                    $baseName = $originalName
                }
                $restoredName = "$baseName - [Restored]"
                $newPolicyName = if ($IsImport) { "$baseName - [New Policy]" } else { "$baseName - [Restored]" }

                Write-ColoredOutput "  - $baseName" -Color $script:Colors.Default -NoNewline

                $status = Get-PolicyExistenceStatus -BackupPolicy $scriptData `
                    -ApiEndpoint "deviceManagement/deviceShellScripts" `
                    -RestoredName $restoredName -OriginalName $originalName

                $shouldRestore = Write-RestoreStatus -Status $status -Preview $Preview -RestoreStats $restoreStats -IsImport $IsImport
                if (-not $shouldRestore) {
                    if (-not $Preview -and -not $IsImport -and $status.Status -in @('NameConflict', 'Modified', 'Renamed')) {
                        [void]$conflictPolicies.Add(@{
                            PolicyType = 'MacOSScripts'; PolicyName = $baseName; BackupFile = $scriptFile.FullName
                            Status = $status; ApiEndpoint = 'deviceManagement/deviceShellScripts'
                            NameField = 'displayName'; RestoredName = $restoredName; OriginalName = $originalName
                        })
                    }
                    continue
                }

                $scriptData.displayName = $newPolicyName
                $scriptData.PSObject.Properties.Remove('id')
                $scriptData.PSObject.Properties.Remove('createdDateTime')
                $scriptData.PSObject.Properties.Remove('lastModifiedDateTime')
                $scriptData.PSObject.Properties.Remove('assignments')
                $scriptData.PSObject.Properties.Remove('scriptContentMd5Hash')
                $scriptData.PSObject.Properties.Remove('scriptContentSha256Hash')

                try {
                    $body = $scriptData | ConvertTo-Json -Depth 100
                    $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceShellScripts" `
                        -Method POST -Body $body -ContentType "application/json"

                    Write-ColoredOutput " [SUCCESS]" -Color $script:Colors.Success
                    $restoreStats.Success++
                }
                catch {
                    Write-ColoredOutput " [FAILED: $_]" -Color $script:Colors.Error
                    $restoreStats.Failed++
                }
            }
        }
    }

    # 17. Restore macOS Custom Attributes
    if ($componentsToRestore -contains "MacOSCustomAttributes") {
        $attrPath = Join-Path $BackupPath "MacOSCustomAttributes"
        if (Test-Path $attrPath) {
            Write-ColoredOutput "`n$actionVerb macOS Custom Attributes..." -Color $script:Colors.Info
            $attrs = Get-ChildItem -Path $attrPath -Filter "*.json"

            foreach ($attrFile in $attrs) {
                $attr = Get-Content $attrFile.FullName | ConvertFrom-Json
                $originalName = $attr.displayName

                # Strip existing restore markers (backward compat: old prefix and new suffix)
                if ($originalName -like "[Restored] *") {
                    $baseName = $originalName -replace '^\[Restored\] ', ''
                } elseif ($originalName -match ' - \[(Restored|New Policy)\]$') {
                    $baseName = $originalName -replace ' - \[(Restored|New Policy)\]$', ''
                } else {
                    $baseName = $originalName
                }
                $restoredName = "$baseName - [Restored]"
                $newPolicyName = if ($IsImport) { "$baseName - [New Policy]" } else { "$baseName - [Restored]" }

                Write-ColoredOutput "  - $baseName" -Color $script:Colors.Default -NoNewline

                $status = Get-PolicyExistenceStatus -BackupPolicy $attr `
                    -ApiEndpoint "deviceManagement/deviceCustomAttributeShellScripts" `
                    -RestoredName $restoredName -OriginalName $originalName

                $shouldRestore = Write-RestoreStatus -Status $status -Preview $Preview -RestoreStats $restoreStats -IsImport $IsImport
                if (-not $shouldRestore) {
                    if (-not $Preview -and -not $IsImport -and $status.Status -in @('NameConflict', 'Modified', 'Renamed')) {
                        [void]$conflictPolicies.Add(@{
                            PolicyType = 'MacOSCustomAttributes'; PolicyName = $baseName; BackupFile = $attrFile.FullName
                            Status = $status; ApiEndpoint = 'deviceManagement/deviceCustomAttributeShellScripts'
                            NameField = 'displayName'; RestoredName = $restoredName; OriginalName = $originalName
                        })
                    }
                    continue
                }

                $attr.displayName = $newPolicyName
                $attr.PSObject.Properties.Remove('id')
                $attr.PSObject.Properties.Remove('createdDateTime')
                $attr.PSObject.Properties.Remove('lastModifiedDateTime')
                $attr.PSObject.Properties.Remove('assignments')
                $attr.PSObject.Properties.Remove('scriptContentMd5Hash')
                $attr.PSObject.Properties.Remove('scriptContentSha256Hash')

                try {
                    $body = $attr | ConvertTo-Json -Depth 100
                    $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCustomAttributeShellScripts" `
                        -Method POST -Body $body -ContentType "application/json"

                    Write-ColoredOutput " [SUCCESS]" -Color $script:Colors.Success
                    $restoreStats.Success++
                }
                catch {
                    Write-ColoredOutput " [FAILED: $_]" -Color $script:Colors.Error
                    $restoreStats.Failed++
                }
            }
        }
    }

    # 18. Restore Legacy Intune Intents (deviceManagementIntent format - old Security Baselines API)
    if ($componentsToRestore -contains "LegacyIntents") {
        $intentsPath = Join-Path $BackupPath "LegacyIntents"
        if (Test-Path $intentsPath) {
            Write-ColoredOutput "`n$actionVerb Legacy Intune Intents (Security Baselines)..." -Color $script:Colors.Info
            $intentFiles = Get-ChildItem -Path $intentsPath -Filter "*.json"

            foreach ($intentFile in $intentFiles) {
                $intent = Get-Content $intentFile.FullName | ConvertFrom-Json
                $originalName = if ($intent.displayName) { $intent.displayName } else { $intent.name }

                if ($originalName -like "[Restored] *") {
                    $baseName = $originalName -replace '^\[Restored\] ', ''
                } elseif ($originalName -match ' - \[(Restored|New Policy)\]$') {
                    $baseName = $originalName -replace ' - \[(Restored|New Policy)\]$', ''
                } else {
                    $baseName = $originalName
                }
                $newPolicyName = if ($IsImport) { "$baseName - [New Policy]" } else { "$baseName - [Restored]" }

                Write-ColoredOutput "  - $baseName" -Color $script:Colors.Default -NoNewline

                # Clean settings: strip id and OData navigation properties (keep @odata.type)
                $cleanSettings = @()
                if ($intent.settings) {
                    foreach ($setting in @($intent.settings)) {
                        $settingJson = $setting | ConvertTo-Json -Depth 50 -Compress | ConvertFrom-Json
                        $propsToRemove = @($settingJson.PSObject.Properties.Name | Where-Object {
                            $_ -eq 'id' -or ($_ -like '*@odata*' -and $_ -ne '@odata.type')
                        })
                        foreach ($prop in $propsToRemove) {
                            $settingJson.PSObject.Properties.Remove($prop)
                        }
                        $cleanSettings += $settingJson
                    }
                }

                $newIntent = @{
                    displayName = $newPolicyName
                    description = if ($intent.description) { $intent.description } else { "" }
                    templateId  = $intent.templateId
                    settings    = $cleanSettings
                }

                try {
                    $body = $newIntent | ConvertTo-Json -Depth 50
                    $null = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/intents" `
                        -Method POST -Body $body -ContentType "application/json"

                    Write-ColoredOutput " [CREATED]" -Color $script:Colors.Success
                    $restoreStats.Success++
                }
                catch {
                    Write-ColoredOutput " [FAILED: $_]" -Color $script:Colors.Error
                    $restoreStats.Failed++
                }
            }
        }
    }

    # ── Conflict Resolution Phase ──
    if (-not $Preview -and $conflictPolicies.Count -gt 0) {
        Write-Host ""
        Write-Host "  ════════════════════════════════════════════════════════════════════" -ForegroundColor Yellow
        Write-Host "              POLICIES WITH CHANGES DETECTED                        " -ForegroundColor Yellow
        Write-Host "  ════════════════════════════════════════════════════════════════════" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  The following $($conflictPolicies.Count) policies have changes compared to the backup:" -ForegroundColor White
        Write-Host ""

        for ($i = 0; $i -lt $conflictPolicies.Count; $i++) {
            $cp = $conflictPolicies[$i]
            $st = $cp.Status
            $num = $i + 1

            # Build status description
            $statusParts = @()
            if ($st.IdChanged) { $statusParts += "Policy ID Changed" }
            if ($st.NameChanged) { $statusParts += "Name Changed to '$($st.TenantName)'" }
            if ($st.SettingsChanged) { $statusParts += "Settings Modified ($($st.DifferenceCount) changes)" }
            if (-not $st.SettingsChanged -and -not $st.NameChanged) { $statusParts += "Settings Unchanged" }
            $statusDesc = $statusParts -join ' + '

            $typeColor = switch ($st.Status) {
                'NameConflict' { if ($st.SettingsChanged) { 'Red' } else { 'Yellow' } }
                'Modified' { 'Yellow' }
                'Renamed' { 'Yellow' }
                default { 'White' }
            }

            Write-Host "  [$num] " -ForegroundColor Cyan -NoNewline
            Write-Host "$($cp.PolicyType) / " -ForegroundColor Gray -NoNewline
            Write-Host "$($cp.PolicyName)" -ForegroundColor White
            Write-Host "      Status: $statusDesc" -ForegroundColor $typeColor

            # Show changed properties if settings changed
            if ($st.SettingsChanged -and $st.Differences.Count -gt 0) {
                $diffList = ($st.Differences | Select-Object -First 5) -join ', '
                if ($st.Differences.Count -gt 5) { $diffList += ", ... +$($st.Differences.Count - 5) more" }
                Write-Host "      Changed: $diffList" -ForegroundColor DarkGray
            }
        }

        Write-Host ""
        Write-Host "  ────────────────────────────────────────────────────────────────────" -ForegroundColor Yellow
        Write-Host "  Options:" -ForegroundColor White
        Write-Host "    [A] Restore ALL listed policies from backup (creates '- [Restored]' copies)" -ForegroundColor Green
        Write-Host "    [S] Select individual policies to restore" -ForegroundColor Cyan
        Write-Host "    [D] View detailed difference reports for dirty policies" -ForegroundColor White
        Write-Host "    [N] Skip all - do not restore any (make a new backup later)" -ForegroundColor Yellow
        Write-Host ""
        $conflictChoice = Read-Host "  Select option (A/S/D/N)"

        $policiesToRestore = @()

        switch ($conflictChoice.ToUpper()) {
            'A' {
                $policiesToRestore = $conflictPolicies
                Write-Host ""
                Write-ColoredOutput "  $actionVerb all $($conflictPolicies.Count) conflicting policies from backup..." -Color $script:Colors.Info
            }
            'S' {
                Write-Host ""
                foreach ($cp in $conflictPolicies) {
                    $choice = Read-Host "  Restore '$($cp.PolicyName)' ($($cp.PolicyType))? (Y/N)"
                    if ($choice.ToUpper() -eq 'Y') {
                        $policiesToRestore += $cp
                    }
                }
                if ($policiesToRestore.Count -gt 0) {
                    Write-Host ""
                    Write-ColoredOutput "  $actionVerb $($policiesToRestore.Count) selected policies from backup..." -Color $script:Colors.Info
                } else {
                    Write-ColoredOutput "  No policies selected for restore." -Color $script:Colors.Warning
                }
            }
            'D' {
                # Show detailed difference reports for all dirty policies
                Write-Host ""
                foreach ($cp in $conflictPolicies) {
                    $st = $cp.Status
                    if ($st.SettingsChanged -and $st.DifferenceCount -gt 0) {
                        Write-DifferenceReport -PolicyName $cp.PolicyName -PolicyType $cp.PolicyType -Status $st
                    }
                }

                # After viewing reports, ask again
                Write-Host ""
                Write-Host "  ────────────────────────────────────────────────────────────────────" -ForegroundColor Yellow
                Write-Host "  Options:" -ForegroundColor White
                Write-Host "    [A] Restore ALL listed policies from backup (creates '- [Restored]' copies)" -ForegroundColor Green
                Write-Host "    [S] Select individual policies to restore" -ForegroundColor Cyan
                Write-Host "    [N] Skip all - do not restore any" -ForegroundColor Yellow
                Write-Host ""
                $followUpChoice = Read-Host "  Select option (A/S/N)"

                switch ($followUpChoice.ToUpper()) {
                    'A' {
                        $policiesToRestore = $conflictPolicies
                        Write-Host ""
                        Write-ColoredOutput "  $actionVerb all $($conflictPolicies.Count) conflicting policies from backup..." -Color $script:Colors.Info
                    }
                    'S' {
                        Write-Host ""
                        foreach ($cp2 in $conflictPolicies) {
                            $choice2 = Read-Host "  Restore '$($cp2.PolicyName)' ($($cp2.PolicyType))? (Y/N)"
                            if ($choice2.ToUpper() -eq 'Y') {
                                $policiesToRestore += $cp2
                            }
                        }
                        if ($policiesToRestore.Count -gt 0) {
                            Write-Host ""
                            Write-ColoredOutput "  $actionVerb $($policiesToRestore.Count) selected policies from backup..." -Color $script:Colors.Info
                        } else {
                            Write-ColoredOutput "  No policies selected for restore." -Color $script:Colors.Warning
                        }
                    }
                    default {
                        Write-ColoredOutput "  Skipping all conflicting policies." -Color $script:Colors.Warning
                    }
                }
            }
            default {
                Write-ColoredOutput "  Skipping all conflicting policies." -Color $script:Colors.Warning
            }
        }

        # Process selected conflict resolutions
        foreach ($cp in $policiesToRestore) {
            Write-ColoredOutput "  - $($cp.PolicyName)" -Color $script:Colors.Default -NoNewline

            $restoreResult = Restore-ConflictPolicy -ConflictInfo $cp

            if ($restoreResult.Success) {
                Write-ColoredOutput " [CREATED]" -Color $script:Colors.Success
                $restoreStats.ConflictResolved++
                $restoreStats.Success++
                $restoreStats.Skipped--
            } else {
                Write-ColoredOutput " [FAILED: $($restoreResult.Error)]" -Color $script:Colors.Error
                $restoreStats.Failed++
            }
        }
    }

    # Display summary
    Write-Host ""
    Write-Host "  ════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    $completedTitle = if ($IsImport) {
        if ($Preview) { "                  IMPORT PREVIEW COMPLETED                          " }
        else          { "                  IMPORT COMPLETED                                  " }
    } else {
        if ($Preview) { "                  RESTORE PREVIEW COMPLETED                         " }
        else          { "                  RESTORE COMPLETED                                 " }
    }
    Write-Host $completedTitle -ForegroundColor Cyan
    Write-Host "  ════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Unchanged:          " -ForegroundColor Green -NoNewline
    Write-Host "$($restoreStats.Unchanged)" -ForegroundColor White
    Write-Host "  Changed:            " -ForegroundColor Yellow -NoNewline
    Write-Host "$($restoreStats.Modified)" -ForegroundColor White
    Write-Host "  Renamed:            " -ForegroundColor Cyan -NoNewline
    Write-Host "$($restoreStats.Renamed)" -ForegroundColor White
    $restoredLabel = if ($IsImport) { "  Imported:           " } else { "  Restored:           " }
    Write-Host $restoredLabel -ForegroundColor Green -NoNewline
    Write-Host "$($restoreStats.Deleted)" -ForegroundColor White
    if ($restoreStats.NameConflict -gt 0) {
        $ncLabel = if ($IsImport) { "  New:                " } else { "  Prev. Restored:     " }
        Write-Host $ncLabel -ForegroundColor Yellow -NoNewline
        Write-Host "$($restoreStats.NameConflict)" -ForegroundColor White
    }
    if ($restoreStats.ConflictResolved -gt 0) {
        Write-Host "  Conflicts Resolved: " -ForegroundColor Green -NoNewline
        Write-Host "$($restoreStats.ConflictResolved)" -ForegroundColor White
    }
    if ($restoreStats.Failed -gt 0) {
        Write-Host "  Failed:             " -ForegroundColor Red -NoNewline
        Write-Host "$($restoreStats.Failed)" -ForegroundColor White
    }
    Write-Host ""
}

#endregion

#region Comparison Functions

function Remove-MetadataForComparison {
    param (
        [PSObject]$Policy
    )

    # Create a deep copy of the policy
    $cleanPolicy = $Policy | ConvertTo-Json -Depth 100 | ConvertFrom-Json

    # Remove metadata fields that change between backups or are not relevant for comparison
    $metadataFields = @(
        # Standard metadata
        'id',
        'createdDateTime',
        'lastModifiedDateTime',
        'version',
        'modifiedDateTime',

        # OData metadata
        '@odata.context',
        '@odata.type',
        '@odata.id',
        '@odata.etag',

        # Scope and role tags
        'roleScopeTagIds',
        'supportsScopeTags',
        'deviceManagementApplicabilityRuleOsEdition',
        'deviceManagementApplicabilityRuleOsVersion',
        'deviceManagementApplicabilityRuleDeviceMode',

        # Creation/modification metadata
        'creationSource',
        'isAssigned',
        'priority',
        'createdBy',
        'lastModifiedBy',

        # Template references (can include GUIDs)
        'templateReference',
        'templateId',
        'templateDisplayName',
        'templateDisplayVersion',
        'templateFamily',

        # Settings Catalog specific GUIDs
        'settingDefinitionId',
        'settingInstanceTemplateId',
        'settingValueTemplateId',
        'settingValueTemplateReference',

        # Additional GUID-based fields
        'secretReferenceValueId',
        'deviceConfigurationId',
        'groupId',
        'sourceId',
        'payloadId',

        # Autopilot and enrollment specific
        'deviceNameTemplate',
        'azureAdJoinType',

        # App protection specific
        'targetedAppManagementLevels',
        'appGroupType',

        # Script specific
        'scriptContentMd5Hash',
        'scriptContentSha256Hash'
    )

    foreach ($field in $metadataFields) {
        if ($cleanPolicy.PSObject.Properties[$field]) {
            $cleanPolicy.PSObject.Properties.Remove($field)
        }
    }

    # Also clean nested objects if they exist
    if ($cleanPolicy.assignments) {
        $cleanPolicy.PSObject.Properties.Remove('assignments')
    }

    # Clean settings array from Settings Catalog policies (remove GUIDs from nested settings)
    if ($cleanPolicy.settings) {
        $cleanPolicy.settings = Clean-SettingsForComparison -Settings $cleanPolicy.settings
    }

    # Remove scheduledActionsForRule IDs but keep the configuration values
    if ($cleanPolicy.scheduledActionsForRule) {
        foreach ($rule in $cleanPolicy.scheduledActionsForRule) {
            if ($rule.PSObject.Properties['id']) {
                $rule.PSObject.Properties.Remove('id')
            }
            if ($rule.scheduledActionConfigurations) {
                foreach ($config in $rule.scheduledActionConfigurations) {
                    if ($config.PSObject.Properties['id']) {
                        $config.PSObject.Properties.Remove('id')
                    }
                }
            }
        }
    }

    # Sort properties to ensure consistent ordering
    $sortedPolicy = [PSCustomObject]@{}
    $cleanPolicy.PSObject.Properties | Sort-Object Name | ForEach-Object {
        $sortedPolicy | Add-Member -NotePropertyName $_.Name -NotePropertyValue $_.Value -Force
    }

    return $sortedPolicy
}

function Clean-SettingsForComparison {
    param (
        [object]$Settings
    )

    if ($null -eq $Settings) { return $null }

    # Fields to remove from settings objects
    $settingsMetadataFields = @(
        'settingDefinitionId',
        'settingInstanceTemplateId',
        'settingValueTemplateId',
        'settingValueTemplateReference',
        'id'
    )

    if ($Settings -is [array]) {
        $cleanedSettings = @()
        foreach ($setting in $Settings) {
            $cleanedSettings += Clean-SettingsForComparison -Settings $setting
        }
        return $cleanedSettings
    }
    elseif ($Settings -is [PSCustomObject]) {
        $cleanSetting = $Settings | ConvertTo-Json -Depth 50 | ConvertFrom-Json

        foreach ($field in $settingsMetadataFields) {
            if ($cleanSetting.PSObject.Properties[$field]) {
                $cleanSetting.PSObject.Properties.Remove($field)
            }
        }

        # Recursively clean nested settingInstance
        if ($cleanSetting.settingInstance) {
            $cleanSetting.settingInstance = Clean-SettingsForComparison -Settings $cleanSetting.settingInstance
        }

        # Recursively clean children
        if ($cleanSetting.children) {
            $cleanSetting.children = Clean-SettingsForComparison -Settings $cleanSetting.children
        }

        # Recursively clean choiceSettingValue
        if ($cleanSetting.choiceSettingValue) {
            if ($cleanSetting.choiceSettingValue.children) {
                $cleanSetting.choiceSettingValue.children = Clean-SettingsForComparison -Settings $cleanSetting.choiceSettingValue.children
            }
        }

        # Recursively clean groupSettingCollectionValue
        if ($cleanSetting.groupSettingCollectionValue) {
            $cleanSetting.groupSettingCollectionValue = Clean-SettingsForComparison -Settings $cleanSetting.groupSettingCollectionValue
        }

        return $cleanSetting
    }
    else {
        return $Settings
    }
}

function Compare-DeepObject {
    <#
    .SYNOPSIS
    Performs deep comparison of two objects, handling various edge cases.

    .DESCRIPTION
    Compares objects recursively, handling:
    - Null/empty equivalence (null, empty string, empty array are treated equally)
    - Type differences (hashtable vs PSCustomObject)
    - Array comparison (by content, not order for simple values)
    - Boolean and string equivalence
    #>
    param (
        [Parameter(Mandatory = $false)]
        $Object1,

        [Parameter(Mandatory = $false)]
        $Object2
    )

    # Helper to check if value is "empty" (null, empty string, empty array)
    $IsEmpty = {
        param($obj)
        if ($null -eq $obj) { return $true }
        if ($obj -is [string] -and [string]::IsNullOrWhiteSpace($obj)) { return $true }
        if ($obj -is [array] -and $obj.Count -eq 0) { return $true }
        return $false
    }

    # If both are "empty", they're equal
    if ((&$IsEmpty $Object1) -and (&$IsEmpty $Object2)) { return $true }

    # If only one is empty, they're different
    if ((&$IsEmpty $Object1) -or (&$IsEmpty $Object2)) { return $false }

    # Handle arrays (including System.Object[])
    if ($Object1 -is [System.Collections.IList] -and $Object2 -is [System.Collections.IList]) {
        if ($Object1.Count -ne $Object2.Count) { return $false }

        # For arrays of simple values, compare sorted to handle order differences
        $allSimple1 = $true
        $allSimple2 = $true
        foreach ($item in $Object1) {
            if ($item -is [PSCustomObject] -or $item -is [hashtable] -or $item -is [System.Collections.IList]) {
                $allSimple1 = $false
                break
            }
        }
        foreach ($item in $Object2) {
            if ($item -is [PSCustomObject] -or $item -is [hashtable] -or $item -is [System.Collections.IList]) {
                $allSimple2 = $false
                break
            }
        }

        if ($allSimple1 -and $allSimple2) {
            # Compare simple arrays by sorted content
            $sorted1 = $Object1 | Sort-Object
            $sorted2 = $Object2 | Sort-Object
            for ($i = 0; $i -lt $sorted1.Count; $i++) {
                if ($sorted1[$i] -ne $sorted2[$i]) { return $false }
            }
            return $true
        }

        # For complex arrays, compare by index (order matters)
        for ($i = 0; $i -lt $Object1.Count; $i++) {
            if (-not (Compare-DeepObject -Object1 $Object1[$i] -Object2 $Object2[$i])) {
                return $false
            }
        }
        return $true
    }

    # Handle hashtables and PSCustomObjects (treat them equivalently)
    if (($Object1 -is [hashtable] -or $Object1 -is [PSCustomObject]) -and
        ($Object2 -is [hashtable] -or $Object2 -is [PSCustomObject])) {

        $props1 = if ($Object1 -is [hashtable]) { $Object1.Keys } else { $Object1.PSObject.Properties.Name }
        $props2 = if ($Object2 -is [hashtable]) { $Object2.Keys } else { $Object2.PSObject.Properties.Name }

        # Filter out empty properties from comparison
        $props1Array = @($props1 | Where-Object {
            $val = if ($Object1 -is [hashtable]) { $Object1[$_] } else { $Object1.$_ }
            -not (&$IsEmpty $val)
        })
        $props2Array = @($props2 | Where-Object {
            $val = if ($Object2 -is [hashtable]) { $Object2[$_] } else { $Object2.$_ }
            -not (&$IsEmpty $val)
        })

        # Check if non-empty property count is same
        if ($props1Array.Count -ne $props2Array.Count) { return $false }

        # Check each non-empty property
        foreach ($prop in $props1Array) {
            if ($props2Array -notcontains $prop) { return $false }

            $val1 = if ($Object1 -is [hashtable]) { $Object1[$prop] } else { $Object1.$prop }
            $val2 = if ($Object2 -is [hashtable]) { $Object2[$prop] } else { $Object2.$prop }

            if (-not (Compare-DeepObject -Object1 $val1 -Object2 $val2)) {
                return $false
            }
        }
        return $true
    }

    # Handle boolean comparison (convert to same type)
    if ($Object1 -is [bool] -or $Object2 -is [bool]) {
        try {
            return [bool]$Object1 -eq [bool]$Object2
        }
        catch {
            return $false
        }
    }

    # Handle numeric comparison (int vs long vs double)
    if (($Object1 -is [int] -or $Object1 -is [long] -or $Object1 -is [double] -or $Object1 -is [decimal]) -and
        ($Object2 -is [int] -or $Object2 -is [long] -or $Object2 -is [double] -or $Object2 -is [decimal])) {
        return [double]$Object1 -eq [double]$Object2
    }

    # For strings and other primitives, use standard comparison
    return "$Object1" -eq "$Object2"
}

function Compare-Backups {
    param (
        [string]$Backup1Path,
        [string]$Backup2Path
    )

    Write-ColoredOutput "`nComparing backups..." -Color $script:Colors.Header

    # Load metadata
    $metadata1 = Get-Content (Join-Path $Backup1Path "metadata.json") | ConvertFrom-Json
    $metadata2 = Get-Content (Join-Path $Backup2Path "metadata.json") | ConvertFrom-Json

    $itemsText1 = "$($metadata1.TotalItems) items"
    $itemsText2 = "$($metadata2.TotalItems) items"
    $backup1Label = $metadata1.BackupDate
    $backup2Label = $metadata2.BackupDate
    Write-ColoredOutput "Backup 1: $backup1Label ($itemsText1)" -Color $script:Colors.Info
    Write-ColoredOutput "Backup 2: $backup2Label ($itemsText2)" -Color $script:Colors.Info

    $differences = @{
        OnlyInBackup1 = @()
        OnlyInBackup2 = @()
        Modified      = @()
        Unchanged     = 0
    }

    # Compare each policy type
    $policyTypes = @("DeviceConfigurations", "CompliancePolicies", "SettingsCatalogPolicies",
                     "AppProtectionPolicies", "PowerShellScripts", "AdministrativeTemplates",
                     "AutopilotProfiles", "AutopilotDevicePrep", "EnrollmentStatusPage",
                     "RemediationScripts", "WUfBPolicies", "AssignmentFilters",
                     "AppConfigManagedDevices", "AppConfigManagedApps")

    foreach ($policyType in $policyTypes) {
        $path1 = Join-Path $Backup1Path $policyType
        $path2 = Join-Path $Backup2Path $policyType

        if ((Test-Path $path1) -or (Test-Path $path2)) {
            Write-ColoredOutput "`nComparing $policyType..." -Color $script:Colors.Info

            $files1 = if (Test-Path $path1) { Get-ChildItem -Path $path1 -Filter "*.json" } else { @() }
            $files2 = if (Test-Path $path2) { Get-ChildItem -Path $path2 -Filter "*.json" } else { @() }

            # Load all policies from both backups
            $policies1 = @{}
            foreach ($file1 in $files1) {
                $p = Get-Content $file1.FullName | ConvertFrom-Json
                $policies1[$file1.Name] = $p
            }
            $policies2 = @{}
            foreach ($file2 in $files2) {
                $p = Get-Content $file2.FullName | ConvertFrom-Json
                $policies2[$file2.Name] = $p
            }

            # Track which policies from backup 1 have been matched
            $matched1Keys = @()

            # Check each policy in backup 2 against backup 1
            foreach ($key2 in $policies2.Keys) {
                $policy2 = $policies2[$key2]
                $policy2Name = if ($policy2.displayName) { $policy2.displayName } else { $policy2.name }
                $matched = $false
                $matchedKey1 = $null

                # Step 1: Try to match by ID first (handles renames)
                if ($policy2.id) {
                    foreach ($key1 in $policies1.Keys) {
                        $policy1 = $policies1[$key1]
                        if ($policy1.id -and $policy1.id -eq $policy2.id) {
                            $matched = $true
                            $matchedKey1 = $key1
                            break
                        }
                    }
                }

                # Step 2: Fallback to name match
                if (-not $matched) {
                    foreach ($key1 in $policies1.Keys) {
                        $policy1 = $policies1[$key1]
                        $displayNameMatch = $policy1.displayName -and $policy2.displayName -and ($policy1.displayName -eq $policy2.displayName)
                        $nameMatch = $policy1.name -and $policy2.name -and ($policy1.name -eq $policy2.name)
                        if ($displayNameMatch -or $nameMatch) {
                            $matched = $true
                            $matchedKey1 = $key1
                            break
                        }
                    }
                }

                if ($matched) {
                    $matched1Keys += $matchedKey1
                    $policy1 = $policies1[$matchedKey1]
                    $policy1Name = if ($policy1.displayName) { $policy1.displayName } else { $policy1.name }

                    # Check if name changed between backups
                    $nameChanged = ($policy1Name -ne $policy2Name)

                    # Check if settings changed using Compare-PolicyContent
                    $comparison = Compare-PolicyContent -BackupPolicy $policy1 -TenantPolicy $policy2
                    $settingsDiffs = @($comparison.Differences | Where-Object { $_ -ne 'name' -and $_ -ne 'displayName' })

                    if ($settingsDiffs.Count -gt 0) {
                        # Settings changed - [CHANGED]
                        $settingWord = if ($settingsDiffs.Count -eq 1) { 'Setting' } else { 'Settings' }
                        $changeMsg = "  - $policy2Name [CHANGED] - [$($settingsDiffs.Count) $settingWord Differ]"
                        if ($nameChanged) { $changeMsg += " [RENAMED] -> [$policy2Name]" }
                        $differences.Modified += [PSCustomObject]@{
                            PolicyType   = $policyType
                            PolicyName   = $policy2Name
                            Created      = ConvertFrom-JsonDate -DateString $policy2.createdDateTime
                            LastModified = ConvertFrom-JsonDate -DateString $policy2.lastModifiedDateTime
                            IsAssigned   = if ($null -ne $policy2.isAssigned) { if ($policy2.isAssigned) { "Yes" } else { "No" } } else { "N/A" }
                        }
                        Write-ColoredOutput $changeMsg -Color $script:Colors.Warning
                    }
                    elseif ($nameChanged) {
                        # Only name changed - [RENAMED]
                        Write-ColoredOutput "  - $policy1Name [RENAMED] -> [$policy2Name]" -Color $script:Colors.Info
                        $differences.Unchanged++
                    }
                    else {
                        # Identical
                        $differences.Unchanged++
                        if ($VerbosePreference -eq 'Continue') {
                            Write-Verbose "  - $policy2Name [UNCHANGED]"
                        }
                    }
                }
                else {
                    $differences.OnlyInBackup2 += [PSCustomObject]@{
                        PolicyType   = $policyType
                        PolicyName   = $policy2Name
                        Created      = ConvertFrom-JsonDate -DateString $policy2.createdDateTime
                        LastModified = ConvertFrom-JsonDate -DateString $policy2.lastModifiedDateTime
                        IsAssigned   = if ($null -ne $policy2.isAssigned) { if ($policy2.isAssigned) { "Yes" } else { "No" } } else { "N/A" }
                    }
                    Write-ColoredOutput "  - $policy2Name [ONLY IN BACKUP 2]" -Color $script:Colors.Warning
                }
            }

            # Check for policies only in backup 1 (not matched)
            foreach ($key1 in $policies1.Keys) {
                if ($key1 -notin $matched1Keys) {
                    $policy1 = $policies1[$key1]
                    $policyName = if ($policy1.displayName) { $policy1.displayName } else { $policy1.name }
                    $differences.OnlyInBackup1 += [PSCustomObject]@{
                        PolicyType   = $policyType
                        PolicyName   = $policyName
                        Created      = ConvertFrom-JsonDate -DateString $policy1.createdDateTime
                        LastModified = ConvertFrom-JsonDate -DateString $policy1.lastModifiedDateTime
                        IsAssigned   = if ($null -ne $policy1.isAssigned) { if ($policy1.isAssigned) { "Yes" } else { "No" } } else { "N/A" }
                    }
                    Write-ColoredOutput "  - $policyName [ONLY IN BACKUP 1]" -Color $script:Colors.Warning
                }
            }
        }
    }

    # Summary
    Write-ColoredOutput "`n════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
    Write-ColoredOutput "Comparison Summary:" -Color $script:Colors.Header
    Write-ColoredOutput "Only in Backup 1 ($backup1Label): $($differences.OnlyInBackup1.Count) policies" -Color $script:Colors.Warning
    Write-ColoredOutput "Only in Backup 2 ($backup2Label): $($differences.OnlyInBackup2.Count) policies" -Color $script:Colors.Warning
    Write-ColoredOutput "Changed (different settings): $($differences.Modified.Count) policies" -Color $script:Colors.Warning
    Write-ColoredOutput "Unchanged (identical): $($differences.Unchanged) policies" -Color $script:Colors.Success
    Write-ColoredOutput "════════════════════════════════════════════════════════════════" -Color $script:Colors.Header

    # Show detailed lists if there are changes
    if ($differences.OnlyInBackup1.Count -gt 0) {
        Write-ColoredOutput "`nOnly in Backup 1 ($backup1Label):" -Color $script:Colors.Warning
        foreach ($policy in $differences.OnlyInBackup1) {
            Write-ColoredOutput "  - $($policy.PolicyType)/$($policy.PolicyName)" -Color $script:Colors.Warning
        }
    }

    if ($differences.OnlyInBackup2.Count -gt 0) {
        Write-ColoredOutput "`nOnly in Backup 2 ($backup2Label):" -Color $script:Colors.Warning
        foreach ($policy in $differences.OnlyInBackup2) {
            Write-ColoredOutput "  - $($policy.PolicyType)/$($policy.PolicyName)" -Color $script:Colors.Warning
        }
    }

    if ($differences.Modified.Count -gt 0) {
        Write-ColoredOutput "`nChanged Policies (different settings between backups):" -Color $script:Colors.Warning
        foreach ($policy in $differences.Modified) {
            Write-ColoredOutput "  - $($policy.PolicyType)/$($policy.PolicyName)" -Color $script:Colors.Warning
        }
    }

    return $differences
}

function Export-ComparisonToMarkdown {
    <#
    .SYNOPSIS
    Exports backup comparison results to a Markdown file.

    .DESCRIPTION
    Creates a formatted Markdown report of backup comparison results,
    listing each policy on a separate line for clear visibility.

    .PARAMETER Differences
    The comparison results hashtable from Compare-Backups

    .PARAMETER Backup1Path
    Path to the first backup

    .PARAMETER Backup2Path
    Path to the second backup

    .PARAMETER OutputPath
    Path for the output Markdown file
    #>
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Differences,

        [Parameter(Mandatory = $true)]
        [string]$Backup1Path,

        [Parameter(Mandatory = $true)]
        [string]$Backup2Path,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    # Load metadata from backups
    $metadata1 = if (Test-Path (Join-Path $Backup1Path "metadata.json")) {
        Get-Content (Join-Path $Backup1Path "metadata.json") | ConvertFrom-Json
    } else { @{BackupDate = "Unknown"} }

    $metadata2 = if (Test-Path (Join-Path $Backup2Path "metadata.json")) {
        Get-Content (Join-Path $Backup2Path "metadata.json") | ConvertFrom-Json
    } else { @{BackupDate = "Unknown"} }

    $md = @"
# UniFy-Endpoint Backup Comparison Report

**Generated:** $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

## Backups Compared

| Backup | Date | Path |
|--------|------|------|
| Backup 1 | $($metadata1.BackupDate) | ``$Backup1Path`` |
| Backup 2 | $($metadata2.BackupDate) | ``$Backup2Path`` |

---

## Summary

| Status | Count |
|--------|-------|
| Only in Backup 1 | $($Differences.OnlyInBackup1.Count) |
| Only in Backup 2 | $($Differences.OnlyInBackup2.Count) |
| Modified | $($Differences.Modified.Count) |
| Unchanged | $($Differences.Unchanged) |

---

"@

    # Only in Backup 1
    $md += "## Only in Backup 1 - $($metadata1.BackupDate) ($($Differences.OnlyInBackup1.Count))`n`n"
    if ($Differences.OnlyInBackup1.Count -gt 0) {
        $md += "| Policy Type | Policy Name | Created | Last Modified | Assigned |`n"
        $md += "|-------------|-------------|---------|--------------|----------|`n"
        foreach ($policy in ($Differences.OnlyInBackup1 | Sort-Object -Property PolicyName)) {
            $md += "| $($policy.PolicyType) | $($policy.PolicyName) | $($policy.Created) | $($policy.LastModified) | $($policy.IsAssigned) |`n"
        }
    } else {
        $md += "*No policies exclusive to Backup 1*`n"
    }
    $md += "`n"

    # Only in Backup 2
    $md += "## Only in Backup 2 - $($metadata2.BackupDate) ($($Differences.OnlyInBackup2.Count))`n`n"
    if ($Differences.OnlyInBackup2.Count -gt 0) {
        $md += "| Policy Type | Policy Name | Created | Last Modified | Assigned |`n"
        $md += "|-------------|-------------|---------|--------------|----------|`n"
        foreach ($policy in ($Differences.OnlyInBackup2 | Sort-Object -Property PolicyName)) {
            $md += "| $($policy.PolicyType) | $($policy.PolicyName) | $($policy.Created) | $($policy.LastModified) | $($policy.IsAssigned) |`n"
        }
    } else {
        $md += "*No policies exclusive to Backup 2*`n"
    }
    $md += "`n"

    # Modified Policies
    $md += "## Modified Policies - Different Settings ($($Differences.Modified.Count))`n`n"
    if ($Differences.Modified.Count -gt 0) {
        $md += "| Policy Type | Policy Name | Created | Last Modified | Assigned |`n"
        $md += "|-------------|-------------|---------|--------------|----------|`n"
        foreach ($policy in ($Differences.Modified | Sort-Object -Property PolicyName)) {
            $md += "| $($policy.PolicyType) | $($policy.PolicyName) | $($policy.Created) | $($policy.LastModified) | $($policy.IsAssigned) |`n"
        }
    } else {
        $md += "*No policies have different settings*`n"
    }
    $md += "`n"

    $md += @"
---

*Generated by UniFy-Endpoint v$($script:Version)*

**Note:** The comparison excludes metadata fields (id, createdDateTime, lastModifiedDateTime, version, assignments, etc.)
to focus on actual configuration differences.
"@

    # Save to file
    $md | Out-File -FilePath $OutputPath -Encoding UTF8
    Write-ColoredOutput "Comparison report saved to: $OutputPath" -Color $script:Colors.Success
}

#endregion

#region Drift Detection Functions

function Get-CurrentPolicyHash {
    param (
        [object]$Policy
    )

    # Comprehensive list of properties to ignore (must match Compare-PolicyContent)
    $propertiesToRemove = @(
        'id', 'createdDateTime', 'lastModifiedDateTime', 'modifiedDateTime', 'version',
        '@odata.context', '@odata.type', '@odata.id', '@odata.etag',
        'roleScopeTagIds', 'supportsScopeTags',
        'deviceManagementApplicabilityRuleOsEdition', 'deviceManagementApplicabilityRuleOsVersion',
        'deviceManagementApplicabilityRuleDeviceMode',
        'creationSource', 'isAssigned', 'priority', 'createdBy', 'lastModifiedBy',
        'assignments', 'assignmentFilterEvaluationStatusDetails',
        'templateReference', 'templateId', 'templateDisplayName', 'templateDisplayVersion', 'templateFamily',
        'settingCount', 'settingDefinitionId', 'settingInstanceTemplateId', 'settingValueTemplateId',
        'settingValueTemplateReference', 'settingInstanceTemplateReference', 'odataType',
        'secretReferenceValueId', 'deviceConfigurationId', 'groupId', 'sourceId', 'payloadId',
        'deviceNameTemplate', 'azureAdJoinType', 'managementServiceAppId',
        'targetedAppManagementLevels', 'appGroupType', 'deployedAppCount', 'apps', 'deploymentSummary',
        'minimumRequiredPatchVersion', 'minimumRequiredSdkVersion', 'minimumWipeSdkVersion',
        'minimumWipePatchVersion', 'minimumRequiredAppVersion', 'minimumWipeAppVersion',
        'minimumRequiredOsVersion', 'minimumWipeOsVersion',
        'scriptContentMd5Hash', 'scriptContentSha256Hash',
        'detectionScriptContentMd5Hash', 'detectionScriptContentSha256Hash',
        'remediationScriptContentMd5Hash', 'remediationScriptContentSha256Hash',
        'scheduledActionsForRule', 'validOperatingSystemBuildRanges',
        'runSummary', 'deviceRunStates', 'isGlobalScript', 'deviceHealthScriptType', 'runSchedule',
        'fileName', 'runAsAccount', 'enforceSignatureCheck', 'runAs32Bit',
        'customSettings', 'encodedSettingXml', 'targetedManagedAppGroupType',
        'configurationDeploymentSummaryPerApp', 'customBrowserProtocol', 'customBrowserPackageId',
        'customBrowserDisplayName', 'customDialerAppProtocol', 'customDialerAppPackageId',
        'customDialerAppDisplayName', 'allowedAndroidDeviceManufacturers',
        'appActionIfAndroidDeviceManufacturerNotAllowed', 'exemptedAppProtocols', 'exemptedAppPackages',
        'settingDefinitions',
        'technologies', 'platformType', 'settingsCount', 'priorityMetaData', 'assignedAppsCount'
    )

    # Recursive helper to remove ignored properties at all nesting levels
    function Remove-HashIgnoredProperties {
        param ($obj, $ignoreList)

        if ($null -eq $obj) { return $null }

        if ($obj -is [array]) {
            $cleanArray = @()
            foreach ($item in $obj) {
                $cleanArray += Remove-HashIgnoredProperties -obj $item -ignoreList $ignoreList
            }
            return $cleanArray
        }
        elseif ($obj -is [PSCustomObject] -or $obj -is [System.Collections.IDictionary]) {
            $cleanObj = [PSCustomObject]@{}
            $props = if ($obj -is [System.Collections.IDictionary]) { $obj.Keys } else { $obj.PSObject.Properties.Name }

            foreach ($propName in $props) {
                if ($propName -notin $ignoreList -and $propName -notlike '*@odata*') {
                    $value = if ($obj -is [System.Collections.IDictionary]) { $obj[$propName] } else { $obj.$propName }
                    $cleanValue = Remove-HashIgnoredProperties -obj $value -ignoreList $ignoreList
                    $cleanObj | Add-Member -NotePropertyName $propName -NotePropertyValue $cleanValue -Force
                }
            }
            return $cleanObj
        }
        else {
            return $obj
        }
    }

    # Remove dynamic properties that change but don't represent actual configuration drift
    $cleanPolicy = $Policy | ConvertTo-Json -Depth 50 -Compress | ConvertFrom-Json
    $cleanPolicy = Remove-HashIgnoredProperties -obj $cleanPolicy -ignoreList $propertiesToRemove

    # Normalize Base64 script content fields
    # Backup stores decoded text (human-readable), API returns Base64-encoded.
    # Decode any Base64 values so both sides are in plain text for comparison.
    $base64Fields = @('scriptContent', 'detectionScriptContent', 'remediationScriptContent')
    foreach ($field in $base64Fields) {
        if ($null -ne $cleanPolicy -and $cleanPolicy.PSObject.Properties[$field]) {
            $val = $cleanPolicy.$field
            if ($val -and $val -is [string]) {
                try {
                    $decoded = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($val))
                    $cleanPolicy | Add-Member -NotePropertyName $field -NotePropertyValue $decoded -Force
                } catch {
                    # Not valid Base64 — already decoded text, leave as-is
                }
            }
        }
    }

    # Sort properties alphabetically at all nesting levels to ensure order-independent comparison.
    # Backup items (from ConvertFrom-Json) have properties in JSON file order, while API items
    # (from Invoke-MgGraphRequest Hashtable) have non-deterministic enumeration order.
    # Without sorting, identical content produces different JSON strings and different hashes.
    function Sort-ObjectProperties {
        param ($obj)
        if ($null -eq $obj) { return $null }
        if ($obj -is [array]) {
            $sorted = @()
            foreach ($item in $obj) {
                $sorted += Sort-ObjectProperties -obj $item
            }
            return $sorted
        }
        elseif ($obj -is [PSCustomObject]) {
            $sortedObj = [PSCustomObject]@{}
            $obj.PSObject.Properties.Name | Sort-Object | ForEach-Object {
                $sortedValue = Sort-ObjectProperties -obj $obj.$_
                $sortedObj | Add-Member -NotePropertyName $_ -NotePropertyValue $sortedValue -Force
            }
            return $sortedObj
        }
        else {
            return $obj
        }
    }

    $cleanPolicy = Sort-ObjectProperties -obj $cleanPolicy

    # Create hash of the cleaned policy
    $policyJson = $cleanPolicy | ConvertTo-Json -Depth 50 -Compress
    $stringAsStream = [System.IO.MemoryStream]::new()
    $writer = [System.IO.StreamWriter]::new($stringAsStream)
    $writer.write($policyJson)
    $writer.Flush()
    $stringAsStream.Position = 0
    $hash = Get-FileHash -InputStream $stringAsStream -Algorithm SHA256

    return $hash.Hash
}

function Get-PolicyBaseName {
    <#
    .SYNOPSIS
    Strips restore/import suffixes from a policy name to get the original base name.
    Handles: "- [Restored]", "- [New Policy]" suffixes and "[Restored] " prefix (backward compat).
    #>
    param ([string]$Name)
    if ($Name -like "[Restored] *") {
        return $Name -replace '^\[Restored\] ', ''
    } elseif ($Name -match ' - \[(Restored|New Policy)\]$') {
        return $Name -replace ' - \[(Restored|New Policy)\]$', ''
    }
    return $Name
}

function Compare-PolicyType {
    <#
    .SYNOPSIS
    Helper function to compare a specific policy type for drift detection.

    .DESCRIPTION
    Compares backup policies with current Intune policies for a specific type.
    Returns drift results (Added, Removed, Modified, Unchanged).
    #>
    param (
        [string]$BackupPath,
        [string]$FolderName,
        [string]$TypeName,
        [string]$GraphUri,
        [string]$DetailUriTemplate,  # Use {id} as placeholder
        [string]$NameProperty = "displayName",  # Some use 'name' instead
        [hashtable]$DriftResults,
        [string]$OdataTypeFilter = '',         # Filter current items to this @odata.type
        [string]$BackupOdataTypeFilter = '',    # Filter backup items to this @odata.type
        [string[]]$ExcludeOdataTypes = @(),    # Exclude current items matching these @odata.type values
        [switch]$SuppressHeader                 # Suppress "Checking X..." header (for split types)
    )

    $backupFolderPath = Join-Path $BackupPath $FolderName

    if (-not (Test-Path $backupFolderPath)) {
        return
    }

    if (-not $SuppressHeader) {
        Write-ColoredOutput "`nChecking $TypeName..." -Color $script:Colors.Info
    }

    try {
        $currentItems = Get-GraphData -Uri $GraphUri

        # Filter current items by @odata.type if specified (e.g., ESP only wants windows10EnrollmentCompletionPageConfiguration)
        if ($OdataTypeFilter) {
            $currentItems = @($currentItems | Where-Object { $_.'@odata.type' -eq $OdataTypeFilter })
        }

        # Exclude current items by @odata.type (e.g., exclude WUfB from DeviceConfigurations since they have their own category)
        if ($ExcludeOdataTypes.Count -gt 0) {
            $currentItems = @($currentItems | Where-Object { $_.'@odata.type' -notin $ExcludeOdataTypes })
        }

        $backupItems = Get-ChildItem -Path $backupFolderPath -Filter "*.json" -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                Get-Content $_.FullName -Raw | ConvertFrom-Json
            }
            catch {
                Write-Log "Failed to parse $($_.FullName): $_" -Level Warning
                $null
            }
        } | Where-Object { $_ -ne $null }

        # Filter backup items by @odata.type if specified (e.g., App Protection split by platform)
        if ($BackupOdataTypeFilter -and $backupItems) {
            $backupItems = @($backupItems | Where-Object { $_.'@odata.type' -eq $BackupOdataTypeFilter })
        }

        if (-not $backupItems -or $backupItems.Count -eq 0) {
            Write-ColoredOutput "  No backup items found in $FolderName" -Color $script:Colors.Info
            return
        }

        # Check for removed, renamed, or changed policies
        foreach ($backupItem in $backupItems) {
            $itemName = $backupItem.$NameProperty
            if (-not $itemName) { continue }

            # Step 1: Try to match by ID first (handles renames)
            $currentItem = $null
            $foundById = $false
            if ($backupItem.id) {
                $currentItem = $currentItems | Where-Object { $_.id -eq $backupItem.id } | Select-Object -First 1
                if ($currentItem) { $foundById = $true }
            }

            # Step 2: Fallback to name match
            if (-not $currentItem) {
                $currentMatches = @($currentItems | Where-Object { $_.$NameProperty -eq $itemName })

                # For ESP and similar endpoints, filter by @odata.type if backup has it
                if ($currentMatches.Count -gt 1 -and $backupItem.'@odata.type') {
                    $typeFiltered = @($currentMatches | Where-Object { $_.'@odata.type' -eq $backupItem.'@odata.type' })
                    if ($typeFiltered.Count -gt 0) {
                        $currentMatches = $typeFiltered
                    }
                }

                $currentItem = if ($currentMatches.Count -gt 0) { $currentMatches[0] } else { $null }
            }

            # Step 3: Fallback - check if a tenant policy with restore/import suffix matches this backup policy
            $matchedByBaseName = $false
            if (-not $currentItem) {
                $currentItem = $currentItems | Where-Object {
                    (Get-PolicyBaseName $_.$NameProperty) -eq $itemName
                } | Select-Object -First 1
                if ($currentItem) { $matchedByBaseName = $true }
            }

            if (-not $currentItem) {
                # Not found by ID, name, or base name - truly removed
                $DriftResults.Removed += [PSCustomObject]@{
                    Type       = $TypeName
                    Name       = $itemName
                    BackupPath = $backupFolderPath
                    BackupData = $backupItem
                }
                Write-ColoredOutput "  - $itemName" -Color $script:Colors.Default -NoNewline
                Write-ColoredOutput " [REMOVED]" -Color $script:Colors.Error
            }
            else {
                # Get full current item if detail URI provided
                $fullCurrentItem = $currentItem
                if ($DetailUriTemplate) {
                    try {
                        $detailUri = $DetailUriTemplate -replace '\{id\}', $currentItem.id
                        $fullCurrentItem = Get-GraphData -Uri $detailUri
                    }
                    catch {
                        $fullCurrentItem = $currentItem
                    }
                }

                # Check if name changed
                $tenantName = $fullCurrentItem.$NameProperty
                $nameChanged = ($tenantName -ne $itemName)

                # Check if content changed using hash (fast path) + Compare-PolicyContent (authoritative)
                $backupHash = Get-CurrentPolicyHash -Policy $backupItem
                $currentHash = Get-CurrentPolicyHash -Policy $fullCurrentItem
                $contentChanged = $false
                $differences = @()

                if ($backupHash -ne $currentHash) {
                    # Hash differs - run detailed comparison to confirm
                    $comparison = Compare-PolicyContent -BackupPolicy $backupItem -TenantPolicy $fullCurrentItem
                    # Filter out name fields - a name change is a rename, not a settings change
                    $settingsDiffs = @($comparison.Differences | Where-Object { $_ -ne 'name' -and $_ -ne 'displayName' })
                    if ($settingsDiffs.Count -gt 0) {
                        $contentChanged = $true
                        $differences = $settingsDiffs
                    }
                }

                if ($contentChanged) {
                    # Settings changed - [CHANGED]
                    $DriftResults.Changed += [PSCustomObject]@{
                        Type         = $TypeName
                        Name         = $itemName
                        TenantName   = $tenantName
                        NameChanged  = $nameChanged
                        BackupPath   = $backupFolderPath
                        BackupData   = $backupItem
                        CurrentData  = $fullCurrentItem
                        Differences  = $differences
                        GraphUri     = $GraphUri
                        NameProperty = $NameProperty
                    }
                    $settingWord = if ($differences.Count -eq 1) { 'Setting' } else { 'Settings' }
                    $changeMsg = " [CHANGED] - [$($differences.Count) $settingWord Differ]"
                    if ($nameChanged) { $changeMsg += " [RENAMED] -> [$tenantName]" }
                    Write-ColoredOutput "  - $itemName" -Color $script:Colors.Default -NoNewline
                    Write-ColoredOutput $changeMsg -Color $script:Colors.Warning
                }
                elseif ($matchedByBaseName) {
                    # Matched by base name - previously restored/imported policy
                    $DriftResults.PrevRestored += [PSCustomObject]@{
                        Type         = $TypeName
                        Name         = $itemName
                        TenantName   = $tenantName
                        BackupPath   = $backupFolderPath
                        BackupData   = $backupItem
                        CurrentData  = $fullCurrentItem
                        GraphUri     = $GraphUri
                        NameProperty = $NameProperty
                    }
                    Write-ColoredOutput "  - $itemName" -Color $script:Colors.Default -NoNewline
                    Write-ColoredOutput " [RESTORED]" -Color $script:Colors.Success
                }
                elseif ($nameChanged) {
                    # Only name changed - [RENAMED]
                    $DriftResults.Renamed += [PSCustomObject]@{
                        Type         = $TypeName
                        Name         = $itemName
                        TenantName   = $tenantName
                        BackupPath   = $backupFolderPath
                        BackupData   = $backupItem
                        CurrentData  = $fullCurrentItem
                        GraphUri     = $GraphUri
                        NameProperty = $NameProperty
                    }
                    Write-ColoredOutput "  - $itemName" -Color $script:Colors.Default -NoNewline
                    Write-ColoredOutput " [RENAMED] -> [$tenantName]" -Color $script:Colors.Info
                }
                else {
                    # No changes
                    $DriftResults.Unchanged += [PSCustomObject]@{
                        Type = $TypeName
                        Name = $itemName
                    }
                    Write-ColoredOutput "  - $itemName" -Color $script:Colors.Default -NoNewline
                    Write-ColoredOutput " [UNCHANGED]" -Color $script:Colors.Success
                }
            }
        }

        # Collect backup IDs and @odata.types for filtering added items
        $backupIds = @($backupItems | Where-Object { $_.id } | ForEach-Object { $_.id })
        $backupOdataTypes = @($backupItems | Where-Object { $_.'@odata.type' } | ForEach-Object { $_.'@odata.type' } | Select-Object -Unique)

        # Check for added policies (exist on tenant but not in backup)
        foreach ($currentItem in $currentItems) {
            $itemName = $currentItem.$NameProperty
            if (-not $itemName) { continue }

            # Skip items whose @odata.type doesn't match any backup item type (prevents cross-category false positives)
            if ($backupOdataTypes.Count -gt 0 -and $currentItem.'@odata.type' -and $currentItem.'@odata.type' -notin $backupOdataTypes) {
                continue
            }

            # Skip if this tenant policy matches a backup item by ID (handles renames)
            if ($currentItem.id -and $currentItem.id -in $backupIds) {
                continue
            }

            # Also check by name
            $backupItem = $backupItems | Where-Object { $_.$NameProperty -eq $itemName }

            # Also check by base name (handles restored/imported copies with suffix)
            if (-not $backupItem) {
                $tenantBaseName = Get-PolicyBaseName $itemName
                if ($tenantBaseName -ne $itemName) {
                    $backupItem = $backupItems | Where-Object { $_.$NameProperty -eq $tenantBaseName }
                }
            }

            if (-not $backupItem) {
                $DriftResults.Added += [PSCustomObject]@{
                    Type        = $TypeName
                    Name        = $itemName
                    CurrentData = $currentItem
                }
                Write-ColoredOutput "  - $itemName" -Color $script:Colors.Default -NoNewline
                Write-ColoredOutput " [ADDED]" -Color $script:Colors.Success
            }
        }
    }
    catch {
        Write-ColoredOutput "  Error checking $TypeName : $_" -Color $script:Colors.Warning
        Write-Log "Drift detection error for $TypeName : $_" -Level Warning
    }
}

function Detect-ConfigurationDrift {
    param (
        [string]$BackupPath
    )

    Write-ColoredOutput "`nDetecting configuration drift..." -Color $script:Colors.Header

    if (-not (Test-Path $BackupPath)) {
        Write-ColoredOutput "Backup path not found: $BackupPath" -Color $script:Colors.Error
        return
    }

    $metadataPath = Join-Path $BackupPath "metadata.json"
    if (-not (Test-Path $metadataPath)) {
        Write-ColoredOutput "Invalid backup: metadata.json not found" -Color $script:Colors.Error
        return
    }

    $metadata = Get-Content $metadataPath | ConvertFrom-Json
    Write-ColoredOutput "Comparing current configuration with backup from: $($metadata.BackupDate)" -Color $script:Colors.Info

    $driftResults = @{
        Added        = @()
        Removed      = @()
        Changed      = @()
        Renamed      = @()
        PrevRestored = @()
        Unchanged    = @()
    }

    # Define all policy types to check with their configurations
    # Format: FolderName, TypeName, GraphUri, DetailUriTemplate, NameProperty
    $policyTypes = @(
        @{
            FolderName = "DeviceConfigurations"
            TypeName = "DeviceConfiguration"
            GraphUri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations"
            DetailUri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/{id}"
            NameProperty = "displayName"
            ExcludeOdataTypes = @('#microsoft.graph.windowsUpdateForBusinessConfiguration')
        },
        @{
            FolderName = "CompliancePolicies"
            TypeName = "CompliancePolicy"
            GraphUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies"
            DetailUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies/{id}"
            NameProperty = "displayName"
        },
        @{
            FolderName = "SettingsCatalogPolicies"
            TypeName = "SettingsCatalogPolicy"
            GraphUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
            DetailUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/{id}?`$expand=settings"
            NameProperty = "name"
        },
        @{
            FolderName = "PowerShellScripts"
            TypeName = "PowerShellScript"
            GraphUri = "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts"
            DetailUri = "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts/{id}"
            NameProperty = "displayName"
        },
        @{
            FolderName = "RemediationScripts"
            TypeName = "RemediationScript"
            GraphUri = "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts"
            DetailUri = "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts/{id}"
            NameProperty = "displayName"
        },
        @{
            FolderName = "AutopilotProfiles"
            TypeName = "AutopilotProfile"
            GraphUri = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles"
            DetailUri = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles/{id}"
            NameProperty = "displayName"
        },
        @{
            FolderName = "EnrollmentStatusPage"
            TypeName = "EnrollmentStatusPage"
            GraphUri = "https://graph.microsoft.com/v1.0/deviceManagement/deviceEnrollmentConfigurations"
            DetailUri = "https://graph.microsoft.com/v1.0/deviceManagement/deviceEnrollmentConfigurations/{id}"
            NameProperty = "displayName"
            OdataTypeFilter = '#microsoft.graph.windows10EnrollmentCompletionPageConfiguration'
        },
        @{
            FolderName = "AssignmentFilters"
            TypeName = "AssignmentFilter"
            GraphUri = "https://graph.microsoft.com/beta/deviceManagement/assignmentFilters"
            DetailUri = "https://graph.microsoft.com/beta/deviceManagement/assignmentFilters/{id}"
            NameProperty = "displayName"
        },
        @{
            FolderName = "AdminTemplates"
            TypeName = "AdministrativeTemplate"
            GraphUri = "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations"
            DetailUri = "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations/{id}"
            NameProperty = "displayName"
        },
        @{
            FolderName = "AppProtectionPolicies"
            TypeName = "AppProtectionPolicy"
            GraphUri = "https://graph.microsoft.com/beta/deviceAppManagement/iosManagedAppProtections"
            DetailUri = "https://graph.microsoft.com/beta/deviceAppManagement/iosManagedAppProtections/{id}"
            NameProperty = "displayName"
            BackupOdataTypeFilter = '#microsoft.graph.iosManagedAppProtection'
        },
        @{
            FolderName = "AppProtectionPolicies"
            TypeName = "AppProtectionPolicy"
            GraphUri = "https://graph.microsoft.com/beta/deviceAppManagement/androidManagedAppProtections"
            DetailUri = "https://graph.microsoft.com/beta/deviceAppManagement/androidManagedAppProtections/{id}"
            NameProperty = "displayName"
            BackupOdataTypeFilter = '#microsoft.graph.androidManagedAppProtection'
        },
        @{
            FolderName = "AppProtectionPolicies"
            TypeName = "AppProtectionPolicy"
            GraphUri = "https://graph.microsoft.com/beta/deviceAppManagement/windowsManagedAppProtections"
            DetailUri = "https://graph.microsoft.com/beta/deviceAppManagement/windowsManagedAppProtections/{id}"
            NameProperty = "displayName"
            BackupOdataTypeFilter = '#microsoft.graph.windowsManagedAppProtection'
        },
        @{
            FolderName = "AppConfigManagedApps"
            TypeName = "AppConfigManagedApp"
            GraphUri = "https://graph.microsoft.com/beta/deviceAppManagement/targetedManagedAppConfigurations"
            DetailUri = "https://graph.microsoft.com/beta/deviceAppManagement/targetedManagedAppConfigurations/{id}?`$expand=apps,assignments"
            NameProperty = "displayName"
        },
        @{
            FolderName = "MacOSScripts"
            TypeName = "MacOSScript"
            GraphUri = "https://graph.microsoft.com/beta/deviceManagement/deviceShellScripts"
            DetailUri = "https://graph.microsoft.com/beta/deviceManagement/deviceShellScripts/{id}"
            NameProperty = "displayName"
        },
        @{
            FolderName = "MacOSCustomAttributes"
            TypeName = "MacOSCustomAttribute"
            GraphUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCustomAttributeShellScripts"
            DetailUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCustomAttributeShellScripts/{id}"
            NameProperty = "displayName"
        },
        @{
            FolderName = "AppConfigManagedDevices"
            TypeName = "AppConfigManagedDevice"
            GraphUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations"
            DetailUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations/{id}"
            NameProperty = "displayName"
        },
        @{
            FolderName = "AutopilotDevicePrep"
            TypeName = "AutopilotDevicePrep"
            GraphUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?`$filter=templateReference/templateId eq '20f9c2d1-e508-4122-8692-7f284b3956f1'"
            DetailUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/{id}?`$expand=settings"
            NameProperty = "name"
        },
        @{
            FolderName = "WUfBPolicies"
            TypeName = "WUfBPolicy"
            GraphUri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations"
            DetailUri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/{id}"
            NameProperty = "displayName"
            OdataTypeFilter = '#microsoft.graph.windowsUpdateForBusinessConfiguration'
        }
    )

    # Check each policy type
    $lastTypeName = ''
    foreach ($policyType in $policyTypes) {
        $params = @{
            BackupPath        = $BackupPath
            FolderName        = $policyType.FolderName
            TypeName          = $policyType.TypeName
            GraphUri          = $policyType.GraphUri
            DetailUriTemplate = $policyType.DetailUri
            NameProperty      = $policyType.NameProperty
            DriftResults      = $driftResults
        }
        if ($policyType.OdataTypeFilter) {
            $params['OdataTypeFilter'] = $policyType.OdataTypeFilter
        }
        if ($policyType.BackupOdataTypeFilter) {
            $params['BackupOdataTypeFilter'] = $policyType.BackupOdataTypeFilter
        }
        if ($policyType.ExcludeOdataTypes) {
            $params['ExcludeOdataTypes'] = $policyType.ExcludeOdataTypes
        }
        # Suppress duplicate headers when same TypeName appears multiple times (e.g., App Protection split by platform)
        if ($policyType.TypeName -eq $lastTypeName) {
            $params['SuppressHeader'] = $true
        }
        $lastTypeName = $policyType.TypeName

        Compare-PolicyType @params
    }

    # Summary
    Write-ColoredOutput "`n════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
    Write-ColoredOutput "Drift Detection Summary:" -Color $script:Colors.Header
    Write-ColoredOutput "Added: $($driftResults.Added.Count) policies" -Color $script:Colors.Success
    Write-ColoredOutput "Removed: $($driftResults.Removed.Count) policies" -Color $script:Colors.Error
    Write-ColoredOutput "Changed: $($driftResults.Changed.Count) policies" -Color $script:Colors.Warning
    Write-ColoredOutput "Renamed: $($driftResults.Renamed.Count) policies" -Color $script:Colors.Info
    if ($driftResults.PrevRestored.Count -gt 0) {
        Write-ColoredOutput "Prev. Restored: $($driftResults.PrevRestored.Count) policies" -Color $script:Colors.Success
    }
    Write-ColoredOutput "Unchanged: $($driftResults.Unchanged.Count) policies" -Color $script:Colors.Info
    Write-ColoredOutput "════════════════════════════════════════════════════════════════" -Color $script:Colors.Header

    return $driftResults
}

function Rename-PolicyOnTenant {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Policy
    )

    $backupName = $Policy.Name
    $currentName = $Policy.TenantName
    $policyId = $Policy.BackupData.id
    $baseUri = ($Policy.GraphUri -split '\?')[0]
    $patchUri = "$baseUri/$policyId"
    $nameProperty = $Policy.NameProperty

    Write-ColoredOutput "  Reverting: '$currentName' -> '$backupName'" -Color $script:Colors.Default -NoNewline

    try {
        $body = @{ $nameProperty = $backupName } | ConvertTo-Json
        Invoke-MgGraphRequest -Uri $patchUri -Method PATCH -Body $body -ContentType "application/json"
        Write-ColoredOutput " [SUCCESS]" -Color $script:Colors.Success
        return $true
    }
    catch {
        Write-ColoredOutput " [FAILED: $($_.Exception.Message)]" -Color $script:Colors.Error
        return $false
    }
}

function Restore-SelectedPolicies {
    param (
        [object]$DriftResults,
        [switch]$Preview
    )

    if (-not $DriftResults) {
        Write-ColoredOutput "No drift results to process" -Color $script:Colors.Warning
        return
    }

    while ($true) {

    # Show policies that can be restored
    Write-ColoredOutput "`nPolicies that can be restored:" -Color $script:Colors.Header
    Write-ColoredOutput "══════════════════════════════════════════════════════════════════════" -Color $script:Colors.Info

    $restorablePolicies = @()
    $index = 1

    # Removed policies can be restored
    foreach ($policy in $DriftResults.Removed) {
        Write-ColoredOutput "[$index] [REMOVED] $($policy.Name) ($($policy.Type))" -Color $script:Colors.Error
        $restorablePolicies += $policy
        $index++
    }

    # Changed policies can be restored to backup version
    foreach ($policy in $DriftResults.Changed) {
        $diffCount = if ($policy.Differences) { $policy.Differences.Count } else { 0 }
        $changeMsg = "[$index] [CHANGED] $($policy.Name) ($($policy.Type)) - $diffCount settings differ"
        if ($policy.NameChanged) { $changeMsg += " [also renamed to '$($policy.TenantName)']" }
        Write-ColoredOutput $changeMsg -Color $script:Colors.Warning
        $restorablePolicies += $policy
        $index++
    }

    # Show renamed policies with indices for revert
    $renamablePolicies = @()
    if ($DriftResults.Renamed -and $DriftResults.Renamed.Count -gt 0) {
        Write-ColoredOutput "`nRenamed policies (Settings Unchanged):" -Color $script:Colors.Info
        $renameIndex = 1
        foreach ($policy in $DriftResults.Renamed) {
            Write-ColoredOutput "  [R$renameIndex] [RENAMED] $($policy.Name) -> $($policy.TenantName) ($($policy.Type))" -Color $script:Colors.Info
            $renamablePolicies += $policy
            $renameIndex++
        }
    }

    if ($restorablePolicies.Count -eq 0 -and $renamablePolicies.Count -eq 0) {
        Write-ColoredOutput "`nNo policies need restoration" -Color $script:Colors.Success
        return
    }

    # Menu options
    Write-ColoredOutput "" -Color $script:Colors.Default
    if ($restorablePolicies.Count -gt 0) {
        Write-ColoredOutput "[A] Restore all changed/removed policies" -Color $script:Colors.Success
    }
    if ($DriftResults.Changed -and $DriftResults.Changed.Count -gt 0) {
        Write-ColoredOutput "[D] Generate drift comparison report" -Color $script:Colors.Info
    }
    if ($renamablePolicies.Count -gt 0) {
        Write-ColoredOutput "[R] Revert ALL renamed policies to backup names" -Color $script:Colors.Info
        Write-ColoredOutput "    (or R1,R3 to select specific renamed policies)" -Color "Gray"
    }
    Write-ColoredOutput "[0] Back to main menu" -Color $script:Colors.Error

    $selection = Read-Host "`nSelect option (1,3,5 = restore, A = all, D = report, R = revert names, 0 = back)"

    if ($selection -eq "0") {
        return
    }

    # Handle [D] - Generate drift comparison report
    if ($selection -eq "D" -or $selection -eq "d") {
        if (-not $DriftResults.Changed -or $DriftResults.Changed.Count -eq 0) {
            Write-ColoredOutput "No changed policies to report on." -Color $script:Colors.Warning
            continue
        }
        if (-not (Test-Path $script:ReportsLocation)) {
            New-Item -ItemType Directory -Path $script:ReportsLocation -Force | Out-Null
        }
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

        $formatChoice = Read-Host "Report format? [H]TML or [M]arkdown (default: HTML)"
        $format = if ($formatChoice -eq "M" -or $formatChoice -eq "m") { "Markdown" } else { "HTML" }
        $ext = if ($format -eq "Markdown") { ".md" } else { ".html" }
        $outputPath = Join-Path $script:ReportsLocation "drift_report_$timestamp$ext"

        Write-ColoredOutput "`nGenerating drift report..." -Color $script:Colors.Info
        Export-DriftReport -DriftResults $DriftResults -Format $format -OutputPath $outputPath

        if (Test-Path $outputPath) {
            Write-ColoredOutput "Drift report saved to: $outputPath" -Color $script:Colors.Success
            $openReport = Read-Host "Open report now? (yes/no)"
            if ($openReport -eq "yes") { Start-Process $outputPath }
        }
        continue
    }

    # Handle [R] - Revert all renamed policies
    if ($selection -eq "R" -or $selection -eq "r") {
        if ($renamablePolicies.Count -eq 0) {
            Write-ColoredOutput "No renamed policies to revert." -Color $script:Colors.Warning
            continue
        }
        Write-ColoredOutput "`nThis will PATCH $($renamablePolicies.Count) policies to revert their names on the Intune tenant." -Color $script:Colors.Warning
        $confirm = Read-Host "Are you sure? (yes/no)"
        if ($confirm -ne "yes") {
            Write-ColoredOutput "Revert cancelled." -Color $script:Colors.Warning
            continue
        }
        Write-ColoredOutput "`nReverting policy names..." -Color $script:Colors.Header
        $successCount = 0
        $failCount = 0
        foreach ($policy in $renamablePolicies) {
            $result = Rename-PolicyOnTenant -Policy $policy
            if ($result) { $successCount++ } else { $failCount++ }
        }
        Write-ColoredOutput "`nRevert completed. Success: $successCount, Failed: $failCount" -Color $script:Colors.Success
        continue
    }

    # Parse selection for R-prefixed and numeric indices
    $rSelections = @()
    $numSelections = @()
    $parts = $selection -split ',' | ForEach-Object { $_.Trim() }
    foreach ($part in $parts) {
        if ($part -match '^[Rr](\d+)$') {
            $rSelections += [int]$Matches[1]
        }
        elseif ($part -match '^\d+$') {
            $numSelections += [int]$part
        }
    }

    # Process R-selections (rename reverts)
    if ($rSelections.Count -gt 0) {
        $selectedRenames = @()
        foreach ($idx in $rSelections) {
            $num = $idx - 1
            if ($num -ge 0 -and $num -lt $renamablePolicies.Count) {
                $selectedRenames += $renamablePolicies[$num]
            }
        }
        if ($selectedRenames.Count -gt 0) {
            Write-ColoredOutput "`nThis will PATCH $($selectedRenames.Count) policies to revert their names." -Color $script:Colors.Warning
            $confirm = Read-Host "Are you sure? (yes/no)"
            if ($confirm -eq "yes") {
                Write-ColoredOutput "`nReverting policy names..." -Color $script:Colors.Header
                $successCount = 0
                $failCount = 0
                foreach ($policy in $selectedRenames) {
                    $result = Rename-PolicyOnTenant -Policy $policy
                    if ($result) { $successCount++ } else { $failCount++ }
                }
                Write-ColoredOutput "`nRevert completed. Success: $successCount, Failed: $failCount" -Color $script:Colors.Success
            }
        }
        # If only R-selections and no numeric, back to menu
        if ($numSelections.Count -eq 0 -and ($selection -ne "A" -and $selection -ne "a")) {
            continue
        }
    }

    # Process numeric selections or [A] for restore
    $selectedPolicies = @()

    if ($selection -eq "A" -or $selection -eq "a") {
        $selectedPolicies = $restorablePolicies
    }
    elseif ($numSelections.Count -gt 0) {
        foreach ($idx in $numSelections) {
            $num = $idx - 1
            if ($num -ge 0 -and $num -lt $restorablePolicies.Count) {
                $selectedPolicies += $restorablePolicies[$num]
            }
        }
    }

    if ($selectedPolicies.Count -eq 0) {
        if ($rSelections.Count -eq 0) {
            Write-ColoredOutput "No policies selected for restore." -Color $script:Colors.Warning
        }
        continue
    }

    if (-not $Preview) {
        Write-ColoredOutput "`nWARNING: This will create $($selectedPolicies.Count) NEW policies with '- [Restored]' suffix!" -Color $script:Colors.Warning
        Write-ColoredOutput "Original policies will NOT be modified or deleted." -Color $script:Colors.Info
        $confirm = Read-Host "`nAre you sure you want to continue? (yes/no)"
        if ($confirm -ne "yes") {
            Write-ColoredOutput "Restore cancelled." -Color $script:Colors.Warning
            continue
        }
    }

    # Restore selected policies
    Write-ColoredOutput "`nRestoring selected policies..." -Color $script:Colors.Header
    $restoreStats = @{
        Success = 0
        Failed  = 0
    }

    foreach ($policy in $selectedPolicies) {
        $originalName = $policy.Name
        $newPolicyName = "$originalName - [Restored]"
        Write-ColoredOutput "  - $originalName -> $newPolicyName" -Color $script:Colors.Default -NoNewline

        if ($Preview) {
            Write-ColoredOutput " [PREVIEW]" -Color $script:Colors.Warning
            $restoreStats.Success++
            continue
        }

        try {
            $policyData = $policy.BackupData

            # Update the display name with new naming convention
            if ($policyData.displayName) {
                $policyData.displayName = $newPolicyName
            }
            elseif ($policyData.name) {
                $policyData.name = $newPolicyName
            }

            # Remove properties that shouldn't be in create/update request
            $policyData.PSObject.Properties.Remove('id')
            $policyData.PSObject.Properties.Remove('createdDateTime')
            $policyData.PSObject.Properties.Remove('lastModifiedDateTime')
            $policyData.PSObject.Properties.Remove('version')
            $policyData.PSObject.Properties.Remove('assignments')

            # Define API endpoints for each policy type
            $restoreSuccess = $false
            $apiUri = $null

            switch ($policy.Type) {
                "DeviceConfiguration" {
                    $apiUri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations"
                }
                "CompliancePolicy" {
                    $apiUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies"
                    # Handle scheduledActionsForRule: strip nested IDs and ensure a block action exists
                    if ($policyData.scheduledActionsForRule -and ($policyData.scheduledActionsForRule | Measure-Object).Count -gt 0) {
                        foreach ($rule in $policyData.scheduledActionsForRule) {
                            if ($rule.PSObject.Properties['id']) { $rule.PSObject.Properties.Remove('id') }
                            if ($rule.scheduledActionConfigurations) {
                                foreach ($config in $rule.scheduledActionConfigurations) {
                                    if ($config.PSObject.Properties['id']) { $config.PSObject.Properties.Remove('id') }
                                }
                            }
                        }
                        $hasBlock = $policyData.scheduledActionsForRule | Where-Object {
                            $_.scheduledActionConfigurations | Where-Object { $_.actionType -eq "block" }
                        }
                        if (-not $hasBlock) {
                            $blockConfig = [PSCustomObject]@{ actionType = "block"; gracePeriodHours = 0; notificationTemplateId = ""; notificationMessageCCList = @() }
                            $policyData.scheduledActionsForRule[0].scheduledActionConfigurations = @($blockConfig) + @($policyData.scheduledActionsForRule[0].scheduledActionConfigurations)
                        }
                    } else {
                        $defaultAction = [PSCustomObject]@{
                            ruleName = "PasswordRequired"
                            scheduledActionConfigurations = @(
                                [PSCustomObject]@{ actionType = "block"; gracePeriodHours = 0; notificationTemplateId = ""; notificationMessageCCList = @() }
                            )
                        }
                        $policyData | Add-Member -NotePropertyName "scheduledActionsForRule" -NotePropertyValue @($defaultAction) -Force
                    }
                }
                "SettingsCatalogPolicy" {
                    $apiUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
                }
                "ConfigurationPolicy" {
                    # Legacy name compatibility
                    $apiUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
                }
                "PowerShellScript" {
                    $apiUri = "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts"
                    # Scripts need special handling - remove script content hash
                    $policyData.PSObject.Properties.Remove('scriptContentMd5Hash')
                    $policyData.PSObject.Properties.Remove('scriptContentSha256Hash')
                }
                "RemediationScript" {
                    $apiUri = "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts"
                    # Remove hash properties
                    $policyData.PSObject.Properties.Remove('detectionScriptMd5Hash')
                    $policyData.PSObject.Properties.Remove('detectionScriptSha256Hash')
                    $policyData.PSObject.Properties.Remove('remediationScriptMd5Hash')
                    $policyData.PSObject.Properties.Remove('remediationScriptSha256Hash')
                }
                "AutopilotProfile" {
                    $apiUri = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles"
                }
                "EnrollmentStatusPage" {
                    $apiUri = "https://graph.microsoft.com/v1.0/deviceManagement/deviceEnrollmentConfigurations"
                }
                "AssignmentFilter" {
                    $apiUri = "https://graph.microsoft.com/beta/deviceManagement/assignmentFilters"
                }
                "AdministrativeTemplate" {
                    $apiUri = "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations"
                    # Remove definition values - they need to be created separately
                    $policyData.PSObject.Properties.Remove('definitionValues')
                }
                "AppProtectionPolicy" {
                    # App Protection policies have multiple endpoints based on platform
                    $odataType = $policyData.'@odata.type'
                    if ($odataType -like "*iosManagedAppProtection*") {
                        $apiUri = "https://graph.microsoft.com/beta/deviceAppManagement/iosManagedAppProtections"
                    }
                    elseif ($odataType -like "*androidManagedAppProtection*") {
                        $apiUri = "https://graph.microsoft.com/beta/deviceAppManagement/androidManagedAppProtections"
                    }
                    elseif ($odataType -like "*windowsManagedAppProtection*") {
                        $apiUri = "https://graph.microsoft.com/beta/deviceAppManagement/windowsManagedAppProtections"
                    }
                    else {
                        $apiUri = "https://graph.microsoft.com/beta/deviceAppManagement/managedAppPolicies"
                    }
                    # Remove apps array - needs to be assigned separately
                    $policyData.PSObject.Properties.Remove('apps')
                }
                "AppConfigManagedApp" {
                    $apiUri = "https://graph.microsoft.com/beta/deviceAppManagement/targetedManagedAppConfigurations"
                    $policyData.PSObject.Properties.Remove('apps')
                }
                "MacOSScript" {
                    $apiUri = "https://graph.microsoft.com/beta/deviceManagement/deviceShellScripts"
                    $policyData.PSObject.Properties.Remove('scriptContentMd5Hash')
                    $policyData.PSObject.Properties.Remove('scriptContentSha256Hash')
                }
                "MacOSCustomAttribute" {
                    $apiUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCustomAttributeShellScripts"
                    $policyData.PSObject.Properties.Remove('scriptContentMd5Hash')
                    $policyData.PSObject.Properties.Remove('scriptContentSha256Hash')
                }
                default {
                    Write-ColoredOutput " [SKIPPED - Unsupported type: $($policy.Type)]" -Color $script:Colors.Warning
                    $restoreStats.Failed++
                    continue
                }
            }

            if ($apiUri) {
                # Remove additional read-only properties that may cause issues
                $policyData.PSObject.Properties.Remove('isAssigned')
                $policyData.PSObject.Properties.Remove('roleScopeTagIds')
                $policyData.PSObject.Properties.Remove('supportsScopeTags')
                $policyData.PSObject.Properties.Remove('creationSource')

                $body = $policyData | ConvertTo-Json -Depth 100
                $result = Invoke-MgGraphRequest -Uri $apiUri -Method POST -Body $body -ContentType "application/json"
                Write-ColoredOutput " [SUCCESS]" -Color $script:Colors.Success
                $restoreStats.Success++
            }
        }
        catch {
            Write-ColoredOutput " [FAILED: $_]" -Color $script:Colors.Error
            Write-Log "Restore failed for $($policy.Name): $_" -Level Error
            $restoreStats.Failed++
        }
    }

    # Summary
    Write-ColoredOutput "`n════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
    if ($Preview) {
        Write-ColoredOutput "Preview completed!" -Color $script:Colors.Success
        Write-ColoredOutput "Policies that would be created: $($restoreStats.Success)" -Color $script:Colors.Success
    }
    else {
        Write-ColoredOutput "Restore completed!" -Color $script:Colors.Success
        Write-ColoredOutput "Successfully restored: $($restoreStats.Success)" -Color $script:Colors.Success
        Write-ColoredOutput "Failed: $($restoreStats.Failed)" -Color $script:Colors.Error
    }
    Write-ColoredOutput "════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
    continue
    } # end while menu loop
}

function Flatten-Object {
    param ($Obj, [string]$Prefix = '')
    $result = [ordered]@{}
    if ($null -eq $Obj) { return $result }
    if ($Obj -is [array]) {
        for ($i = 0; $i -lt $Obj.Count; $i++) {
            $childPrefix = if ($Prefix) { "$Prefix[$i]" } else { "[$i]" }
            $childResult = Flatten-Object -Obj $Obj[$i] -Prefix $childPrefix
            foreach ($key in $childResult.Keys) { $result[$key] = $childResult[$key] }
        }
    }
    elseif ($Obj -is [PSCustomObject]) {
        foreach ($prop in $Obj.PSObject.Properties) {
            $childPrefix = if ($Prefix) { "$Prefix.$($prop.Name)" } else { $prop.Name }
            if ($prop.Value -is [PSCustomObject] -or $prop.Value -is [array]) {
                $childResult = Flatten-Object -Obj $prop.Value -Prefix $childPrefix
                foreach ($key in $childResult.Keys) { $result[$key] = $childResult[$key] }
            }
            else {
                $result[$childPrefix] = $prop.Value
            }
        }
    }
    else {
        if ($Prefix) { $result[$Prefix] = $Obj }
    }
    return $result
}

function Extract-SettingInstanceValues {
    param (
        [object]$Instance,
        [hashtable]$Result
    )
    if (-not $Instance) { return }
    $defId = $Instance.settingDefinitionId
    if (-not $defId) { return }

    $odataType = $Instance.'@odata.type'
    $children = @()

    switch -Wildcard ($odataType) {
        '*choiceSettingInstance' {
            $Result[$defId] = $Instance.choiceSettingValue.value
            if ($Instance.choiceSettingValue.children) {
                $children = @($Instance.choiceSettingValue.children)
            }
        }
        '*simpleSettingInstance' {
            $val = $Instance.simpleSettingValue.value
            $Result[$defId] = if ($null -ne $val) { "$val" } else { $null }
        }
        '*choiceSettingCollectionInstance' {
            $vals = @($Instance.choiceSettingCollectionValue | ForEach-Object { $_.value })
            $Result[$defId] = $vals -join '||'
            foreach ($csv in $Instance.choiceSettingCollectionValue) {
                if ($csv.children) { $children += @($csv.children) }
            }
        }
        '*simpleSettingCollectionInstance' {
            $vals = @($Instance.simpleSettingCollectionValue | ForEach-Object { "$($_.value)" })
            $Result[$defId] = $vals -join '||'
        }
        '*groupSettingInstance' {
            if ($Instance.groupSettingValue.children) {
                $children = @($Instance.groupSettingValue.children)
            }
        }
        '*groupSettingCollectionInstance' {
            foreach ($gsv in $Instance.groupSettingCollectionValue) {
                if ($gsv.children) { $children += @($gsv.children) }
            }
        }
    }

    foreach ($child in $children) {
        if ($child) {
            Extract-SettingInstanceValues -Instance $child -Result $Result
        }
    }
}

function Resolve-SettingDisplayValue {
    param (
        [object]$RawValue,
        [hashtable]$OptionsLookup
    )
    if ($null -eq $RawValue) { return '(not configured)' }

    $strVal = "$RawValue"

    # Handle pipe-separated collection values
    if ($strVal -match '\|\|') {
        $parts = $strVal -split '\|\|'
        $resolved = foreach ($part in $parts) {
            if ($OptionsLookup -and $OptionsLookup.ContainsKey($part)) { $OptionsLookup[$part] } else { $part }
        }
        return ($resolved -join ', ')
    }

    # Single value - check options lookup
    if ($OptionsLookup -and $OptionsLookup.ContainsKey($strVal)) { return $OptionsLookup[$strVal] }

    return $strVal
}

function Resolve-SettingsCatalogDiffs {
    param (
        [object]$BackupPolicy,
        [object]$CurrentPolicy
    )

    $backupSettings = $BackupPolicy.settings
    $currentSettings = $CurrentPolicy.settings

    if (-not $backupSettings) { $backupSettings = @() }
    if (-not $currentSettings) { $currentSettings = @() }

    # Extract settingDefinitionId -> rawValue maps
    $backupValues = @{}
    $currentValues = @{}

    foreach ($s in $backupSettings) {
        if ($s.settingInstance) {
            Extract-SettingInstanceValues -Instance $s.settingInstance -Result $backupValues
        }
    }
    foreach ($s in $currentSettings) {
        if ($s.settingInstance) {
            Extract-SettingInstanceValues -Instance $s.settingInstance -Result $currentValues
        }
    }

    # Find definition IDs where values differ
    $allDefIds = @($backupValues.Keys) + @($currentValues.Keys) | Select-Object -Unique
    $changedDefIds = @()
    foreach ($defId in $allDefIds) {
        $bv = if ($backupValues.ContainsKey($defId)) { "$($backupValues[$defId])" } else { '__MISSING__' }
        $cv = if ($currentValues.ContainsKey($defId)) { "$($currentValues[$defId])" } else { '__MISSING__' }
        if ($bv -ne $cv) { $changedDefIds += $defId }
    }

    if ($changedDefIds.Count -eq 0) { return @() }

    # Fetch setting definitions from Graph API for friendly names
    Write-ColoredOutput "  Resolving $($changedDefIds.Count) setting definition(s) from Intune..." -Color $script:Colors.Info
    $definitions = @{}
    foreach ($defId in $changedDefIds) {
        try {
            $encodedId = [System.Uri]::EscapeDataString($defId)
            $uri = "https://graph.microsoft.com/beta/deviceManagement/configurationSettings('$encodedId')"
            $def = Invoke-MgGraphRequest -Uri $uri -Method GET
            $definitions[$defId] = $def
        }
        catch {
            $definitions[$defId] = $null
        }
    }

    # Build display diffs
    $diffs = @()
    foreach ($defId in $changedDefIds) {
        $def = $definitions[$defId]
        $displayName = if ($def -and $def.displayName) { $def.displayName } else { $defId }
        $description = if ($def -and $def.description) { $def.description } else { '' }

        # Build options lookup for choice values
        $optionsLookup = @{}
        if ($def -and $def.options) {
            foreach ($opt in $def.options) {
                if ($opt.itemId -and $opt.displayName) {
                    $optionsLookup[$opt.itemId] = $opt.displayName
                }
            }
        }

        $bRaw = if ($backupValues.ContainsKey($defId)) { $backupValues[$defId] } else { $null }
        $cRaw = if ($currentValues.ContainsKey($defId)) { $currentValues[$defId] } else { $null }

        $backupDisplay = Resolve-SettingDisplayValue -RawValue $bRaw -OptionsLookup $optionsLookup
        $currentDisplay = Resolve-SettingDisplayValue -RawValue $cRaw -OptionsLookup $optionsLookup

        $diffs += [PSCustomObject]@{
            Name        = $displayName
            BackupValue = $backupDisplay
            TenantValue = $currentDisplay
            Description = $description
        }
    }

    return $diffs
}

function Export-DriftReport {
    param (
        [Parameter(Mandatory = $true)]
        [object]$DriftResults,

        [Parameter(Mandatory = $true)]
        [ValidateSet('HTML', 'Markdown')]
        [string]$Format,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    $reportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $changedCount = if ($DriftResults.Changed) { $DriftResults.Changed.Count } else { 0 }
    $renamedCount = if ($DriftResults.Renamed) { $DriftResults.Renamed.Count } else { 0 }
    $removedCount = if ($DriftResults.Removed) { $DriftResults.Removed.Count } else { 0 }
    $addedCount = if ($DriftResults.Added) { $DriftResults.Added.Count } else { 0 }
    $prevRestoredCount = if ($DriftResults.PrevRestored) { $DriftResults.PrevRestored.Count } else { 0 }

    # Helper to format a value for display
    function Format-Value {
        param ($Value)
        if ($null -eq $Value) { return '(not set)' }
        if ($Value -is [string] -or $Value -is [bool] -or $Value -is [int] -or $Value -is [long] -or $Value -is [double]) {
            return "$Value"
        }
        try {
            return ($Value | ConvertTo-Json -Depth 5 -Compress)
        }
        catch {
            return "$Value"
        }
    }

    if ($Format -eq 'HTML') {
        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>UniFy-Endpoint Drift Report - $reportDate</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 0;
            background: linear-gradient(135deg, #0078d4 0%, #005a9e 100%);
            min-height: 100vh;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }
        .header {
            background: white;
            border-radius: 10px;
            padding: 30px;
            margin-bottom: 30px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        }
        .header h1 {
            color: #333;
            margin: 0 0 10px 0;
            font-size: 2.5em;
        }
        .header .subtitle {
            color: #666;
            font-size: 1.1em;
        }
        .metadata {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 20px;
            margin-top: 20px;
        }
        .metadata-item {
            background: #f8f9fa;
            padding: 15px;
            border-radius: 8px;
            border-left: 4px solid #0078d4;
        }
        .metadata-item .label {
            color: #666;
            font-size: 0.9em;
            margin-bottom: 5px;
        }
        .metadata-item .value {
            color: #333;
            font-size: 1.3em;
            font-weight: bold;
        }
        .policy-section {
            background: white;
            border-radius: 10px;
            padding: 25px;
            margin-bottom: 20px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.08);
        }
        .policy-section h2 {
            color: #333;
            margin: 0 0 5px 0;
            padding-bottom: 10px;
            border-bottom: 2px solid #f0f0f0;
        }
        .policy-section h3 {
            color: #444;
            margin: 20px 0 5px 0;
            font-size: 1.1em;
        }
        .policy-section h3:first-of-type {
            margin-top: 15px;
        }
        .policy-section .policy-meta {
            color: #666;
            font-size: 0.9em;
            margin-bottom: 15px;
        }
        .diff-table {
            width: 100%;
            border-collapse: collapse;
        }
        .diff-table th {
            background: #0078d4;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: 500;
        }
        .diff-table td {
            padding: 12px;
            border-bottom: 1px solid #f0f0f0;
            word-break: break-word;
            max-width: 400px;
        }
        .diff-table tr:hover {
            background: #f8f9fa;
        }
        .val-backup {
            background: #e8f5e9;
        }
        .val-tenant {
            background: #e3f2fd;
        }
        .status-changed {
            border-left: 4px solid #ff9800;
        }
        .status-renamed {
            border-left: 4px solid #2196f3;
        }
        .status-removed {
            border-left: 4px solid #f44336;
        }
        .status-added {
            border-left: 4px solid #4caf50;
        }
        .badge {
            display: inline-block;
            padding: 3px 10px;
            border-radius: 12px;
            font-size: 0.8em;
            font-weight: bold;
            color: white;
        }
        .badge-changed { background: #ff9800; }
        .badge-renamed { background: #2196f3; }
        .badge-removed { background: #f44336; }
        .badge-added { background: #4caf50; }
        .footer {
            text-align: center;
            color: rgba(255,255,255,0.7);
            padding: 20px;
            font-size: 0.9em;
        }
        code {
            background: #f5f5f5;
            padding: 2px 6px;
            border-radius: 3px;
            font-size: 0.9em;
        }
        .desc-cell {
            max-width: 300px;
            font-size: 0.85em;
            color: #555;
            cursor: pointer;
        }
        .desc-cell .desc-short { display: inline; }
        .desc-cell .desc-full { display: none; }
        .desc-cell.expanded .desc-short { display: none; }
        .desc-cell.expanded .desc-full { display: inline; }
        .desc-toggle {
            color: #0078d4;
            font-size: 0.8em;
            margin-left: 4px;
        }
    </style>
    <script>
        function toggleDesc(el) { el.closest('.desc-cell').classList.toggle('expanded'); }
    </script>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>UniFy-Endpoint Drift Report</h1>
        <div class="subtitle">Configuration drift analysis between backup and Intune tenant</div>
        <div class="metadata">
            <div class="metadata-item">
                <div class="label">Report Date</div>
                <div class="value">$reportDate</div>
            </div>
            <div class="metadata-item">
                <div class="label">Changed</div>
                <div class="value">$changedCount</div>
            </div>
            <div class="metadata-item">
                <div class="label">Renamed</div>
                <div class="value">$renamedCount</div>
            </div>
            <div class="metadata-item">
                <div class="label">Removed</div>
                <div class="value">$removedCount</div>
            </div>
            <div class="metadata-item">
                <div class="label">Added</div>
                <div class="value">$addedCount</div>
            </div>
            <div class="metadata-item">
                <div class="label">Prev. Restored</div>
                <div class="value">$prevRestoredCount</div>
            </div>
        </div>
    </div>
"@

        # Changed policies with diff tables (only changed settings)
        if ($changedCount -gt 0) {
            $html += @"

    <div class="policy-section status-changed">
        <h2>Changed Policies ($changedCount)</h2>
"@
            foreach ($policy in $DriftResults.Changed) {
                $policyName = [System.Web.HttpUtility]::HtmlEncode($policy.Name)
                $policyType = [System.Web.HttpUtility]::HtmlEncode($policy.Type)
                $tenantName = if ($policy.TenantName) { [System.Web.HttpUtility]::HtmlEncode($policy.TenantName) } else { $policyName }

                # Detect Settings Catalog policies for friendly name resolution
                $isSettingsCatalog = ($policy.Type -in @('SettingsCatalogPolicy', 'ConfigurationPolicy')) -and
                                     $policy.BackupData.settings -is [array] -and
                                     $policy.BackupData.settings.Count -gt 0

                # Build unified diff list: Name, BackupValue, TenantValue, Description
                $allDiffs = @()

                # For Settings Catalog: resolve friendly names via Graph API
                if ($isSettingsCatalog -and 'settings' -in $policy.Differences) {
                    $catalogDiffs = Resolve-SettingsCatalogDiffs -BackupPolicy $policy.BackupData -CurrentPolicy $policy.CurrentData
                    $allDiffs += $catalogDiffs
                }

                # For non-settings properties (or non-Settings Catalog), use Flatten-Object
                $propsToFlatten = if ($isSettingsCatalog) {
                    @($policy.Differences | Where-Object { $_ -ne 'settings' })
                } else {
                    $policy.Differences
                }

                if ($propsToFlatten.Count -gt 0) {
                    $allIgnore = $script:ComparisonIgnoreProperties
                    $cleanBackup = Remove-IgnoredProperties -obj ($policy.BackupData | ConvertTo-Json -Depth 50 -Compress | ConvertFrom-Json) -ignoreList $allIgnore
                    $cleanTenant = Remove-IgnoredProperties -obj ($policy.CurrentData | ConvertTo-Json -Depth 50 -Compress | ConvertFrom-Json) -ignoreList $allIgnore

                    foreach ($propName in $propsToFlatten) {
                        $flatBackup = Flatten-Object -Obj $cleanBackup.$propName -Prefix $propName
                        $flatTenant = Flatten-Object -Obj $cleanTenant.$propName -Prefix $propName
                        $allKeys = @($flatBackup.Keys) + @($flatTenant.Keys) | Select-Object -Unique | Sort-Object
                        foreach ($key in $allKeys) {
                            $bStr = if ($null -eq $flatBackup[$key]) { '(not set)' } else { "$($flatBackup[$key])" }
                            $tStr = if ($null -eq $flatTenant[$key]) { '(not set)' } else { "$($flatTenant[$key])" }
                            if ($bStr -ne $tStr) {
                                $allDiffs += [PSCustomObject]@{ Name = $key; BackupValue = $flatBackup[$key]; TenantValue = $flatTenant[$key]; Description = '' }
                            }
                        }
                    }
                }

                $diffCount = $allDiffs.Count
                $backupCreated   = ConvertFrom-JsonDate -DateString $policy.BackupData.createdDateTime
                $currentModified = ConvertFrom-JsonDate -DateString $policy.CurrentData.lastModifiedDateTime
                $driftAssigned   = if ($null -ne $policy.CurrentData.isAssigned) { if ($policy.CurrentData.isAssigned) { "Yes" } else { "No" } } else { "N/A" }

                $html += @"
        <h3>$policyName</h3>
        <div class="policy-meta">Type: $policyType | Created: $backupCreated | Last Modified: $currentModified | Assigned: $driftAssigned | Differences: $diffCount setting(s)$(if ($policy.NameChanged) { " | Also renamed to: $tenantName" })</div>
        <table class="diff-table">
            <tr><th>Setting</th><th>Backup Value</th><th>Tenant Value</th><th>Description</th></tr>
"@

                foreach ($diff in $allDiffs) {
                    $backupVal = Format-Value -Value $diff.BackupValue
                    $tenantVal = Format-Value -Value $diff.TenantValue
                    $encodedName = [System.Web.HttpUtility]::HtmlEncode($diff.Name)
                    $encodedBackup = [System.Web.HttpUtility]::HtmlEncode($backupVal)
                    $encodedTenant = [System.Web.HttpUtility]::HtmlEncode($tenantVal)
                    $fullDesc = [System.Web.HttpUtility]::HtmlEncode($diff.Description)

                    $descHtml = if ($diff.Description.Length -gt 150) {
                        $shortDesc = [System.Web.HttpUtility]::HtmlEncode($diff.Description.Substring(0, 150) + '...')
                        "<span class=`"desc-short`">$shortDesc <span class=`"desc-toggle`" onclick=`"toggleDesc(this)`">[show more]</span></span><span class=`"desc-full`">$fullDesc <span class=`"desc-toggle`" onclick=`"toggleDesc(this)`">[show less]</span></span>"
                    } elseif ($diff.Description) {
                        $fullDesc
                    } else { '-' }

                    $html += "            <tr><td>$encodedName</td><td class=`"val-backup`">$encodedBackup</td><td class=`"val-tenant`">$encodedTenant</td><td class=`"desc-cell`">$descHtml</td></tr>`n"
                }

                $html += @"
        </table>
"@
            }

            $html += @"
    </div>
"@
        }

        # Renamed policies
        if ($renamedCount -gt 0) {
            $html += @"

    <div class="policy-section status-renamed">
        <h2>Renamed Policies ($renamedCount)</h2>
        <table class="diff-table">
            <tr><th>Policy Type</th><th>Backup Name</th><th>Tenant Name</th><th>Created</th><th>Last Modified</th><th>Assigned</th></tr>
"@
            foreach ($policy in $DriftResults.Renamed) {
                $encodedName = [System.Web.HttpUtility]::HtmlEncode($policy.Name)
                $encodedType = [System.Web.HttpUtility]::HtmlEncode($policy.Type)
                $encodedTenant = [System.Web.HttpUtility]::HtmlEncode($policy.TenantName)
                $rCreated  = ConvertFrom-JsonDate -DateString $policy.BackupData.createdDateTime
                $rModified = ConvertFrom-JsonDate -DateString $policy.CurrentData.lastModifiedDateTime
                $rAssigned = if ($null -ne $policy.CurrentData.isAssigned) { if ($policy.CurrentData.isAssigned) { "Yes" } else { "No" } } else { "N/A" }
                $html += "            <tr><td>$encodedType</td><td class=`"val-backup`">$encodedName</td><td class=`"val-tenant`">$encodedTenant</td><td>$rCreated</td><td>$rModified</td><td>$rAssigned</td></tr>`n"
            }
            $html += @"
        </table>
    </div>
"@
        }

        # Removed policies
        if ($removedCount -gt 0) {
            $html += @"

    <div class="policy-section status-removed">
        <h2>Removed Policies ($removedCount)</h2>
        <table class="diff-table">
            <tr><th>Policy Type</th><th>Backup Name</th><th>Status</th><th>Created</th><th>Last Modified</th><th>Assigned</th></tr>
"@
            foreach ($policy in $DriftResults.Removed) {
                $encodedName = [System.Web.HttpUtility]::HtmlEncode($policy.Name)
                $encodedType = [System.Web.HttpUtility]::HtmlEncode($policy.Type)
                $remCreated  = ConvertFrom-JsonDate -DateString $policy.BackupData.createdDateTime
                $remModified = ConvertFrom-JsonDate -DateString $policy.BackupData.lastModifiedDateTime
                $remAssigned = if ($null -ne $policy.BackupData.isAssigned) { if ($policy.BackupData.isAssigned) { "Yes" } else { "No" } } else { "N/A" }
                $html += "            <tr><td>$encodedType</td><td class=`"val-backup`">$encodedName</td><td class=`"val-tenant`">Removed</td><td>$remCreated</td><td>$remModified</td><td>$remAssigned</td></tr>`n"
            }
            $html += @"
        </table>
    </div>
"@
        }

        # Added policies
        if ($addedCount -gt 0) {
            $html += @"

    <div class="policy-section status-added">
        <h2>Added Policies ($addedCount)</h2>
        <table class="diff-table">
            <tr><th>Policy Type</th><th>Policy Name</th><th>Status</th><th>Created</th><th>Last Modified</th><th>Assigned</th></tr>
"@
            foreach ($policy in $DriftResults.Added) {
                $encodedName = [System.Web.HttpUtility]::HtmlEncode($policy.Name)
                $encodedType = [System.Web.HttpUtility]::HtmlEncode($policy.Type)
                $addCreated  = ConvertFrom-JsonDate -DateString $policy.CurrentData.createdDateTime
                $addModified = ConvertFrom-JsonDate -DateString $policy.CurrentData.lastModifiedDateTime
                $addAssigned = if ($null -ne $policy.CurrentData.isAssigned) { if ($policy.CurrentData.isAssigned) { "Yes" } else { "No" } } else { "N/A" }
                $html += "            <tr><td>$encodedType</td><td class=`"val-backup`">$encodedName</td><td class=`"val-tenant`">Added after backup</td><td>$addCreated</td><td>$addModified</td><td>$addAssigned</td></tr>`n"
            }
            $html += @"
        </table>
    </div>
"@
        }

        if ($prevRestoredCount -gt 0) {
            $html += @"

    <div class="policy-section status-added">
        <h2>Previously Restored Policies ($prevRestoredCount)</h2>
        <table class="diff-table">
            <tr><th>Policy Type</th><th>Backup Name</th><th>Tenant Name</th><th>Created</th><th>Last Modified</th><th>Assigned</th></tr>
"@
            foreach ($policy in $DriftResults.PrevRestored) {
                $encodedName = [System.Web.HttpUtility]::HtmlEncode($policy.Name)
                $encodedTenantName = [System.Web.HttpUtility]::HtmlEncode($policy.TenantName)
                $encodedType = [System.Web.HttpUtility]::HtmlEncode($policy.Type)
                $prCreated  = ConvertFrom-JsonDate -DateString $policy.BackupData.createdDateTime
                $prModified = ConvertFrom-JsonDate -DateString $policy.CurrentData.lastModifiedDateTime
                $prAssigned = if ($null -ne $policy.CurrentData.isAssigned) { if ($policy.CurrentData.isAssigned) { "Yes" } else { "No" } } else { "N/A" }
                $html += "            <tr><td>$encodedType</td><td class=`"val-backup`">$encodedName</td><td class=`"val-tenant`">$encodedTenantName</td><td>$prCreated</td><td>$prModified</td><td>$prAssigned</td></tr>`n"
            }
            $html += @"
        </table>
    </div>
"@
        }

        $html += @"

    <div class="footer">
        Generated by UniFy-Endpoint v2.0 | $reportDate
    </div>
</div>
</body>
</html>
"@

        $html | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
    }
    else {
        # Markdown format
        $md = @"
# UniFy-Endpoint Drift Report

**Report Date:** $reportDate
**Changed:** $changedCount | **Renamed:** $renamedCount | **Removed:** $removedCount | **Added:** $addedCount | **Prev. Restored:** $prevRestoredCount

---

"@

        # Changed policies (only changed settings)
        if ($changedCount -gt 0) {
            $md += "## Changed Policies ($changedCount)`n`n"

            foreach ($policy in $DriftResults.Changed) {
                # Detect Settings Catalog policies
                $isSettingsCatalog = ($policy.Type -in @('SettingsCatalogPolicy', 'ConfigurationPolicy')) -and
                                     $policy.BackupData.settings -is [array] -and
                                     $policy.BackupData.settings.Count -gt 0

                # Build unified diff list
                $allDiffs = @()

                if ($isSettingsCatalog -and 'settings' -in $policy.Differences) {
                    $catalogDiffs = Resolve-SettingsCatalogDiffs -BackupPolicy $policy.BackupData -CurrentPolicy $policy.CurrentData
                    $allDiffs += $catalogDiffs
                }

                $propsToFlatten = if ($isSettingsCatalog) {
                    @($policy.Differences | Where-Object { $_ -ne 'settings' })
                } else {
                    $policy.Differences
                }

                if ($propsToFlatten.Count -gt 0) {
                    $allIgnore = $script:ComparisonIgnoreProperties
                    $cleanBackup = Remove-IgnoredProperties -obj ($policy.BackupData | ConvertTo-Json -Depth 50 -Compress | ConvertFrom-Json) -ignoreList $allIgnore
                    $cleanTenant = Remove-IgnoredProperties -obj ($policy.CurrentData | ConvertTo-Json -Depth 50 -Compress | ConvertFrom-Json) -ignoreList $allIgnore

                    foreach ($propName in $propsToFlatten) {
                        $flatBackup = Flatten-Object -Obj $cleanBackup.$propName -Prefix $propName
                        $flatTenant = Flatten-Object -Obj $cleanTenant.$propName -Prefix $propName
                        $allKeys = @($flatBackup.Keys) + @($flatTenant.Keys) | Select-Object -Unique | Sort-Object
                        foreach ($key in $allKeys) {
                            $bStr = if ($null -eq $flatBackup[$key]) { '(not set)' } else { "$($flatBackup[$key])" }
                            $tStr = if ($null -eq $flatTenant[$key]) { '(not set)' } else { "$($flatTenant[$key])" }
                            if ($bStr -ne $tStr) {
                                $allDiffs += [PSCustomObject]@{ Name = $key; BackupValue = $flatBackup[$key]; TenantValue = $flatTenant[$key]; Description = '' }
                            }
                        }
                    }
                }

                $mdCreated   = ConvertFrom-JsonDate -DateString $policy.BackupData.createdDateTime
                $mdModified  = ConvertFrom-JsonDate -DateString $policy.CurrentData.lastModifiedDateTime
                $mdAssigned  = if ($null -ne $policy.CurrentData.isAssigned) { if ($policy.CurrentData.isAssigned) { "Yes" } else { "No" } } else { "N/A" }
                $md += "### $($policy.Name) ($($policy.Type))`n"
                $md += "**Created:** $mdCreated | **Last Modified:** $mdModified | **Assigned:** $mdAssigned | **Differences:** $($allDiffs.Count) setting(s)"
                if ($policy.NameChanged) { $md += " | **Also renamed to:** $($policy.TenantName)" }
                $md += "`n`n"

                $md += "| Setting | Backup Value | Tenant Value | Description |`n"
                $md += "|---------|-------------|-------------|-------------|`n"

                foreach ($diff in $allDiffs) {
                    $backupVal = (Format-Value -Value $diff.BackupValue) -replace '\|', '\|'
                    $tenantVal = (Format-Value -Value $diff.TenantValue) -replace '\|', '\|'
                    $descVal = if ($diff.Description) {
                        $shortDesc = if ($diff.Description.Length -gt 80) { $diff.Description.Substring(0, 80) + '...' } else { $diff.Description }
                        ($shortDesc -replace '\|', '\|' -replace "`n", ' ' -replace "`r", '')
                    } else { '-' }
                    $md += "| ``$($diff.Name)`` | ``$backupVal`` | ``$tenantVal`` | $descVal |`n"
                }
                $md += "`n"
            }
        }

        # Renamed policies
        if ($renamedCount -gt 0) {
            $md += "## Renamed Policies ($renamedCount)`n`n"
            $md += "| Policy Type | Backup Name | Tenant Name | Created | Last Modified | Assigned |`n"
            $md += "|-------------|-------------|-------------|---------|--------------|----------|`n"
            foreach ($policy in $DriftResults.Renamed) {
                $bName = ($policy.Name) -replace '\|', '\|'
                $tName = ($policy.TenantName) -replace '\|', '\|'
                $rnCreated  = ConvertFrom-JsonDate -DateString $policy.BackupData.createdDateTime
                $rnModified = ConvertFrom-JsonDate -DateString $policy.CurrentData.lastModifiedDateTime
                $rnAssigned = if ($null -ne $policy.CurrentData.isAssigned) { if ($policy.CurrentData.isAssigned) { "Yes" } else { "No" } } else { "N/A" }
                $md += "| $($policy.Type) | $bName | $tName | $rnCreated | $rnModified | $rnAssigned |`n"
            }
            $md += "`n"
        }

        # Removed policies
        if ($removedCount -gt 0) {
            $md += "## Removed Policies ($removedCount)`n`n"
            $md += "| Policy Type | Backup Name | Status | Created | Last Modified | Assigned |`n"
            $md += "|-------------|-------------|--------|---------|--------------|----------|`n"
            foreach ($policy in $DriftResults.Removed) {
                $bName = ($policy.Name) -replace '\|', '\|'
                $rmCreated  = ConvertFrom-JsonDate -DateString $policy.BackupData.createdDateTime
                $rmModified = ConvertFrom-JsonDate -DateString $policy.BackupData.lastModifiedDateTime
                $rmAssigned = if ($null -ne $policy.BackupData.isAssigned) { if ($policy.BackupData.isAssigned) { "Yes" } else { "No" } } else { "N/A" }
                $md += "| $($policy.Type) | $bName | **Removed** | $rmCreated | $rmModified | $rmAssigned |`n"
            }
            $md += "`n"
        }

        # Added policies
        if ($addedCount -gt 0) {
            $md += "## Added Policies ($addedCount)`n`n"
            $md += "| Policy Type | Policy Name | Status | Created | Last Modified | Assigned |`n"
            $md += "|-------------|-------------|--------|---------|--------------|----------|`n"
            foreach ($policy in $DriftResults.Added) {
                $pName = ($policy.Name) -replace '\|', '\|'
                $adCreated  = ConvertFrom-JsonDate -DateString $policy.CurrentData.createdDateTime
                $adModified = ConvertFrom-JsonDate -DateString $policy.CurrentData.lastModifiedDateTime
                $adAssigned = if ($null -ne $policy.CurrentData.isAssigned) { if ($policy.CurrentData.isAssigned) { "Yes" } else { "No" } } else { "N/A" }
                $md += "| $($policy.Type) | $pName | **Added after backup** | $adCreated | $adModified | $adAssigned |`n"
            }
            $md += "`n"
        }

        # Previously restored policies
        if ($prevRestoredCount -gt 0) {
            $md += "## Previously Restored Policies ($prevRestoredCount)`n`n"
            $md += "| Policy Type | Backup Name | Tenant Name | Created | Last Modified | Assigned |`n"
            $md += "|-------------|-------------|-------------|---------|--------------|----------|`n"
            foreach ($policy in $DriftResults.PrevRestored) {
                $bName = ($policy.Name) -replace '\|', '\|'
                $tName = ($policy.TenantName) -replace '\|', '\|'
                $psCreated  = ConvertFrom-JsonDate -DateString $policy.BackupData.createdDateTime
                $psModified = ConvertFrom-JsonDate -DateString $policy.CurrentData.lastModifiedDateTime
                $psAssigned = if ($null -ne $policy.CurrentData.isAssigned) { if ($policy.CurrentData.isAssigned) { "Yes" } else { "No" } } else { "N/A" }
                $md += "| $($policy.Type) | $bName | $tName | $psCreated | $psModified | $psAssigned |`n"
            }
            $md += "`n"
        }

        $md += "---`n*Generated by UniFy-Endpoint v2.0 | $reportDate*`n"

        $md | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
    }
}

#endregion

function Build-SettingsAnalysis {
    param ([array]$AllSections)

    # Policy types included in analysis — Device Configurations and Compliance Policies
    # are excluded by design (too noisy / not meaningful to cross-compare).
    $allIncluded = @(
        'Settings Catalog Policies',
        'App Protection Policies',
        'App Config (Managed Devices)',
        'App Config (Managed Apps)'
    )
    $notConfigured = @('(not set)', '(not configured)', '')

    # Generic sub-property names that appear across many unrelated nested settings
    # (e.g. a firewall rule's "Name" and an OneDrive folder's "Name" are completely
    # different things). Including them produces false conflicts, so they are skipped.
    $genericNames = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]@('Name', 'Description', 'Id', 'RuleId', 'DisplayName'),
        [System.StringComparer]::OrdinalIgnoreCase
    )

    # rawByPlatform[platform]["TypeDisplay|||SettingName"] → list of entries
    $rawByPlatform = @{}

    foreach ($section in $AllSections) {
        if ($section.DisplayName -notin $allIncluded) { continue }

        foreach ($card in $section.Cards) {
            $platform = $card.Platform
            if (-not $rawByPlatform.ContainsKey($platform)) { $rawByPlatform[$platform] = @{} }

            $cardGroups = if ($card.Groups) { @($card.Groups) } else { @() }

            foreach ($setting in $card.Settings) {
                $val = if ($null -eq $setting.Value) { '' } else { "$($setting.Value)".Trim() }
                if ($val -in $notConfigured -or [string]::IsNullOrEmpty($val)) { continue }

                $sName = "$($setting.Name)".Trim()
                if ([string]::IsNullOrEmpty($sName)) { continue }
                if ($genericNames.Contains($sName)) { continue }  # skip structural sub-property names

                $compoundKey = "$($section.DisplayName)|||$sName"
                if (-not $rawByPlatform[$platform].ContainsKey($compoundKey)) {
                    $rawByPlatform[$platform][$compoundKey] = [System.Collections.Generic.List[object]]::new()
                }
                $rawByPlatform[$platform][$compoundKey].Add([pscustomobject]@{
                    PolicyName   = $card.Name
                    PolicyType   = $section.DisplayName
                    Value        = $val
                    Groups       = $cardGroups
                    IsUnassigned = ($cardGroups.Count -eq 0)
                })
            }
        }
    }

    $platformAnalysis = @{}
    foreach ($platform in $rawByPlatform.Keys) {
        $settingMap = @{}
        foreach ($compoundKey in $rawByPlatform[$platform].Keys) {
            $entries = $rawByPlatform[$platform][$compoundKey]
            if ($entries.Count -lt 2) { continue }

            $included        = [System.Collections.Generic.HashSet[int]]::new()
            $hasRealConflict = $false

            for ($i = 0; $i -lt $entries.Count; $i++) {
                for ($j = $i + 1; $j -lt $entries.Count; $j++) {
                    # Never compare a policy against itself (guards against duplicate setting
                    # names returned by Resolve-PolicyAllSettings for the same policy).
                    if ($entries[$i].PolicyName -eq $entries[$j].PolicyName) { continue }
                    $iUnassigned = $entries[$i].IsUnassigned
                    $jUnassigned = $entries[$j].IsUnassigned

                    if ($iUnassigned -or $jUnassigned) {
                        # At least one unassigned: include as forced duplicate — never a conflict
                        [void]$included.Add($i)
                        [void]$included.Add($j)
                    } else {
                        # Both assigned — require shared group to compare
                        $iGroups = @($entries[$i].Groups)
                        $jGroups = @($entries[$j].Groups)
                        $shared  = $iGroups | Where-Object { $_ -in $jGroups }
                        if (@($shared).Count -gt 0) {
                            [void]$included.Add($i)
                            [void]$included.Add($j)
                            if ($entries[$i].Value -ne $entries[$j].Value) {
                                $hasRealConflict = $true
                            }
                        }
                        # else: different groups — skip (no flag)
                    }
                }
            }

            if ($included.Count -lt 2) { continue }

            $comparableEntries = [System.Collections.Generic.List[object]]::new()
            foreach ($idx in ($included | Sort-Object)) { $comparableEntries.Add($entries[$idx]) }

            $status = if ($hasRealConflict) { 'Conflict' } else { 'Duplicate' }
            $settingMap[$compoundKey] = [pscustomobject]@{
                Policies        = $comparableEntries
                Status          = $status
                HasRealConflict = $hasRealConflict
                Count           = $comparableEntries.Count
            }
        }

        if ($settingMap.Count -gt 0) {
            $platformAnalysis[$platform] = [pscustomobject]@{ SettingMap = $settingMap }
        }
    }

    return $platformAnalysis
}

function Render-AnalysisSectionHTML {
    param ([hashtable]$PlatformAnalysis)

    if (-not $PlatformAnalysis -or $PlatformAnalysis.Count -eq 0) { return '' }

    $activePlatforms = @($PlatformAnalysis.Keys | Where-Object { $PlatformAnalysis[$_].SettingMap.Count -gt 0 } | Sort-Object)
    if ($activePlatforms.Count -eq 0) { return '' }

    $sb            = [System.Text.StringBuilder]::new()
    $firstPlatform = $activePlatforms[0]

    [void]$sb.AppendLine('<div class="analysis-section" id="settings-analysis">')
    [void]$sb.AppendLine('    <h2>Duplicate &amp; Conflict Settings Analysis</h2>')
    [void]$sb.AppendLine('    <p class="section-desc">Policy-centric analysis per platform. Shows policies with settings that conflict (same setting, different values) or are duplicated (same setting, same value) across multiple policies. Expand a policy row to view the specific settings involved.</p>')
    [void]$sb.AppendLine('    <div class="analysis-tabs">')
    foreach ($plat in $activePlatforms) {
        $activeClass = if ($plat -eq $firstPlatform) { ' active' } else { '' }
        $encPlat     = [System.Net.WebUtility]::HtmlEncode($plat)
        [void]$sb.AppendLine("        <button class=`"analysis-tab-btn$activeClass`" data-platform=`"$encPlat`" onclick=`"switchAnalysisTab('$encPlat')`">$encPlat</button>")
    }
    [void]$sb.AppendLine('    </div>')

    foreach ($plat in $activePlatforms) {
        $activeClass     = if ($plat -eq $firstPlatform) { ' active' } else { '' }
        $safeId          = $plat.ToLower() -replace '[^a-z0-9]', '-'
        $panelId         = "analysis-panel-$safeId"
        $encPlat         = [System.Net.WebUtility]::HtmlEncode($plat)
        $data       = $PlatformAnalysis[$plat]
        $settingMap = $data.SettingMap

        # Build policy-centric map (Conflict and Duplicate only)
        $policyMap = @{}
        foreach ($sEntry in $settingMap.GetEnumerator()) {
            # Strip the "PolicyType|||" compound key prefix to get the display setting name
            $sepIdx = $sEntry.Key.IndexOf('|||')
            $sName  = if ($sepIdx -ge 0) { $sEntry.Key.Substring($sepIdx + 3) } else { $sEntry.Key }
            $sData  = $sEntry.Value
            $status = $sData.Status
            foreach ($pol in $sData.Policies) {
                $polKey = "$($pol.PolicyName)||$($pol.PolicyType)"
                if (-not $policyMap.ContainsKey($polKey)) {
                    $policyMap[$polKey] = [pscustomobject]@{
                        Name       = $pol.PolicyName
                        Type       = ($pol.PolicyType -replace ' Policies$', '')
                        Conflicts  = [System.Collections.Generic.List[object]]::new()
                        Duplicates = [System.Collections.Generic.List[object]]::new()
                    }
                }
                $otherPols = [System.Collections.Generic.List[object]]::new()
                if ($pol.IsUnassigned -or $status -eq 'Duplicate') {
                    # Duplicate/unassigned context: show other policies with same value
                    foreach ($op in $sData.Policies) {
                        if ($op.PolicyName -ne $pol.PolicyName -and $op.Value -eq $pol.Value) { $otherPols.Add($op) }
                    }
                } else {
                    # Conflict context: show other assigned policies only (exclude unassigned)
                    foreach ($op in $sData.Policies) {
                        if ($op.PolicyName -ne $pol.PolicyName -and -not $op.IsUnassigned) { $otherPols.Add($op) }
                    }
                }
                $sItem = [pscustomobject]@{
                    SettingName   = $sName
                    OwnValue      = $pol.Value
                    OtherPolicies = $otherPols
                }
                if ($pol.IsUnassigned) {
                    $policyMap[$polKey].Duplicates.Add($sItem)
                } elseif ($status -eq 'Conflict') {
                    $policyMap[$polKey].Conflicts.Add($sItem)
                } else {
                    $policyMap[$polKey].Duplicates.Add($sItem)
                }
            }
        }

        $conflictPolicies      = @($policyMap.Values | Where-Object { $_.Conflicts.Count -gt 0 } | Sort-Object Name)
        $duplicatePolicies     = @($policyMap.Values | Where-Object { $_.Duplicates.Count -gt 0 } | Sort-Object Name)
        $duplicateOnlyPolicies = @($policyMap.Values | Where-Object { $_.Conflicts.Count -eq 0 -and $_.Duplicates.Count -gt 0 } | Sort-Object Name)
        $sortedPolicies        = @($conflictPolicies) + @($duplicateOnlyPolicies)
        $conflictSettingCount  = @($settingMap.Values | Where-Object { $_.Status -eq 'Conflict' }).Count
        $duplicateSettingCount = @($settingMap.Values | Where-Object { $_.Status -eq 'Duplicate' }).Count

        [void]$sb.AppendLine("    <div class=`"analysis-platform-panel$activeClass`" id=`"$panelId`" data-platform=`"$encPlat`">")
        [void]$sb.AppendLine('        <div class="analysis-summary-bar">')
        if ($conflictPolicies.Count -gt 0) {
            [void]$sb.AppendLine("            <button class=`"analysis-filter-pill status-conflict`" data-filter=`"conflict`" onclick=`"filterAnalysisPolicies('$panelId',this)`">Conflict $conflictSettingCount Settings</button>")
        }
        if ($duplicatePolicies.Count -gt 0) {
            [void]$sb.AppendLine("            <button class=`"analysis-filter-pill status-duplicate`" data-filter=`"duplicate`" onclick=`"filterAnalysisPolicies('$panelId',this)`">Duplicated $duplicateSettingCount Settings</button>")
        }
        [void]$sb.AppendLine('        </div>')

        if ($sortedPolicies.Count -eq 0) {
            [void]$sb.AppendLine('        <p class="analysis-no-data">No policies with conflict or duplicate settings found for this platform.</p>')
        } else {
            [void]$sb.AppendLine('        <table class="analysis-table">')
            [void]$sb.AppendLine('            <thead><tr>')
            [void]$sb.AppendLine('                <th style="width:38%">Policy Name</th>')
            [void]$sb.AppendLine('                <th style="width:27%">Type</th>')
            [void]$sb.AppendLine('                <th style="width:12%;text-align:center">Conflicts</th>')
            [void]$sb.AppendLine('                <th style="width:12%;text-align:center">Duplicates</th>')
            [void]$sb.AppendLine('                <th style="width:11%">Details</th>')
            [void]$sb.AppendLine('            </tr></thead>')
            [void]$sb.AppendLine('            <tbody>')

            $rowIdx = 0
            foreach ($polData in $sortedPolicies) {
                $filterAttr  = if ($polData.Conflicts.Count -gt 0 -and $polData.Duplicates.Count -gt 0) { 'both' }
                               elseif ($polData.Conflicts.Count -gt 0) { 'conflict' }
                               else { 'duplicate' }
                $detailRowId = "adr-$safeId-$rowIdx"
                $encPN       = [System.Net.WebUtility]::HtmlEncode($polData.Name)
                $encPT       = [System.Net.WebUtility]::HtmlEncode($polData.Type)
                $cCount      = $polData.Conflicts.Count
                $dCount      = $polData.Duplicates.Count
                $cBadge      = if ($cCount -gt 0) { "<span class='status-badge status-conflict'>$cCount</span>" } else { '<span style="color:#bbb">-</span>' }
                $dBadge      = if ($dCount -gt 0) { "<span class='status-badge status-duplicate'>$dCount</span>" } else { '<span style="color:#bbb">-</span>' }

                [void]$sb.AppendLine("                <tr class=`"analysis-policy-row`" data-filter=`"$filterAttr`" data-detail-id=`"$detailRowId`">")
                [void]$sb.AppendLine("                    <td class=`"setting-name-cell`">$encPN</td>")
                [void]$sb.AppendLine("                    <td>$encPT</td>")
                [void]$sb.AppendLine("                    <td style=`"text-align:center`">$cBadge</td>")
                [void]$sb.AppendLine("                    <td style=`"text-align:center`">$dBadge</td>")
                [void]$sb.AppendLine("                    <td><span class=`"expand-arrow`" onclick=`"toggleAnalysisRow(this,'$detailRowId')`">&#9654; Details</span></td>")
                [void]$sb.AppendLine('                </tr>')
                [void]$sb.AppendLine("                <tr class=`"analysis-detail-row`" id=`"$detailRowId`">")
                [void]$sb.AppendLine('                    <td colspan="5"><div class="analysis-detail-inner">')

                if ($cCount -gt 0) {
                    [void]$sb.AppendLine('                        <div class="detail-group-label conflict-label">Conflicting Settings</div>')
                    # Collect unique other-policy names across every conflict setting for dynamic columns
                    $conflictOtherPolNames = @($polData.Conflicts |
                        ForEach-Object { $_.OtherPolicies } |
                        Select-Object -ExpandProperty PolicyName |
                        Sort-Object -Unique)
                    $encPN_c = [System.Net.WebUtility]::HtmlEncode($polData.Name)
                    [void]$sb.Append('                        <table class="analysis-detail-table"><thead><tr>')
                    [void]$sb.Append("<th>Setting Name</th><th>$encPN_c (Value)</th>")
                    foreach ($opName in $conflictOtherPolNames) {
                        [void]$sb.Append("<th>$([System.Net.WebUtility]::HtmlEncode($opName)) (Value)</th>")
                    }
                    [void]$sb.AppendLine('</tr></thead><tbody>')
                    foreach ($s in ($polData.Conflicts | Sort-Object SettingName)) {
                        $encSN = [System.Net.WebUtility]::HtmlEncode($s.SettingName)
                        $encOV = [System.Net.WebUtility]::HtmlEncode($s.OwnValue)
                        [void]$sb.Append("                            <tr><td>$encSN</td><td>$encOV</td>")
                        foreach ($opName in $conflictOtherPolNames) {
                            $opEntry = $s.OtherPolicies | Where-Object { $_.PolicyName -eq $opName }
                            if ($opEntry) {
                                [void]$sb.Append("<td>$([System.Net.WebUtility]::HtmlEncode($opEntry.Value))</td>")
                            } else {
                                [void]$sb.Append('<td style="color:#bbb;text-align:center">&#8212;</td>')
                            }
                        }
                        [void]$sb.AppendLine('</tr>')
                    }
                    [void]$sb.AppendLine('                        </tbody></table>')
                }

                if ($dCount -gt 0) {
                    [void]$sb.AppendLine('                        <div class="detail-group-label duplicate-label">Duplicate Settings (same value in other policies)</div>')

                    # Collect all unique other-policy names across every duplicate setting for this policy
                    $otherPolNames = @($polData.Duplicates |
                        ForEach-Object { $_.OtherPolicies } |
                        Select-Object -ExpandProperty PolicyName |
                        Sort-Object -Unique)

                    # Dynamic header: Setting Name | Type | [This Policy] (Value) | [OtherPol1] (Value) | ...
                    $encPT = [System.Net.WebUtility]::HtmlEncode($polData.Type)
                    $encPN = [System.Net.WebUtility]::HtmlEncode($polData.Name)
                    [void]$sb.Append('                        <table class="analysis-detail-table"><thead><tr>')
                    [void]$sb.Append("<th>Setting Name</th><th>Type</th><th>$encPN (Value)</th>")
                    foreach ($opName in $otherPolNames) {
                        [void]$sb.Append("<th>$([System.Net.WebUtility]::HtmlEncode($opName)) (Value)</th>")
                    }
                    [void]$sb.AppendLine('</tr></thead><tbody>')

                    foreach ($s in ($polData.Duplicates | Sort-Object SettingName)) {
                        $encSN = [System.Net.WebUtility]::HtmlEncode($s.SettingName)
                        $encOV = [System.Net.WebUtility]::HtmlEncode($s.OwnValue)
                        [void]$sb.Append("                            <tr><td>$encSN</td><td>$encPT</td><td>$encOV</td>")
                        foreach ($opName in $otherPolNames) {
                            $opHasIt = $s.OtherPolicies | Where-Object { $_.PolicyName -eq $opName }
                            if ($opHasIt) {
                                [void]$sb.Append("<td>$encOV</td>")
                            } else {
                                [void]$sb.Append('<td style="color:#bbb;text-align:center">&#8212;</td>')
                            }
                        }
                        [void]$sb.AppendLine('</tr>')
                    }
                    [void]$sb.AppendLine('                        </tbody></table>')
                }

                [void]$sb.AppendLine('                    </div></td>')
                [void]$sb.AppendLine('                </tr>')
                $rowIdx++
            }

            [void]$sb.AppendLine('            </tbody>')
            [void]$sb.AppendLine('        </table>')
        }

        [void]$sb.AppendLine('    </div>')
    }

    [void]$sb.AppendLine('</div>')
    return $sb.ToString()
}

function Get-MarkdownAnalysisSection {
    param ([hashtable]$PlatformAnalysis)

    if (-not $PlatformAnalysis -or $PlatformAnalysis.Count -eq 0) { return '' }

    $activePlatforms = @($PlatformAnalysis.Keys | Where-Object { $PlatformAnalysis[$_].SettingMap.Count -gt 0 } | Sort-Object)
    if ($activePlatforms.Count -eq 0) { return '' }

    $md  = "## Settings Analysis`n`n"
    $md += "_Policy-centric duplicate and conflict detection. Shows only policies with conflicting or duplicate settings. Unique settings excluded._`n`n"
    $md += "---`n`n"

    foreach ($plat in $activePlatforms) {
        $data       = $PlatformAnalysis[$plat]
        $settingMap = $data.SettingMap

        # Build policy-centric map
        $policyMap = @{}
        foreach ($sEntry in $settingMap.GetEnumerator()) {
            # Strip the "PolicyType|||" compound key prefix to get the display setting name
            $sepIdx = $sEntry.Key.IndexOf('|||')
            $sName  = if ($sepIdx -ge 0) { $sEntry.Key.Substring($sepIdx + 3) } else { $sEntry.Key }
            $sData  = $sEntry.Value
            $status = $sData.Status
            foreach ($pol in $sData.Policies) {
                $polKey = "$($pol.PolicyName)||$($pol.PolicyType)"
                if (-not $policyMap.ContainsKey($polKey)) {
                    $policyMap[$polKey] = [pscustomobject]@{
                        Name       = $pol.PolicyName
                        Type       = ($pol.PolicyType -replace ' Policies$', '')
                        Conflicts  = [System.Collections.Generic.List[object]]::new()
                        Duplicates = [System.Collections.Generic.List[object]]::new()
                    }
                }
                $otherPols = [System.Collections.Generic.List[object]]::new()
                if ($pol.IsUnassigned -or $status -eq 'Duplicate') {
                    foreach ($op in $sData.Policies) {
                        if ($op.PolicyName -ne $pol.PolicyName -and $op.Value -eq $pol.Value) { $otherPols.Add($op) }
                    }
                } else {
                    foreach ($op in $sData.Policies) {
                        if ($op.PolicyName -ne $pol.PolicyName -and -not $op.IsUnassigned) { $otherPols.Add($op) }
                    }
                }
                $sItem = [pscustomobject]@{
                    SettingName   = $sName
                    OwnValue      = $pol.Value
                    OtherPolicies = $otherPols
                }
                if ($pol.IsUnassigned) {
                    $policyMap[$polKey].Duplicates.Add($sItem)
                } elseif ($status -eq 'Conflict') {
                    $policyMap[$polKey].Conflicts.Add($sItem)
                } else {
                    $policyMap[$polKey].Duplicates.Add($sItem)
                }
            }
        }

        $conflictPolicies  = @($policyMap.Values | Where-Object { $_.Conflicts.Count -gt 0 } | Sort-Object Name)
        $duplicatePolicies = @($policyMap.Values | Where-Object { $_.Conflicts.Count -eq 0 -and $_.Duplicates.Count -gt 0 } | Sort-Object Name)

        if ($conflictPolicies.Count -eq 0 -and $duplicatePolicies.Count -eq 0) { continue }

        $md += "### $plat`n`n"

        if ($conflictPolicies.Count -gt 0) {
            $md += "**Policies with Conflicting Settings ($($conflictPolicies.Count)):**`n`n"
            $md += "| Policy | Type | Setting Name | Policy (Value) | Other Policy (Values) |`n"
            $md += "|--------|------|--------------|----------------|----------------------|`n"
            foreach ($p in $conflictPolicies) {
                $polName = $p.Name -replace '\|', '&#124;'
                $polType = $p.Type -replace '\|', '&#124;'
                foreach ($s in ($p.Conflicts | Sort-Object SettingName)) {
                    $sName  = $s.SettingName -replace '\|', '&#124;'
                    $ownVal = $s.OwnValue -replace '\|', '&#124;'
                    if ($ownVal.Length -gt 40) { $ownVal = $ownVal.Substring(0, 37) + '...' }
                    $otherVals = ($s.OtherPolicies | ForEach-Object {
                        $n = $_.PolicyName -replace '\|', '&#124;'
                        $v = $_.Value -replace '\|', '&#124;'
                        if ($v.Length -gt 30) { $v = $v.Substring(0, 27) + '...' }
                        "$n ($v)"
                    }) -join ' / '
                    if ($otherVals.Length -gt 80) { $otherVals = $otherVals.Substring(0, 77) + '...' }
                    $md += "| $polName | $polType | $sName | $polName ($ownVal) | $otherVals |`n"
                }
            }
            $md += "`n"
        }

        if ($duplicatePolicies.Count -gt 0) {
            $md += "**Policies with Duplicate Settings ($($duplicatePolicies.Count)):**`n`n"
            $md += "| Setting Name | Type | Policy Name (Value) | Also Configured In (Value) |`n"
            $md += "|--------------|------|---------------------|----------------------------|`n"
            foreach ($p in $duplicatePolicies) {
                $polType = $p.Type -replace '\|', '&#124;'
                $polName = $p.Name -replace '\|', '&#124;'
                foreach ($s in ($p.Duplicates | Sort-Object SettingName)) {
                    $sName  = $s.SettingName -replace '\|', '&#124;'
                    $ownVal = $s.OwnValue -replace '\|', '&#124;'
                    if ($ownVal.Length -gt 30) { $ownVal = $ownVal.Substring(0, 27) + '...' }
                    $others = ($s.OtherPolicies | ForEach-Object {
                        $n = $_.PolicyName -replace '\|', '&#124;'
                        "$n ($ownVal)"
                    }) -join ', '
                    if ($others.Length -gt 80) { $others = $others.Substring(0, 77) + '...' }
                    $md += "| $sName | $polType | $polName ($ownVal) | $others |`n"
                }
            }
            $md += "`n"
        }

        $md += "`n"
    }

    return $md
}

#region Export/Import Functions

function Export-BackupToMarkdown {
    param (
        [string]$BackupPath,
        [string]$OutputPath
    )

    Write-ColoredOutput "`nExporting backup to Markdown..." -Color $script:Colors.Info

    if (-not (Test-Path $BackupPath)) {
        Write-ColoredOutput "Backup path not found: $BackupPath" -Color $script:Colors.Error
        Write-Log "Export failed: Backup path not found" -Level Error
        return
    }

    # Load metadata
    $metadataPath = Join-Path $BackupPath "metadata.json"
    $metadata = if (Test-Path $metadataPath) {
        Get-Content $metadataPath | ConvertFrom-Json
    } else {
        @{BackupDate = "Unknown"; TenantId = "Unknown"; TotalItems = 0}
    }

    # Start building Markdown content
    $md = @"
# UniFy-Endpoint Backup Export

**Export Date:** $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
**Source Backup:** $BackupPath
**Backup Date:** $($metadata.BackupDate)
**Tenant ID:** $($metadata.TenantId)
**Total Items:** $($metadata.TotalItems)

---

"@

    # Build settings analysis and prepend to report
    if (-not $script:SettingDefinitionCache) { $script:SettingDefinitionCache = @{} }
    $mdAnalysisPolicyTypes = @(
        @{Name = "DeviceConfigurations";    DisplayName = "Device Configuration Policies"},
        @{Name = "CompliancePolicies";      DisplayName = "Compliance Policies"},
        @{Name = "SettingsCatalogPolicies"; DisplayName = "Settings Catalog Policies"},
        @{Name = "AppProtectionPolicies";   DisplayName = "App Protection Policies"},
        @{Name = "PowerShellScripts";       DisplayName = "PowerShell Scripts"},
        @{Name = "AdministrativeTemplates"; DisplayName = "Administrative Templates"},
        @{Name = "AutopilotProfiles";       DisplayName = "Autopilot Profiles"},
        @{Name = "AutopilotDevicePrep";     DisplayName = "Autopilot Device Prep"},
        @{Name = "EnrollmentStatusPage";    DisplayName = "Enrollment Status Page"},
        @{Name = "RemediationScripts";      DisplayName = "Remediation Scripts"},
        @{Name = "WUfBPolicies";            DisplayName = "Windows Update for Business"},
        @{Name = "AssignmentFilters";       DisplayName = "Assignment Filters"},
        @{Name = "AppConfigManagedDevices"; DisplayName = "App Config (Managed Devices)"},
        @{Name = "AppConfigManagedApps";    DisplayName = "App Config (Managed Apps)"},
        @{Name = "MacOSScripts";            DisplayName = "macOS Scripts"},
        @{Name = "MacOSCustomAttributes";   DisplayName = "macOS Custom Attributes"}
    )
    $mdAllSections = @()
    foreach ($pt in $mdAnalysisPolicyTypes) {
        $typePath = Join-Path $BackupPath $pt.Name
        if (-not (Test-Path $typePath)) { continue }
        $ptPolicies = @(Get-ChildItem -Path $typePath -Filter "*.json" -File | ForEach-Object {
            try { Get-Content $_.FullName -Raw | ConvertFrom-Json } catch { $null }
        } | Where-Object { $null -ne $_ })
        if ($ptPolicies.Count -eq 0) { continue }
        $ptCards = @()
        foreach ($pol in $ptPolicies) {
            $polName     = if ($pol.displayName) { $pol.displayName } else { $pol.name }
            if (-not $polName) { $polName = "Unnamed" }
            $polPlatform = Get-PolicyPlatform -PolicyName $polName -PolicyType $pt.Name -Policy $pol
            $polSettings = @(Resolve-PolicyAllSettings -Policy $pol -PolicyType $pt.Name)
            $polGroups   = @($pol.assignments | Where-Object {
                $_.target -and $_.target.'@odata.type' -ne '#microsoft.graph.exclusionGroupAssignmentTarget'
            } | ForEach-Object {
                $t = $_.target
                if     ($t.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') { 'all_users' }
                elseif ($t.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget')       { 'all_devices' }
                elseif ($t.groupId) { $t.groupId }
            } | Where-Object { $_ })
            $ptCards    += @{ Name=$polName; Platform=$polPlatform; Settings=$polSettings; Groups=$polGroups }
        }
        $mdAllSections += @{ DisplayName=$pt.DisplayName; TypeName=$pt.Name; Cards=$ptCards }
    }
    $mdPlatformAnalysis = Build-SettingsAnalysis -AllSections $mdAllSections
    $mdAnalysisText     = Get-MarkdownAnalysisSection -PlatformAnalysis $mdPlatformAnalysis
    if ($mdAnalysisText) { $md += $mdAnalysisText }

    # Export each policy type
    $policyTypes = @(
        @{Name = "DeviceConfigurations"; DisplayName = "Device Configuration Policies"},
        @{Name = "CompliancePolicies"; DisplayName = "Compliance Policies"},
        @{Name = "SettingsCatalogPolicies"; DisplayName = "Settings Catalog Policies"},
        @{Name = "AppProtectionPolicies"; DisplayName = "App Protection Policies"},
        @{Name = "PowerShellScripts"; DisplayName = "PowerShell Scripts"},
        @{Name = "AdministrativeTemplates"; DisplayName = "Administrative Templates"},
        @{Name = "AutopilotProfiles"; DisplayName = "Autopilot Profiles"},
        @{Name = "AutopilotDevicePrep"; DisplayName = "Autopilot Device Prep"},
        @{Name = "EnrollmentStatusPage"; DisplayName = "Enrollment Status Page"},
        @{Name = "RemediationScripts"; DisplayName = "Remediation Scripts"},
        @{Name = "WUfBPolicies"; DisplayName = "Windows Update for Business"},
        @{Name = "AssignmentFilters"; DisplayName = "Assignment Filters"},
        @{Name = "AppConfigManagedDevices"; DisplayName = "App Config (Managed Devices)"},
        @{Name = "AppConfigManagedApps"; DisplayName = "App Config (Managed Apps)"}
    )

    $totalExported = 0

    foreach ($policyType in $policyTypes) {
        $typePath = Join-Path $BackupPath $policyType.Name
        if (Test-Path $typePath) {
            $policies = Get-ChildItem -Path $typePath -Filter "*.json" | ForEach-Object {
                Get-Content $_.FullName | ConvertFrom-Json
            }

            if ($policies.Count -gt 0) {
                $md += "## $($policyType.DisplayName) ($($policies.Count))`n`n"
                $md += "| Name | Policy ID | Created | Modified | Platform | Assigned |`n"
                $md += "|------|-----------|---------|----------|----------|----------|`n"

                foreach ($policy in $policies) {
                    $name = if ($policy.displayName) { $policy.displayName } elseif ($policy.name) { $policy.name } else { "N/A" }
                    $id = if ($policy.id) { $policy.id } else { "N/A" }
                    $created = ConvertFrom-JsonDate -DateString $policy.createdDateTime
                    $modified = ConvertFrom-JsonDate -DateString $policy.lastModifiedDateTime
                    $isAssigned = if ($null -ne $policy.isAssigned) { if ($policy.isAssigned) { "Yes" } else { "No" } } else { "N/A" }
                    # Determine platform from various possible properties
                    $platform = "All"
                    if ($policy.platform) { $platform = $policy.platform }
                    elseif ($policy.platforms -and $policy.platforms.Count -gt 0) { $platform = $policy.platforms -join ", " }
                    elseif ($policy.'@odata.type') {
                        if ($policy.'@odata.type' -like "*ios*") { $platform = "iOS" }
                        elseif ($policy.'@odata.type' -like "*android*") { $platform = "Android" }
                        elseif ($policy.'@odata.type' -like "*windows*" -or $policy.'@odata.type' -like "*win*") { $platform = "Windows" }
                        elseif ($policy.'@odata.type' -like "*macOS*") { $platform = "macOS" }
                    }

                    $md += "| $name | ``$id`` | $created | $modified | $platform | $isAssigned |`n"
                }

                $md += "`n"
                $totalExported += $policies.Count
                Write-ColoredOutput "  - Exported $($policies.Count) $($policyType.DisplayName)" -Color $script:Colors.Success
            }
        }
    }

    $md += @"
---

*Generated by UniFy-Endpoint v$($script:Version) on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")*
"@

    # Save to file
    $md | Out-File -FilePath $OutputPath -Encoding UTF8
    Write-ColoredOutput "Markdown export completed: $OutputPath" -Color $script:Colors.Success
}

function Export-BackupToCSV {
    param (
        [string]$BackupPath,
        [string]$OutputFolder
    )

    Write-ColoredOutput "`nExporting backup to CSV..." -Color $script:Colors.Info

    if (-not (Test-Path $BackupPath)) {
        Write-ColoredOutput "Backup path not found: $BackupPath" -Color $script:Colors.Error
        Write-Log "Export failed: Backup path not found" -Level Error
        return
    }

    if (-not (Test-Path $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    }

    # Export summary CSV
    $summaryData = @()

    $policyTypes = @(
        @{Name = "DeviceConfigurations"; DisplayName = "Device Configuration Policies"},
        @{Name = "CompliancePolicies"; DisplayName = "Compliance Policies"},
        @{Name = "SettingsCatalogPolicies"; DisplayName = "Settings Catalog Policies"},
        @{Name = "AppProtectionPolicies"; DisplayName = "App Protection Policies"},
        @{Name = "PowerShellScripts"; DisplayName = "PowerShell Scripts"},
        @{Name = "AdministrativeTemplates"; DisplayName = "Administrative Templates"},
        @{Name = "AutopilotProfiles"; DisplayName = "Autopilot Profiles"},
        @{Name = "AutopilotDevicePrep"; DisplayName = "Autopilot Device Prep"},
        @{Name = "EnrollmentStatusPage"; DisplayName = "Enrollment Status Page"},
        @{Name = "RemediationScripts"; DisplayName = "Remediation Scripts"},
        @{Name = "WUfBPolicies"; DisplayName = "Windows Update for Business"},
        @{Name = "AssignmentFilters"; DisplayName = "Assignment Filters"},
        @{Name = "AppConfigManagedDevices"; DisplayName = "App Config (Managed Devices)"},
        @{Name = "AppConfigManagedApps"; DisplayName = "App Config (Managed Apps)"}
    )

    foreach ($policyType in $policyTypes) {
        $typePath = Join-Path $BackupPath $policyType.Name
        if (Test-Path $typePath) {
            $policies = Get-ChildItem -Path $typePath -Filter "*.json" | ForEach-Object {
                $policy = Get-Content $_.FullName | ConvertFrom-Json

                # Determine platform from various possible properties
                $platform = "All"
                if ($policy.platform) { $platform = $policy.platform }
                elseif ($policy.platforms -and $policy.platforms.Count -gt 0) { $platform = $policy.platforms -join ", " }
                elseif ($policy.'@odata.type') {
                    if ($policy.'@odata.type' -like "*ios*") { $platform = "iOS" }
                    elseif ($policy.'@odata.type' -like "*android*") { $platform = "Android" }
                    elseif ($policy.'@odata.type' -like "*windows*" -or $policy.'@odata.type' -like "*win*") { $platform = "Windows" }
                    elseif ($policy.'@odata.type' -like "*macOS*") { $platform = "macOS" }
                }

                $summaryData += [PSCustomObject]@{
                    PolicyType = $policyType.DisplayName
                    PolicyName = if ($policy.displayName) { $policy.displayName } elseif ($policy.name) { $policy.name } else { "N/A" }
                    PolicyId = if ($policy.id) { $policy.id } else { "N/A" }
                    CreatedDate = ConvertFrom-JsonDate -DateString $policy.createdDateTime
                    ModifiedDate = ConvertFrom-JsonDate -DateString $policy.lastModifiedDateTime
                    Platform = $platform
                    IsAssigned = if ($null -ne $policy.isAssigned) { if ($policy.isAssigned) { "Yes" } else { "No" } } else { "N/A" }
                }
            }
        }
    }

    if ($summaryData.Count -gt 0) {
        $csvPath = Join-Path $OutputFolder "backup_summary.csv"
        $summaryData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-ColoredOutput "  - Summary exported to: $csvPath" -Color $script:Colors.Success
    }

    Write-Log "CSV export completed" -Level Info
}

function Get-PolicyPlatform {
    param (
        [string]$PolicyName,
        [string]$PolicyType,
        [object]$Policy = $null
    )

    # 1. Check @odata.type (most reliable for DeviceConfigurations, CompliancePolicies, AppConfig)
    if ($Policy -and $Policy.'@odata.type') {
        $odataType = $Policy.'@odata.type'.ToLower()
        if ($odataType -match 'windows|win10|win32') { return 'Windows' }
        if ($odataType -match '\.ios') { return 'iOS' }
        if ($odataType -match 'android') { return 'Android' }
        if ($odataType -match 'macos') { return 'macOS' }
    }

    # 2. Check platforms property (SettingsCatalog, AssignmentFilters)
    if ($Policy -and $Policy.platforms) {
        $platforms = "$($Policy.platforms)".ToLower()
        if ($platforms -match 'windows') { return 'Windows' }
        if ($platforms -match '\bios\b') { return 'iOS' }
        if ($platforms -match 'android') { return 'Android' }
        if ($platforms -match 'macos') { return 'macOS' }
    }

    # 3. Check platform property (some policy types use singular)
    if ($Policy -and $Policy.platform) {
        $platform = "$($Policy.platform)".ToLower()
        if ($platform -match 'windows') { return 'Windows' }
        if ($platform -match '\bios\b') { return 'iOS' }
        if ($platform -match 'android') { return 'Android' }
        if ($platform -match 'macos') { return 'macOS' }
    }

    # 4. Check policy type name
    if ($PolicyType -in @('MacOSScripts','MacOSCustomAttributes')) { return 'macOS' }
    if ($PolicyType -in @('PowerShellScripts','RemediationScripts','WUfBPolicies','AdministrativeTemplates')) { return 'Windows' }
    if ($PolicyType -in @('AutopilotProfiles','AutopilotDevicePrep','EnrollmentStatusPage')) { return 'Windows' }

    # 5. Check policy name patterns
    if ($PolicyName -match '^(WIN |WIN-|WIN_)') { return 'Windows' }
    if ($PolicyName -match '^(ADN |AND |ADN-|AND-)') { return 'Android' }
    if ($PolicyName -match '^(iOS |iOS-|\[CIS\] - L\d - iOS)') { return 'iOS' }
    if ($PolicyName -match '(macOS|macos|\[CIS\].*macOS)') { return 'macOS' }

    return 'General'
}

function ConvertFrom-CamelCase {
    param ([string]$Name)
    # Handle dot-notation paths: convert last segment only, keep path for context
    if ($Name -match '\.') {
        $parts = $Name -split '\.'
        $converted = foreach ($part in $parts) {
            # Strip array indices for display
            $clean = $part -replace '\[\d+\]', ''
            if ($clean) {
                # Insert space before each uppercase letter that follows a lowercase letter or digit
                $spaced = [regex]::Replace($clean, '(?<=[a-z0-9])(?=[A-Z])', ' ')
                # Capitalize first letter
                if ($spaced.Length -gt 0) {
                    $spaced.Substring(0,1).ToUpper() + $spaced.Substring(1)
                } else { $spaced }
            }
        }
        return ($converted -join ' > ')
    }

    # Simple name: convert camelCase to Title Case
    $clean = $Name -replace '\[\d+\]', ''
    $spaced = [regex]::Replace($clean, '(?<=[a-z0-9])(?=[A-Z])', ' ')
    if ($spaced.Length -gt 0) {
        return $spaced.Substring(0,1).ToUpper() + $spaced.Substring(1)
    }
    return $spaced
}

function Resolve-PolicyAllSettings {
    param (
        [object]$Policy,
        [string]$PolicyType
    )

    $settings = @()

    if ($PolicyType -eq 'SettingsCatalogPolicies' -and $Policy.settings) {
        # Extract all settingDefinitionId -> rawValue pairs
        $values = @{}
        foreach ($s in $Policy.settings) {
            if ($s.settingInstance) {
                Extract-SettingInstanceValues -Instance $s.settingInstance -Result $values
            }
        }

        if ($values.Count -eq 0) { return $settings }

        # Resolve each definition via Graph API (with caching)
        foreach ($defId in $values.Keys) {
            $rawValue = $values[$defId]
            $displayName = $defId
            $description = ''
            $displayValue = if ($null -ne $rawValue) { "$rawValue" } else { '(not configured)' }

            if (-not $script:SettingDefinitionCache) { $script:SettingDefinitionCache = @{} }

            if ($script:SettingDefinitionCache.ContainsKey($defId)) {
                $cached = $script:SettingDefinitionCache[$defId]
                if ($cached) {
                    $displayName = if ($cached.DisplayName) { $cached.DisplayName } else { $defId }
                    $description = if ($cached.Description) { $cached.Description } else { '' }
                    $displayValue = Resolve-SettingDisplayValue -RawValue $rawValue -OptionsLookup $cached.Options
                }
            }
            else {
                try {
                    $encodedId = [System.Uri]::EscapeDataString($defId)
                    $uri = "https://graph.microsoft.com/beta/deviceManagement/configurationSettings('$encodedId')"
                    $def = Invoke-MgGraphRequest -Uri $uri -Method GET

                    $optionsLookup = @{}
                    if ($def.options) {
                        foreach ($opt in $def.options) {
                            if ($opt.itemId -and $opt.displayName) {
                                $optionsLookup[$opt.itemId] = $opt.displayName
                            }
                        }
                    }

                    $script:SettingDefinitionCache[$defId] = @{
                        DisplayName = $def.displayName
                        Description = $def.description
                        Options = $optionsLookup
                    }

                    $displayName = if ($def.displayName) { $def.displayName } else { $defId }
                    $description = if ($def.description) { $def.description } else { '' }
                    $displayValue = Resolve-SettingDisplayValue -RawValue $rawValue -OptionsLookup $optionsLookup
                }
                catch {
                    $script:SettingDefinitionCache[$defId] = $null
                }
            }

            $settings += @{
                Name        = $displayName
                Value       = $displayValue
                Description = $description
            }
        }
    }
    else {
        # Non-Settings Catalog: flatten the policy object
        $metadataFields = @('id', 'createdDateTime', 'lastModifiedDateTime', 'modifiedDateTime', 'version',
            '@odata.context', '@odata.type', '@odata.id', '@odata.etag',
            'roleScopeTagIds', 'supportsScopeTags', 'isAssigned', 'assignments',
            'creationSource', 'priority', 'createdBy', 'lastModifiedBy',
            'templateReference', 'templateId', 'templateDisplayName', 'templateDisplayVersion', 'templateFamily',
            'settingCount', 'displayName', 'name', 'description')

        $clone = $Policy | ConvertTo-Json -Depth 20 | ConvertFrom-Json
        foreach ($field in $metadataFields) {
            if ($clone.PSObject.Properties[$field]) {
                $clone.PSObject.Properties.Remove($field)
            }
        }

        $flat = Flatten-Object -Obj $clone
        foreach ($key in $flat.Keys) {
            $val = $flat[$key]
            $displayVal = if ($null -eq $val) { '(not set)' }
                          elseif ($val -is [string] -or $val -is [bool] -or $val -is [int] -or $val -is [long] -or $val -is [double]) { "$val" }
                          else { try { $val | ConvertTo-Json -Depth 5 -Compress } catch { "$val" } }

            $friendlyName = ConvertFrom-CamelCase -Name $key

            $settings += @{
                Name        = $friendlyName
                Value       = $displayVal
                Description = ''
            }
        }
    }

    return $settings
}

function Export-CurrentTenantConfigToHTML {
    param (
        [string]$OutputPath
    )

    $startTime = Get-Date
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $tempFolder = Join-Path $env:TEMP "UniFy-Endpoint_LiveReport_$timestamp"

    try {
        New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null

        # Save and override platform filter to get all policies
        $savedPlatform = $script:SelectedPlatform
        $script:SelectedPlatform = "All"

        Write-ColoredOutput "`nFetching live configuration from Intune tenant..." -Color $script:Colors.Info
        Write-ColoredOutput "This may take a few minutes depending on tenant size.`n" -Color $script:Colors.Warning

        $script:IsLiveReport = $true

        # Export all 16 component types to temp folder
        $counts = @{}
        $counts["DeviceConfigurations"]    = Export-DeviceConfigurations -BackupPath $tempFolder
        $counts["CompliancePolicies"]      = Export-CompliancePolicies -BackupPath $tempFolder
        $counts["SettingsCatalogPolicies"] = Export-SettingsCatalogPolicies -BackupPath $tempFolder
        $counts["AppProtectionPolicies"]   = Export-AppProtectionPolicies -BackupPath $tempFolder
        $counts["PowerShellScripts"]       = Export-PowerShellScripts -BackupPath $tempFolder
        $counts["AdministrativeTemplates"] = Export-AdministrativeTemplates -BackupPath $tempFolder
        $counts["AutopilotProfiles"]       = Export-AutopilotProfiles -BackupPath $tempFolder
        $counts["AutopilotDevicePrep"]     = Export-AutopilotDevicePrep -BackupPath $tempFolder
        $counts["EnrollmentStatusPage"]    = Export-EnrollmentStatusPage -BackupPath $tempFolder
        $counts["RemediationScripts"]      = Export-RemediationScripts -BackupPath $tempFolder
        $counts["WUfBPolicies"]            = Export-WUfBPolicies -BackupPath $tempFolder
        $counts["AssignmentFilters"]       = Export-AssignmentFilters -BackupPath $tempFolder
        $counts["AppConfigManagedDevices"] = Export-AppConfigManagedDevices -BackupPath $tempFolder
        $counts["AppConfigManagedApps"]    = Export-AppConfigManagedApps -BackupPath $tempFolder
        $counts["MacOSScripts"]            = Export-MacOSScripts -BackupPath $tempFolder
        $counts["MacOSCustomAttributes"]   = Export-MacOSCustomAttributes -BackupPath $tempFolder

        # Merge Autopilot Device Prep into Settings Catalog folder for unified HTML display
        # (Device Prep policies are Settings Catalog policies with a specific templateId)
        $devicePrepPath     = Join-Path $tempFolder "AutopilotDevicePrep"
        $settingsCatalogPath = Join-Path $tempFolder "SettingsCatalogPolicies"
        if (Test-Path $devicePrepPath) {
            if (-not (Test-Path $settingsCatalogPath)) {
                New-Item -ItemType Directory -Path $settingsCatalogPath -Force | Out-Null
            }
            Get-ChildItem -Path $devicePrepPath -Filter "*.json" |
                Move-Item -Destination $settingsCatalogPath -Force
            Remove-Item -Path $devicePrepPath -Recurse -Force -ErrorAction SilentlyContinue
            $counts["SettingsCatalogPolicies"] = $counts["SettingsCatalogPolicies"] + $counts["AutopilotDevicePrep"]
            $counts.Remove("AutopilotDevicePrep")
        }

        # Restore original platform filter
        $script:SelectedPlatform = $savedPlatform

        # Calculate duration
        $endTime = Get-Date
        $duration = $endTime - $startTime
        $durationSeconds = [Math]::Round($duration.TotalSeconds)
        $durationFormatted = if ($durationSeconds -lt 60) { "${durationSeconds}s" } else {
            $minutes = [Math]::Floor($durationSeconds / 60)
            $seconds = $durationSeconds % 60
            "${minutes}m ${seconds}s"
        }

        $totalItems = ($counts.Values | Measure-Object -Sum).Sum

        # Create metadata for the report
        $metadata = @{
            BackupDate      = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            TenantId        = (Get-MgContext).TenantId
            Duration        = $durationFormatted
            TotalItems      = $totalItems
            ReportDateLabel = "Snapshot Date"
            ItemCounts      = $counts
        }
        $metadataPath = Join-Path $tempFolder "metadata.json"
        $metadata | ConvertTo-Json -Depth 10 | Out-File -FilePath $metadataPath -Encoding UTF8

        # Initialize settings cache and generate HTML report
        $script:SettingDefinitionCache = @{}

        Write-ColoredOutput "`nGenerating HTML report from live data..." -Color $script:Colors.Info

        $tenantDisplayName = try {
            (Get-MgOrganization -ErrorAction Stop | Select-Object -First 1).DisplayName
        } catch { (Get-MgContext).TenantId }

        Export-BackupToHTML -BackupPath $tempFolder -OutputPath $OutputPath `
            -ReportTitle "$tenantDisplayName Intune Tenant Report" `
            -ReportSubtitle "Current Intune Configuration Snapshot - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"

        Write-ColoredOutput "Fetched $totalItems policies in $durationFormatted" -Color $script:Colors.Success
    }
    finally {
        $script:IsLiveReport = $false
        # Clean up temp folder
        if (Test-Path $tempFolder) {
            Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Export-BackupToHTML {
    param (
        [string]$BackupPath,
        [string]$OutputPath,
        [string]$ReportTitle = "UniFy-Endpoint Backup Report",
        [string]$ReportSubtitle = "Comprehensive Intune Configuration Backup Analysis"
    )

    if (-not (Test-Path $BackupPath)) {
        Write-ColoredOutput "Backup path not found: $BackupPath" -Color $script:Colors.Error
        Write-Log "Export failed: Backup path not found" -Level Error
        return
    }

    # Load metadata
    $metadataPath = Join-Path $BackupPath "metadata.json"
    $metadata = if (Test-Path $metadataPath) {
        Get-Content $metadataPath | ConvertFrom-Json
    } else {
        @{BackupDate = "Unknown"; TenantId = "Unknown"; TotalItems = 0; Duration = "N/A"}
    }

    # Initialize settings definition cache for Settings Catalog resolution
    if (-not $script:SettingDefinitionCache) { $script:SettingDefinitionCache = @{} }

    # Policy types to process
    $policyTypes = @(
        @{Name = "DeviceConfigurations"; DisplayName = "Device Configuration Policies"},
        @{Name = "CompliancePolicies"; DisplayName = "Compliance Policies"},
        @{Name = "SettingsCatalogPolicies"; DisplayName = "Settings Catalog Policies"},
        @{Name = "AppProtectionPolicies"; DisplayName = "App Protection Policies"},
        @{Name = "PowerShellScripts"; DisplayName = "PowerShell Scripts"},
        @{Name = "AdministrativeTemplates"; DisplayName = "Administrative Templates"},
        @{Name = "AutopilotProfiles"; DisplayName = "Autopilot Profiles"},
        @{Name = "AutopilotDevicePrep"; DisplayName = "Autopilot Device Prep"},
        @{Name = "EnrollmentStatusPage"; DisplayName = "Enrollment Status Page"},
        @{Name = "RemediationScripts"; DisplayName = "Remediation Scripts"},
        @{Name = "WUfBPolicies"; DisplayName = "Windows Update for Business"},
        @{Name = "AssignmentFilters"; DisplayName = "Assignment Filters"},
        @{Name = "AppConfigManagedDevices"; DisplayName = "App Config (Managed Devices)"},
        @{Name = "AppConfigManagedApps"; DisplayName = "App Config (Managed Apps)"},
        @{Name = "MacOSScripts"; DisplayName = "macOS Scripts"},
        @{Name = "MacOSCustomAttributes"; DisplayName = "macOS Custom Attributes"}
    )

    # First pass: collect all policy data with settings
    $allSections = @()
    $totalPolicies = 0
    $totalSettings = 0
    $summaryCounts = @{}

    foreach ($policyType in $policyTypes) {
        $typePath = Join-Path $BackupPath $policyType.Name
        if (-not (Test-Path $typePath)) { continue }

        $policies = @(Get-ChildItem -Path $typePath -Filter "*.json" -File | ForEach-Object {
            try {
                Get-Content $_.FullName -Raw | ConvertFrom-Json
            } catch {
                Write-ColoredOutput "  Warning: Failed to read $($_.Name)" -Color $script:Colors.Warning
                $null
            }
        } | Where-Object { $null -ne $_ })

        if ($policies.Count -eq 0) { continue }

        $summaryCounts[$policyType.DisplayName] = $policies.Count
        $totalPolicies += $policies.Count

        Write-ColoredOutput "  Processing $($policyType.DisplayName) ($($policies.Count))..." -Color $script:Colors.Info

        $policyCards = @()
        foreach ($policy in $policies) {
            $name = if ($policy.displayName) { $policy.displayName } else { $policy.name }
            if (-not $name) { $name = "Unnamed Policy" }
            $id = if ($policy.id) { $policy.id } else { "N/A" }
            $created = ConvertFrom-JsonDate -DateString $policy.createdDateTime
            $modified = ConvertFrom-JsonDate -DateString $policy.lastModifiedDateTime
            $platform = Get-PolicyPlatform -PolicyName $name -PolicyType $policyType.Name -Policy $policy

            $settings = @(Resolve-PolicyAllSettings -Policy $policy -PolicyType $policyType.Name)
            $totalSettings += $settings.Count

            $policyGroups = @($policy.assignments | Where-Object {
                $_.target -and $_.target.'@odata.type' -ne '#microsoft.graph.exclusionGroupAssignmentTarget'
            } | ForEach-Object {
                $t = $_.target
                if     ($t.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') { 'all_users' }
                elseif ($t.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget')       { 'all_devices' }
                elseif ($t.groupId) { $t.groupId }
            } | Where-Object { $_ })
            $policyCards += @{
                Name       = $name
                Id         = $id
                Created    = $created
                Modified   = $modified
                Platform   = $platform
                Settings   = $settings
                IsAssigned = if ($null -ne $policy.isAssigned) { if ($policy.isAssigned) { "Yes" } else { "No" } } else { "N/A" }
                Groups     = $policyGroups
            }
        }

        $allSections += @{
            DisplayName = $policyType.DisplayName
            TypeName    = $policyType.Name
            Cards       = $policyCards
        }
    }

    # Helper for HTML encoding
    function HtmlEncode { param([string]$Text) return [System.Net.WebUtility]::HtmlEncode($Text) }

    # Build settings analysis data (no additional API calls — uses already-resolved $allSections)
    Write-ColoredOutput "  Analyzing settings for duplicates and conflicts..." -Color $script:Colors.Info
    $platformAnalysis = Build-SettingsAnalysis -AllSections $allSections

    $analysisCSS = @'
        /* ===== Settings Analysis Section ===== */
        .analysis-section { background: white; border-radius: 10px; padding: 25px; margin-bottom: 20px; box-shadow: 0 5px 15px rgba(0,0,0,0.08); }
        .analysis-section h2 { color: #1a1a2e; margin: 0 0 12px 0; padding: 10px 14px; border: 1px solid #c7ddf5; border-left: 4px solid #0078d4; border-radius: 6px; font-size: 1.1em; font-weight: 600; background: #f5f9fe; }
        .section-desc { color: #666; font-size: 0.88em; margin: 0 0 14px 0; }
        .analysis-tabs { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 14px; }
        .analysis-tab-btn { padding: 6px 16px; border: 2px solid #e0e0e0; border-radius: 8px; background: white; cursor: pointer; font-size: 0.86em; font-family: inherit; transition: border-color 0.2s, background 0.2s, color 0.2s; }
        .analysis-tab-btn:hover { border-color: #0078d4; }
        .analysis-tab-btn.active { background: #0078d4; color: white; border-color: #0078d4; }
        .analysis-platform-panel { display: none; }
        .analysis-platform-panel.active { display: block; }
        .analysis-summary-bar { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 14px; align-items: center; }
        .status-badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 0.78em; font-weight: 700; }
        .status-conflict  { background: #fdecea; color: #c62828; border: 1px solid #ef9a9a; }
        .status-duplicate { background: #0078d4; color: white; border: 1px solid #0078d4; }
        .analysis-filter-pill { padding: 5px 14px; border-radius: 8px; cursor: pointer; font-size: 0.84em; font-family: inherit; transition: opacity 0.2s, box-shadow 0.2s; }
        .analysis-filter-pill:hover { opacity: 0.85; }
        .analysis-filter-pill.filter-active { box-shadow: 0 0 0 3px rgba(0,0,0,0.2); }
        .analysis-table { width: 100%; border-collapse: collapse; font-size: 0.87em; margin-bottom: 6px; }
        .analysis-table th { background: #0078d4; color: white; padding: 8px 12px; text-align: left; font-weight: 500; position: sticky; top: 0; z-index: 1; }
        .analysis-table td { padding: 8px 12px; border-bottom: 1px solid #f0f0f0; vertical-align: top; word-break: break-word; }
        .analysis-table tr:hover td { background: #f5faff; }
        .setting-name-cell { max-width: 300px; font-weight: 500; }
        .expand-arrow { cursor: pointer; color: #0078d4; font-size: 0.83em; user-select: none; white-space: nowrap; }
        .expand-arrow:hover { text-decoration: underline; }
        .analysis-detail-row { display: none; }
        .analysis-detail-row.visible { display: table-row; }
        .analysis-detail-row td { background: #f0f7ff; padding: 0; }
        .analysis-detail-inner { padding: 10px 16px; }
        .detail-group-label { font-size: 0.82em; font-weight: 700; margin: 8px 0 4px 0; padding: 3px 8px; border-radius: 4px; display: inline-block; }
        .conflict-label { background: #fdecea; color: #c62828; }
        .duplicate-label { background: #fff8e1; color: #e65100; margin-top: 10px; }
        .analysis-detail-table { width: 100%; border-collapse: collapse; font-size: 0.84em; margin-bottom: 4px; }
        .analysis-detail-table th { background: #e5f1fb; color: #005a9e; padding: 5px 10px; text-align: left; font-weight: 600; }
        .analysis-detail-table td { padding: 5px 10px; border-bottom: 1px solid #c7e0f4; vertical-align: top; word-break: break-word; }
        .analysis-no-data { color: #999; font-style: italic; font-size: 0.9em; padding: 10px 0; }
        /* ===== Main Tab Navigation ===== */
        .main-tabs { display:flex; gap:0; padding:0 20px; border-bottom:3px solid #0078d4; background:#fff; position:sticky; top:0; z-index:10; box-shadow:0 2px 8px rgba(0,0,0,0.08); }
        .main-tab-btn { padding:13px 30px; border:none; border-radius:6px 6px 0 0; cursor:pointer; font-size:13px; font-weight:600; background:transparent; color:#0078d4; margin-right:2px; transition:background-color 0.2s, color 0.2s; font-family:inherit; letter-spacing:0.2px; }
        .main-tab-btn:hover { background:#e5f1fb; color:#005a9e; }
        .main-tab-btn.active { background:#0078d4; color:white; box-shadow:inset 0 -3px 0 rgba(255,255,255,0.3); }
        .main-tab-btn.active:hover { background:#005a9e; color:white; }
        .expand-group-label { font-size:0.78em; color:#888; font-weight:600; text-transform:uppercase; letter-spacing:0.5px; padding:0 2px; align-self:center; }
        .expand-group-divider { width:1px; background:#e0e0e0; align-self:stretch; margin:4px 4px; }
        .main-panel { display:none; }
        .main-panel.active { display:block; }
'@

    $analysisJS = @'
        // ===== Settings Analysis =====
        function switchAnalysisTab(platform) {
            document.querySelectorAll('.analysis-tab-btn').forEach(function(btn) {
                btn.classList.toggle('active', btn.getAttribute('data-platform') === platform);
            });
            document.querySelectorAll('.analysis-platform-panel').forEach(function(panel) {
                panel.classList.toggle('active', panel.getAttribute('data-platform') === platform);
            });
        }
        function toggleAnalysisRow(triggerEl, rowId) {
            var detailRow = document.getElementById(rowId);
            if (!detailRow) return;
            var isVisible = detailRow.classList.contains('visible');
            detailRow.classList.toggle('visible', !isVisible);
            triggerEl.innerHTML = isVisible ? '&#9654; Details' : '&#9660; Hide';
        }
        function filterAnalysisPolicies(panelId, btn) {
            var panel = document.getElementById(panelId);
            if (!panel) return;
            var filter = btn.getAttribute('data-filter');
            var isActive = btn.classList.contains('filter-active');
            panel.querySelectorAll('.analysis-filter-pill[data-filter]').forEach(function(b) { b.classList.remove('filter-active'); });
            var activeFilter = isActive ? 'all' : filter;
            if (!isActive) btn.classList.add('filter-active');
            panel.querySelectorAll('tr.analysis-policy-row').forEach(function(row) {
                var rowFilter = row.getAttribute('data-filter');
                var show = activeFilter === 'all' || rowFilter === activeFilter || rowFilter === 'both';
                row.style.display = show ? '' : 'none';
                if (!show) {
                    var detailId = row.getAttribute('data-detail-id');
                    if (detailId) {
                        var dr = document.getElementById(detailId);
                        if (dr) { dr.style.display = 'none'; dr.classList.remove('visible'); }
                    }
                }
            });
        }
        function switchMainTab(tabName) {
            document.querySelectorAll('.main-tab-btn').forEach(function(btn) {
                btn.classList.toggle('active', btn.getAttribute('data-tab') === tabName);
            });
            document.querySelectorAll('.main-panel').forEach(function(panel) {
                panel.classList.toggle('active', panel.id === 'main-panel-' + tabName);
            });
        }
'@

    # Build HTML
    $backupDateString = $metadata.BackupDate
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$ReportTitle - $backupDateString</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 0;
            background: #f0f4f8;
            min-height: 100vh;
        }
        .container {
            max-width: 1400px;
            margin: 0 auto;
            padding: 20px;
        }
        .header {
            background: white;
            border-radius: 10px;
            padding: 30px;
            margin-bottom: 0;
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        }
        .header h1 {
            color: #0078d4;
            margin: 0 0 6px 0;
            font-size: 2em;
            font-weight: 600;
        }
        .header .subtitle {
            color: #555;
            font-size: 1em;
        }
        .toolbar {
            display: flex;
            gap: 12px;
            padding: 14px 20px;
            flex-wrap: wrap;
            align-items: center;
            background: white;
            box-shadow: 0 2px 8px rgba(0,0,0,0.06);
            margin-bottom: 20px;
        }
        .search-box {
            flex: 1;
            min-width: 250px;
            padding: 10px 15px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            font-size: 1em;
            outline: none;
            transition: border-color 0.2s;
        }
        .search-box:focus {
            border-color: #0078d4;
        }
        .platform-filters {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }
        .platform-btn {
            padding: 8px 16px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            background: white;
            cursor: pointer;
            font-size: 0.9em;
            transition: border-color 0.2s, background 0.2s, color 0.2s;
        }
        .platform-btn:hover {
            border-color: #0078d4;
        }
        .platform-btn.active {
            background: #0078d4;
            color: white;
            border-color: #0078d4;
        }
        .expand-controls {
            display: flex;
            gap: 8px;
        }
        .expand-btn {
            padding: 8px 14px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            background: white;
            cursor: pointer;
            font-size: 0.85em;
            transition: all 0.2s;
        }
        .expand-btn:hover {
            border-color: #0078d4;
            background: #f0f7ff;
        }
        .metadata {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 15px;
            margin-top: 20px;
        }
        .metadata-item {
            background: #f8f9fa;
            padding: 15px;
            border-radius: 8px;
            border-left: 4px solid #0078d4;
        }
        .metadata-item .label {
            color: #666;
            font-size: 0.9em;
            margin-bottom: 5px;
        }
        .metadata-item .value {
            color: #333;
            font-size: 1.3em;
            font-weight: bold;
        }
        .component-section {
            background: white;
            border-radius: 10px;
            padding: 25px;
            margin-bottom: 20px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.08);
        }
        .component-section h2 {
            background: #0078d4;
            color: white;
            margin: -25px -25px 15px -25px;
            padding: 14px 20px;
            border-radius: 8px 8px 0 0;
            font-size: 1em;
            font-weight: 600;
            cursor: pointer;
            display: flex;
            justify-content: space-between;
            align-items: center;
            user-select: none;
        }
        .component-section h2:hover { background: #005a9e; }
        .section-toggle-icon { font-size: 0.85em; transition: transform 0.2s; flex-shrink: 0; }
        .component-section.collapsed .section-toggle-icon { transform: rotate(-90deg); }
        .section-cards { display: block; }
        .component-section.collapsed .section-cards { display: none; }
        .policy-card {
            border: 1px solid #e8e8e8;
            border-radius: 8px;
            margin-bottom: 10px;
            overflow: hidden;
            transition: box-shadow 0.2s;
        }
        .policy-card:hover {
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        .policy-header {
            padding: 15px 20px;
            cursor: pointer;
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: #fafafa;
            transition: background 0.2s;
        }
        .policy-header:hover {
            background: #f0f7ff;
        }
        .policy-name {
            font-weight: 600;
            color: #333;
            font-size: 1em;
        }
        .policy-meta {
            display: flex;
            gap: 12px;
            font-size: 0.85em;
            color: #666;
            margin-top: 4px;
            flex-wrap: wrap;
        }
        .policy-meta span {
            white-space: nowrap;
        }
        .policy-info {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 10px;
            padding: 15px 20px;
            background: #f8f9fa;
            border-bottom: 1px solid #e8e8e8;
        }
        .policy-info-item {
            display: flex;
            flex-direction: column;
        }
        .policy-info-item .info-label {
            font-size: 0.75em;
            color: #888;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 2px;
        }
        .policy-info-item .info-value {
            font-size: 0.9em;
            color: #333;
            word-break: break-all;
        }
        .platform-badge {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 4px;
            font-size: 0.75em;
            font-weight: 600;
            color: white;
        }
        .platform-windows { background: #0078d4; }
        .platform-ios { background: #8e8e93; }
        .platform-android { background: #3ddc84; color: #1b5e20; }
        .platform-macos { background: #005a9e; }
        .toggle-icon {
            font-size: 1.2em;
            color: #999;
            transition: transform 0.2s;
            flex-shrink: 0;
            margin-left: 10px;
        }
        .policy-card.expanded .toggle-icon {
            transform: rotate(90deg);
        }
        .policy-details {
            display: none;
            padding: 0;
        }
        .policy-card.expanded .policy-details {
            display: block;
        }
        .diff-table {
            width: 100%;
            border-collapse: collapse;
        }
        .diff-table th {
            background: #0078d4;
            color: white;
            padding: 10px 12px;
            text-align: left;
            font-weight: 500;
            font-size: 0.9em;
        }
        .diff-table td {
            padding: 10px 12px;
            border-bottom: 1px solid #f0f0f0;
            word-break: break-word;
            font-size: 0.9em;
        }
        .diff-table tr:hover {
            background: #f8f9fa;
        }
        .desc-cell {
            max-width: 300px;
            font-size: 0.85em;
            color: #555;
            cursor: pointer;
        }
        .desc-cell .desc-short { display: inline; }
        .desc-cell .desc-full { display: none; }
        .desc-cell.expanded .desc-short { display: none; }
        .desc-cell.expanded .desc-full { display: inline; }
        .desc-toggle {
            color: #0078d4;
            font-size: 0.8em;
            margin-left: 4px;
        }
        .no-settings {
            padding: 15px 20px;
            color: #999;
            font-style: italic;
            font-size: 0.9em;
        }
        .settings-toolbar {
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 8px 12px;
            background: #f8f9fa;
            border-radius: 6px;
            margin-bottom: 8px;
        }
        .settings-filter-toggle {
            display: flex;
            align-items: center;
            gap: 6px;
            cursor: pointer;
            font-size: 0.85em;
            color: #555;
            user-select: none;
        }
        .settings-filter-toggle input[type="checkbox"] {
            accent-color: #0078d4;
            cursor: pointer;
        }
        .settings-filter-toggle .filter-count {
            color: #0078d4;
            font-weight: 600;
        }
        .diff-table tr.not-configured { }
        .diff-table tr.not-configured.filter-hidden { display: none; }
        .settings-count {
            background: #0078d4;
            color: white;
            padding: 2px 8px;
            border-radius: 4px;
            font-size: 0.75em;
            font-weight: 600;
        }
        .summary {
            background: white;
            border-radius: 10px;
            padding: 25px;
            margin-top: 30px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.08);
        }
        .summary h3 {
            color: #333;
            margin: 0 0 15px 0;
        }
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 15px;
        }
        .summary-item {
            text-align: center;
            padding: 20px;
            background: #f8f9fa;
            border-radius: 8px;
        }
        .summary-item .count {
            font-size: 2em;
            font-weight: bold;
            color: #0078d4;
        }
        .summary-item .label {
            color: #666;
            margin-top: 5px;
        }
        .footer {
            text-align: center;
            color: #555;
            padding: 18px 30px;
            margin-top: 24px;
            font-size: 0.88em;
            background: white;
            border-radius: 10px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        }
        .hidden { display: none !important; }
        .platform-hidden { display: none !important; }
    </style>
    <script>
        function filterReport() {
            var query = document.getElementById('searchBox').value.toLowerCase();
            var cards = document.querySelectorAll('.policy-card');
            var sections = document.querySelectorAll('.component-section');

            cards.forEach(function(card) {
                var text = card.textContent.toLowerCase();
                if (query === '' || text.indexOf(query) !== -1) {
                    card.classList.remove('hidden');
                } else {
                    card.classList.add('hidden');
                }
            });

            // Hide sections with no visible cards
            sections.forEach(function(section) {
                var visibleCards = section.querySelectorAll('.policy-card:not(.hidden)');
                if (visibleCards.length === 0) {
                    section.classList.add('hidden');
                } else {
                    section.classList.remove('hidden');
                }
            });
        }

        var activePlatform = 'All';
        function filterByPlatform(platform) {
            activePlatform = platform;
            document.querySelectorAll('.platform-btn').forEach(function(btn) {
                btn.classList.toggle('active', btn.getAttribute('data-platform') === platform);
            });

            var cards = document.querySelectorAll('.policy-card');
            var sections = document.querySelectorAll('.component-section');

            cards.forEach(function(card) {
                var cardPlatform = card.getAttribute('data-platform');
                if (platform === 'All' || cardPlatform === platform) {
                    card.classList.remove('platform-hidden');
                } else {
                    card.classList.add('platform-hidden');
                }
                // Combine with search filter
                applyFilters(card);
            });

            sections.forEach(function(section) {
                var visibleCards = section.querySelectorAll('.policy-card:not(.hidden):not(.platform-hidden)');
                if (visibleCards.length === 0) {
                    section.classList.add('hidden');
                } else {
                    section.classList.remove('hidden');
                }
            });
        }

        function applyFilters(card) {
            var query = document.getElementById('searchBox').value.toLowerCase();
            var text = card.textContent.toLowerCase();
            var cardPlatform = card.getAttribute('data-platform');
            var searchMatch = (query === '' || text.indexOf(query) !== -1);
            var platformMatch = (activePlatform === 'All' || cardPlatform === activePlatform);

            if (searchMatch) { card.classList.remove('hidden'); } else { card.classList.add('hidden'); }
            if (platformMatch) { card.classList.remove('platform-hidden'); } else { card.classList.add('platform-hidden'); }
        }

        function filterAll() {
            document.querySelectorAll('.policy-card').forEach(function(card) { applyFilters(card); });
            document.querySelectorAll('.component-section').forEach(function(section) {
                var visibleCards = section.querySelectorAll('.policy-card:not(.hidden):not(.platform-hidden)');
                section.classList.toggle('hidden', visibleCards.length === 0);
            });
        }

        function togglePolicy(el) {
            var card = el.closest('.policy-card');
            card.classList.toggle('expanded');
        }

        function toggleSection(h2el) {
            h2el.parentElement.classList.toggle('collapsed');
        }

        function toggleAllSections(expand) {
            document.querySelectorAll('.component-section').forEach(function(section) {
                if (expand) { section.classList.remove('collapsed'); } else { section.classList.add('collapsed'); }
            });
        }

        function toggleAll(expand) {
            document.querySelectorAll('.policy-card').forEach(function(card) {
                if (expand) { card.classList.add('expanded'); } else { card.classList.remove('expanded'); }
            });
        }

        function toggleDesc(el) { el.closest('.desc-cell').classList.toggle('expanded'); }

        function toggleConfigured(checkbox) {
            var card = checkbox.closest('.policy-details');
            var rows = card.querySelectorAll('.diff-table tr.not-configured');
            rows.forEach(function(row) {
                if (checkbox.checked) {
                    row.classList.add('filter-hidden');
                } else {
                    row.classList.remove('filter-hidden');
                }
            });
        }
    </script>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>$ReportTitle</h1>
        <div class="subtitle">$ReportSubtitle</div>
        <div class="metadata">
            <div class="metadata-item">
                <div class="label">$(if ($metadata.ReportDateLabel) { $metadata.ReportDateLabel } else { "Backup Date" })</div>
                <div class="value">$($metadata.BackupDate)</div>
            </div>
            <div class="metadata-item">
                <div class="label">Tenant ID</div>
                <div class="value">$($metadata.TenantId)</div>
            </div>
            <div class="metadata-item">
                <div class="label">Total Policies</div>
                <div class="value">$totalPolicies</div>
            </div>
            <div class="metadata-item">
                <div class="label">Total Settings</div>
                <div class="value">$totalSettings</div>
            </div>
            <div class="metadata-item">
                <div class="label">Duration</div>
                <div class="value">$($metadata.Duration)</div>
            </div>
        </div>
    </div>
"@

    # Inject analysis CSS and JS into the <head> block
    $html = $html.Replace('    </style>', "$analysisCSS`n    </style>")
    $html = $html.Replace('    </script>', "$analysisJS`n    </script>")

    # Main tab strip + open Policy Configuration panel (toolbar lives here)
    $html += @"
<div class="main-tabs">
    <button class="main-tab-btn active" data-tab="policies" onclick="switchMainTab('policies')">Policy Configuration</button>
    <button class="main-tab-btn" data-tab="analysis" onclick="switchMainTab('analysis')">Settings Analysis</button>
</div>
<div class="main-panel active" id="main-panel-policies">
    <div class="toolbar">
        <input type="text" id="searchBox" class="search-box" placeholder="Search policies or settings..." oninput="filterAll()">
        <div class="platform-filters">
            <button class="platform-btn active" data-platform="All" onclick="filterByPlatform('All')">All</button>
            <button class="platform-btn" data-platform="Windows" onclick="filterByPlatform('Windows')">Windows</button>
            <button class="platform-btn" data-platform="iOS" onclick="filterByPlatform('iOS')">iOS</button>
            <button class="platform-btn" data-platform="Android" onclick="filterByPlatform('Android')">Android</button>
            <button class="platform-btn" data-platform="macOS" onclick="filterByPlatform('macOS')">macOS</button>
        </div>
        <div class="expand-controls">
            <span class="expand-group-label">Policies</span>
            <button class="expand-btn" onclick="toggleAll(true)">Expand All</button>
            <button class="expand-btn" onclick="toggleAll(false)">Collapse All</button>
            <div class="expand-group-divider"></div>
            <span class="expand-group-label">Sections</span>
            <button class="expand-btn" onclick="toggleAllSections(true)">Expand All</button>
            <button class="expand-btn" onclick="toggleAllSections(false)">Collapse All</button>
        </div>
    </div>
"@

    # Generate per-section HTML
    foreach ($section in $allSections) {
        $sectionCount = $section.Cards.Count
        $html += @"

    <div class="component-section">
        <h2 onclick="toggleSection(this)">$($section.DisplayName) ($sectionCount) <span class="section-toggle-icon">&#9660;</span></h2>
        <div class="section-cards">
"@

        foreach ($card in $section.Cards) {
            $encodedName = HtmlEncode $card.Name
            $encodedId = HtmlEncode $card.Id
            $platformClass = "platform-$($card.Platform.ToLower())"
            $settingsCount = $card.Settings.Count

            $html += @"
        <div class="policy-card" data-platform="$($card.Platform)">
            <div class="policy-header" onclick="togglePolicy(this)">
                <div>
                    <div class="policy-name">$encodedName</div>
                    <div class="policy-meta">
                        <span>ID: $encodedId</span>
                        <span>Created: $($card.Created)</span>
                        <span>Modified: $($card.Modified)</span>
                        <span>Assigned: $($card.IsAssigned)</span>
                        <span><span class="platform-badge $platformClass">$($card.Platform)</span></span>
                        <span><span class="settings-count">$settingsCount settings</span></span>
                    </div>
                </div>
                <span class="toggle-icon">&#9654;</span>
            </div>
            <div class="policy-details">
                <div class="policy-info">
                    <div class="policy-info-item">
                        <span class="info-label">Name</span>
                        <span class="info-value">$encodedName</span>
                    </div>
                    <div class="policy-info-item">
                        <span class="info-label">ID</span>
                        <span class="info-value">$encodedId</span>
                    </div>
                    <div class="policy-info-item">
                        <span class="info-label">Created</span>
                        <span class="info-value">$($card.Created)</span>
                    </div>
                    <div class="policy-info-item">
                        <span class="info-label">Modified</span>
                        <span class="info-value">$($card.Modified)</span>
                    </div>
                    <div class="policy-info-item">
                        <span class="info-label">Assigned</span>
                        <span class="info-value">$($card.IsAssigned)</span>
                    </div>
                    <div class="policy-info-item">
                        <span class="info-label">Platform</span>
                        <span class="info-value"><span class="platform-badge $platformClass">$($card.Platform)</span></span>
                    </div>
                </div>
"@

            if ($card.Settings.Count -gt 0) {
                $configuredCount = ($card.Settings | Where-Object { $_.Value -ne '(not set)' -and $_.Value -ne '(not configured)' -and -not [string]::IsNullOrEmpty($_.Value) }).Count
                $notConfiguredCount = $card.Settings.Count - $configuredCount
                $html += @"
                <div class="settings-toolbar">
                    <label class="settings-filter-toggle">
                        <input type="checkbox" onchange="toggleConfigured(this)"> Show only configured settings (<span class="filter-count">$configuredCount</span> of $($card.Settings.Count))
                    </label>
                </div>
                <table class="diff-table">
                    <thead>
                        <tr>
                            <th>Setting</th>
                            <th>Configured Value</th>
                            <th>Description</th>
                        </tr>
                    </thead>
                    <tbody>
"@
                foreach ($setting in $card.Settings) {
                    $encodedSettingName = HtmlEncode $setting.Name
                    $encodedValue = HtmlEncode $setting.Value

                    $descHtml = ''
                    if ($setting.Description) {
                        $encodedDesc = HtmlEncode $setting.Description
                        $shortLen = [Math]::Min(80, $encodedDesc.Length)
                        if ($encodedDesc.Length -gt 80) {
                            $shortText = $encodedDesc.Substring(0, $shortLen) + '...'
                            $descHtml = "<td class=`"desc-cell`" onclick=`"toggleDesc(this)`"><span class=`"desc-short`">$shortText <span class=`"desc-toggle`">[more]</span></span><span class=`"desc-full`">$encodedDesc <span class=`"desc-toggle`">[less]</span></span></td>"
                        } else {
                            $descHtml = "<td class=`"desc-cell`">$encodedDesc</td>"
                        }
                    } else {
                        $descHtml = '<td class="desc-cell"></td>'
                    }

                    $rowClass = ''
                    if ($setting.Value -eq '(not set)' -or $setting.Value -eq '(not configured)' -or [string]::IsNullOrEmpty($setting.Value)) {
                        $rowClass = ' class="not-configured"'
                    }
                    $html += "                        <tr$rowClass><td>$encodedSettingName</td><td>$encodedValue</td>$descHtml</tr>`n"
                }

                $html += @"
                    </tbody>
                </table>
"@
            } else {
                $html += '                <div class="no-settings">No configurable settings found</div>' + "`n"
            }

            $html += @"
            </div>
        </div>
"@
        }

        $html += @"
        </div>
    </div>
"@
    }

    # Close Policy Configuration panel; open/fill/close Settings Analysis panel
    $html += "</div>`n"  # closes main-panel-policies
    $html += "<div class=`"main-panel`" id=`"main-panel-analysis`">`n"
    $html += Render-AnalysisSectionHTML -PlatformAnalysis $platformAnalysis
    $html += "</div>`n"  # closes main-panel-analysis

    # Add summary section
    $html += @"

    <div class="summary">
        <h3>Backup Summary</h3>
        <div class="summary-grid">
"@

    foreach ($type in $summaryCounts.Keys) {
        $html += @"
            <div class="summary-item">
                <div class="count">$($summaryCounts[$type])</div>
                <div class="label">$type</div>
            </div>
"@
    }

    $html += @"
        </div>
    </div>
    <div class="footer">
        <p>Generated by UniFy-Endpoint Script v$($script:Version) &nbsp;|&nbsp; $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
    </div>
</div>
</body>
</html>
"@

    # Save HTML file
    $html | Out-File -FilePath $OutputPath -Encoding UTF8
    Write-Log "HTML report generated: $OutputPath" -Level Info
}

function Import-ConfigurationFromFile {
    param (
        [string]$ImportPath,
        [switch]$Preview
    )

    Write-ColoredOutput "`nAnalyzing import source..." -Color $script:Colors.Info

    if (-not (Test-Path $ImportPath)) {
        Write-ColoredOutput "Import path not found: $ImportPath" -Color $script:Colors.Error
        Write-Log "Import failed: Path not found" -Level Error
        return
    }

    $importType = $null
    $policiesToImport = @()

    try {
        # Determine import type: Directory (backup folder), JSON export file, or individual JSON policy
        if (Test-Path $ImportPath -PathType Container) {
            # It's a directory - could be a backup folder
            $importType = "BackupFolder"
            Write-ColoredOutput "Detected: Backup folder" -Color $script:Colors.Info

            # Check for metadata.json to confirm it's a backup
            $metadataPath = Join-Path $ImportPath "metadata.json"
            if (Test-Path $metadataPath) {
                $metadata = Get-Content $metadataPath | ConvertFrom-Json
                Write-ColoredOutput "Backup Date: $($metadata.BackupDate)" -Color $script:Colors.Info
                Write-ColoredOutput "Total Items: $($metadata.TotalItems)" -Color $script:Colors.Info
            }

            # Scan for policy JSON files in subdirectories
            $policyFolders = Get-ChildItem -Path $ImportPath -Directory
            foreach ($folder in $policyFolders) {
                $jsonFiles = Get-ChildItem -Path $folder.FullName -Filter "*.json" -File
                foreach ($jsonFile in $jsonFiles) {
                    # Skip metadata.json
                    if ($jsonFile.Name -eq "metadata.json") { continue }
                    try {
                        $policy = Get-Content $jsonFile.FullName | ConvertFrom-Json
                        $policyName = if ($policy.displayName) { $policy.displayName } else { $policy.name }
                        if ($policyName) {
                            $policiesToImport += [PSCustomObject]@{
                                Name = $policyName
                                Type = $folder.Name
                                FilePath = $jsonFile.FullName
                                Policy = $policy
                            }
                        }
                    }
                    catch {
                        Write-ColoredOutput "  Warning: Could not parse $($jsonFile.Name)" -Color $script:Colors.Warning
                    }
                }
            }

            # Also scan for JSON files directly in the root folder (for non-UniFy-Endpoint folder structures)
            $rootJsonFiles = Get-ChildItem -Path $ImportPath -Filter "*.json" -File
            if ($rootJsonFiles.Count -gt 0) {
                $folderName = Split-Path $ImportPath -Leaf
                foreach ($jsonFile in $rootJsonFiles) {
                    # Skip metadata.json
                    if ($jsonFile.Name -eq "metadata.json") { continue }
                    try {
                        $policy = Get-Content $jsonFile.FullName | ConvertFrom-Json
                        $policyName = if ($policy.displayName) { $policy.displayName } else { $policy.name }
                        if ($policyName) {
                            $policiesToImport += [PSCustomObject]@{
                                Name = $policyName
                                Type = $folderName
                                FilePath = $jsonFile.FullName
                                Policy = $policy
                            }
                        }
                    }
                    catch {
                        Write-ColoredOutput "  Warning: Could not parse $($jsonFile.Name)" -Color $script:Colors.Warning
                    }
                }
            }
        }
        else {
            # It's a file - check if it's an export JSON or individual policy JSON
            $content = Get-Content $ImportPath -Raw
            $jsonData = $content | ConvertFrom-Json

            if ($jsonData.Policies) {
                # It's an exported JSON file with Policies structure
                $importType = "ExportFile"
                Write-ColoredOutput "Detected: UniFy-Endpoint export file" -Color $script:Colors.Info

                if ($jsonData.ExportDate) {
                    Write-ColoredOutput "Export Date: $($jsonData.ExportDate)" -Color $script:Colors.Info
                }
                if ($jsonData.Metadata) {
                    Write-ColoredOutput "Original Backup: $($jsonData.Metadata.BackupDate)" -Color $script:Colors.Info
                }

                foreach ($policyType in $jsonData.Policies.PSObject.Properties.Name) {
                    foreach ($policy in @($jsonData.Policies.$policyType)) {
                        $policyName = if ($policy.displayName) { $policy.displayName } else { $policy.name }
                        if ($policyName) {
                            $policiesToImport += [PSCustomObject]@{
                                Name = $policyName
                                Type = $policyType
                                FilePath = $ImportPath
                                Policy = $policy
                            }
                        }
                    }
                }
            }
            elseif ($jsonData.displayName -or $jsonData.name -or $jsonData.'@odata.type') {
                # It's an individual policy JSON file
                $importType = "SinglePolicy"
                $policyName = if ($jsonData.displayName) { $jsonData.displayName } else { $jsonData.name }
                Write-ColoredOutput "Detected: Individual policy file" -Color $script:Colors.Info
                Write-ColoredOutput "Policy Name: $policyName" -Color $script:Colors.Info

                # Try to determine policy type from @odata.type
                $policyType = "Unknown"
                if ($jsonData.'@odata.type') {
                    $odataType = $jsonData.'@odata.type'
                    if ($odataType -match 'deviceManagementIntent') { $policyType = "LegacyIntents" }
                    elseif ($odataType -match 'deviceConfiguration') { $policyType = "DeviceConfigurations" }
                    elseif ($odataType -match 'deviceCompliancePolicy') { $policyType = "CompliancePolicies" }
                    elseif ($odataType -match 'deviceManagementConfigurationPolicy') { $policyType = "SettingsCatalogPolicies" }
                    elseif ($odataType -match 'managedAppProtection') { $policyType = "AppProtectionPolicies" }
                    elseif ($odataType -match 'deviceManagementScript') { $policyType = "PowerShellScripts" }
                }

                $policiesToImport += [PSCustomObject]@{
                    Name = $policyName
                    Type = $policyType
                    FilePath = $ImportPath
                    Policy = $jsonData
                }
            }
            else {
                Write-ColoredOutput "Unrecognized file format. Expected:" -Color $script:Colors.Error
                Write-ColoredOutput "  - UniFy-Endpoint export file (with 'Policies' property)" -Color $script:Colors.Default
                Write-ColoredOutput "  - Individual Intune policy JSON file" -Color $script:Colors.Default
                Write-ColoredOutput "  - Backup folder path" -Color $script:Colors.Default
                Write-Log "Import failed: Unrecognized file format" -Level Error
                return
            }
        }

        if ($policiesToImport.Count -eq 0) {
            Write-ColoredOutput "No policies found to import." -Color $script:Colors.Warning
            return
        }

        # Display policies and allow multi-select
        Write-ColoredOutput "`n════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
        Write-ColoredOutput " AVAILABLE POLICIES TO IMPORT" -Color $script:Colors.Info
        Write-ColoredOutput "════════════════════════════════════════════════════════════════" -Color $script:Colors.Header

        $index = 1
        $groupedPolicies = $policiesToImport | Group-Object -Property Type
        foreach ($group in $groupedPolicies) {
            Write-ColoredOutput "`n  $($group.Name):" -Color $script:Colors.Info
            foreach ($policy in $group.Group) {
                Write-Host "    [$index]" -ForegroundColor Yellow -NoNewline
                Write-Host " $($policy.Name)" -ForegroundColor White
                $index++
            }
        }

        Write-ColoredOutput "`n  [A] Select all policies" -Color $script:Colors.Success
        Write-ColoredOutput "  [0] Cancel" -Color $script:Colors.Error

        $selection = Read-Host "`nSelect policies to import (comma-separated like 1,3,5 or A for all, or range like 1-5)"

        if ($selection -eq "0") {
            Write-ColoredOutput "Import cancelled." -Color $script:Colors.Warning
            return
        }

        $selectedPolicies = @()

        if ($selection -eq "A" -or $selection -eq "a") {
            $selectedPolicies = $policiesToImport
        }
        else {
            # Parse selection (support comma-separated and ranges)
            $indices = @()
            $parts = $selection -split ','
            foreach ($part in $parts) {
                $part = $part.Trim()
                if ($part -match '^(\d+)-(\d+)$') {
                    # Range like 1-5
                    $start = [int]$Matches[1]
                    $end = [int]$Matches[2]
                    for ($i = $start; $i -le $end; $i++) {
                        $indices += $i
                    }
                }
                elseif ($part -match '^\d+$') {
                    $indices += [int]$part
                }
            }

            foreach ($idx in $indices) {
                $num = $idx - 1
                if ($num -ge 0 -and $num -lt $policiesToImport.Count) {
                    $selectedPolicies += $policiesToImport[$num]
                }
            }
        }

        if ($selectedPolicies.Count -eq 0) {
            Write-ColoredOutput "No policies selected." -Color $script:Colors.Warning
            return
        }

        Write-ColoredOutput "`nSelected $($selectedPolicies.Count) policy(s) for import" -Color $script:Colors.Info

        if ($Preview) {
            Write-ColoredOutput "`n[PREVIEW MODE] No changes will be made" -Color $script:Colors.Warning
            foreach ($policy in $selectedPolicies) {
                Write-ColoredOutput "  [WOULD IMPORT] $($policy.Name) ($($policy.Type))" -Color $script:Colors.Info
            }
            return
        }

        # Confirm import
        Write-ColoredOutput "`nWARNING: This will create $($selectedPolicies.Count) new policy(s) in your Intune tenant!" -Color $script:Colors.Warning
        Write-ColoredOutput "Policies will be created with '- [New Policy]' or '- [Imported]' suffix." -Color $script:Colors.Info
        $confirm = Read-Host "Are you sure you want to continue? (yes/no)"
        if ($confirm -ne "yes") {
            Write-ColoredOutput "Import cancelled." -Color $script:Colors.Warning
            Write-Log "Import cancelled by user" -Level Info
            return
        }

        # Create temporary backup structure for selected policies
        $tempBackupPath = Join-Path $env:TEMP "UniFy-Endpoint_Import_$(Get-Date -Format 'yyyyMMddHHmmss')"
        New-Item -ItemType Directory -Path $tempBackupPath -Force | Out-Null

        # Valid UniFy-Endpoint folder names that Restore-IntuneConfiguration recognizes
        $validFolderNames = @(
            "DeviceConfigurations", "CompliancePolicies", "SettingsCatalogPolicies",
            "AdministrativeTemplates", "AppProtectionPolicies", "PowerShellScripts",
            "RemediationScripts", "MacOSScripts", "MacOSCustomAttributes",
            "AutopilotProfiles", "AutopilotDevicePrep", "EnrollmentStatusPage",
            "WUfBPolicies", "AssignmentFilters", "AppConfigManagedDevices", "AppConfigManagedApps",
            "LegacyIntents"
        )

        foreach ($policy in $selectedPolicies) {
            # Use existing Type if it's a valid UniFy-Endpoint folder name, otherwise detect from JSON
            $folderName = if ($validFolderNames -contains $policy.Type) {
                $policy.Type
            } else {
                Get-PolicyTypeFromJson -Policy $policy.Policy
            }

            $typePath = Join-Path $tempBackupPath $folderName
            if (-not (Test-Path $typePath)) {
                New-Item -ItemType Directory -Path $typePath -Force | Out-Null
            }

            $fileName = Get-SafeFileName -FileName "$($policy.Name)_imported"
            $filePath = Join-Path $typePath "$fileName.json"
            $policy.Policy | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding UTF8
        }

        # Save metadata
        $tempMetadata = @{
            BackupDate = Get-Date -Format "yyyy-MM-dd-HHmmss"
            TotalItems = $selectedPolicies.Count
        }
        $tempMetadata | ConvertTo-Json -Depth 10 | Out-File -FilePath (Join-Path $tempBackupPath "metadata.json") -Encoding UTF8

        # Use existing restore function
        Write-ColoredOutput "`nImporting policies..." -Color $script:Colors.Info
        $script:SelectedPlatform = "All"
        $SelectedComponents = Get-SelectedComponents -ComponentList @("All")
        Restore-IntuneConfiguration -BackupPath $tempBackupPath -SelectedComponents $SelectedComponents -IsImport

        # Cleanup temp files
        Remove-Item -Path $tempBackupPath -Recurse -Force

        Write-ColoredOutput "`nImport completed successfully" -Color $script:Colors.Success
        Write-Log "Import completed successfully" -Level Info
    }
    catch {
        Write-ColoredOutput "Import failed: $_" -Color $script:Colors.Error
        Write-Log "Import failed: $_" -Level Error
    }
}

#endregion

#region Cleanup Functions

function Remove-OldBackups {
    param (
        [int]$RetentionDays
    )

    Write-ColoredOutput "`nCleaning up backups older than $RetentionDays days..." -Color $script:Colors.Info

    $cutoffDate = (Get-Date).AddDays(-$RetentionDays)
    $backups = Get-AvailableBackups
    $removedCount = 0

    foreach ($backup in $backups) {
        $backupDate = [DateTime]::ParseExact($backup.Date, "yyyy-MM-dd-HHmmss", $null)

        if ($backupDate -lt $cutoffDate) {
            Write-ColoredOutput "  Removing: $($backup.Name) (Created: $($backup.Date))" -Color $script:Colors.Warning
            Remove-Item -Path $backup.Path -Recurse -Force
            $removedCount++
        }
    }

    if ($removedCount -gt 0) {
        Write-ColoredOutput "Removed $removedCount old backup(s)" -Color $script:Colors.Success
    }
    else {
        Write-ColoredOutput "No old backups to remove" -Color $script:Colors.Info
    }
}

function Remove-OldLogs {
    param (
        [int]$RetentionDays
    )

    Write-ColoredOutput "`nCleaning up log files older than $RetentionDays days..." -Color $script:Colors.Info

    if (-not (Test-Path $script:LogLocation)) {
        Write-ColoredOutput "Log folder does not exist: $($script:LogLocation)" -Color $script:Colors.Warning
        return
    }

    $cutoffDate = (Get-Date).AddDays(-$RetentionDays)
    $logFiles = Get-ChildItem -Path $script:LogLocation -Filter "*.log" -File | Where-Object { $_.LastWriteTime -lt $cutoffDate }
    $removedCount = 0
    $totalSize = 0

    foreach ($logFile in $logFiles) {
        $totalSize += $logFile.Length
        Write-ColoredOutput "  Removing: $($logFile.Name) (Last Modified: $($logFile.LastWriteTime.ToString('yyyy-MM-dd HH:mm')))" -Color $script:Colors.Warning
        Remove-Item -Path $logFile.FullName -Force
        $removedCount++
    }

    if ($removedCount -gt 0) {
        $sizeInMB = [math]::Round($totalSize / 1MB, 2)
        Write-ColoredOutput "Removed $removedCount old log file(s) ($sizeInMB MB freed)" -Color $script:Colors.Success
    }
    else {
        Write-ColoredOutput "No old log files to remove" -Color $script:Colors.Info
    }

    # Show remaining log files
    $remainingLogs = Get-ChildItem -Path $script:LogLocation -Filter "*.log" -File
    if ($remainingLogs.Count -gt 0) {
        Write-ColoredOutput "`nRemaining log files: $($remainingLogs.Count)" -Color $script:Colors.Info
    }
}

#endregion

#region Backup & Restore Helper Functions

function Get-AvailableBackups {
    $backups = @()

    if (Test-Path $script:BackupLocation) {
        $backupFolders = Get-ChildItem -Path $script:BackupLocation -Directory | Where-Object { $_.Name -like "backup-*" }

        foreach ($folder in $backupFolders) {
            $metadataPath = Join-Path $folder.FullName "metadata.json"
            if (Test-Path $metadataPath) {
                $metadata = Get-Content $metadataPath | ConvertFrom-Json
                $backups += [PSCustomObject]@{
                    Name       = $folder.Name
                    Path       = $folder.FullName
                    Date       = $metadata.BackupDate
                    TotalItems = $metadata.TotalItems
                    Duration   = $metadata.Duration
                    TenantId   = $metadata.TenantId
                }
            }
        }
    }

    # Force array return to ensure .Count works in PowerShell 5.1
    return @($backups | Sort-Object Date -Descending)
}

function Show-BackupList {
    Write-Host ""
    Write-Host "  AVAILABLE BACKUPS" -ForegroundColor Cyan
    Write-Host "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan

    $backups = @(Get-AvailableBackups)

    if ($backups.Count -eq 0) {
        Write-ColoredOutput "No backups found in $script:BackupLocation" -Color $script:Colors.Warning
        return @()
    }

    $index = 1
    foreach ($backup in $backups) {
        Write-Host "  [$index]" -ForegroundColor Yellow -NoNewline
        Write-Host " $($backup.Name)" -ForegroundColor White
        Write-Host "      Date: " -ForegroundColor DarkGray -NoNewline
        Write-Host "$($backup.Date)" -ForegroundColor Gray -NoNewline
        Write-Host " | " -ForegroundColor DarkGray -NoNewline
        Write-Host "Items: " -ForegroundColor DarkGray -NoNewline
        Write-Host "$($backup.TotalItems) items" -ForegroundColor Gray -NoNewline
        Write-Host " | " -ForegroundColor DarkGray -NoNewline
        Write-Host "Time: " -ForegroundColor DarkGray -NoNewline
        Write-Host "$($backup.Duration)" -ForegroundColor Gray
        $index++
    }

    Write-Host ""
    Write-Host "  [0]" -ForegroundColor Red -NoNewline
    Write-Host " Return to Main Menu" -ForegroundColor White

    return @($backups)
}

#endregion

#region Menu Functions

function Show-MainMenu {
    Show-Banner

    Write-Host "  📋 MAIN MENU" -ForegroundColor Cyan
    Write-Host "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan
    Write-Host ""
    Write-Host "  === Backup & Restore ===" -ForegroundColor DarkCyan
    Write-Host "  [1]" -ForegroundColor Yellow -NoNewline
    Write-Host " Backup Intune Configuration" -ForegroundColor White -NoNewline
    Write-Host " (Read-Only)" -ForegroundColor Green
    Write-Host "  [2]" -ForegroundColor Yellow -NoNewline
    Write-Host " Preview Restore (Dry Run)" -ForegroundColor White
    Write-Host "  [3]" -ForegroundColor Yellow -NoNewline
    Write-Host " Drift Detection & Selective Restore" -ForegroundColor White
    Write-Host "  [4]" -ForegroundColor Yellow -NoNewline
    Write-Host " Full Restore Intune Configuration" -ForegroundColor White -NoNewline
    Write-Host " (Creates New Policies)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  === Reports & Export ===" -ForegroundColor DarkCyan
    Write-Host "  [5]" -ForegroundColor Yellow -NoNewline
    Write-Host " Generate HTML Report" -ForegroundColor White -NoNewline
    Write-Host " (Backup / Current Intune Tenant)" -ForegroundColor Cyan
    Write-Host "  [6]" -ForegroundColor Yellow -NoNewline
    Write-Host " Export Backup (MD/CSV)" -ForegroundColor White
    Write-Host "  [7]" -ForegroundColor Yellow -NoNewline
    Write-Host " Import Configuration from File" -ForegroundColor White
    Write-Host ""
    Write-Host "  === Management ===" -ForegroundColor DarkCyan
    Write-Host "  [8]" -ForegroundColor Yellow -NoNewline
    Write-Host " List Available Backups" -ForegroundColor White
    Write-Host "  [9]" -ForegroundColor Yellow -NoNewline
    Write-Host " Compare Two Backups" -ForegroundColor White
    Write-Host "  [10]" -ForegroundColor Yellow -NoNewline
    Write-Host " Cleanup Old Backups" -ForegroundColor White
    Write-Host "  [11]" -ForegroundColor Yellow -NoNewline
    Write-Host " Cleanup Old Log Files" -ForegroundColor White
    Write-Host ""
    Write-Host "  === Advanced Options ===" -ForegroundColor DarkCyan
    Write-Host "  [12]" -ForegroundColor Yellow -NoNewline
    Write-Host " Backup/Restore by Platform" -ForegroundColor White
    Write-Host "  [13]" -ForegroundColor Yellow -NoNewline
    Write-Host " Backup/Restore by Components" -ForegroundColor White
    Write-Host ""
    Write-Host "  === Settings ===" -ForegroundColor DarkCyan
    Write-Host "  [14]" -ForegroundColor Yellow -NoNewline
    Write-Host " Change Storage Locations" -ForegroundColor White
    Write-Host "  [15]" -ForegroundColor Yellow -NoNewline
    Write-Host " View/Open Log Files" -ForegroundColor White
    Write-Host "  [16]" -ForegroundColor Yellow -NoNewline
    Write-Host " Configure Logging" -ForegroundColor White
    Write-Host ""
    Write-Host "  [Q]" -ForegroundColor Magenta -NoNewline
    Write-Host " Disconnect & Switch Tenant" -ForegroundColor White
    Write-Host "  [0]" -ForegroundColor Red -NoNewline
    Write-Host " Exit" -ForegroundColor White
    Write-Host ""
    Write-Host "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan
}

function Show-PlatformMenu {
    Show-Banner

    Write-Host "  BACKUP BY PLATFORM" -ForegroundColor Cyan
    Write-Host "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan
    Write-Host ""
    Write-Host "  [1]" -ForegroundColor Yellow -NoNewline
    Write-Host " Windows" -ForegroundColor White
    Write-Host "  [2]" -ForegroundColor Yellow -NoNewline
    Write-Host " iOS" -ForegroundColor White
    Write-Host "  [3]" -ForegroundColor Yellow -NoNewline
    Write-Host " Android" -ForegroundColor White
    Write-Host "  [4]" -ForegroundColor Yellow -NoNewline
    Write-Host " macOS" -ForegroundColor White
    Write-Host ""
    Write-Host "  [0]" -ForegroundColor Red -NoNewline
    Write-Host " Return to Main Menu" -ForegroundColor White
    Write-Host ""
    Write-Host "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan
}

function Show-ComponentMenu {
    Show-Banner

    Write-Host "  BACKUP BY COMPONENTS" -ForegroundColor Cyan
    Write-Host "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan
    Write-Host ""
    Write-Host "  === Device Policies ===" -ForegroundColor DarkCyan
    Write-Host "  [1]" -ForegroundColor Yellow -NoNewline
    Write-Host " Device Configurations (Administrative Templates)" -ForegroundColor White
    Write-Host "  [2]" -ForegroundColor Yellow -NoNewline
    Write-Host " Compliance Policies" -ForegroundColor White
    Write-Host "  [3]" -ForegroundColor Yellow -NoNewline
    Write-Host " Settings Catalog" -ForegroundColor White
    Write-Host ""
    Write-Host "  === Autopilot & Enrollment (Windows) ===" -ForegroundColor DarkCyan
    Write-Host "  [4]" -ForegroundColor Yellow -NoNewline
    Write-Host " Autopilot Profiles" -ForegroundColor White
    Write-Host "  [5]" -ForegroundColor Yellow -NoNewline
    Write-Host " Autopilot Device Preparation" -ForegroundColor White
    Write-Host "  [6]" -ForegroundColor Yellow -NoNewline
    Write-Host " Enrollment Status Page (ESP)" -ForegroundColor White
    Write-Host ""
    Write-Host "  === Scripts & Remediations (Windows) ===" -ForegroundColor DarkCyan
    Write-Host "  [7]" -ForegroundColor Yellow -NoNewline
    Write-Host " PowerShell Scripts" -ForegroundColor White
    Write-Host "  [8]" -ForegroundColor Yellow -NoNewline
    Write-Host " Remediation Scripts" -ForegroundColor White
    Write-Host ""
    Write-Host "  === Scripts (macOS) ===" -ForegroundColor DarkCyan
    Write-Host "  [9]" -ForegroundColor Yellow -NoNewline
    Write-Host " macOS Shell Scripts" -ForegroundColor White
    Write-Host "  [10]" -ForegroundColor Yellow -NoNewline
    Write-Host " macOS Custom Attributes" -ForegroundColor White
    Write-Host ""
    Write-Host "  === App Management ===" -ForegroundColor DarkCyan
    Write-Host "  [11]" -ForegroundColor Yellow -NoNewline
    Write-Host " App Protection Policies (MAM)" -ForegroundColor White
    Write-Host "  [12]" -ForegroundColor Yellow -NoNewline
    Write-Host " App Configuration (Managed Devices)" -ForegroundColor White
    Write-Host "  [13]" -ForegroundColor Yellow -NoNewline
    Write-Host " App Configuration (Managed Apps)" -ForegroundColor White
    Write-Host ""
    Write-Host "  === Assignment Filters ===" -ForegroundColor DarkCyan
    Write-Host "  [14]" -ForegroundColor Yellow -NoNewline
    Write-Host " Assignment Filters" -ForegroundColor White
    Write-Host ""
    Write-Host "  [A]" -ForegroundColor Green -NoNewline
    Write-Host " Select All Components" -ForegroundColor White
    Write-Host "  [0]" -ForegroundColor Red -NoNewline
    Write-Host " Return to Main Menu" -ForegroundColor White
    Write-Host ""
    Write-Host "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan
}

function Start-InteractiveMode {
    $continue = $true

    while ($continue) {
        Show-MainMenu
        $choice = Read-Host "`nSelected option"

        switch ($choice) {
            "1" {
                # Backup
                Write-ColoredOutput "`n════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
                Write-ColoredOutput " FULL BACKUP OPERATION" -Color $script:Colors.Success
                Write-ColoredOutput "════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
                Write-ColoredOutput "This operation will:" -Color $script:Colors.Info
                Write-ColoredOutput "  + READ policies from your Intune tenant" -Color $script:Colors.Success
                Write-ColoredOutput "  + SAVE them to local JSON files" -Color $script:Colors.Success
                Write-ColoredOutput "  - NOT modify anything in Intune" -Color $script:Colors.Warning
                Write-ColoredOutput "  - NOT delete any policies" -Color $script:Colors.Warning
                Write-ColoredOutput "════════════════════════════════════════════════════════════════" -Color $script:Colors.Header

                $backupName = Read-Host "`nEnter backup name (optional, press Enter to skip, 0 to cancel)"

                # Check if user wants to cancel
                if ($backupName -eq "0") {
                    Write-ColoredOutput "`nOperation cancelled." -Color $script:Colors.Warning
                }
                else {
                    $script:SelectedPlatform = "All"
                    $SelectedComponents = Get-SelectedComponents -ComponentList @("All")
                    Backup-IntuneConfiguration -BackupName $backupName -SelectedComponents $SelectedComponents
                }

                $continuePrompt = Read-Host "`nPress Enter to return to main menu (or 0)"
                # No need to check the value, just return to menu either way
            }
            "4" {
                # Full Restore
                Write-ColoredOutput "`n════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
                Write-ColoredOutput " FULL RESTORE OPERATION - CREATES NEW POLICIES" -Color $script:Colors.Info
                Write-ColoredOutput "════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
                Write-ColoredOutput "This operation will:" -Color $script:Colors.Info
                Write-ColoredOutput "  + CREATE new policies with '- [Restored]' suffix" -Color $script:Colors.Success
                Write-ColoredOutput "  + SKIP policies that already exist" -Color $script:Colors.Success
                Write-ColoredOutput "  - NOT modify existing policies" -Color $script:Colors.Warning
                Write-ColoredOutput "  - NOT delete any policies" -Color $script:Colors.Warning
                Write-ColoredOutput "  - NOT overwrite your current configuration" -Color $script:Colors.Warning
                Write-ColoredOutput "════════════════════════════════════════════════════════════════" -Color $script:Colors.Header

                $backups = Show-BackupList
                if ($backups) {
                    $selection = Read-Host "`nSelect backup number (or 0 to cancel)"

                    if ($selection -eq "0") {
                        Write-ColoredOutput "`nOperation cancelled." -Color $script:Colors.Warning
                    }
                    elseif ($selection -le $backups.Count) {
                        $selectedBackup = $backups[$selection - 1]
                        $script:SelectedPlatform = "All"
                        $SelectedComponents = Get-SelectedComponents -ComponentList @("All")
                        Restore-IntuneConfiguration -BackupPath $selectedBackup.Path -SelectedComponents $SelectedComponents
                    }
                    else {
                        Write-ColoredOutput "`nInvalid selection." -Color $script:Colors.Error
                    }
                }

                $continuePrompt = Read-Host "`nPress Enter to return to main menu (or 0)"
                # No need to check the value, just return to menu either way
            }
            "2" {
                # Preview Restore (Dry Run)
                Write-ColoredOutput "`n════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
                Write-ColoredOutput " PREVIEW MODE - DRY RUN ONLY" -Color $script:Colors.Info
                Write-ColoredOutput "════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
                Write-ColoredOutput "This will show you what WOULD happen without making changes:" -Color $script:Colors.Info
                Write-ColoredOutput "  - Which policies would be created" -Color $script:Colors.Default
                Write-ColoredOutput "  - Which policies already exist and would be skipped" -Color $script:Colors.Default
                Write-ColoredOutput "  - NO actual changes will be made to Intune" -Color $script:Colors.Warning
                Write-ColoredOutput "════════════════════════════════════════════════════════════════" -Color $script:Colors.Header

                $backups = @(Show-BackupList)
                if ($backups.Count -gt 0) {
                    $selection = Read-Host "`nSelect backup to preview"

                    if ($selection -eq "0" -or $selection -eq "") {
                        # Return to main menu
                    }
                    elseif ($selection -match '^\d+$' -and [int]$selection -gt 0 -and [int]$selection -le $backups.Count) {
                        $selectedBackup = $backups[[int]$selection - 1]
                        $script:SelectedPlatform = "All"
                        $SelectedComponents = Get-SelectedComponents -ComponentList @("All")
                        Write-ColoredOutput "`nStarting preview restore..." -Color $script:Colors.Info
                        Restore-IntuneConfiguration -BackupPath $selectedBackup.Path -SelectedComponents $SelectedComponents -Preview
                    }
                    else {
                        Write-ColoredOutput "Invalid selection. Please enter a number between 1 and $($backups.Count)" -Color $script:Colors.Warning
                    }
                }
                else {
                    Write-ColoredOutput "No backups available to preview." -Color $script:Colors.Warning
                }

                Read-Host "`nPress Enter to return to main menu (or 0)"
            }
            "3" {
                # Drift Detection & Selective Restore
                Write-ColoredOutput "`n════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
                Write-ColoredOutput " DRIFT DETECTION & SELECTIVE RESTORE" -Color $script:Colors.Info
                Write-ColoredOutput "════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
                Write-ColoredOutput "This operation will:" -Color $script:Colors.Info
                Write-ColoredOutput "  1. COMPARE backup with current Intune configuration" -Color $script:Colors.Default
                Write-ColoredOutput "  2. IDENTIFY removed or modified policies" -Color $script:Colors.Default
                Write-ColoredOutput "  3. Allow you to CREATE new copies with '- [Restored]' suffix" -Color $script:Colors.Success
                Write-ColoredOutput "  - Original policies remain untouched" -Color $script:Colors.Warning
                Write-ColoredOutput "  - No policies will be deleted or modified" -Color $script:Colors.Warning
                Write-ColoredOutput "════════════════════════════════════════════════════════════════" -Color $script:Colors.Header

                $backups = @(Show-BackupList)
                if ($backups.Count -gt 0) {
                    $selection = Read-Host "`nSelect backup for drift detection"

                    if ($selection -eq "0" -or $selection -eq "") {
                        # Return to main menu
                    }
                    elseif ($selection -match '^\d+$' -and [int]$selection -gt 0 -and [int]$selection -le $backups.Count) {
                        $selectedBackup = $backups[[int]$selection - 1]
                        Write-ColoredOutput "`nStarting drift detection..." -Color $script:Colors.Info
                        $driftResults = Detect-ConfigurationDrift -BackupPath $selectedBackup.Path

                        if ($driftResults -and ($driftResults.Removed.Count -gt 0 -or $driftResults.Changed.Count -gt 0 -or $driftResults.Renamed.Count -gt 0)) {
                            $restore = Read-Host "`nDo you want to restore/revert selected policies? (yes/no)"
                            if ($restore -eq "yes") {
                                Restore-SelectedPolicies -DriftResults $driftResults
                            }
                        }
                        elseif ($driftResults -and $driftResults.Added.Count -gt 0) {
                            # Only Added policies detected - these exist on tenant but not in backup
                            Write-ColoredOutput "`n════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
                            Write-ColoredOutput "POLICIES ADDED SINCE BACKUP" -Color $script:Colors.Header
                            Write-ColoredOutput "════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
                            Write-ColoredOutput "The following $($driftResults.Added.Count) policies exist on tenant but were NOT in the backup:" -Color $script:Colors.Info
                            Write-ColoredOutput "(These were created after the backup was made)`n" -Color "Gray"

                            foreach ($addedPolicy in $driftResults.Added) {
                                Write-ColoredOutput "  • $($addedPolicy.Type)/$($addedPolicy.Name)" -Color $script:Colors.Success
                            }

                            Write-ColoredOutput "`n════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
                            Write-ColoredOutput "TIP: Create a new backup to include these policies." -Color $script:Colors.Info
                            Write-ColoredOutput "════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
                        }
                        elseif ($driftResults) {
                            Write-ColoredOutput "`nNo drift detected - all policies match the backup." -Color $script:Colors.Success
                        }
                    }
                    else {
                        Write-ColoredOutput "Invalid selection. Please enter a number between 1 and $($backups.Count)" -Color $script:Colors.Warning
                    }
                }
                else {
                    Write-ColoredOutput "No backups available for drift detection." -Color $script:Colors.Warning
                }

                Read-Host "`nPress Enter to return to main menu (or 0)"
            }
            "5" {
                # Generate HTML Report - Sub-menu
                $reportMenuLoop = $true
                while ($reportMenuLoop) {
                    Write-Host ""
                    Write-Host "  === Generate HTML Report ===" -ForegroundColor DarkCyan
                    Write-Host "  [1]" -ForegroundColor Yellow -NoNewline
                    Write-Host " From Backup" -ForegroundColor White -NoNewline
                    Write-Host " (select a saved backup)" -ForegroundColor Gray
                    Write-Host "  [2]" -ForegroundColor Yellow -NoNewline
                    Write-Host " From Current Intune Tenant" -ForegroundColor White -NoNewline
                    Write-Host " (fetch current configuration)" -ForegroundColor Gray
                    Write-Host ""
                    Write-Host "  [0]" -ForegroundColor Red -NoNewline
                    Write-Host " Return to Main Menu" -ForegroundColor White
                    Write-Host ""

                    $reportChoice = Read-Host "  Select option"

                    switch ($reportChoice) {
                        "1" {
                            # From Backup
                            $backups = @(Show-BackupList)
                            if ($backups.Count -gt 0) {
                                $selection = Read-Host "`nSelect backup to generate HTML report"

                                if ($selection -eq "0" -or $selection -eq "") {
                                    # Return to sub-menu
                                }
                                elseif ($selection -match '^\d+$' -and [int]$selection -gt 0 -and [int]$selection -le $backups.Count) {
                                    $selectedBackup = $backups[[int]$selection - 1]

                                    if (-not (Test-Path $script:ReportsLocation)) {
                                        New-Item -ItemType Directory -Path $script:ReportsLocation -Force | Out-Null
                                    }

                                    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                                    $outputPath = Join-Path $script:ReportsLocation "report_$timestamp.html"

                                    Write-ColoredOutput "`nGenerating detailed HTML report..." -Color $script:Colors.Info
                                    Write-ColoredOutput "Resolving settings from backup (this may take a moment)..." -Color $script:Colors.Info
                                    $script:SettingDefinitionCache = @{}
                                    Export-BackupToHTML -BackupPath $selectedBackup.Path -OutputPath $outputPath

                                    if (Test-Path $outputPath) {
                                        Write-ColoredOutput "`nHTML report saved to: $outputPath" -Color $script:Colors.Success
                                        $openReport = Read-Host "Open report in browser? (yes/no)"
                                        if ($openReport -eq "yes") {
                                            Start-Process $outputPath
                                        }
                                    }
                                    else {
                                        Write-ColoredOutput "Failed to generate HTML report." -Color $script:Colors.Error
                                    }
                                }
                                else {
                                    Write-ColoredOutput "Invalid selection. Please enter a number between 1 and $($backups.Count)" -Color $script:Colors.Warning
                                }
                            }
                            else {
                                Write-ColoredOutput "No backups available to generate report." -Color $script:Colors.Warning
                            }
                            $reportMenuLoop = $false
                        }
                        "2" {
                            # From Current Intune Tenant
                            if (-not (Test-Path $script:ReportsLocation)) {
                                New-Item -ItemType Directory -Path $script:ReportsLocation -Force | Out-Null
                            }

                            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                            $outputPath = Join-Path $script:ReportsLocation "tenant_report_$timestamp.html"

                            Export-CurrentTenantConfigToHTML -OutputPath $outputPath

                            if (Test-Path $outputPath) {
                                Write-ColoredOutput "`nLive tenant report saved to: $outputPath" -Color $script:Colors.Success
                                $openReport = Read-Host "Open report in browser? (yes/no)"
                                if ($openReport -eq "yes") {
                                    Start-Process $outputPath
                                }
                            }
                            else {
                                Write-ColoredOutput "Failed to generate live tenant report." -Color $script:Colors.Error
                            }
                            $reportMenuLoop = $false
                        }
                        "0" { $reportMenuLoop = $false }
                        default {
                            Write-ColoredOutput "Invalid option. Please select 1, 2, or 0." -Color $script:Colors.Warning
                        }
                    }
                }

                Read-Host "`nPress Enter to return to main menu (or 0)"
            }
            "6" {
                # Export Backup (MD/CSV)
                $backups = @(Show-BackupList)
                if ($backups.Count -gt 0) {
                    $selection = Read-Host "`nSelect backup to export"

                    if ($selection -eq "0" -or $selection -eq "") {
                        # Return to main menu
                    }
                    elseif ($selection -match '^\d+$' -and [int]$selection -gt 0 -and [int]$selection -le $backups.Count) {
                        $selectedBackup = $backups[[int]$selection - 1]

                        Write-Host "`n  Export Format:" -ForegroundColor Cyan
                        Write-Host "  [1] Markdown (.md)" -ForegroundColor Yellow
                        Write-Host "  [2] CSV" -ForegroundColor Yellow
                        Write-Host "  [3] Both Markdown and CSV" -ForegroundColor Yellow
                        Write-Host "  [0] Cancel" -ForegroundColor Red

                        $formatChoice = Read-Host "`nSelect export format"

                        if ($formatChoice -ne "0") {
                            if (-not (Test-Path $script:ExportsLocation)) {
                                New-Item -ItemType Directory -Path $script:ExportsLocation -Force | Out-Null
                            }

                            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

                            # Create timestamped export folder for all export types
                            $exportFolder = Join-Path $script:ExportsLocation "export_$timestamp"
                            if (-not (Test-Path $exportFolder)) {
                                New-Item -ItemType Directory -Path $exportFolder -Force | Out-Null
                            }

                            switch ($formatChoice) {
                                "1" {
                                    $outputPath = Join-Path $exportFolder "backup_report.md"
                                    Export-BackupToMarkdown -BackupPath $selectedBackup.Path -OutputPath $outputPath
                                }
                                "2" {
                                    Export-BackupToCSV -BackupPath $selectedBackup.Path -OutputFolder $exportFolder
                                }
                                "3" {
                                    $mdPath = Join-Path $exportFolder "backup_report.md"

                                    Export-BackupToMarkdown -BackupPath $selectedBackup.Path -OutputPath $mdPath
                                    Export-BackupToCSV -BackupPath $selectedBackup.Path -OutputFolder $exportFolder
                                }
                                default {
                                    Write-ColoredOutput "Invalid selection. Please enter 1, 2, or 3." -Color $script:Colors.Warning
                                }
                            }
                        }
                    }
                    else {
                        Write-ColoredOutput "Invalid selection. Please enter a number between 1 and $($backups.Count)" -Color $script:Colors.Warning
                    }
                }
                else {
                    Write-ColoredOutput "No backups available to export." -Color $script:Colors.Warning
                }

                Read-Host "`nPress Enter to return to main menu (or 0)"
            }
            "7" {
                # Import Configuration from File
                Write-ColoredOutput "`n════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
                Write-ColoredOutput " IMPORT CONFIGURATION" -Color $script:Colors.Info
                Write-ColoredOutput "════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
                Write-ColoredOutput "Import from:" -Color $script:Colors.Info
                Write-Host "  [1]" -ForegroundColor Yellow -NoNewline
                Write-Host " JSON file (import file or individual policy)" -ForegroundColor White
                Write-Host "  [2]" -ForegroundColor Yellow -NoNewline
                Write-Host " Backup folder (select multiple policies)" -ForegroundColor White
                Write-Host "  [3]" -ForegroundColor Yellow -NoNewline
                Write-Host " Available backups list" -ForegroundColor White
                Write-Host "  [0]" -ForegroundColor Red -NoNewline
                Write-Host " Cancel" -ForegroundColor White

                $importChoice = Read-Host "`nSelect import source"

                $importPath = $null

                switch ($importChoice) {
                    "1" {
                        Write-Host "`n  Opening file browser..." -ForegroundColor Gray
                        $importPath = Show-OpenFileDialog
                    }
                    "2" {
                        Write-Host "`n  Opening folder browser..." -ForegroundColor Gray
                        $importPath = Show-FolderBrowserDialog -Description "Select backup folder to import from"
                    }
                    "3" {
                        $backups = @(Show-BackupList)
                        if ($backups.Count -gt 0) {
                            $selection = Read-Host "`nSelect backup to import from"
                            if ($selection -eq "0" -or $selection -eq "") {
                                # Return to main menu
                            }
                            elseif ($selection -match '^\d+$' -and [int]$selection -gt 0 -and [int]$selection -le $backups.Count) {
                                $importPath = $backups[[int]$selection - 1].Path
                            }
                            else {
                                Write-ColoredOutput "Invalid selection. Please enter a number between 1 and $($backups.Count)" -Color $script:Colors.Warning
                            }
                        }
                        else {
                            Write-ColoredOutput "No backups available." -Color $script:Colors.Warning
                        }
                    }
                    "0" {
                        Write-ColoredOutput "`nOperation cancelled." -Color $script:Colors.Warning
                    }
                    default {
                        Write-ColoredOutput "Invalid selection." -Color $script:Colors.Warning
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($importPath)) {
                    Write-ColoredOutput "`nSelected: $importPath" -Color $script:Colors.Info

                    $preview = Read-Host "`nPreview import without making changes? (yes/no, or 0 to cancel)"

                    if ($preview -eq "0") {
                        Write-ColoredOutput "`nOperation cancelled." -Color $script:Colors.Warning
                    }
                    else {
                        $previewMode = $preview -eq "yes"
                        Import-ConfigurationFromFile -ImportPath $importPath -Preview:$previewMode
                    }
                }

                Read-Host "`nPress Enter to return to main menu (or 0)"
            }
            "8" {
                # List Available Backups
                Show-BackupList | Out-Null
                Read-Host "`nPress Enter to return to main menu (or 0)"
            }
            "9" {
                # Compare Two Backups
                Write-ColoredOutput "`nSelect first backup:" -Color $script:Colors.Info
                $backups = @(Show-BackupList)

                if ($backups.Count -ge 2) {
                    $selection1 = Read-Host "`nSelect first backup"

                    if ($selection1 -eq "0" -or $selection1 -eq "") {
                        # Return to main menu
                    }
                    elseif ($selection1 -match '^\d+$' -and [int]$selection1 -gt 0 -and [int]$selection1 -le $backups.Count) {
                        Write-ColoredOutput "`nSelect second backup:" -Color $script:Colors.Info
                        Show-BackupList | Out-Null
                        $selection2 = Read-Host "`nSelect second backup"

                        if ($selection2 -eq "0" -or $selection2 -eq "") {
                            # Return to main menu
                        }
                        elseif ($selection2 -match '^\d+$' -and [int]$selection2 -gt 0 -and [int]$selection2 -le $backups.Count) {
                            $backup1 = $backups[[int]$selection1 - 1]
                            $backup2 = $backups[[int]$selection2 - 1]

                            $differences = Compare-Backups -Backup1Path $backup1.Path -Backup2Path $backup2.Path

                            # Offer to export comparison results
                            if ($differences.OnlyInBackup1.Count -gt 0 -or $differences.OnlyInBackup2.Count -gt 0 -or $differences.Modified.Count -gt 0) {
                                Write-Host ""
                                $exportChoice = Read-Host "Export comparison report to Markdown? (yes/no)"
                                if ($exportChoice -eq "yes" -or $exportChoice -eq "y") {
                                    if (-not (Test-Path $script:ReportsLocation)) {
                                        New-Item -ItemType Directory -Path $script:ReportsLocation -Force | Out-Null
                                    }
                                    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                                    $outputPath = Join-Path $script:ReportsLocation "comparison_$timestamp.md"
                                    Export-ComparisonToMarkdown -Differences $differences -Backup1Path $backup1.Path -Backup2Path $backup2.Path -OutputPath $outputPath
                                }
                            }
                        }
                        else {
                            Write-ColoredOutput "Invalid selection for second backup." -Color $script:Colors.Warning
                        }
                    }
                    else {
                        Write-ColoredOutput "Invalid selection for first backup." -Color $script:Colors.Warning
                    }
                }
                elseif ($backups.Count -eq 1) {
                    Write-ColoredOutput "Need at least 2 backups to compare (only 1 backup found)" -Color $script:Colors.Warning
                }
                else {
                    Write-ColoredOutput "No backups available to compare." -Color $script:Colors.Warning
                }

                Read-Host "`nPress Enter to return to main menu (or 0)"
            }
            "10" {
                # Cleanup Old Backups
                Write-Host "`n  Backup Cleanup Options" -ForegroundColor Cyan
                Write-Host "  [1] Delete specific backup(s)" -ForegroundColor Yellow
                Write-Host "  [2] Delete backups by retention days" -ForegroundColor Yellow
                Write-Host "  [0] Return to Main Menu" -ForegroundColor Red

                $cleanupChoice = Read-Host "`nSelect option"

                switch ($cleanupChoice) {
                    "1" {
                        # Delete specific backups
                        $backups = @(Get-AvailableBackups)

                        if ($backups.Count -eq 0) {
                            Write-ColoredOutput "`nNo backups found in $script:BackupLocation" -Color $script:Colors.Warning
                        }
                        else {
                            Write-Host "`n  AVAILABLE BACKUPS" -ForegroundColor Cyan
                            Write-Host "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan

                            $index = 1
                            foreach ($backup in $backups) {
                                Write-Host "  [$index]" -ForegroundColor Yellow -NoNewline
                                Write-Host " $($backup.Name)" -ForegroundColor White
                                Write-Host "      Date: " -ForegroundColor DarkGray -NoNewline
                                Write-Host "$($backup.Date)" -ForegroundColor Gray -NoNewline
                                Write-Host " | " -ForegroundColor DarkGray -NoNewline
                                Write-Host "Items: " -ForegroundColor DarkGray -NoNewline
                                Write-Host "$($backup.TotalItems) items" -ForegroundColor Gray
                                $index++
                            }

                            Write-Host ""
                            Write-Host "  Enter backup number(s) to delete (comma-separated, e.g., 1,3,5)" -ForegroundColor Cyan
                            Write-Host "  [0] Cancel" -ForegroundColor Red

                            $selection = Read-Host "`nSelect backup(s) to delete"

                            if ($selection -ne "0" -and -not [string]::IsNullOrWhiteSpace($selection)) {
                                # Parse selection
                                $selectedIndices = $selection -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ }
                                $validSelections = $selectedIndices | Where-Object { $_ -ge 1 -and $_ -le $backups.Count }

                                if ($validSelections.Count -eq 0) {
                                    Write-ColoredOutput "`nNo valid selections made." -Color $script:Colors.Warning
                                }
                                else {
                                    # Show selected backups for confirmation
                                    Write-Host "`n  Backups selected for deletion:" -ForegroundColor Yellow
                                    foreach ($idx in $validSelections) {
                                        $selectedBackup = $backups[$idx - 1]
                                        Write-Host "    - $($selectedBackup.Name) ($($selectedBackup.Date))" -ForegroundColor White
                                    }

                                    $confirm = Read-Host "`nAre you sure you want to delete these $($validSelections.Count) backup(s)? (Y/N)"

                                    if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                                        $deletedCount = 0
                                        foreach ($idx in $validSelections) {
                                            $backupToDelete = $backups[$idx - 1]
                                            try {
                                                Write-ColoredOutput "  Deleting: $($backupToDelete.Name)..." -Color $script:Colors.Warning
                                                Remove-Item -Path $backupToDelete.Path -Recurse -Force
                                                $deletedCount++
                                            }
                                            catch {
                                                Write-ColoredOutput "  Failed to delete $($backupToDelete.Name): $($_.Exception.Message)" -Color $script:Colors.Error
                                            }
                                        }
                                        Write-ColoredOutput "`nSuccessfully deleted $deletedCount backup(s)" -Color $script:Colors.Success
                                    }
                                    else {
                                        Write-ColoredOutput "`nDeletion cancelled." -Color $script:Colors.Info
                                    }
                                }
                            }
                        }
                    }
                    "2" {
                        # Delete by retention days (original behavior)
                        $days = Read-Host "`nEnter retention days (default: 30)"
                        if ([string]::IsNullOrWhiteSpace($days)) { $days = 30 }

                        Remove-OldBackups -RetentionDays ([int]$days)
                    }
                    "0" {
                        # Return to main menu
                    }
                    default {
                        Write-ColoredOutput "Invalid selection" -Color $script:Colors.Warning
                    }
                }

                Read-Host "`nPress Enter to return to main menu (or 0)"
            }
            "11" {
                # Cleanup Old Log Files
                $days = Read-Host "`nEnter retention days for log files (default: 30)"
                if ([string]::IsNullOrWhiteSpace($days)) { $days = 30 }

                Remove-OldLogs -RetentionDays ([int]$days)
                Read-Host "`nPress Enter to return to main menu (or 0)"
            }
            "12" {
                # Backup/Restore by Platform - Loop until user chooses to return
                $platformMenuLoop = $true
                while ($platformMenuLoop) {
                    Write-Host "`n  Platform Operations" -ForegroundColor Cyan
                    Write-Host "  [1] Backup by Platform" -ForegroundColor Yellow
                    Write-Host "  [2] Restore by Platform" -ForegroundColor Yellow
                    Write-Host "  [0] Return to Main Menu" -ForegroundColor Red

                    $platformOpChoice = Read-Host "`nSelect option"

                    switch ($platformOpChoice) {
                        "1" {
                            # Backup by Platform
                            Show-PlatformMenu
                            $platformChoice = Read-Host "`nSelected option"

                            switch ($platformChoice) {
                                "1" {
                                    $script:SelectedPlatform = "Windows"
                                    $backupName = Read-Host "`nEnter backup name (optional, press Enter to skip)"
                                    $SelectedComponents = Get-SelectedComponents -ComponentList @("All")
                                    Backup-IntuneConfiguration -BackupName $backupName -SelectedComponents $SelectedComponents
                                    Read-Host "`nPress Enter to continue (or 0)"
                                }
                                "2" {
                                    $script:SelectedPlatform = "iOS"
                                    $backupName = Read-Host "`nEnter backup name (optional, press Enter to skip)"
                                    $SelectedComponents = Get-SelectedComponents -ComponentList @("All")
                                    Backup-IntuneConfiguration -BackupName $backupName -SelectedComponents $SelectedComponents
                                    Read-Host "`nPress Enter to continue (or 0)"
                                }
                                "3" {
                                    $script:SelectedPlatform = "Android"
                                    $backupName = Read-Host "`nEnter backup name (optional, press Enter to skip)"
                                    $SelectedComponents = Get-SelectedComponents -ComponentList @("All")
                                    Backup-IntuneConfiguration -BackupName $backupName -SelectedComponents $SelectedComponents
                                    Read-Host "`nPress Enter to continue (or 0)"
                                }
                                "4" {
                                    $script:SelectedPlatform = "macOS"
                                    $backupName = Read-Host "`nEnter backup name (optional, press Enter to skip)"
                                    $SelectedComponents = Get-SelectedComponents -ComponentList @("All")
                                    Backup-IntuneConfiguration -BackupName $backupName -SelectedComponents $SelectedComponents
                                    Read-Host "`nPress Enter to continue (or 0)"
                                }
                                "0" {
                                    # Return to Platform Operations menu
                                }
                                default {
                                    Write-ColoredOutput "Invalid selection" -Color $script:Colors.Warning
                                }
                            }
                            $script:SelectedPlatform = "All"
                        }
                        "2" {
                            # Restore by Platform
                            $backups = Show-BackupList
                            if ($backups) {
                                $selection = Read-Host "`nSelect backup to restore"
                                if ($selection -ne "0" -and $selection -le $backups.Count) {
                                    $selectedBackup = $backups[$selection - 1]

                                    Show-PlatformMenu
                                    $platformChoice = Read-Host "`nSelected option"

                                    switch ($platformChoice) {
                                        "1" {
                                            $script:SelectedPlatform = "Windows"
                                            $SelectedComponents = Get-SelectedComponents -ComponentList @("All")
                                            Restore-IntuneConfiguration -BackupPath $selectedBackup.Path -SelectedComponents $SelectedComponents
                                            Read-Host "`nPress Enter to continue (or 0)"
                                        }
                                        "2" {
                                            $script:SelectedPlatform = "iOS"
                                            $SelectedComponents = Get-SelectedComponents -ComponentList @("All")
                                            Restore-IntuneConfiguration -BackupPath $selectedBackup.Path -SelectedComponents $SelectedComponents
                                            Read-Host "`nPress Enter to continue (or 0)"
                                        }
                                        "3" {
                                            $script:SelectedPlatform = "Android"
                                            $SelectedComponents = Get-SelectedComponents -ComponentList @("All")
                                            Restore-IntuneConfiguration -BackupPath $selectedBackup.Path -SelectedComponents $SelectedComponents
                                            Read-Host "`nPress Enter to continue (or 0)"
                                        }
                                        "4" {
                                            $script:SelectedPlatform = "macOS"
                                            $SelectedComponents = Get-SelectedComponents -ComponentList @("All")
                                            Restore-IntuneConfiguration -BackupPath $selectedBackup.Path -SelectedComponents $SelectedComponents
                                            Read-Host "`nPress Enter to continue (or 0)"
                                        }
                                        "0" {
                                            # Return to Platform Operations menu
                                        }
                                        default {
                                            Write-ColoredOutput "Invalid selection" -Color $script:Colors.Warning
                                        }
                                    }
                                    $script:SelectedPlatform = "All"
                                }
                            }
                        }
                        "0" {
                            $platformMenuLoop = $false
                        }
                        default {
                            Write-ColoredOutput "Invalid selection" -Color $script:Colors.Warning
                        }
                    }
                }
            }
            "13" {
                # Backup/Restore by Components - Loop until user chooses to return
                $componentMenuLoop = $true
                while ($componentMenuLoop) {
                    Write-ColoredOutput "`n════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
                    Write-ColoredOutput " BACKUP/RESTORE BY COMPONENTS" -Color $script:Colors.Info
                    Write-ColoredOutput "════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
                    Write-Host "  [1]" -ForegroundColor Yellow -NoNewline
                    Write-Host " Backup by Components" -ForegroundColor White
                    Write-Host "  [2]" -ForegroundColor Yellow -NoNewline
                    Write-Host " Restore by Components" -ForegroundColor White
                    Write-Host "  [0]" -ForegroundColor Red -NoNewline
                    Write-Host " Return to Main Menu" -ForegroundColor White

                    $componentOpChoice = Read-Host "`nSelect option"

                    $componentMap = @{
                        "1"  = "DeviceConfigurations"
                        "2"  = "CompliancePolicies"
                        "3"  = "SettingsCatalogPolicies"
                        "4"  = "AutopilotProfiles"
                        "5"  = "AutopilotDevicePrep"
                        "6"  = "EnrollmentStatusPage"
                        "7"  = "PowerShellScripts"
                        "8"  = "RemediationScripts"
                        "9"  = "MacOSScripts"
                        "10" = "MacOSCustomAttributes"
                        "11" = "AppProtectionPolicies"
                        "12" = "AppConfigManagedDevices"
                        "13" = "AppConfigManagedApps"
                        "14" = "AssignmentFilters"
                    }

                    switch ($componentOpChoice) {
                        "1" {
                            # Backup by Components
                            Show-ComponentMenu
                            $componentChoice = Read-Host "`nEnter component number (1-14) or A for all, 0 to return"

                            if ($componentChoice -eq "A" -or $componentChoice -eq "a") {
                                $SelectedComponents = Get-SelectedComponents -ComponentList @("All")
                                $backupName = Read-Host "`nEnter backup name (optional, press Enter to skip)"
                                $script:SelectedPlatform = "All"
                                Backup-IntuneConfiguration -BackupName $backupName -SelectedComponents $SelectedComponents
                                Write-ColoredOutput "`nBackup completed. Returning to Component Menu..." -Color $script:Colors.Success
                            }
                            elseif ($componentChoice -ne "0" -and $componentMap.ContainsKey($componentChoice)) {
                                $SelectedComponents = @($componentMap[$componentChoice])
                                $backupName = Read-Host "`nEnter backup name (optional, press Enter to skip)"
                                $script:SelectedPlatform = "All"
                                Backup-IntuneConfiguration -BackupName $backupName -SelectedComponents $SelectedComponents
                                Write-ColoredOutput "`nBackup completed. Returning to Component Menu..." -Color $script:Colors.Success
                            }
                            elseif ($componentChoice -ne "0") {
                                Write-ColoredOutput "Invalid selection" -Color $script:Colors.Warning
                            }
                            # Stay in component menu loop
                        }
                        "2" {
                            # Restore by Components
                            $backups = Show-BackupList
                            if ($backups -and $backups.Count -gt 0) {
                                $selection = Read-Host "`nSelect backup to restore"
                                if ($selection -match '^\d+$' -and [int]$selection -gt 0 -and [int]$selection -le $backups.Count) {
                                    $selectedBackup = $backups[[int]$selection - 1]

                                    Show-ComponentMenu
                                    $componentChoice = Read-Host "`nEnter component number (1-14) or A for all, 0 to return"

                                    if ($componentChoice -eq "A" -or $componentChoice -eq "a") {
                                        $script:SelectedPlatform = "All"
                                        $SelectedComponents = Get-SelectedComponents -ComponentList @("All")
                                        Restore-IntuneConfiguration -BackupPath $selectedBackup.Path -SelectedComponents $SelectedComponents
                                        Write-ColoredOutput "`nRestore completed. Returning to Component Menu..." -Color $script:Colors.Success
                                    }
                                    elseif ($componentChoice -ne "0" -and $componentMap.ContainsKey($componentChoice)) {
                                        $script:SelectedPlatform = "All"
                                        $selectedComponent = $componentMap[$componentChoice]
                                        Restore-IntuneConfiguration -BackupPath $selectedBackup.Path -SelectedComponents @($selectedComponent)
                                        Write-ColoredOutput "`nRestore completed. Returning to Component Menu..." -Color $script:Colors.Success
                                    }
                                    elseif ($componentChoice -ne "0") {
                                        Write-ColoredOutput "Invalid selection" -Color $script:Colors.Warning
                                    }
                                }
                                elseif ($selection -ne "0") {
                                    Write-ColoredOutput "Invalid selection" -Color $script:Colors.Warning
                                }
                            }
                            else {
                                Write-ColoredOutput "No backups available." -Color $script:Colors.Warning
                            }
                            # Stay in component menu loop
                        }
                        "0" {
                            # Exit component menu loop and return to main menu
                            $componentMenuLoop = $false
                        }
                        default {
                            Write-ColoredOutput "Invalid selection. Please enter 1, 2, or 0." -Color $script:Colors.Warning
                        }
                    }
                }
            }
            "14" {
                # Change Storage Locations — submenu
                $subChoice14 = ""
                do {
                    Write-Host ""
                    Write-ColoredOutput "  === Change Storage Locations ===" -Color $script:Colors.Info
                    Write-Host ""
                    Write-Host "  Current locations:" -ForegroundColor Gray
                    Write-Host ""
                    Write-Host "  [1]" -ForegroundColor Yellow -NoNewline
                    Write-Host " Backup  : " -ForegroundColor White -NoNewline
                    Write-Host $script:BackupLocation -ForegroundColor Cyan
                    Write-Host "  [2]" -ForegroundColor Yellow -NoNewline
                    Write-Host " Reports : " -ForegroundColor White -NoNewline
                    Write-Host $script:ReportsLocation -ForegroundColor Cyan
                    Write-Host "  [3]" -ForegroundColor Yellow -NoNewline
                    Write-Host " Logs    : " -ForegroundColor White -NoNewline
                    Write-Host $script:LogLocation -ForegroundColor Cyan
                    Write-Host "  [4]" -ForegroundColor Yellow -NoNewline
                    Write-Host " Exports : " -ForegroundColor White -NoNewline
                    Write-Host $script:ExportsLocation -ForegroundColor Cyan
                    Write-Host ""
                    Write-Host "  [0]" -ForegroundColor Red -NoNewline
                    Write-Host " Back to Main Menu" -ForegroundColor White
                    Write-Host ""

                    $subChoice14 = Read-Host "  Select location to change"

                    switch ($subChoice14) {
                        "1" {
                            # Change Backup Location
                            Write-Host "  Opening folder browser..." -ForegroundColor Gray
                            $newPath = Show-FolderBrowserDialog -Description "Select new backup location" -SelectedPath $script:BackupLocation
                            if (-not [string]::IsNullOrWhiteSpace($newPath)) {
                                if (-not (Test-Path $newPath)) {
                                    Write-ColoredOutput "Path does not exist. Creating..." -Color $script:Colors.Warning
                                    New-Item -ItemType Directory -Path $newPath -Force | Out-Null
                                }
                                $script:BackupLocation = $newPath
                                Write-ColoredOutput "Backup location changed to: $script:BackupLocation" -Color $script:Colors.Success
                                Write-Log "Backup location changed to: $script:BackupLocation" -Level Info
                            }
                            else {
                                Write-ColoredOutput "No folder selected. Operation cancelled." -Color $script:Colors.Warning
                            }
                        }
                        "2" {
                            # Change Reports Location
                            Write-Host "  Opening folder browser..." -ForegroundColor Gray
                            $newPath = Show-FolderBrowserDialog -Description "Select new reports location" -SelectedPath $script:ReportsLocation
                            if (-not [string]::IsNullOrWhiteSpace($newPath)) {
                                if (-not (Test-Path $newPath)) {
                                    Write-ColoredOutput "Path does not exist. Creating..." -Color $script:Colors.Warning
                                    New-Item -ItemType Directory -Path $newPath -Force | Out-Null
                                }
                                $script:ReportsLocation = $newPath
                                Write-ColoredOutput "Reports location changed to: $script:ReportsLocation" -Color $script:Colors.Success
                                Write-Log "Reports location changed to: $script:ReportsLocation" -Level Info
                            }
                            else {
                                Write-ColoredOutput "No folder selected. Operation cancelled." -Color $script:Colors.Warning
                            }
                        }
                        "3" {
                            # Change Logs Location
                            Write-Host "  Opening folder browser..." -ForegroundColor Gray
                            $newPath = Show-FolderBrowserDialog -Description "Select new logs location" -SelectedPath $script:LogLocation
                            if (-not [string]::IsNullOrWhiteSpace($newPath)) {
                                if (-not (Test-Path $newPath)) {
                                    Write-ColoredOutput "Path does not exist. Creating..." -Color $script:Colors.Warning
                                    New-Item -ItemType Directory -Path $newPath -Force | Out-Null
                                }
                                $script:LogLocation = $newPath
                                # Redirect current session log file to the new location
                                $logTimestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
                                $script:LogFile = Join-Path $script:LogLocation "UniFy-Endpoint_v2_$logTimestamp.log"
                                Write-ColoredOutput "Logs location changed to: $script:LogLocation" -Color $script:Colors.Success
                                Write-Log "Logs location changed to: $script:LogLocation" -Level Info
                            }
                            else {
                                Write-ColoredOutput "No folder selected. Operation cancelled." -Color $script:Colors.Warning
                            }
                        }
                        "4" {
                            # Change Exports Location
                            Write-Host "  Opening folder browser..." -ForegroundColor Gray
                            $newPath = Show-FolderBrowserDialog -Description "Select new exports location" -SelectedPath $script:ExportsLocation
                            if (-not [string]::IsNullOrWhiteSpace($newPath)) {
                                if (-not (Test-Path $newPath)) {
                                    Write-ColoredOutput "Path does not exist. Creating..." -Color $script:Colors.Warning
                                    New-Item -ItemType Directory -Path $newPath -Force | Out-Null
                                }
                                $script:ExportsLocation = $newPath
                                Write-ColoredOutput "Exports location changed to: $script:ExportsLocation" -Color $script:Colors.Success
                                Write-Log "Exports location changed to: $script:ExportsLocation" -Level Info
                            }
                            else {
                                Write-ColoredOutput "No folder selected. Operation cancelled." -Color $script:Colors.Warning
                            }
                        }
                        "0" { }
                        default {
                            Write-ColoredOutput "Invalid option. Please select 1-4 or 0." -Color $script:Colors.Warning
                        }
                    }

                } while ($subChoice14 -ne "0")
            }
            "15" {
                # View/Open Log Files
                if (Test-Path $script:LogLocation) {
                    $logFiles = Get-ChildItem -Path $script:LogLocation -Filter "*.log" | Sort-Object LastWriteTime -Descending

                    if ($logFiles.Count -gt 0) {
                        Write-ColoredOutput "`nRecent log files:" -Color $script:Colors.Info

                        for ($i = 0; $i -lt [Math]::Min(10, $logFiles.Count); $i++) {
                            Write-Host "  [$($i + 1)]" -ForegroundColor Yellow -NoNewline
                            Write-Host " $($logFiles[$i].Name)" -ForegroundColor White -NoNewline
                            Write-Host " - $($logFiles[$i].LastWriteTime)" -ForegroundColor Gray
                        }

                        $logChoice = Read-Host "`nEnter log number to open (or 0 to return)"

                        if ($logChoice -ne "0" -and $logChoice -le $logFiles.Count) {
                            $selectedLog = $logFiles[$logChoice - 1]
                            Start-Process notepad.exe -ArgumentList $selectedLog.FullName
                        }
                    }
                    else {
                        Write-ColoredOutput "No log files found" -Color $script:Colors.Warning
                    }
                }
                else {
                    Write-ColoredOutput "Log directory not found" -Color $script:Colors.Warning
                }

                Read-Host "`nPress Enter to return to main menu (or 0)"
            }
            "16" {
                # Configure Logging
                Write-Host "`n  Logging Configuration" -ForegroundColor Cyan
                Write-Host "  Current Status: " -NoNewline
                if ($script:LoggingEnabled) {
                    Write-Host "Enabled" -ForegroundColor Green
                    Write-Host "  Log Location: $($script:LogLocation)" -ForegroundColor White
                    Write-Host "  Log Level: $($script:CurrentLogLevel)" -ForegroundColor White
                } else {
                    Write-Host "Disabled" -ForegroundColor Yellow
                }

                Write-Host "`n  [1] Enable Logging" -ForegroundColor Yellow
                Write-Host "  [2] Disable Logging" -ForegroundColor Yellow
                Write-Host "  [3] Change Log Level" -ForegroundColor Yellow
                Write-Host "  [4] Change Log Location" -ForegroundColor Yellow
                Write-Host "  [0] Cancel" -ForegroundColor Red

                $logChoice = Read-Host "`nSelect option"

                switch ($logChoice) {
                    "1" {
                        $script:LoggingEnabled = $true
                        Initialize-Logging
                        Write-ColoredOutput "Logging enabled" -Color $script:Colors.Success
                    }
                    "2" {
                        $script:LoggingEnabled = $false
                        Write-ColoredOutput "Logging disabled" -Color $script:Colors.Success
                    }
                    "3" {
                        Write-Host "`n  Select Log Level:" -ForegroundColor Cyan
                        Write-Host "  [1] Verbose" -ForegroundColor Yellow
                        Write-Host "  [2] Info" -ForegroundColor Yellow
                        Write-Host "  [3] Warning" -ForegroundColor Yellow
                        Write-Host "  [4] Error" -ForegroundColor Yellow

                        $levelChoice = Read-Host "`nSelect level"
                        switch ($levelChoice) {
                            "1" { $script:CurrentLogLevel = "Verbose" }
                            "2" { $script:CurrentLogLevel = "Info" }
                            "3" { $script:CurrentLogLevel = "Warning" }
                            "4" { $script:CurrentLogLevel = "Error" }
                        }
                        Write-ColoredOutput "Log level set to: $($script:CurrentLogLevel)" -Color $script:Colors.Success
                        Write-Log "Log level changed to: $($script:CurrentLogLevel)" -Level Info
                    }
                    "4" {
                        $newLogPath = Read-Host "Enter new log location"
                        if (-not [string]::IsNullOrWhiteSpace($newLogPath)) {
                            $script:LogLocation = $newLogPath
                            if (-not (Test-Path $script:LogLocation)) {
                                New-Item -ItemType Directory -Path $script:LogLocation -Force | Out-Null
                            }
                            Write-ColoredOutput "Log location changed to: $($script:LogLocation)" -Color $script:Colors.Success
                            Write-Log "Log location changed to: $($script:LogLocation)" -Level Info
                        }
                    }
                }

                Read-Host "`nPress Enter to return to main menu (or 0)"
            }
            { $_ -eq "Q" -or $_ -eq "q" } {
                # Disconnect and switch tenant
                Write-ColoredOutput "`n════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
                Write-ColoredOutput " DISCONNECT FROM MICROSOFT GRAPH" -Color $script:Colors.Warning
                Write-ColoredOutput "════════════════════════════════════════════════════════════════" -Color $script:Colors.Header
                Write-ColoredOutput "This will:" -Color $script:Colors.Info
                Write-ColoredOutput "  - Disconnect from current Microsoft Graph session" -Color $script:Colors.Default
                Write-ColoredOutput "  - Allow you to connect with different credentials/tenant" -Color $script:Colors.Default

                $confirm = Read-Host "`nDisconnect and switch tenant? (yes/no)"
                if ($confirm -eq "yes") {
                    try {
                        Write-ColoredOutput "`nDisconnecting from Microsoft Graph..." -Color $script:Colors.Info
                        Disconnect-MgGraph -ErrorAction SilentlyContinue
                        Write-ColoredOutput "Disconnected successfully." -Color $script:Colors.Success
                        Write-Log "Disconnected from Microsoft Graph" -Level Info

                        # Reconnect with new credentials
                        Write-ColoredOutput "`nConnecting with new credentials..." -Color $script:Colors.Info
                        $reconnected = Connect-UniFy-Endpoint

                        # Check reconnection result (must check type explicitly due to PowerShell type coercion)
                        if ($reconnected -is [bool] -and $reconnected -eq $true) {
                            Write-ColoredOutput "Connected successfully to new tenant!" -Color $script:Colors.Success
                            Write-Log "Reconnected to new tenant" -Level Info
                        }
                        elseif ($reconnected -is [string] -and $reconnected -eq "cancelled") {
                            Write-ColoredOutput "`nOperation cancelled. You are now disconnected from Microsoft Graph." -Color $script:Colors.Warning
                            Write-ColoredOutput "Please restart UniFy-Endpoint to connect again." -Color $script:Colors.Info
                            $continue = $false
                        }
                        else {
                            Write-ColoredOutput "Failed to reconnect. Please restart UniFy-Endpoint." -Color $script:Colors.Error
                            $continue = $false
                        }
                    }
                    catch {
                        Write-ColoredOutput "Error during disconnect: $_" -Color $script:Colors.Error
                        Write-Log "Disconnect error: $_" -Level Error
                    }
                }
                else {
                    Write-ColoredOutput "Disconnect cancelled." -Color $script:Colors.Info
                }

                if ($continue) {
                    Read-Host "`nPress Enter to return to main menu"
                }
            }
            "0" {
                # Exit
                $continue = $false
                Write-ColoredOutput "`nDisconnecting from Microsoft Graph..." -Color $script:Colors.Info
                try {
                    Disconnect-MgGraph -ErrorAction SilentlyContinue
                }
                catch {
                    # Silently ignore disconnect errors on exit
                }
                Write-ColoredOutput "Exiting UniFy-Endpoint. Goodbye!" -Color $script:Colors.Success
            }
            default {
                Write-ColoredOutput "Invalid selection. Please try again." -Color $script:Colors.Warning
                Read-Host "`nPress Enter to return to main menu"
            }
        }
    }
}

#endregion

#region Main Execution

# Initialize logging
Initialize-Logging

# Connect to Microsoft Graph
$connected = Connect-UniFy-Endpoint

# Check connection result (must check type explicitly due to PowerShell type coercion)
if ($connected -is [bool] -and $connected -eq $true) {
    # Success - continue
}
elseif ($connected -is [string] -and $connected -eq "cancelled") {
    Write-ColoredOutput "`nAuthentication cancelled. Exiting UniFy-Endpoint." -Color $script:Colors.Warning
    Close-Logging
    exit 0
}
else {
    Write-ColoredOutput "Failed to connect. Exiting." -Color $script:Colors.Error
    Close-Logging
    exit 1
}

# Execute requested operation or enter interactive mode
if ($Backup) {
    Ensure-BackupDirectory
    Backup-IntuneConfiguration -BackupName "" -SelectedComponents (Get-SelectedComponents -ComponentList $Components)
}
elseif ($List) {
    Show-BackupList | Out-Null
}
else {
    # Interactive mode
    Start-InteractiveMode
}

# Disconnect from Graph
Write-ColoredOutput "`nDisconnecting from Microsoft Graph..." -Color $script:Colors.Info
Write-Log "Disconnecting from Microsoft Graph" -Level Info
try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
catch {
    # Silently ignore disconnect errors (already disconnected or never connected)
}

# Close logging
Close-Logging

Write-Host ""
Write-Host "  UniFy-Endpoint session completed successfully!" -ForegroundColor Green
Write-Host ""

#endregion



