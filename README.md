[![Download Latest Release](https://img.shields.io/github/v/release/UniFy-Endpoint/UniFy-Endpoint-Tool?label=Download%20Latest&style=for-the-badge&logo=github)](https://github.com/UniFy-Endpoint/UniFy-Endpoint-Tool/releases/latest)

# UniFy-Endpoint v2.0 - Enterprise Intune Backup & Restore Solution

<p align="center">
  <img src="https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg" alt="PowerShell 5.1+">
  <img src="https://img.shields.io/badge/Version-2.0.0-green.svg" alt="Version 2.0.0">
  <img src="https://img.shields.io/badge/Components-16-orange.svg" alt="16 Components">
  <img src="https://img.shields.io/badge/Platforms-Windows%20%7C%20iOS%20%7C%20Android%20%7C%20macOS-purple.svg" alt="Multi-Platform">
</p>

UniFy-Endpoint is a comprehensive PowerShell-based tool for backing up and restoring Microsoft Intune configurations. The tool emphasizes **safety** by implementing read-only backups and non-destructive restore operations that create new policies rather than modifying existing ones.

**Original Author:** [Ugur Koc](https://github.com/ugurkocde)
**Enhanced By:** Yoennis Olmo

---

## What's New in v2.0

| Metric | v1.0 | v2.0 |
|--------|:----:|:----:|
| **Supported Components** | 6 | **16** |
| **Export Functions** | 3 | **22** |
| **Menu Options** | 13 | **16 + sub-menus** |
| **Total Functions** | 30 | **83** |
| **Lines of Code** | 3,112 | **10,390** |

### Key Enhancements

- **10 New Component Types** - Autopilot Profiles, Autopilot Device Prep, ESP, Remediation Scripts, WUfB, Assignment Filters, App Config (MDM & MAM), macOS Scripts, macOS Custom Attributes
- **Current Tenant HTML Report** - Generate live snapshots from current Intune tenant (Option [5] → [2])
- **Client Secret Authentication** - App-only auth with AppId/TenantId/ClientSecret for automated workflows
- **Platform Filtering** - Filter backups by Windows, iOS, Android, or macOS
- **macOS Support** - Full backup/restore for macOS Shell Scripts and Custom Attributes
- **Interactive HTML Reports** - Search, filters, expandable cards, Settings Catalog resolution
- **Enhanced Drift Detection** - All 16 component types with hash-based comparison
- **Clean JSON Protocol** - Exports ready for immediate import (no manual editing required)
- **Nested Assignments** - Policy assignments embedded directly in JSON exports
- **Enterprise Pagination** - Handles 1000+ policies with auto-retry for `$top` errors
- **Base64 Auto-Decoding** - Scripts exported as readable text
- **Enhanced File Naming** - `PolicyName_Platform_ID.json` format

---

## Key Features

### Backup & Restore Capabilities

- Automated backup of **16 Intune configuration types** to local JSON files
- Safe restore operations that create policies with `[Restored]` prefix
- **Preview mode** for dry-run testing before actual changes
- No risk of overwriting production configurations
- **Platform-specific** backup and restore operations
- **Component-selective** backup for granular control

### Configuration Management

- **Drift detection** comparing current tenant state with backups (all 16 component types)
- **Current Tenant HTML Report** - Live snapshot from Graph API without creating a backup
- **Interactive HTML reports** - Search, platform filters, expandable cards, Settings Catalog resolution
- Multiple export formats (Markdown, CSV, HTML)
- Comprehensive logging with configurable verbosity levels
- Timestamped, organized backup structure
- **Log file cleanup** with configurable retention

### Enterprise Features

- Certificate-based authentication for automation
- Enterprise-scale pagination (unlimited policy counts)
- Progress indicators for bulk operations
- Detailed error code handling (403/404/400)

---

## Supported Components (16 Total)

### Device Management

| Component | Description | API Version |
|-----------|-------------|:-----------:|
| Device Configurations | Legacy device configuration profiles | Beta |
| Compliance Policies | Device compliance requirements | Beta |
| Settings Catalog | Modern configuration policies | Beta |
| Administrative Templates | Group Policy-based configurations (ADMX) | Beta |
| Windows Update for Business | Windows update rings and policies | Beta |

### Windows Autopilot & Enrollment

| Component | Description | API Version |
|-----------|-------------|:-----------:|
| Autopilot Profiles | Windows Autopilot deployment profiles | Beta |
| Autopilot Device Preparation | Device preparation policies | Beta |
| Enrollment Status Page (ESP) | Windows enrollment experience | V1.0 |

### Scripts & Remediations

| Component | Description | API Version |
|-----------|-------------|:-----------:|
| PowerShell Scripts | Windows device management scripts | Beta |
| Remediation Scripts | Proactive remediation scripts | Beta |
| macOS Shell Scripts | macOS device management scripts | Beta |
| macOS Custom Attributes | macOS inventory extension scripts | Beta |

### App Management

| Component | Description | API Version |
|-----------|-------------|:-----------:|
| App Protection Policies | iOS/Android/Windows MAM policies | Beta |
| App Config (Managed Devices) | Device-level app configuration | V1.0 |
| App Config (Managed Apps) | App-specific configuration | V1.0 |

### Other

| Component | Description | API Version |
|-----------|-------------|:-----------:|
| Assignment Filters | Dynamic device assignment filters | Beta |

---

## Safety First

UniFy-Endpoint is designed with safety as the primary concern:

1. Read-Only Backups: The backup operation only reads from your Intune tenant. It never modifies anything.

2. Non-Destructive Restores: Restore operations create NEW policies with a "[Restored]" prefix. Your existing policies remain untouched.

3. Preview Mode: Test any restore operation without making actual changes.

4. Explicit Confirmations: All potentially impactful operations require explicit user confirmation.

---

## Requirements & Prerequisites

### Prerequisites

- **Windows PowerShell 5.1** or **PowerShell Core 7.x**
- **Microsoft.Graph.Authentication** module (auto-installed on UniFy-Endpoint first run)
- **Global Administrator** or **Intune Administrator** Role Premissions

---

## Required Permissions

Your account or app registration needs the following Microsoft Graph permissions:

| Permission | Type | Purpose |
|------------|------|---------|
| `DeviceManagementConfiguration.ReadWrite.All` | Delegated/Application | Device configurations, compliance, settings catalog |
| `DeviceManagementServiceConfig.ReadWrite.All` | Delegated/Application | Autopilot, ESP, enrollment |
| `DeviceManagementApps.ReadWrite.All` | Delegated/Application | App protection, app configuration |
| `DeviceManagementManagedDevices.ReadWrite.All` | Delegated/Application | Scripts, remediations |
| `AuditLog.Read.All` | Delegated/Application | Audit logs (fallback for policy timestamps) |

> **Note:** ReadWrite scopes are required even though backup uses read-only API calls. `AuditLog.Read.All` is optional but recommended for accurate timestamps.

---

### Command-Line Mode

```powershell
# Full backup (all components, all platforms)
.\UniFy-Endpoint_v2.0.ps1 -Backup

# Backup with custom name and location
.\UniFy-Endpoint_v2.0.ps1 -Backup -BackupName "Full-Backup-2026-01-24" -BackupPath "C:\UniFy-Endpoint\Backups"

# Platform-specific backup
.\UniFy-Endpoint_v2.0.ps1 -Backup -Platform Windows
.\UniFy-Endpoint_v2.0.ps1 -Backup -Platform iOS
.\UniFy-Endpoint_v2.0.ps1 -Backup -Platform Android
.\UniFy-Endpoint_v2.0.ps1 -Backup -Platform macOS

# Component-specific backup
.\UniFy-Endpoint_v2.0.ps1 -Backup -Components AutopilotProfiles,EnrollmentStatusPage

# Restore from backup
.\UniFy-Endpoint_v2.0.ps1 -Restore -RestorePath "C:\UniFy-Endpoint\Full-Backup-2026-01-24"

# Export to different formats
.\UniFy-Endpoint_v2.0.ps1 -Export -ExportFormat JSON
.\UniFy-Endpoint_v2.0.ps1 -Export -ExportFormat CSV
.\UniFy-Endpoint_v2.0.ps1 -Export -ExportFormat HTML
```

### Certificate Authentication (Recommended for Automation)

```powershell
.\UniFy-Endpoint_v2.0.ps1 -Backup `
  -AppId "12345678-1234-1234-1234-123456789012" `
  -TenantId "contoso.onmicrosoft.com" `
  -CertificateThumbprint "A1B2C3D4E5F6G7H8I9J0K1L2M3N4O5P6Q7R8S9T0"
```

### Client Secret Authentication (Alternative)

```powershell
.\UniFy-Endpoint_v2.0.ps1 -Backup `
  -AppId "12345678-1234-1234-1234-123456789012" `
  -TenantId "contoso.onmicrosoft.com" `
  -ClientSecret "your-client-secret-value"
```

> **Note:** Certificate authentication is recommended for production environments as it is more secure than client secrets. Client Secret authentication requires Microsoft.Graph.Authentication v2.x or later.

---

## Main Menu (16 Options + Sub-menus)

```
=== Backup & Restore ===
[1] Backup Intune Configuration (Read-Only)
[2] Preview Restore (Dry Run)
[3] Drift Detection & Selective Restore
[4] Full Restore Intune Configuration (Creates New Policies)

=== Reports & Export ===
[5] Generate HTML Report (Backup / Current Intune Tenant)  ← Sub-menu
[6] Export Backup (MD/CSV)
[7] Import Configuration from File (Multi-select supported)

=== Management ===
[8] List Available Backups
[9] Compare Two Backups
[10] Cleanup Old Backups
[11] Cleanup Old Log Files

=== Advanced Options ===
[12] Backup/Restore by Platform                            ← Sub-menu
[13] Backup/Restore by Components                          ← Sub-menu

=== Settings ===
[14] Change Backup Location
[15] View/Open Log Files
[16] Configure Logging

[Q] Disconnect & Switch Tenant
[0] Exit
```

### Sub-menu: Option [5] - Generate HTML Report

```
[1] From Backup (select a saved backup)
[2] From Current Intune Tenant (fetch current configuration)  ← NEW
[0] Return to Main Menu
```

---

## Component Selection Menu

When using option **[13] Backup/Restore by Components**, you can select specific components:

```
=== Device Policies ===
[1] Device Configurations
[2] Compliance Policies
[3] Settings Catalog
[4] Administrative Templates (Group Policy ADMX)
[5] Windows Update for Business (WUfB)

=== Autopilot & Enrollment ===
[6] Autopilot Profiles
[7] Autopilot Device Preparation
[8] Enrollment Status Page (ESP)

=== Scripts & Remediations ===
[9] PowerShell Scripts
[10] Remediation Scripts
[11] macOS Shell Scripts
[12] macOS Custom Attributes

=== App Management ===
[13] App Protection Policies (MAM)
[14] App Configuration (Managed Devices)
[15] App Configuration (Managed Apps)

=== Assignment Filters ===
[16] Assignment Filters

[A] Select All Components
[0] Return to Main Menu
```

---

## Platform Filtering

Filter backups by operating system platform:

| Platform | Supported Components |
|----------|---------------------|
| **Windows** | All components except macOS Scripts |
| **iOS** | Device Configs, Compliance, Settings Catalog, App Protection, App Config |
| **Android** | Device Configs, Compliance, Settings Catalog, App Protection, App Config |
| **macOS** | Device Configs, Compliance, Settings Catalog, macOS Scripts, Custom Attributes, Assignment Filters |

### Platform Detection Logic

The tool automatically detects platform from:
- `platform` property
- `appType` property
- `platforms` array
- `@odata.type` hints

---

## Backup Structure

```
backup-2026-01-24-143052/
├── metadata.json
├── DeviceConfigurations/
│   ├── Email_Profile_iOS_efg-123.json
│   └── WiFi_Config_Android_hij-456.json
├── CompliancePolicies/
│   └── Windows_10_Compliance_Windows_nop-012.json
├── SettingsCatalogPolicies/
│   └── Security_Baseline_windows10_tuv-678.json
├── AdministrativeTemplates/
│   └── Office_GPO_Settings_Windows_abc-123.json
├── WUfBPolicies/
│   └── Feature_Update_Ring_Windows_def-456.json
├── AutopilotProfiles/
│   ├── Corporate_OOBE_Windows_ghi-789.json
│   └── Shared_Device_Windows_jkl-012.json
├── AutopilotDevicePrep/
│   └── Device_Prep_Policy_mno-345.json
├── EnrollmentStatusPage/
│   └── Standard_ESP_Windows_pqr-678.json
├── PowerShellScripts/
│   ├── Configure_Bitlocker_Windows_stu-901.json
│   └── Install_Software_Windows_vwx-234.json
├── RemediationScripts/
│   ├── Check_Disk_Space_Windows_yza-567.json
│   └── Fix_Network_Windows_bcd-890.json
├── MacOSScripts/
│   └── Setup_Homebrew_macOS_efg-123.json
├── MacOSCustomAttributes/
│   └── Get_Serial_Number_macOS_hij-456.json
├── AppProtectionPolicies/
│   ├── Outlook_MAM_iOS_klm-789.json
│   └── Teams_MAM_Android_nop-012.json
├── AppConfigManagedDevices/
│   └── VPN_App_Config_iOS_qrs-345.json
├── AppConfigManagedApps/
│   └── Outlook_Settings_Android_tuv-678.json
└── AssignmentFilters/
    └── Surface_Devices_Windows_wxy-901.json
```

---

## Clean JSON Protocol

All exports automatically remove read-only properties to prevent errors during import:

| Property Removed | Description |
|------------------|-------------|
| `id` | System-generated unique identifier |
| `createdDateTime` | Creation timestamp |
| `lastModifiedDateTime` | Last modification timestamp |
| `version` | Policy version number |
| `isAssigned` | Assignment status flag |
| `@odata.context` | OData context URL |
| `@odata.etag` | Entity tag for concurrency |

---

## Nested Assignment Export

All policy exports include embedded assignment information:

```json
{
  "displayName": "Windows 10 Compliance Policy",
  "platform": "windows10AndLater",
  "settings": { ... },
  "assignments": [
    {
      "target": {
        "@odata.type": "#microsoft.graph.groupAssignmentTarget",
        "groupId": "abc-123-def"
      }
    }
  ]
}
```

---

## Scheduling Automated Backups

### Windows Task Scheduler

**Program/script:**
```
powershell.exe
```

**Arguments (Certificate Auth - Recommended):**
```
-ExecutionPolicy Bypass -File "C:\UniFy-Endpoint\UniFy-Endpoint_v2.0.ps1" -Backup -AppId "your-app-id" -TenantId "your-tenant-id" -CertificateThumbprint "your-thumbprint"
```

**Arguments (Client Secret Auth - Alternative):**
```
-ExecutionPolicy Bypass -File "C:\UniFy-Endpoint\UniFy-Endpoint_v2.0.ps1" -Backup -AppId "your-app-id" -TenantId "your-tenant-id" -ClientSecret "your-secret"
```

### PowerShell Scheduled Script

**Option 1: Certificate Authentication (Recommended)**
```powershell
# RunBackup.ps1
$scriptPath = "C:\UniFy-Endpoint\UniFy-Endpoint_v2.0.ps1"
$backupPath = "D:\IntuneBackups"
$appId = "your-app-id"
$tenantId = "your-tenant-id"
$certThumbprint = "your-cert-thumbprint"

& $scriptPath -Backup `
  -BackupPath $backupPath `
  -AppId $appId `
  -TenantId $tenantId `
  -CertificateThumbprint $certThumbprint `
  -LogLevel Info
```

**Option 2: Client Secret Authentication**
```powershell
# RunBackup.ps1
$scriptPath = "C:\UniFy-Endpoint\UniFy-Endpoint_v2.0.ps1"
$backupPath = "D:\IntuneBackups"
$appId = "your-app-id"
$tenantId = "your-tenant-id"
$clientSecret = "your-client-secret"

& $scriptPath -Backup `
  -BackupPath $backupPath `
  -AppId $appId `
  -TenantId $tenantId `
  -ClientSecret $clientSecret `
  -LogLevel Info
```

---

## Best Practices

1. **Schedule regular daily backups** for production environments
2. **Create backups before major configuration changes**
3. **Test restores in development environments first**
4. **Store backups in secure locations** with restricted access
5. **Use platform filtering** to reduce backup scope and execution time
6. **Use certificate authentication** for automated/scheduled backups
7. **Regularly clean up old backup and log files**
8. **Document your backup naming conventions**

---

## Troubleshooting

### "Module not found"

```powershell
# Install Microsoft Graph module
Install-Module -Name Microsoft.Graph.Authentication -Force
Import-Module Microsoft.Graph.Authentication
```

### "Permission denied (403)"

```powershell
# Reconnect with proper scopes
Connect-MgGraph -Scopes @(
    "DeviceManagementConfiguration.ReadWrite.All",
    "DeviceManagementServiceConfig.ReadWrite.All",
    "DeviceManagementApps.ReadWrite.All",
    "DeviceManagementManagedDevices.ReadWrite.All"
)
```

### "Execution policy error"

```powershell
# Option 1: Bypass for this session
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# Option 2: Set for current user
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

### "No policies found for platform"

- Verify policies exist for the selected platform in Intune portal
- Try `-Platform All` to see all available policies
- Check that the correct Graph API permissions are granted

---

## Documentation

| Document | Description |
|----------|-------------|
| [CHANGELOG.md](CHANGELOG.md) | Detailed v1.0 vs v2.0 comparison |
| [QUICKSTART_v2.md](QUICKSTART_v2.md) | Quick start guide |
| [UPGRADE_SUMMARY.md](UPGRADE_SUMMARY.md) | Feature documentation |
| [API_REFERENCE_v2.md](API_REFERENCE_v2.md) | Graph API endpoint reference |
| [DELIVERY_NOTES.md](DELIVERY_NOTES.md) | Implementation summary |

---

## Support & Resources

- **GitHub Issues:** [https://github.com/ugurkocde/UniFy-Endpoint/issues](https://github.com/ugurkocde/UniFy-Endpoint/issues)
- **Microsoft Graph Docs:** [https://learn.microsoft.com/graph](https://learn.microsoft.com/graph)
- **Intune API Reference:** [https://learn.microsoft.com/graph/api/resources/intune-graph-overview](https://learn.microsoft.com/graph/api/resources/intune-graph-overview)

---

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

**UniFy-Endpoint v2.0** - Your Enterprise Intune Configuration Safety Net
