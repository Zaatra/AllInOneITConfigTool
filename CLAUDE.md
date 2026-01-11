# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Trust Palestine IT Configuration Tool (v18.3)** is a Windows Forms GUI application for enterprise IT system management, packaged as a single PowerShell script (`TrustInstaller.ps1`, 2,660 lines). It provides automated software deployment, HP driver management, and system configuration for Windows workstations, with special support for legacy systems (Windows 7/8, PowerShell 2.0).

**Key Features:**
- 25+ application installations via WinGet, Office Deployment Tool, or local installers
- HP driver scanning and installation using HPIA or CMSL (with offline repository support)
- System configuration (timezone, locale, power settings, boot options)
- Backward compatibility with legacy Windows and HP hardware

## Architecture

### Single-File Structure

The entire application is contained in `TrustInstaller.ps1`, organized into 8 sections:

1. **Admin Privilege Elevation** (UAC check and elevation)
2. **Global Configuration** (lines 36-164):
   - `$Global:CrowdStrikeCID` - Falcon Sensor organization ID
   - `$Global:HPLegacyRepoRoot` - Network share path for offline HP driver repository
   - `$Global:AppList` - Master application registry with metadata
   - `$Office2024XML` / `$Office365XML` - Office Deployment Tool configurations
3. **Utility Functions** (version detection, file operations, registry queries)
4. **GUI Form Creation** (Windows Forms with 3 tabs: Installation, Drivers, System Config)
5. **Application Installation Logic** (WinGet, Office ODT, local installers)
6. **Driver Management** (HP scanning, downloading, installation)
7. **Event Handlers** (button clicks, tab changes, UI interactions)
8. **Form Initialization** (display and main loop)

### Data Structures

#### Application Registry (`$Global:AppList`)

Each app is a hashtable with these keys:
```powershell
@{
    Cat              = "Category"              # UI grouping
    Name             = "Display Name"          # Unique identifier
    Id               = "WinGet Package ID"     # For single-arch apps
    Id64/Id86        = "x64/x86 Package IDs"   # For dual-arch apps
    Args             = "Installer arguments"   # Silent install flags
    DualArch         = $true/$false            # Multi-architecture support
    DownloadMode     = "winget|office|direct|localonly|onlineonly"
    FileStem         = "filename-prefix"       # For local file matching
    Pat              = "Display Name Regex"    # Registry detection pattern
    LocalAltNames    = @("alt-exe-names")      # Alternative file names
    VCKey            = "2005|2008|..."         # VC++ version key
}
```

**Categories:**
- **VC++ Redistributables**: 2005-2015+ (dual-arch x64/x86)
- **Core Applications**: Java, Chrome, Firefox, WinRAR, TeamViewer, NAPS2, K-Lite Codec
- **Office**: 2024 LTSC (volume license), 365 Enterprise
- **Driver & Support Tools**: Intel DSA, HP Support Assistant
- **Security & VMS**: FortiClient VPN, iVMS-4200, CrowdStrike Falcon Sensor

#### Driver Data Structure

Driver grid columns (DataGridView):
- `Select` (checkbox), `Source` (HPIA/CMSL/Legacy), `Status`, `Name`, `Category`, `InstalledVer`, `LatestVer`, `SPID`

**Status Values:**
- `Critical` (red) - Security/critical updates
- `Update Available` (orange) - Non-critical updates
- `Not Installed` (default) - Missing drivers
- `Installed` (green) - Current version installed

### Core Functions by Purpose

#### Version Detection & Comparison
- `Normalize-VersionString($ver)` - Converts "1.2" → "1.2.0.0" for comparison
- `To-VersionObj($verStr)` - Converts string to `[version]` object
- `Get-FileProductVersion($path)` - Extracts version from executable metadata
- `Find-InstalledVersion($app)` - Searches registry for installed app version
- `Get-VCInstalledMap()` - Detects installed Visual C++ redistributables (x64/x86)

#### Application Management
- `Check-LatestVersions()` - Queries WinGet/online sources for latest versions
- `Get-LocalInstallerInfo($app)` - Finds cached installer files with version detection
- `Winget-Download-And-Rename($app, $outDir, $arch)` - Downloads via WinGet with version tagging

#### Driver Management
- `Get-HPSystemInfo()` - Detects HP model/SKU/BIOS via WMI (`Win32_BaseBoard`, `Win32_BIOS`)
- `Get-HPIAPath()` / `Download-HPIA()` - HP Image Assistant acquisition
- `Test-HPCMSLInstalled()` / `Install-HPCMSL()` - HP Command Line Shell management
- `Scan-Drivers()` - Scans hardware using HPIA/CMSL or legacy repository
- `Download-SelectedDrivers()` / `Install-SelectedDrivers()` - Driver batch operations
- `Get-LegacyRepoPackages($model, $osVer)` - Offline repository support for legacy systems

#### System Configuration
- `Check-SystemConfiguration()` - Validates current settings (timezone, power, boot, locale)
- `Apply-SystemConfiguration()` - Applies:
  - Timezone: "West Bank Standard Time"
  - Power: High Performance mode
  - Fast Boot: Disabled (registry: `HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power`)
  - Locale: Arabic (ar-SA) / US English
  - Desktop Icons: Enabled

#### UI Utilities
- `Enable-DoubleBuffer($ctrl)` - Reduces DataGridView flicker
- `Set-RowColor($row, $status)` - Color-codes rows by status
- `Safe-FilePart($str)` - Sanitizes filenames for file system operations
- `Get-CompatInstance($className, $filter)` - Abstracts CIM (modern) vs WMI (legacy) for PowerShell 2.0 compatibility

### GUI Components

**3 Tabs:**

1. **Installation Tab** (`$TabInstall`)
   - Dynamically generated checkboxes (grouped by category)
   - Version indicators: Installed vs. Latest
   - Color coding: Green=installed, Orange=update available, Red=missing
   - Buttons: Download, Install, Check Updates, Refresh UI

2. **Drivers Tab** (`$TabDrivers`)
   - DataGridView with sortable columns
   - Auto-sort by status (Critical > Update > Not Installed > Installed)
   - Selection helpers: "Select All", "Select None", "Select Needs Update", "Select Not Installed", "Select Installed"
   - Buttons: Scan Drivers, Download, Install
   - Real-time logging window (operations log)

3. **System Config Tab** (`$TabSystem`)
   - Timezone, Power Settings, Boot Options, Locale Configuration
   - Apply System Settings button
   - Detailed operation logging

## Development Workflow

### Execution

**Prerequisites:**
- Windows OS (7 or later)
- PowerShell 2.0+ (script is backward-compatible)
- Administrator privileges (automatically elevated via UAC)

**Run Script:**
```powershell
# From PowerShell
powershell -ExecutionPolicy Bypass -File TrustInstaller.ps1

# Or right-click → Run with PowerShell (auto-elevates)
```

**No Build Process** - This is a standalone script. No compilation, bundling, or dependency installation required.

### Testing

**No Automated Testing Framework** - All validation is manual through:
- Registry checks (`HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall`)
- WMI/CIM queries (`Win32_BaseBoard`, `Win32_BIOS`, `Win32_Product`)
- File system checks (local installer cache, network shares)
- Version comparisons (installed vs. latest from WinGet/online sources)

**Manual Testing Approach:**
1. UI interactions (button clicks, checkbox selections, tab changes)
2. Installation verification (check registry after app installation)
3. Driver scanning (verify detected hardware matches system)
4. System configuration (verify timezone, power, locale changes applied)

### Code Patterns

#### PowerShell Conventions
- **Functions**: `Verb-Noun` (e.g., `Get-HPSystemInfo`, `Test-HPCMSLInstalled`)
- **Variables**: `camelCase` for local, `$Global:` prefix for shared state
- **UI Controls**: Hungarian notation (`$btnInstall`, `$lblStatus`, `$txtLog`)

#### Version Comparison Pattern
```powershell
# Always normalize versions before comparison
$installedVer = Normalize-VersionString "1.2.3"
$latestVer = Normalize-VersionString "1.2.4"
$v1 = To-VersionObj $installedVer
$v2 = To-VersionObj $latestVer
if ($v1 -lt $v2) { # Update available }
```

#### Registry-Based Detection
```powershell
# Search both 64-bit and 32-bit registry hives
$paths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)
Get-ItemProperty $paths | Where-Object { $_.DisplayName -match $pattern }
```

#### WMI/CIM Compatibility (PowerShell 2.0)
```powershell
# Use Get-CompatInstance for cross-version compatibility
$board = Get-CompatInstance "Win32_BaseBoard" | Select-Object -First 1
# Falls back to Get-WmiObject on PowerShell 2.0, uses Get-CimInstance on modern versions
```

#### Error Handling
```powershell
# Silent failures for optional operations
$result = Get-SomeData -ErrorAction SilentlyContinue

# User-friendly error dialogs
try {
    # risky operation
} catch {
    [System.Windows.Forms.MessageBox]::Show("Operation failed: $_", "Error")
}
```

## Configuration

### Global Configuration (Lines 36-39)

```powershell
# CrowdStrike Falcon Sensor Organization ID
$Global:CrowdStrikeCID = "753B2C62DBF944A59165207E3FE2FF2A-27"

# HP Legacy Driver Repository (network share for offline systems)
$Global:HPLegacyRepoRoot = "\\192.168.168.6\Admin Tools\Drivers-Repo"
```

**To customize for your environment:**
1. Update CrowdStrike CID to match your organization (line 36)
2. Update HP Legacy Repository path to your network share (line 37)
3. Modify Office XML configurations for your licensing (lines 40-90)
4. Adjust application list in `$Global:AppList` (lines 95-164)

### Office Deployment Tool Configurations

Two pre-configured XML templates (lines 40-90):

**Office 2024 LTSC** (Perpetual Volume License):
- Channel: `PerpetualVL2024`
- Languages: English (en-us), Arabic (ar-sa)
- Excludes: Lync, OneDrive, Publisher
- Product Key: Embedded in XML (PIDKEY attribute)

**Office 365 Enterprise** (Current Channel):
- Channel: `Current`
- Languages: English (en-us), Arabic (ar-sa)
- Excludes: Groove, Lync, OneDrive, OneNote, Outlook for Windows, Publisher
- Shared Computer Licensing: Disabled (can be enabled via `SharedComputerLicensing` property)

## Dependencies & External Tools

**External Software (auto-downloaded if missing):**
- **WinGet** - Windows Package Manager (for automated software downloads)
- **HPIA** - HP Image Assistant (for modern HP driver scanning)
- **CMSL** - HP Command Line Shell (PowerShell module for HP driver management)
- **Office Deployment Tool** - Microsoft tool for Office installation

**Windows Components:**
- `System.Windows.Forms` - GUI framework
- `System.Drawing` - Graphics/colors
- WMI/CIM - Hardware detection (`Win32_BaseBoard`, `Win32_BIOS`, `Win32_Product`)
- Registry API - Application detection and configuration
- PowerShell scripting engine (2.0+)

**Network Resources:**
- HP online repositories (driver updates)
- WinGet repositories (software packages)
- `$Global:HPLegacyRepoRoot` - Offline driver repository for legacy systems (optional)

## Key Design Decisions

### Backward Compatibility
- **PowerShell 2.0 Support**: Uses WMI instead of CIM, avoids modern PowerShell features
- **Legacy HP Systems**: Offline repository support for Windows 7/8 and old hardware
- **Dual-Architecture**: x64/x86 handling for Visual C++ redistributables

### UI Design
- **Color-Coded Status**: Visual indicators (green/orange/red) for quick assessment
- **Auto-Sorting**: Drivers sorted by priority (Critical > Update > Not Installed > Installed)
- **Smart Selection**: Buttons for "Select Needs Update", "Select Not Installed"
- **Real-Time Logging**: Operation logs displayed in UI (not persisted to disk)

### Installation Strategy
- **Priority Order**: Local files → WinGet → Direct download
- **Version Tracking**: Downloaded files renamed with version suffix (e.g., `chrome_123.45.exe`)
- **Silent Installation**: All apps installed with `/silent` or equivalent flags
- **Dual-Arch Support**: Separate x64/x86 installations for VC++ redistributables

## Important File Locations

**Script Directory** (`$ScriptPath` / `$PSScriptRoot`):
- All local installers cached here
- `setup.exe` - Office Deployment Tool (auto-extracted)
- `OfficeSetup/` - Temporary folder for Office installation
- No logging to disk (logs only displayed in UI)

**HP Driver Locations:**
- `C:\Windows\Temp\HPIA\` - HP Image Assistant download location
- `$Global:HPLegacyRepoRoot` - Network share for offline repository
- Driver installers downloaded to `$ScriptPath`

## Common Customizations

### Adding New Applications

Edit `$Global:AppList` (lines 95-164):

```powershell
$Global:AppList = @(
    # ... existing apps ...
    @{
        Cat = "Utilities"
        Name = "7-Zip"
        Id = "7zip.7zip"
        Args = "/S"
        DownloadMode = "winget"
        FileStem = "7zip"
        Pat = "7-Zip"
    }
)
```

### Modifying System Configuration

Edit `Apply-SystemConfiguration()` function to change:
- Timezone (default: "West Bank Standard Time")
- Power plan (default: High Performance)
- Locale settings (default: Arabic ar-SA / US English)
- Desktop icon visibility
- Fast boot behavior

### Changing Office Configuration

Modify XML templates (lines 40-90):
- Change `Channel` attribute for different update channels
- Add/remove languages in `<Language ID="..."/>` tags
- Modify excluded apps in `<ExcludeApp ID="..."/>` tags
- Update product keys in `PIDKEY` attribute (Office 2024 only)
- Toggle shared computer licensing via `SharedComputerLicensing` property

## Limitations & Constraints

1. **Single File Architecture**: All code in one 2,660-line file (difficult to modularize)
2. **Global State**: Heavy use of `$Global:` variables (UI state, app list, configuration)
3. **No Logging Persistence**: Logs only displayed in UI, not saved to disk
4. **Hardcoded Values**: IP addresses, CrowdStrike CID, timezone settings
5. **Windows-Only**: Requires Windows OS and PowerShell (not cross-platform)
6. **Manual Testing**: No automated test framework (Pester, etc.)
7. **PowerShell 2.0 Constraints**: Cannot use modern PowerShell features (classes, `foreach-object -parallel`, etc.)

## Security Considerations

- **Admin Privileges Required**: Script auto-elevates via UAC (line 15)
- **CrowdStrike CID Embedded**: Organization ID in plaintext (line 36)
- **Network Share Access**: UNC path to driver repository (line 37)
- **Silent Installation**: All apps installed without user prompts (potential for unwanted software)
- **Office Product Keys**: Embedded in XML (line 43, PIDKEY attribute)
