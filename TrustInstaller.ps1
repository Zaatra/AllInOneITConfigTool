<#
    Trust Palestine IT Configuration Tool (v18.3)
    - NEW: Offline Repository support for Legacy HP Systems (Win7/8/Old Hardware)
    - NEW: PowerShell 2.0 / WMI Compatibility for older OS
    - FIX: HPCMSL install with proper PowerShellGet detection
    - FIX: Driver tab UI - larger info panel, proper scrolling
    - FEATURE: Detect installed driver versions from registry/WMI
    - FEATURE: Color coding - Green=Installed, Orange=Update, Red=Missing
    - FEATURE: Auto-sort by status (Critical > Update > Not Installed > Installed)
#>

# -------------------------
# 1) Admin Check
# -------------------------
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $p = New-Object System.Diagnostics.ProcessStartInfo
    $p.FileName  = "powershell.exe"
    $p.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    $p.Verb      = "runas"
    [System.Diagnostics.Process]::Start($p) | Out-Null
    exit
}

try {

# -------------------------
# 2) Config
# -------------------------
$ScriptPath = $PSScriptRoot
Set-Location $ScriptPath

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- GLOBAL CONFIG ---
$Global:CrowdStrikeCID = "753B2C62DBF944A59165207E3FE2FF2A-27"
$Global:HPLegacyRepoRoot = "\\192.168.168.6\Admin Tools\Drivers-Repo"

# --- Office XMLs ---
$Office2024XML = @"
<Configuration ID="917267f9-54e9-4c93-8a91-b6df3e3f2b1d">
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024" MigrateArch="TRUE">
    <Product ID="ProPlus2024Volume" PIDKEY="BQG8N-4Y7XV-92QPV-K26X6-QYM2Q">
      <Language ID="en-us"/>
      <Language ID="ar-sa"/>
      <ExcludeApp ID="Lync"/>
      <ExcludeApp ID="OneDrive"/>
      <ExcludeApp ID="Publisher"/>
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0"/>
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE"/>
  <Property Name="AUTOACTIVATE" Value="1"/>
  <Updates Enabled="TRUE"/>
  <Remove All="TRUE"/>
  <Display Level="None" AcceptEULA="TRUE"/>
</Configuration>
"@

$Office365XML = @"
<Configuration ID="eb27020c-89bd-4519-96d9-7c652bf678eb">
<Info Description=""/>
<Add OfficeClientEdition="64" Channel="Current" MigrateArch="TRUE">
<Product ID="O365ProPlusRetail">
<Language ID="en-us"/>
<Language ID="ar-sa"/>
<ExcludeApp ID="Groove"/>
<ExcludeApp ID="Lync"/>
<ExcludeApp ID="OneDrive"/>
<ExcludeApp ID="OneNote"/>
<ExcludeApp ID="OutlookForWindows"/>
<ExcludeApp ID="Publisher"/>
</Product>
</Add>
<Property Name="SharedComputerLicensing" Value="0"/>
<Property Name="FORCEAPPSHUTDOWN" Value="TRUE"/>
<Property Name="DeviceBasedLicensing" Value="0"/>
<Property Name="SCLCacheOverride" Value="0"/>
<Updates Enabled="TRUE"/>
<Remove All="TRUE"/>
<RemoveMSI/>
<AppSettings>
<Setup Name="Company" Value="Trust"/>
<User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas"/>
<User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas"/>
<User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas"/>
</AppSettings>
<Display Level="None" AcceptEULA="TRUE"/>
</Configuration>
"@

# -------------------------
# 2.1 MASTER APP LIST
# -------------------------
$Global:AppList = @(
    # VC++ (Dual-Arch)
    @{ Cat="VC++ Redistributables"; Name="VC++ 2005";  Id64="Microsoft.VCRedist.2005.x64";  Id86="Microsoft.VCRedist.2005.x86";  Args="/q"; DualArch=$true; VCKey="2005";
       FileStem64="vcpp2005_x64"; FileStem86="vcpp2005_x86"; DownloadMode="winget" },

    @{ Cat="VC++ Redistributables"; Name="VC++ 2008";  Id64="Microsoft.VCRedist.2008.x64";  Id86="Microsoft.VCRedist.2008.x86";  Args="/q"; DualArch=$true; VCKey="2008";
       FileStem64="vcpp2008_x64"; FileStem86="vcpp2008_x86"; DownloadMode="winget" },

    @{ Cat="VC++ Redistributables"; Name="VC++ 2010";  Id64="Microsoft.VCRedist.2010.x64";  Id86="Microsoft.VCRedist.2010.x86";  Args="/passive /norestart"; DualArch=$true; VCKey="2010";
       FileStem64="vcpp2010_x64"; FileStem86="vcpp2010_x86"; DownloadMode="winget" },

    @{ Cat="VC++ Redistributables"; Name="VC++ 2012";  Id64="Microsoft.VCRedist.2012.x64";  Id86="Microsoft.VCRedist.2012.x86";  Args="/passive /norestart"; DualArch=$true; VCKey="2012";
       FileStem64="vcpp2012_x64"; FileStem86="vcpp2012_x86"; DownloadMode="winget" },

    @{ Cat="VC++ Redistributables"; Name="VC++ 2013";  Id64="Microsoft.VCRedist.2013.x64";  Id86="Microsoft.VCRedist.2013.x86";  Args="/passive /norestart"; DualArch=$true; VCKey="2013";
       FileStem64="vcpp2013_x64"; FileStem86="vcpp2013_x86"; DownloadMode="winget" },

    @{ Cat="VC++ Redistributables"; Name="VC++ 2015+"; Id64="Microsoft.VCRedist.2015+.x64"; Id86="Microsoft.VCRedist.2015+.x86"; Args="/passive /norestart"; DualArch=$true; VCKey="2015+";
       FileStem64="vcpp2015plus_x64"; FileStem86="vcpp2015plus_x86"; DownloadMode="winget" },

    # Core
    @{ Cat="Core Applications"; Name="Java 8";     Id="Oracle.JavaRuntimeEnvironment"; Args="/s"; Pat="Java.*8 Update"; DualArch=$false;
       FileStem="java8"; DownloadMode="winget" },

    @{ Cat="Core Applications"; Name="Chrome";     Id="Google.Chrome"; Args=""; Pat="Google Chrome"; DualArch=$false;
       FileStem="chrome"; DownloadMode="winget" },

    @{ Cat="Core Applications"; Name="Firefox";    Id="Mozilla.Firefox"; Args=""; Pat="Mozilla Firefox"; DualArch=$false;
       FileStem="firefox"; DownloadMode="winget" },

    @{ Cat="Core Applications"; Name="WinRAR";     Id="RARLab.WinRAR"; Args=""; Pat="WinRAR"; DualArch=$false;
       FileStem="winrar"; DownloadMode="winget" },

    @{ Cat="Core Applications"; Name="TeamViewer"; Id="TeamViewer.TeamViewer.Host"; Args="/qn"; Pat="TeamViewer"; DualArch=$false;
       FileStem="teamviewer_host"; DownloadMode="winget" },

    @{ Cat="Core Applications"; Name="NAPS2";      Id="Cyanfish.NAPS2"; Args=""; Pat="NAPS2"; DualArch=$false;
       FileStem="naps2"; DownloadMode="winget" },

    @{ Cat="Core Applications"; Name="K-Lite Mega";Id="CodecGuide.K-LiteCodecPack.Mega"; Args=""; Pat="K-Lite Mega Codec"; DualArch=$false;
       FileStem="klite_mega"; DownloadMode="winget" },

    # Office (ODT)
    @{ Cat="Core Applications"; Name="Office 2024 LTSC"; Id="Microsoft.OfficeDeploymentTool"; Args=""; Pat="Microsoft Office.*2024"; DualArch=$false;
       DownloadMode="office" },

    @{ Cat="Core Applications"; Name="Office 365 Ent";   Id="Microsoft.OfficeDeploymentTool"; Args=""; Pat="Microsoft 365 Apps"; DualArch=$false;
       DownloadMode="office" },

    # Drivers
    @{ Cat="Driver & Support Tools"; Name="Intel DSA"; Id="Intel.IntelDriverAndSupportAssistant"; Args=""; Pat="Intel.*Driver.*Support"; DualArch=$false;
       FileStem="intel_dsa"; DownloadMode="winget" },

    # HP Support Assistant (MS Store - online only)
    @{ Cat="Driver & Support Tools"; Name="HP Support Asst"; Id="9WZDNCRDRBFB"; Source="msstore"; Args=""; Pat="HP Support (Assistant|Solutions)"; DualArch=$false;
       DownloadMode="onlineonly" },

    # Security
    @{ Cat="Security & VMS"; Name="FortiClient VPN"; Id="Fortinet.FortiClientVPN"; Args=""; Pat="FortiClient VPN"; DualArch=$false;
       FileStem="forticlient_vpn"; DownloadMode="winget" },

    @{ Cat="Security & VMS"; Name="iVMS-4200"; Id=""; Args="/silent /norestart"; Pat="iVMS-4200"; DualArch=$false;
       FileStem="ivms4200"; DownloadMode="direct" },

    # CrowdStrike (Local Only) - CID included in Args
    @{ Cat="Security & VMS"; Name="CrowdStrike Falcon Sensor"; Id=""; Args="/install /quiet /norestart CID=$Global:CrowdStrikeCID"; Pat="CrowdStrike Falcon Sensor|CrowdStrike Windows Sensor"; DualArch=$false;
       FileStem="crowdstrike_falcon_sensor"; DownloadMode="localonly";
       LocalAltNames=@("WindowsSensor.exe","FalconSensor.exe","CrowdStrikeWindowsSensor.exe","CSFalconSensor.exe","CSFalconInstaller.exe") }
)

# -------------------------
# 3) Helpers
# -------------------------
$FontHeader = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$FontItem   = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$FontSmall  = New-Object System.Drawing.Font("Segoe UI", 9,  [System.Drawing.FontStyle]::Regular)
$FontBadge  = New-Object System.Drawing.Font("Segoe UI", 8,  [System.Drawing.FontStyle]::Bold)

function Log-Msg {
    param([string]$Message)
    $txtLog.AppendText("$(Get-Date -Format 'HH:mm:ss') $Message`r`n")
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Log-Driver {
    param([string]$Message)
    if ($txtDriverLog) {
        $txtDriverLog.AppendText("$(Get-Date -Format 'HH:mm:ss') $Message`r`n")
        $txtDriverLog.ScrollToCaret()
    }
    [System.Windows.Forms.Application]::DoEvents()
}

# --- Compatibility Wrapper (Win7/PS2.0 Support) ---
function Get-CompatInstance {
    param(
        [Parameter(Mandatory=$true)] [string] $ClassName,
        [string] $Namespace
    )
    if (Get-Command Get-CimInstance -ErrorAction SilentlyContinue) {
        if ($Namespace) { return Get-CimInstance -ClassName $ClassName -Namespace $Namespace -ErrorAction SilentlyContinue }
        return Get-CimInstance -ClassName $ClassName -ErrorAction SilentlyContinue
    } else {
        # Fallback to WMI for older PowerShell/OS
        if ($Namespace) { return Get-WmiObject -Class $ClassName -Namespace $Namespace -ErrorAction SilentlyContinue }
        return Get-WmiObject -Class $ClassName -ErrorAction SilentlyContinue
    }
}

function Normalize-VersionString {
    param([string]$v)
    if ([string]::IsNullOrWhiteSpace($v)) { return $null }
    $t = $v.Trim()
    if ($t -match '^\d+\.\d+$')             { return ($t + ".0.0") }
    if ($t -match '^\d+\.\d+\.\d+$')       { return ($t + ".0") }
    if ($t -match '^\d+\.\d+\.\d+\.\d+$')  { return $t }
    $m = [regex]::Match($t, '(\d+\.\d+(?:\.\d+){0,2})')
    if ($m.Success) { return (Normalize-VersionString $m.Groups[1].Value) }
    return $t
}

function To-VersionObj {
    param([string]$v)
    $n = Normalize-VersionString $v
    if ([string]::IsNullOrWhiteSpace($n)) { return $null }
    try { return [version]$n } catch { return $null }
}

function Get-FileProductVersion {
    param([string]$Path)
    try {
        if (-not (Test-Path $Path)) { return $null }
        $vi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($Path)
        if ($vi -and $vi.ProductVersion) { return (Normalize-VersionString $vi.ProductVersion) }
        if ($vi -and $vi.FileVersion)    { return (Normalize-VersionString $vi.FileVersion) }
    } catch {}
    return $null
}

function Enable-DoubleBuffer {
    param($ctrl)
    try {
        $prop = $ctrl.GetType().GetProperty("DoubleBuffered", [Reflection.BindingFlags]"NonPublic,Instance")
        if ($prop) { $prop.SetValue($ctrl, $true, $null) }
    } catch {}
}

function Set-RowColor {
    param($UI, [System.Drawing.Color]$Color)
    $UI.Chk.ForeColor     = $Color
    $UI.LblFile.ForeColor = $Color
    $UI.LblInst.ForeColor = $Color
    $UI.LblLat.ForeColor  = $Color
}

function Safe-FilePart {
    param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return "unknown" }
    $t = $s.Trim()
    $t = $t -replace '\s+','_'
    $t = $t -replace '[^a-zA-Z0-9._-]',''
    if ([string]::IsNullOrWhiteSpace($t)) { return "unknown" }
    return $t
}

# ---- Registry Uninstall Enumeration (robust) ----
function Get-AllUninstallEntries {
    $results = @()
    $pairs = @(
        @{ View=[Microsoft.Win32.RegistryView]::Registry64; Path="SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" },
        @{ View=[Microsoft.Win32.RegistryView]::Registry64; Path="SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" },
        @{ View=[Microsoft.Win32.RegistryView]::Registry32; Path="SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" }
    )

    foreach ($p in $pairs) {
        try {
            $bk = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $p.View)
            $sk = $bk.OpenSubKey($p.Path)
            if ($sk) {
                foreach ($n in $sk.GetSubKeyNames()) {
                    try {
                        $i = $sk.OpenSubKey($n)
                        if (-not $i) { continue }
                        $dn = [string]$i.GetValue("DisplayName")
                        if ([string]::IsNullOrWhiteSpace($dn)) { $i.Close(); continue }
                        $dv = [string]$i.GetValue("DisplayVersion")
                        $results += [PSCustomObject]@{
                            DisplayName    = $dn
                            DisplayVersion = $dv
                            RegPath        = "$($p.Path)\$n"
                        }
                        $i.Close()
                    } catch {}
                }
                $sk.Close()
            }
            $bk.Close()
        } catch {}
    }
    return $results
}

function Get-RegVer {
    param([string]$Pat)
    $all = Get-AllUninstallEntries
    $hits = $all | Where-Object { $_.DisplayName -match $Pat -and $_.DisplayVersion }
    if (-not $hits) { return $null }

    $best = $null
    foreach ($h in $hits) {
        $nv = Normalize-VersionString $h.DisplayVersion
        $vo = To-VersionObj $nv
        if (-not $vo) { continue }
        if (-not $best -or $vo -gt $best.V) { $best = [PSCustomObject]@{ S=$nv; V=$vo } }
    }
    if ($best) { return $best.S }
    return (Normalize-VersionString ($hits | Select-Object -First 1).DisplayVersion)
}

# ---- VC++ Installed (per arch) ----
function Get-VCInstalledMap {
    param([string]$VCKey)

    $rx = $null
    switch ($VCKey) {
        "2005"  { $rx = '^Microsoft Visual C\+\+ 2005 Redistributable' }
        "2008"  { $rx = '^Microsoft Visual C\+\+ 2008 Redistributable' }
        "2010"  { $rx = '^Microsoft Visual C\+\+ 2010' }
        "2012"  { $rx = '^Microsoft Visual C\+\+ 2012 Redistributable' }
        "2013"  { $rx = '^Microsoft Visual C\+\+ 2013 Redistributable' }
        "2015+" { $rx = '^Microsoft Visual C\+\+.*((2015-2022)|(2015|2017|2019|2022)|v14).*(Redistributable)' }
        default { return @{ x86=$null; x64=$null } }
    }

    $all = Get-AllUninstallEntries
    $hits = $all | Where-Object { $_.DisplayName -match $rx -and $_.DisplayVersion }
    if (-not $hits) { return @{ x86=$null; x64=$null } }

    $best86 = $null
    $best64 = $null

    foreach ($h in $hits) {
        $dn = $h.DisplayName
        $dv = Normalize-VersionString $h.DisplayVersion
        $vo = To-VersionObj $dv
        if (-not $vo) { continue }

        $arch = $null
        if ($dn -match '\(x64\)|64-bit') { $arch = "x64" }
        elseif ($dn -match '\(x86\)|32-bit|x32') { $arch = "x86" }
        else {
            if ($VCKey -eq "2005") { $arch = "x86" }
        }

        if ($arch -eq "x86") {
            if (-not $best86 -or $vo -gt $best86.V) { $best86 = [PSCustomObject]@{ S=$dv; V=$vo } }
        } elseif ($arch -eq "x64") {
            if (-not $best64 -or $vo -gt $best64.V) { $best64 = [PSCustomObject]@{ S=$dv; V=$vo } }
        }
    }

    return @{
        x86 = if ($best86) { $best86.S } else { $null }
        x64 = if ($best64) { $best64.S } else { $null }
    }
}

function Get-VCCompositeText {
    param([string]$VCKey)
    $m = Get-VCInstalledMap $VCKey
    $str86 = if ($m.x86) { $m.x86 } else { "Missing" }
    $str64 = if ($m.x64) { $m.x64 } else { "Missing" }
    return "x86: $str86 | x64: $str64"
}

# ---- Local Installer Discovery (versioned naming) ----
function Get-BestLocalByPattern {
    param(
        [string]$Pattern,
        [string[]]$ExactNames
    )

    $files = @()
    try { $files += Get-ChildItem -Path $ScriptPath -File -Filter $Pattern -ErrorAction SilentlyContinue } catch {}

    if ($ExactNames) {
        foreach ($n in $ExactNames) {
            $p = Join-Path $ScriptPath $n
            if (Test-Path $p) { $files += Get-Item $p -ErrorAction SilentlyContinue }
        }
    }

    if (-not $files -or $files.Count -eq 0) { return $null }

    $best = $null
    foreach ($f in ($files | Sort-Object -Property Name -Unique)) {
        $name = $f.Name
        $verToken = $null

        $m = [regex]::Match($name, '_([0-9]+(?:\.[0-9]+){1,3})\.(exe|msi)$', 'IgnoreCase')
        if ($m.Success) { $verToken = Normalize-VersionString $m.Groups[1].Value }
        if (-not $verToken) { $verToken = Get-FileProductVersion $f.FullName }

        $vo = To-VersionObj $verToken
        if (-not $vo) {
            if (-not $best) { $best = [PSCustomObject]@{ File=$f; V=$null } }
            continue
        }

        if (-not $best -or (-not $best.V) -or $vo -gt $best.V) {
            $best = [PSCustomObject]@{ File=$f; V=$vo }
        }
    }

    return $best.File
}

function Get-LocalInstallerInfo {
    param($App)

    if ($App.DownloadMode -eq "onlineonly") {
        return [PSCustomObject]@{ Exists=$false; Text="N/A (Online Only)"; Path=$null; Path86=$null; Path64=$null }
    }

    if ($App.DownloadMode -eq "office") {
        $setupPath = Join-Path $ScriptPath "setup.exe"
        if (Test-Path $setupPath) {
            return [PSCustomObject]@{ Exists=$true; Text="FILES: YES (setup.exe)"; Path=(Get-Item $setupPath); Path86=$null; Path64=$null }
        } else {
            return [PSCustomObject]@{ Exists=$false; Text="FILES: NO"; Path=$null; Path86=$null; Path64=$null }
        }
    }

    if ($App.DualArch) {
        $p86 = Get-BestLocalByPattern -Pattern ("{0}_*.exe" -f $App.FileStem86) -ExactNames $null
        $p64 = Get-BestLocalByPattern -Pattern ("{0}_*.exe" -f $App.FileStem64) -ExactNames $null

        $hasAny = ($p86 -or $p64)
        if ($hasAny) {
            $bits = @()
            if ($p86) { $bits += ("x86: " + $p86.Name) } else { $bits += "x86: Missing" }
            if ([Environment]::Is64BitOperatingSystem) {
                if ($p64) { $bits += ("x64: " + $p64.Name) } else { $bits += "x64: Missing" }
            }
            return [PSCustomObject]@{ Exists=$true; Text=("FILES: YES (" + ($bits -join ", ") + ")"); Path=$null; Path86=$p86; Path64=$p64 }
        } else {
            return [PSCustomObject]@{ Exists=$false; Text="FILES: NO"; Path=$null; Path86=$null; Path64=$null }
        }
    }

    # Single-arch
    $pattern = ("{0}_*.exe" -f $App.FileStem)
    $exact = $null
    if ($App.DownloadMode -eq "localonly" -and $App.LocalAltNames) { $exact = $App.LocalAltNames }

    $best = Get-BestLocalByPattern -Pattern $pattern -ExactNames $exact
    if ($best) {
        return [PSCustomObject]@{ Exists=$true; Text=("FILES: YES (" + $best.Name + ")"); Path=$best; Path86=$null; Path64=$null }
    } else {
        return [PSCustomObject]@{ Exists=$false; Text="FILES: NO"; Path=$null; Path86=$null; Path64=$null }
    }
}

# --- Winget Locator ---
function Get-WingetPath {
    if (Get-Command winget -ErrorAction SilentlyContinue) { return "winget" }

    $paths = @(
        "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe",
        "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe"
    )
    foreach ($p in $paths) {
        $r = Resolve-Path $p -ErrorAction SilentlyContinue
        if ($r) { return ($r.Path | Select-Object -First 1) }
    }
    return $null
}

function Get-WingetVer {
    param([string]$Id, [string]$Source="winget")
    $exe = Get-WingetPath
    if (-not $exe) { return "Winget Missing" }

    try {
        $out = & $exe show --id $Id --source $Source --accept-source-agreements --locale en-US 2>&1 | Out-String -Stream
        foreach ($line in $out) {
            if ($line -match 'Version\s*:\s*([0-9a-zA-Z._-]+)') {
                return (Normalize-VersionString $Matches[1].Trim())
            }
        }
    } catch { return "Error" }

    return "N/A"
}

function Get-OfficeVerWeb {
    param($Url)
    try {
        $c = (Invoke-WebRequest $Url -UseBasicParsing -TimeoutSec 10).Content
        if ($c -match 'Version\s+(\d+)\s+\(Build\s+([\d.]+)\)') {
            return "$($Matches[1]) ($($Matches[2]))"
        }
    } catch {}
    return "N/A"
}


function Connect-Repo-WithAuth {
    # Check if we can already see the folder
    if (Test-Path $Global:HPLegacyRepoRoot) { 
        return $true 
    }

    # If not, show the Pop-up
    try {
        # This opens the standard Windows Security Dialog
        $cred = Get-Credential -UserName "TRUST.COM\" -Message "Authentication Required for Drivers Repo`nPath: $Global:HPLegacyRepoRoot"
        
        if (-not $cred) { return $false } # User clicked Cancel

        $userName = $cred.UserName
        $password = $cred.GetNetworkCredential().Password

        # Determine the share root (net use prefers the root share, e.g., \\192.168.168.6\Admin Tools)
        # We split the path to find the share root
        $pathParts = $Global:HPLegacyRepoRoot.Split('\')
        $shareRoot = "\\" + $pathParts[2] + "\" + $pathParts[3] 

        # Use 'net use' to authenticate
        # We use Start-Process to handle special characters in passwords safely
        $proc = Start-Process "net.exe" -ArgumentList "use", "`"$shareRoot`"", "`"$password`"", "/user:`"$userName`"" -WindowStyle Hidden -Wait -PassThru

        if ($proc.ExitCode -eq 0) {
            return $true
        } else {
            [System.Windows.Forms.MessageBox]::Show("Login failed. Please check your password.", "Authentication Error", "OK", "Error")
            return $false
        }
    } catch {
        return $false
    }
}

function Get-HPSALatest {
    try {
        $c = (Invoke-WebRequest "https://hpsa-redirectors.hpcloud.hp.com/common/hpsaredirector.js" -UseBasicParsing).Content
        if ($c -match '(?m)^\s*return\s+"([0-9.]+)"') { return (Normalize-VersionString $Matches[1]) }
    } catch {}
    return "N/A"
}

function Get-iVMSLatest {
    $src = "https://www.hikvision.com/us-en/support/download/software/"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $headers = @{
        "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        "Accept-Language" = "en-US,en;q=0.9"
    }

    try {
        $html = (Invoke-WebRequest -Uri $src -Headers $headers -UseBasicParsing -TimeoutSec 20).Content
        $verStrings = @()

        foreach ($m in [regex]::Matches($html, 'iVMS-4200V(\d+\.\d+\.\d+\.\d+)_E(?:\.exe)?', 'IgnoreCase')) {
            $verStrings += (Normalize-VersionString $m.Groups[1].Value)
        }
        foreach ($m in [regex]::Matches($html, 'iVMS-4200\s+V(\d+\.\d+\.\d+\.\d+)', 'IgnoreCase')) {
            $verStrings += (Normalize-VersionString $m.Groups[1].Value)
        }

        $verStrings = $verStrings | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' } | Select-Object -Unique
        if (-not $verStrings) { return $null }

        $best = $null
        foreach ($vs in $verStrings) {
            $vo = To-VersionObj $vs
            if (-not $vo) { continue }
            if (-not $best -or $vo -gt $best.V) { $best = [PSCustomObject]@{ S=$vs; V=$vo } }
        }
        if (-not $best) { return $null }

        $v = $best.S
        $dash = "v" + ($v -replace '\.','-')
        $dl = "https://www.hikvision.com/content/dam/hikvision/en/support/download/vms/ivms4200-series/software-download/$dash/iVMS-4200V$v" + "_E.exe"

        return [PSCustomObject]@{ Version=$v; Url=$dl; SourcePage=$src }
    } catch {
        return $null
    }
}

# -------------------------
# 3.1) HP Driver Helpers (HPIA + CMSL + LEGACY)
# -------------------------
$Global:HPDriverList = @()
$Global:HPSystemInfo = $null
$Global:InstalledDriversCache = $null

function Get-HPSystemInfo {
    $info = [PSCustomObject]@{
        IsHP = $false
        Manufacturer = ""
        Model = ""
        ProductCode = ""
        SerialNumber = ""
        SKU = ""
        Generation = $null
        SupportsHPIA = $false
        SupportsCMSL = $false
        OSVersion = ""
        OSBuild = ""
    }

    try {
        $cs = Get-CompatInstance Win32_ComputerSystem
        $bios = Get-CompatInstance Win32_BIOS
        $bb = Get-CompatInstance Win32_BaseBoard
        $os = Get-CompatInstance Win32_OperatingSystem

        $info.Manufacturer = $cs.Manufacturer
        $info.Model = $cs.Model
        $info.SerialNumber = $bios.SerialNumber
        $info.OSVersion = $os.Caption
        $info.OSBuild = $os.BuildNumber

        # Check if HP
        if ($cs.Manufacturer -match 'HP|Hewlett|Packard') {
            $info.IsHP = $true

            # Get Product Code (4-char platform ID) from baseboard
            $info.ProductCode = $bb.Product
            
            # Get SKU from registry (more reliable for HP Support lookups)
            $regPath = "HKLM:\HARDWARE\DESCRIPTION\System\BIOS"
            $biosReg = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
            if ($biosReg) {
                if (-not $info.ProductCode -or $info.ProductCode.Length -lt 4) {
                    $info.ProductCode = $biosReg.SystemSKU
                }
                $info.SKU = $biosReg.SystemSKU
            }
            
            # Also try WMI for SKU if not found
            if (-not $info.SKU) {
                try {
                    $csProduct = Get-CompatInstance Win32_ComputerSystemProduct
                    if ($csProduct.SKUNumber) { $info.SKU = $csProduct.SKUNumber }
                } catch {}
            }

            # Detect generation from model name
            if ($info.Model -match 'G(\d+)') {
                $info.Generation = [int]$Matches[1]
            }

            # HPIA supports G3 (2016) and newer
            $info.SupportsHPIA = ($info.Generation -ge 3) -or ($info.Model -match 'Z[0-9]+ G|ZBook.*G[3-9]|Elite.*G[3-9]|Pro.*G[3-9]')

            # CMSL works with most HP systems, but not very old ones
            # If Win7/Win8, likely CMSL won't work well
            if ($info.OSVersion -match "Windows 7|Windows 8") {
                $info.SupportsCMSL = $false
            } else {
                $info.SupportsCMSL = $true
            }

            # Check for very old systems explicitly
            if ($info.Model -match 'Compaq|Pro3?500|dc\d{4}|8[0-3]00') {
                $info.SupportsHPIA = $false
                $info.SupportsCMSL = $false # Prefer repo for these
            }
        }
    } catch {}

    return $info
}

# ---- Get Installed Drivers/Software for comparison ----
function Get-InstalledDriversAndSoftware {
    $installed = @{}

    # Get from Registry (Uninstall entries)
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($path in $regPaths) {
        try {
            Get-ItemProperty $path -ErrorAction SilentlyContinue | ForEach-Object {
                if ($_.DisplayName -and $_.DisplayVersion) {
                    $key = $_.DisplayName.ToLower().Trim()
                    if (-not $installed.ContainsKey($key)) {
                        $installed[$key] = @{
                            Name = $_.DisplayName
                            Version = $_.DisplayVersion
                            Publisher = $_.Publisher
                        }
                    }
                }
            }
        } catch {}
    }

    # Get from Win32_PnPSignedDriver (actual drivers)
    try {
        Get-CompatInstance Win32_PnPSignedDriver | ForEach-Object {
            if ($_.DeviceName -and $_.DriverVersion) {
                $key = $_.DeviceName.ToLower().Trim()
                if (-not $installed.ContainsKey($key)) {
                    $installed[$key] = @{
                        Name = $_.DeviceName
                        Version = $_.DriverVersion
                        Publisher = $_.Manufacturer
                    }
                }
            }
        }
    } catch {}

    # Get BIOS version
    try {
        $bios = Get-CompatInstance Win32_BIOS
        if ($bios) {
            $installed["system bios"] = @{
                Name = "System BIOS"
                Version = $bios.SMBIOSBIOSVersion
                Publisher = $bios.Manufacturer
            }
        }
    } catch {}

    return $installed
}

function Find-InstalledVersion {
    param([string]$DriverName, [string]$Category)

    if (-not $Global:InstalledDriversCache) {
        $Global:InstalledDriversCache = Get-InstalledDriversAndSoftware
    }

    $cache = $Global:InstalledDriversCache
    $driverLower = $DriverName.ToLower()

    # Extract key terms from driver name
    $searchTerms = @()

    # Common patterns to extract
    if ($driverLower -match 'intel') { $searchTerms += 'intel' }
    if ($driverLower -match 'realtek') { $searchTerms += 'realtek' }
    if ($driverLower -match 'nvidia') { $searchTerms += 'nvidia' }
    if ($driverLower -match 'amd') { $searchTerms += 'amd' }
    if ($driverLower -match 'bluetooth') { $searchTerms += 'bluetooth' }
    if ($driverLower -match 'wireless|wlan|wifi') { $searchTerms += 'wireless', 'wlan', 'wifi' }
    if ($driverLower -match 'graphics|video|display') { $searchTerms += 'graphics', 'video', 'display' }
    if ($driverLower -match 'audio|sound') { $searchTerms += 'audio', 'sound' }
    if ($driverLower -match 'ethernet|nic|network') { $searchTerms += 'ethernet', 'network' }
    if ($driverLower -match 'chipset') { $searchTerms += 'chipset' }
    if ($driverLower -match 'storage|raid|rst|rapid') { $searchTerms += 'storage', 'rapid', 'rst' }
    if ($driverLower -match 'bios') { $searchTerms += 'bios' }
    if ($driverLower -match 'firmware') { $searchTerms += 'firmware' }
    if ($driverLower -match 'management engine|me driver') { $searchTerms += 'management engine' }
    if ($driverLower -match 'thunderbolt') { $searchTerms += 'thunderbolt' }
    if ($driverLower -match 'serial io|serialio') { $searchTerms += 'serial' }
    if ($driverLower -match 'arc|a380|a770') { $searchTerms += 'arc' }
    if ($driverLower -match 'usb 3') { $searchTerms += 'usb 3' }

    # Search installed items
    $bestMatch = $null
    $bestScore = 0

    foreach ($item in $cache.GetEnumerator()) {
        $itemName = $item.Key
        $score = 0

        foreach ($term in $searchTerms) {
            if ($itemName -match [regex]::Escape($term)) {
                $score++
            }
        }

        # Category-specific matching
        if ($Category -match 'Graphics' -and $itemName -match 'graphics|display|video') { $score += 2 }
        if ($Category -match 'Audio' -and $itemName -match 'audio|sound|realtek') { $score += 2 }
        if ($Category -match 'Network' -and $itemName -match 'network|ethernet|wireless|wifi|bluetooth') { $score += 2 }
        if ($Category -match 'Chipset' -and $itemName -match 'chipset|serial|management|usb') { $score += 2 }
        if ($Category -match 'Storage' -and $itemName -match 'storage|rapid|rst|raid|optane') { $score += 2 }
        if ($Category -match 'BIOS|Firmware' -and $itemName -match 'bios|firmware') { $score += 2 }

        if ($score -gt $bestScore) {
            $bestScore = $score
            $bestMatch = $item.Value
        }
    }

    if ($bestMatch -and $bestScore -ge 2) {
        return $bestMatch.Version
    }

    return $null
}

function Get-DriverInstallStatus {
    param(
        [string]$DriverName,
        [string]$Category,
        [string]$AvailableVersion
    )

    $installedVer = Find-InstalledVersion -DriverName $DriverName -Category $Category

    if (-not $installedVer) {
        return @{ Status = "Not Installed"; InstalledVersion = "Not Found"; NeedsAction = $true }
    }

    # Compare versions
    $instObj = To-VersionObj $installedVer
    $availObj = To-VersionObj $AvailableVersion

    if ($instObj -and $availObj) {
        if ($availObj -gt $instObj) {
            return @{ Status = "Update Available"; InstalledVersion = $installedVer; NeedsAction = $true }
        } else {
            return @{ Status = "Up to Date"; InstalledVersion = $installedVer; NeedsAction = $false }
        }
    }

    # Can't compare, assume needs checking
    return @{ Status = "Installed"; InstalledVersion = $installedVer; NeedsAction = $false }
}

function Test-HPCMSLInstalled {
    try {
        $mod = Get-Module -ListAvailable -Name "HPCMSL" -ErrorAction SilentlyContinue
        return ($null -ne $mod)
    } catch { return $false }
}

function Install-HPCMSL {
    Log-Driver "Installing HP Client Management Script Library (HPCMSL)..."
    try {
        # Ensure NuGet provider
        $nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
        if (-not $nuget -or $nuget.Version -lt [version]"2.8.5.201") {
            Log-Driver "Installing NuGet provider..."
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers | Out-Null
        }

        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue

        # Check PowerShellGet version to determine install method
        $psGet = Get-Module PowerShellGet -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1

        if ($psGet -and $psGet.Version -ge [version]"2.0.0") {
            # PowerShellGet 2.x - use -AcceptLicense
            Log-Driver "Using PowerShellGet $($psGet.Version) with -AcceptLicense..."
            Install-Module -Name HPCMSL -Force -Scope AllUsers -AllowClobber -SkipPublisherCheck -AcceptLicense -ErrorAction Stop
        } else {
            # PowerShellGet 1.x - use older HPCMSL version
            Log-Driver "PowerShellGet is $($psGet.Version). Installing older HPCMSL 1.6.10..."
            Install-Module -Name HPCMSL -RequiredVersion 1.6.10 -Force -Scope AllUsers -AllowClobber -SkipPublisherCheck -ErrorAction Stop
        }

        Import-Module HPCMSL -Force -ErrorAction Stop
        Log-Driver "HPCMSL installed successfully."
        return $true
    } catch {
        Log-Driver "Failed to install HPCMSL: $_"
        return $false
    }
}

function Get-HPIAPath {
    $paths = @(
        "$env:ProgramFiles\HP\HPIA\HPImageAssistant.exe",
        "$env:ProgramFiles(x86)\HP\HPIA\HPImageAssistant.exe",
        (Join-Path $ScriptPath "HPIA\HPImageAssistant.exe"),
        (Join-Path $ScriptPath "HPImageAssistant.exe")
    )

    foreach ($p in $paths) {
        if (Test-Path $p) { return $p }
    }
    return $null
}

function Download-HPIA {
    Log-Driver "Downloading HP Image Assistant..."
    $hpiaDir = Join-Path $ScriptPath "HPIA"
    New-Item -ItemType Directory -Force -Path $hpiaDir | Out-Null

    try {
        # Try CMSL method first
        if (Test-HPCMSLInstalled -or (Install-HPCMSL)) {
            Import-Module HPCMSL -Force -ErrorAction SilentlyContinue
            if (Get-Command Get-HPImageAssistant -ErrorAction SilentlyContinue) {
                Get-HPImageAssistant -DownloadPath $hpiaDir -Extract -ErrorAction Stop
                $exe = Get-ChildItem $hpiaDir -Recurse -Filter "HPImageAssistant.exe" | Select-Object -First 1
                if ($exe) {
                    Log-Driver "HPIA downloaded via CMSL: $($exe.FullName)"
                    return $exe.FullName
                }
            }
        }

        # Fallback: direct download
        $hpiaUrl = "https://hpia.hpcloud.hp.com/downloads/hpia/hp-hpia-5.2.1.exe"
        $dlPath = Join-Path $hpiaDir "hp-hpia-setup.exe"
        Invoke-WebRequest -Uri $hpiaUrl -OutFile $dlPath -UseBasicParsing
        Start-Process $dlPath -ArgumentList "/s /e /f `"$hpiaDir`"" -Wait -WindowStyle Hidden

        $exe = Get-ChildItem $hpiaDir -Recurse -Filter "HPImageAssistant.exe" | Select-Object -First 1
        if ($exe) {
            Log-Driver "HPIA extracted: $($exe.FullName)"
            return $exe.FullName
        }
    } catch {
        Log-Driver "HPIA download failed: $_"
    }

    return $null
}

# --- Legacy Repository Functions ---
function Get-LegacyRepoPackages {
    param(
        [string] $Platform, 
        [string] $Model
    )

    $potentialPaths = @()
    
    # 1. Try Platform ID (e.g. "3396")
    if ($Platform) { $potentialPaths += Join-Path $Global:HPLegacyRepoRoot $Platform }

    if ($Model) { 
        # 2. Try Exact Model Name (e.g. "HP Compaq Elite 8300 CMT")
        $potentialPaths += Join-Path $Global:HPLegacyRepoRoot $Model 

        # 3. Try Model without "HP" prefix (e.g. "Compaq Elite 8300 CMT")
        # This fixes your specific issue!
        $cleanModel = $Model -replace '^(HP\s+|Hewlett-Packard\s+)', ''
        if ($cleanModel -ne $Model) {
            $potentialPaths += Join-Path $Global:HPLegacyRepoRoot $cleanModel
        }
    }

    foreach ($p in $potentialPaths) {
        $manifestPath = Join-Path $p "manifest.json"
        
        if (Test-Path $manifestPath) {
            try {
                Log-Driver "Loading legacy repository from: $p"
                $json = Get-Content $manifestPath -Raw | ConvertFrom-Json
                
                # Attach the base path to each item for easier processing
                foreach ($item in $json) {
                    Add-Member -InputObject $item -MemberType NoteProperty -Name "_RepoPath" -Value $p -Force
                }
                return $json
            } catch {
                Log-Driver "Failed to parse manifest at $p : $_"
            }
        }
    }
    
    # Debugging help if it still fails
    Log-Driver "Debug: Could not find manifest.json. Checked paths:"
    foreach ($p in $potentialPaths) { Log-Driver " [x] $p" }
    
    return @()
}

function Scan-HPDrivers-HPIA {
    param([string]$HPIAPath)

    Log-Driver "Running HPIA scan..."
    $reportDir = Join-Path $ScriptPath "HPIAReport"
    if (Test-Path $reportDir) { Remove-Item $reportDir -Recurse -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Force -Path $reportDir | Out-Null

    $drivers = @()

    try {
        $args = "/Operation:Analyze /Category:All /Selection:All /Action:List /Silent /ReportFolder:`"$reportDir`""
        $proc = Start-Process $HPIAPath -ArgumentList $args -Wait -PassThru -WindowStyle Hidden

        $jsonFile = Get-ChildItem $reportDir -Filter "*.json" -Recurse | Select-Object -First 1
        if ($jsonFile) {
            $report = Get-Content $jsonFile.FullName -Raw | ConvertFrom-Json

            if ($report.HPIA.Recommendations) {
                foreach ($rec in $report.HPIA.Recommendations) {
                    $drivers += [PSCustomObject]@{
                        Source = "HPIA"
                        Name = $rec.Name
                        Category = $rec.Category
                        Version = $rec.Version
                        InstalledVersion = $rec.CurrentVersion
                        SoftPaqId = $rec.SoftPaqId
                        Url = $rec.ReleaseNotesUrl
                        Size = $rec.Size
                        Status = if ($rec.RecommendationValue -eq "Critical") { "Critical" }
                                 elseif ($rec.RecommendationValue -eq "Recommended") { "Recommended" }
                                 else { "Optional" }
                        Selected = ($rec.RecommendationValue -in @("Critical","Recommended"))
                        FilePath = $null
                    }
                }
            }
        }
        Log-Driver "HPIA found $($drivers.Count) updates."
    } catch {
        Log-Driver "HPIA scan error: $_"
    }

    return $drivers
}

function Scan-HPDrivers-CMSL {
    param([string]$Platform)
    
    Log-Driver "Running CMSL scan..."
    $drivers = @()
    
    try {
        if (-not (Test-HPCMSLInstalled)) {
            if (-not (Install-HPCMSL)) { return $drivers }
        }
        Import-Module HPCMSL -Force -ErrorAction Stop

        # Try multiple OS builds to find catalog
        $osVersions = @(
            @{Os="Win11"; Ver="24H2"}, @{Os="Win11"; Ver="23H2"}, @{Os="Win11"; Ver="22H2"},
            @{Os="Win10"; Ver="22H2"}, @{Os="Win10"; Ver="21H2"}, @{Os="Win10"; Ver="2004"}
        )

        $softpaqs = $null
        foreach ($ov in $osVersions) {
            try {
                $softpaqs = Get-SoftpaqList -Platform $Platform -Os $ov.Os -OsVer $ov.Ver -ErrorAction Stop
                if ($softpaqs) { Log-Driver "Found catalog for $($ov.Os) $($ov.Ver)"; break }
            } catch {}
        }

        if ($softpaqs) {
            foreach ($sp in $softpaqs) {
                if ($sp.Category -match "Driver|BIOS|Firmware|Utility") {
                    $status = Get-DriverInstallStatus -DriverName $sp.Name -Category $sp.Category -AvailableVersion $sp.Version
                    $drivers += [PSCustomObject]@{
                        Source = "CMSL"
                        Name = $sp.Name
                        Category = $sp.Category
                        Version = $sp.Version
                        InstalledVersion = $status.InstalledVersion
                        SoftPaqId = $sp.Id
                        Url = $sp.Url
                        Size = $sp.Size
                        Status = $status.Status
                        Selected = $false
                        FilePath = $null
                    }
                }
            }
        }
    } catch { Log-Driver "CMSL error: $_" }
    
    return $drivers
}

function Download-HPSoftPaq {
    param(
        [string]$SoftPaqId,
        [string]$Url,
        [string]$DestDir
    )

    $dlDir = if ($DestDir) { $DestDir } else { Join-Path $ScriptPath "HPDrivers" }
    New-Item -ItemType Directory -Force -Path $dlDir | Out-Null

    try {
        # Check if URL is actually a local/network file path (Legacy Repo)
        if ($Url -match '^\\\\|^[a-z]:') {
            if (Test-Path $Url) {
                $fileName = [System.IO.Path]::GetFileName($Url)
                $destPath = Join-Path $dlDir $fileName
                Copy-Item -Path $Url -Destination $destPath -Force
                Log-Driver "Copied from repo: $fileName"
                return $destPath
            } else {
                Log-Driver "Repo file missing: $Url"
                return $null
            }
        }

        # Otherwise use CMSL or Web Download
        if (Test-HPCMSLInstalled) {
            Import-Module HPCMSL -Force -ErrorAction SilentlyContinue
            if (Get-Command Get-Softpaq -ErrorAction SilentlyContinue) {
                $spPath = Join-Path $dlDir "$SoftPaqId.exe"
                Get-Softpaq -Number $SoftPaqId -SaveAs $spPath -Overwrite -ErrorAction Stop
                if (Test-Path $spPath) { return $spPath }
            }
        }

        if ($Url -and $Url.StartsWith("http")) {
            $fileName = [System.IO.Path]::GetFileName($Url)
            if (-not $fileName -or $fileName -notmatch '\.exe$') { $fileName = "$SoftPaqId.exe" }
            $dlPath = Join-Path $dlDir $fileName
            Invoke-WebRequest -Uri $Url -OutFile $dlPath -UseBasicParsing
            if (Test-Path $dlPath) { return $dlPath }
        }
    } catch {
        Log-Driver "Download failed for $SoftPaqId : $_"
    }

    return $null
}

function Install-HPSoftPaq {
    param([string]$Path)

    if (-not (Test-Path $Path)) { return $false }

    try {
        Log-Driver "Installing: $([System.IO.Path]::GetFileName($Path))"
        $proc = Start-Process $Path -ArgumentList "/s" -Wait -PassThru -WindowStyle Hidden

        if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
            Log-Driver "Installation completed (Exit: $($proc.ExitCode))"
            return $true
        } else {
            Log-Driver "Installation may have failed (Exit: $($proc.ExitCode))"
            return $false
        }
    } catch {
        Log-Driver "Install error: $_"
        return $false
    }
}

# -------------------------
# 4) GUI Construction
# -------------------------
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Trust Palestine - IT Tool v18.3 (Legacy Support)"
$Form.Size = New-Object System.Drawing.Size(1350, 950)
$Form.StartPosition = "CenterScreen"
$Form.FormBorderStyle = "Sizable"
$Form.BackColor = [System.Drawing.Color]::WhiteSmoke

Enable-DoubleBuffer $Form

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "TRUST PALESTINE - IT CONFIGURATION"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$lblTitle.Size = New-Object System.Drawing.Size(1200, 45)
$lblTitle.Location = New-Object System.Drawing.Point(20, 10)
$lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(0, 100, 0)
$lblTitle.Anchor = "Top, Left, Right"
$Form.Controls.Add($lblTitle)

$TabControl = New-Object System.Windows.Forms.TabControl
$TabControl.Location = New-Object System.Drawing.Point(20, 65)
$TabControl.Size = New-Object System.Drawing.Size(1290, 830)
$TabControl.Anchor = "Top, Bottom, Left, Right"
$Form.Controls.Add($TabControl)

# --- Tab 1: Installation ---
$TabInstall = New-Object System.Windows.Forms.TabPage
$TabInstall.Text = "Installation & Downloads"
$TabInstall.BackColor = [System.Drawing.Color]::White
$TabControl.Controls.Add($TabInstall)

# --- Tab 2: Drivers (PROPERLY DOCKED) ---
$TabDrivers = New-Object System.Windows.Forms.TabPage
$TabDrivers.Text = "HP Drivers (HPIA + CMSL)"
$TabDrivers.BackColor = [System.Drawing.Color]::White
$TabDrivers.Padding = New-Object System.Windows.Forms.Padding(0)
$TabControl.Controls.Add($TabDrivers)

# --- Tab 3: System Config ---
$TabSystem = New-Object System.Windows.Forms.TabPage
$TabSystem.Text = "System Config & Logs"
$TabSystem.BackColor = [System.Drawing.Color]::WhiteSmoke
$TabControl.Controls.Add($TabSystem)

# ============ TAB 1: Installation ============
$pnlButtons = New-Object System.Windows.Forms.Panel
$pnlButtons.Height = 70
$pnlButtons.Dock = "Bottom"
$pnlButtons.BackColor = [System.Drawing.Color]::WhiteSmoke
$TabInstall.Controls.Add($pnlButtons)

$pnlApps = New-Object System.Windows.Forms.Panel
$pnlApps.Dock = "Fill"
$pnlApps.AutoScroll = $true
$pnlApps.BackColor = [System.Drawing.Color]::White
$TabInstall.Controls.Add($pnlApps)

Enable-DoubleBuffer $pnlApps

[int]$script:hY = 10
[int]$script:col1 = 10
[int]$script:col2 = 400
[int]$script:col3 = 600
[int]$script:col4 = 900

function New-HeaderLabel {
    param([string]$Txt, [int]$X)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Txt
    $l.Font = $FontHeader
    $l.Location = New-Object System.Drawing.Point($X, $script:hY)
    $l.AutoSize = $true
    $l.BackColor = $pnlApps.BackColor
    return $l
}

$pnlApps.Controls.Add((New-HeaderLabel "Application Selection" $script:col1))
$pnlApps.Controls.Add((New-HeaderLabel "Local File?"            $script:col2))
$pnlApps.Controls.Add((New-HeaderLabel "Installed Version"      $script:col3))
$pnlApps.Controls.Add((New-HeaderLabel "Latest Available"       $script:col4))
$script:hY += 40

$Global:UIControls = @{}
$LastCat = ""

foreach ($App in $Global:AppList) {
    if ($App.Cat -ne $LastCat) {
        $script:hY += 10
        $grpChk = New-Object System.Windows.Forms.CheckBox
        $grpChk.Text = $App.Cat
        $grpChk.Font = $FontHeader
        $grpChk.ForeColor = [System.Drawing.Color]::FromArgb(0, 100, 0)
        $grpChk.Location = New-Object System.Drawing.Point($script:col1, $script:hY)
        $grpChk.Size = New-Object System.Drawing.Size(300, 30)
        $grpChk.Checked = $true
        $grpChk.Tag = $App.Cat
        $grpChk.BackColor = $pnlApps.BackColor

        $grpChk.Add_CheckedChanged({
            $c = $this.Checked
            $catName = $this.Tag
            foreach ($a in $Global:AppList) {
                if ($a.Cat -eq $catName) { $Global:UIControls[$a.Name].Chk.Checked = $c }
            }
        })

        $pnlApps.Controls.Add($grpChk)
        $script:hY += 35
        $LastCat = $App.Cat
    }

    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = $App.Name
    $chk.Font = $FontItem
    $chk.Location = New-Object System.Drawing.Point(($script:col1 + 20), $script:hY)
    $chk.Size = New-Object System.Drawing.Size(350, 25)
    $chk.Tag = $App
    $chk.Checked = $true
    $chk.BackColor = $pnlApps.BackColor
    $pnlApps.Controls.Add($chk)

    $lblFile = New-Object System.Windows.Forms.Label
    $lblFile.Location = New-Object System.Drawing.Point($script:col2, $script:hY)
    $lblFile.Size = New-Object System.Drawing.Size(190, 25)
    $lblFile.Font = $FontItem
    $lblFile.Text = "Checking..."
    $lblFile.BackColor = $pnlApps.BackColor
    $pnlApps.Controls.Add($lblFile)

    $lblInst = New-Object System.Windows.Forms.Label
    $lblInst.Location = New-Object System.Drawing.Point($script:col3, $script:hY)
    $lblInst.Size = New-Object System.Drawing.Size(280, 25)
    $lblInst.Font = $FontItem
    $lblInst.Text = "Pending..."
    $lblInst.BackColor = $pnlApps.BackColor
    $pnlApps.Controls.Add($lblInst)

    $lblLat = New-Object System.Windows.Forms.Label
    $lblLat.Location = New-Object System.Drawing.Point($script:col4, $script:hY)
    $lblLat.Size = New-Object System.Drawing.Size(320, 25)
    $lblLat.Font = $FontItem
    $lblLat.Text = ""
    $lblLat.BackColor = $pnlApps.BackColor
    $pnlApps.Controls.Add($lblLat)

    $Global:UIControls[$App.Name] = @{
        Chk=$chk; LblFile=$lblFile; LblInst=$lblInst; LblLat=$lblLat;
        LocalInfo=$null
    }

    $script:hY += 30
}

$btnFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

$btnInstall = New-Object System.Windows.Forms.Button
$btnInstall.Text="INSTALL SELECTED"
$btnInstall.Font=$btnFont
$btnInstall.Location=New-Object System.Drawing.Point(10, 15)
$btnInstall.Size=New-Object System.Drawing.Size(200, 45)
$btnInstall.BackColor=[System.Drawing.Color]::ForestGreen
$btnInstall.ForeColor="White"
$pnlButtons.Controls.Add($btnInstall)

$btnDownload = New-Object System.Windows.Forms.Button
$btnDownload.Text="DOWNLOAD INSTALLERS"
$btnDownload.Font=$btnFont
$btnDownload.Location=New-Object System.Drawing.Point(220, 15)
$btnDownload.Size=New-Object System.Drawing.Size(200, 45)
$btnDownload.BackColor=[System.Drawing.Color]::RoyalBlue
$btnDownload.ForeColor="White"
$pnlButtons.Controls.Add($btnDownload)

$btnCheckUpdate = New-Object System.Windows.Forms.Button
$btnCheckUpdate.Text="CHECK FOR UPDATES"
$btnCheckUpdate.Font=$btnFont
$btnCheckUpdate.Location=New-Object System.Drawing.Point(430, 15)
$btnCheckUpdate.Size=New-Object System.Drawing.Size(200, 45)
$btnCheckUpdate.BackColor=[System.Drawing.Color]::DarkOrange
$btnCheckUpdate.ForeColor="White"
$pnlButtons.Controls.Add($btnCheckUpdate)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text="Refresh Files"
$btnRefresh.Font=$btnFont
$btnRefresh.Location=New-Object System.Drawing.Point(640, 15)
$btnRefresh.Size=New-Object System.Drawing.Size(150, 45)
$pnlButtons.Controls.Add($btnRefresh)

# ============ TAB 2: HP DRIVERS (COMPLETE OVERHAUL - PROPER LAYOUT) ============

# Use a main container panel to handle all layout
$pnlDriverMain = New-Object System.Windows.Forms.Panel
$pnlDriverMain.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlDriverMain.BackColor = [System.Drawing.Color]::White
$TabDrivers.Controls.Add($pnlDriverMain)

# === SYSTEM INFO PANEL (TOP - Fixed Height) ===
$pnlDriverInfo = New-Object System.Windows.Forms.Panel
$pnlDriverInfo.Height = 140
$pnlDriverInfo.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlDriverInfo.BackColor = [System.Drawing.Color]::FromArgb(240, 248, 255)
$pnlDriverInfo.Padding = New-Object System.Windows.Forms.Padding(15, 8, 15, 8)

# Left side - System Info
$lblSysInfo = New-Object System.Windows.Forms.Label
$lblSysInfo.Text = "System Information: Detecting..."
$lblSysInfo.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$lblSysInfo.Location = New-Object System.Drawing.Point(15, 8)
$lblSysInfo.AutoSize = $true
$pnlDriverInfo.Controls.Add($lblSysInfo)

$lblSysDetails = New-Object System.Windows.Forms.Label
$lblSysDetails.Text = ""
$lblSysDetails.Font = $FontSmall
$lblSysDetails.Location = New-Object System.Drawing.Point(15, 35)
$lblSysDetails.Size = New-Object System.Drawing.Size(700, 40)
$pnlDriverInfo.Controls.Add($lblSysDetails)

# Driver Count Labels
$lblDriverCounts = New-Object System.Windows.Forms.Label
$lblDriverCounts.Text = ""
$lblDriverCounts.Font = $FontSmall
$lblDriverCounts.Location = New-Object System.Drawing.Point(15, 80)
$lblDriverCounts.Size = New-Object System.Drawing.Size(700, 50)
$pnlDriverInfo.Controls.Add($lblDriverCounts)

# Right side - Tool Support Box (anchored to right)
$grpToolSupport = New-Object System.Windows.Forms.GroupBox
$grpToolSupport.Text = "Tool Support"
$grpToolSupport.Font = $FontSmall
$grpToolSupport.Size = New-Object System.Drawing.Size(180, 90)
$grpToolSupport.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$grpToolSupport.Location = New-Object System.Drawing.Point(1080, 5)
$pnlDriverInfo.Controls.Add($grpToolSupport)

# HPIA Badge & Status
$lblHPIABadge = New-Object System.Windows.Forms.Label
$lblHPIABadge.Text = " HPIA "
$lblHPIABadge.Font = $FontBadge
$lblHPIABadge.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 204)
$lblHPIABadge.ForeColor = [System.Drawing.Color]::White
$lblHPIABadge.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$lblHPIABadge.Size = New-Object System.Drawing.Size(50, 22)
$lblHPIABadge.Location = New-Object System.Drawing.Point(10, 22)
$grpToolSupport.Controls.Add($lblHPIABadge)

$lblHPIAStatus = New-Object System.Windows.Forms.Label
$lblHPIAStatus.Text = "Not Checked"
$lblHPIAStatus.Font = $FontSmall
$lblHPIAStatus.Location = New-Object System.Drawing.Point(65, 24)
$lblHPIAStatus.Size = New-Object System.Drawing.Size(105, 18)
$grpToolSupport.Controls.Add($lblHPIAStatus)

# CMSL Badge & Status
$lblCMSLBadge = New-Object System.Windows.Forms.Label
$lblCMSLBadge.Text = " CMSL "
$lblCMSLBadge.Font = $FontBadge
$lblCMSLBadge.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
$lblCMSLBadge.ForeColor = [System.Drawing.Color]::White
$lblCMSLBadge.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$lblCMSLBadge.Size = New-Object System.Drawing.Size(50, 22)
$lblCMSLBadge.Location = New-Object System.Drawing.Point(10, 52)
$grpToolSupport.Controls.Add($lblCMSLBadge)

$lblCMSLStatus = New-Object System.Windows.Forms.Label
$lblCMSLStatus.Text = "Not Checked"
$lblCMSLStatus.Font = $FontSmall
$lblCMSLStatus.Location = New-Object System.Drawing.Point(65, 54)
$lblCMSLStatus.Size = New-Object System.Drawing.Size(105, 18)
$grpToolSupport.Controls.Add($lblCMSLStatus)

# === DRIVER LOG PANEL (BOTTOM - Fixed Height) ===
$pnlDriverLogContainer = New-Object System.Windows.Forms.Panel
$pnlDriverLogContainer.Height = 110
$pnlDriverLogContainer.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlDriverLogContainer.BackColor = [System.Drawing.Color]::Black
$pnlDriverLogContainer.Padding = New-Object System.Windows.Forms.Padding(0)

$txtDriverLog = New-Object System.Windows.Forms.TextBox
$txtDriverLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtDriverLog.Multiline = $true
$txtDriverLog.ScrollBars = "Vertical"
$txtDriverLog.ReadOnly = $true
$txtDriverLog.BackColor = [System.Drawing.Color]::Black
$txtDriverLog.ForeColor = [System.Drawing.Color]::Cyan
$txtDriverLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$pnlDriverLogContainer.Controls.Add($txtDriverLog)

# === BUTTONS PANEL (BOTTOM - Fixed Height, above log) ===
$pnlDriverButtons = New-Object System.Windows.Forms.Panel
$pnlDriverButtons.Height = 95
$pnlDriverButtons.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlDriverButtons.BackColor = [System.Drawing.Color]::WhiteSmoke
$pnlDriverButtons.Padding = New-Object System.Windows.Forms.Padding(10, 5, 10, 5)

# Row 1 buttons - FlowLayoutPanel
$flowBtnRow1 = New-Object System.Windows.Forms.FlowLayoutPanel
$flowBtnRow1.Location = New-Object System.Drawing.Point(5, 5)
$flowBtnRow1.Size = New-Object System.Drawing.Size(800, 48)
$flowBtnRow1.WrapContents = $false
$flowBtnRow1.BackColor = [System.Drawing.Color]::Transparent
$pnlDriverButtons.Controls.Add($flowBtnRow1)

$btnScanDrivers = New-Object System.Windows.Forms.Button
$btnScanDrivers.Text = "SCAN FOR DRIVERS"
$btnScanDrivers.Font = $btnFont
$btnScanDrivers.Size = New-Object System.Drawing.Size(170, 40)
$btnScanDrivers.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
$btnScanDrivers.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 204)
$btnScanDrivers.ForeColor = [System.Drawing.Color]::White
$btnScanDrivers.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnScanDrivers.FlatAppearance.BorderSize = 0
$flowBtnRow1.Controls.Add($btnScanDrivers)

$btnDownloadDrivers = New-Object System.Windows.Forms.Button
$btnDownloadDrivers.Text = "DOWNLOAD SELECTED"
$btnDownloadDrivers.Font = $btnFont
$btnDownloadDrivers.Size = New-Object System.Drawing.Size(190, 40)
$btnDownloadDrivers.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
$btnDownloadDrivers.BackColor = [System.Drawing.Color]::RoyalBlue
$btnDownloadDrivers.ForeColor = [System.Drawing.Color]::White
$btnDownloadDrivers.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnDownloadDrivers.FlatAppearance.BorderSize = 0
$flowBtnRow1.Controls.Add($btnDownloadDrivers)

$btnInstallDrivers = New-Object System.Windows.Forms.Button
$btnInstallDrivers.Text = "INSTALL SELECTED"
$btnInstallDrivers.Font = $btnFont
$btnInstallDrivers.Size = New-Object System.Drawing.Size(170, 40)
$btnInstallDrivers.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
$btnInstallDrivers.BackColor = [System.Drawing.Color]::ForestGreen
$btnInstallDrivers.ForeColor = [System.Drawing.Color]::White
$btnInstallDrivers.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnInstallDrivers.FlatAppearance.BorderSize = 0
$flowBtnRow1.Controls.Add($btnInstallDrivers)

# Row 2 buttons - FlowLayoutPanel
$flowBtnRow2 = New-Object System.Windows.Forms.FlowLayoutPanel
$flowBtnRow2.Location = New-Object System.Drawing.Point(5, 52)
$flowBtnRow2.Size = New-Object System.Drawing.Size(900, 35)
$flowBtnRow2.WrapContents = $false
$flowBtnRow2.BackColor = [System.Drawing.Color]::Transparent
$pnlDriverButtons.Controls.Add($flowBtnRow2)

$lblSelectLabel = New-Object System.Windows.Forms.Label
$lblSelectLabel.Text = "Select by Status:"
$lblSelectLabel.Font = $FontSmall
$lblSelectLabel.Size = New-Object System.Drawing.Size(95, 26)
$lblSelectLabel.Margin = New-Object System.Windows.Forms.Padding(0, 4, 5, 0)
$lblSelectLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$flowBtnRow2.Controls.Add($lblSelectLabel)

$btnSelectNeedsUpdate = New-Object System.Windows.Forms.Button
$btnSelectNeedsUpdate.Text = "Updates Available"
$btnSelectNeedsUpdate.Font = $FontSmall
$btnSelectNeedsUpdate.Size = New-Object System.Drawing.Size(115, 26)
$btnSelectNeedsUpdate.Margin = New-Object System.Windows.Forms.Padding(0, 0, 5, 0)
$btnSelectNeedsUpdate.BackColor = [System.Drawing.Color]::DarkOrange
$btnSelectNeedsUpdate.ForeColor = [System.Drawing.Color]::White
$btnSelectNeedsUpdate.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSelectNeedsUpdate.FlatAppearance.BorderSize = 0
$flowBtnRow2.Controls.Add($btnSelectNeedsUpdate)

$btnSelectNotInstalled = New-Object System.Windows.Forms.Button
$btnSelectNotInstalled.Text = "Not Installed"
$btnSelectNotInstalled.Font = $FontSmall
$btnSelectNotInstalled.Size = New-Object System.Drawing.Size(90, 26)
$btnSelectNotInstalled.Margin = New-Object System.Windows.Forms.Padding(0, 0, 5, 0)
$btnSelectNotInstalled.BackColor = [System.Drawing.Color]::IndianRed
$btnSelectNotInstalled.ForeColor = [System.Drawing.Color]::White
$btnSelectNotInstalled.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSelectNotInstalled.FlatAppearance.BorderSize = 0
$flowBtnRow2.Controls.Add($btnSelectNotInstalled)

$btnSelectInstalled = New-Object System.Windows.Forms.Button
$btnSelectInstalled.Text = "Up to Date"
$btnSelectInstalled.Font = $FontSmall
$btnSelectInstalled.Size = New-Object System.Drawing.Size(80, 26)
$btnSelectInstalled.Margin = New-Object System.Windows.Forms.Padding(0, 0, 5, 0)
$btnSelectInstalled.BackColor = [System.Drawing.Color]::ForestGreen
$btnSelectInstalled.ForeColor = [System.Drawing.Color]::White
$btnSelectInstalled.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSelectInstalled.FlatAppearance.BorderSize = 0
$flowBtnRow2.Controls.Add($btnSelectInstalled)

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = "Select All"
$btnSelectAll.Font = $FontSmall
$btnSelectAll.Size = New-Object System.Drawing.Size(75, 26)
$btnSelectAll.Margin = New-Object System.Windows.Forms.Padding(0, 0, 5, 0)
$btnSelectAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSelectAll.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
$flowBtnRow2.Controls.Add($btnSelectAll)

$btnSelectNone = New-Object System.Windows.Forms.Button
$btnSelectNone.Text = "Select None"
$btnSelectNone.Font = $FontSmall
$btnSelectNone.Size = New-Object System.Drawing.Size(80, 26)
$btnSelectNone.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
$btnSelectNone.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSelectNone.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
$flowBtnRow2.Controls.Add($btnSelectNone)

# === DRIVER DATAGRIDVIEW (FILL remaining space) ===
$dgvDrivers = New-Object System.Windows.Forms.DataGridView
$dgvDrivers.Dock = [System.Windows.Forms.DockStyle]::Fill
$dgvDrivers.AllowUserToAddRows = $false
$dgvDrivers.AllowUserToDeleteRows = $false
$dgvDrivers.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dgvDrivers.MultiSelect = $true
$dgvDrivers.RowHeadersVisible = $false
$dgvDrivers.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
$dgvDrivers.BackgroundColor = [System.Drawing.Color]::White
$dgvDrivers.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$dgvDrivers.Font = $FontSmall
$dgvDrivers.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$dgvDrivers.AllowUserToResizeRows = $false
$dgvDrivers.ColumnHeadersVisible = $true

# Column headers config
$dgvDrivers.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$dgvDrivers.ColumnHeadersHeight = 32

# Row configuration
$dgvDrivers.RowTemplate.Height = 28
$dgvDrivers.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
$dgvDrivers.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(3, 2, 3, 2)
$dgvDrivers.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::None

Enable-DoubleBuffer $dgvDrivers

# Disable selection highlight (keep row colors visible)
$dgvDrivers.DefaultCellStyle.SelectionBackColor = $dgvDrivers.DefaultCellStyle.BackColor
$dgvDrivers.DefaultCellStyle.SelectionForeColor = $dgvDrivers.DefaultCellStyle.ForeColor

# === DEFINE COLUMNS ===
$colSelect = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colSelect.Name = "Select"
$colSelect.HeaderText = ""
$colSelect.Width = 35
$colSelect.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
$colSelect.Resizable = [System.Windows.Forms.DataGridViewTriState]::False
$dgvDrivers.Columns.Add($colSelect) | Out-Null

$colSource = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSource.Name = "Source"
$colSource.HeaderText = "Source"
$colSource.Width = 55
$colSource.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
$colSource.ReadOnly = $true
$dgvDrivers.Columns.Add($colSource) | Out-Null

$colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colStatus.Name = "Status"
$colStatus.HeaderText = "Status"
$colStatus.Width = 100
$colStatus.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
$colStatus.ReadOnly = $true
$dgvDrivers.Columns.Add($colStatus) | Out-Null

$colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.Name = "Name"
$colName.HeaderText = "Driver/Update Name"
$colName.MinimumWidth = 180
$colName.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$colName.FillWeight = 100
$colName.ReadOnly = $true
$dgvDrivers.Columns.Add($colName) | Out-Null

$colCategory = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colCategory.Name = "Category"
$colCategory.HeaderText = "Category"
$colCategory.Width = 100
$colCategory.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
$colCategory.ReadOnly = $true
$dgvDrivers.Columns.Add($colCategory) | Out-Null

$colInstalled = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colInstalled.Name = "Installed"
$colInstalled.HeaderText = "Installed"
$colInstalled.Width = 95
$colInstalled.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
$colInstalled.ReadOnly = $true
$dgvDrivers.Columns.Add($colInstalled) | Out-Null

$colLatest = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colLatest.Name = "Latest"
$colLatest.HeaderText = "Available"
$colLatest.Width = 95
$colLatest.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
$colLatest.ReadOnly = $true
$dgvDrivers.Columns.Add($colLatest) | Out-Null

$colSPID = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSPID.Name = "SoftPaqId"
$colSPID.HeaderText = "SoftPaq"
$colSPID.Width = 70
$colSPID.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
$colSPID.ReadOnly = $true
$dgvDrivers.Columns.Add($colSPID) | Out-Null

# Style the grid headers
$dgvDrivers.EnableHeadersVisualStyles = $false
$dgvDrivers.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(0, 100, 0)
$dgvDrivers.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$dgvDrivers.ColumnHeadersDefaultCellStyle.Font = $FontItem
$dgvDrivers.ColumnHeadersDefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft

# === ADD CONTROLS TO MAIN CONTAINER IN CORRECT ORDER ===
# Order matters for docking: Bottom items first, then Top, then Fill
$pnlDriverMain.Controls.Add($dgvDrivers)           # Fill - added first, gets remaining space
$pnlDriverMain.Controls.Add($pnlDriverLogContainer) # Bottom
$pnlDriverMain.Controls.Add($pnlDriverButtons)     # Bottom (above log)
$pnlDriverMain.Controls.Add($pnlDriverInfo)        # Top

# Ensure proper z-order by sending to back/front
$pnlDriverInfo.SendToBack()
$dgvDrivers.SendToBack()
$pnlDriverButtons.BringToFront()
$pnlDriverLogContainer.BringToFront()

# ============ TAB 3: System Config ============
$grpSys = New-Object System.Windows.Forms.GroupBox
$grpSys.Text="System Settings"
$grpSys.Font=$FontHeader
$grpSys.Location=New-Object System.Drawing.Point(20,20)
$grpSys.Size=New-Object System.Drawing.Size(500,350)
$TabSystem.Controls.Add($grpSys)

$btnApplySys = New-Object System.Windows.Forms.Button
$btnApplySys.Text="APPLY SYSTEM SETTINGS"
$btnApplySys.Font=$btnFont
$btnApplySys.Location=New-Object System.Drawing.Point(540, 40)
$btnApplySys.Size=New-Object System.Drawing.Size(240, 45)
$btnApplySys.BackColor=[System.Drawing.Color]::SteelBlue
$btnApplySys.ForeColor="White"
$TabSystem.Controls.Add($btnApplySys)

function Add-SysConfig {
    param($Txt, $Y, $Ref)

    $c = New-Object System.Windows.Forms.CheckBox
    $c.Text = $Txt
    $c.Font = $FontItem
    $c.Checked = $true
    $c.Location = New-Object System.Drawing.Point(20, $Y)
    $c.Size = New-Object System.Drawing.Size(360, 25)
    $grpSys.Controls.Add($c)
    Set-Variable $Ref $c -Scope Global

    $l = New-Object System.Windows.Forms.Label
    $l.Location = New-Object System.Drawing.Point(40, ($Y + 25))
    $l.Size = New-Object System.Drawing.Size(420, 20)
    $l.Font = $FontSmall
    $grpSys.Controls.Add($l)
    return $l
}

$lblTime  = Add-SysConfig "Set TimeZone (West Bank)" 40  "chkTime"
$lblPwr   = Add-SysConfig "High Performance Power"   90  "chkPwr"
$lblBoot  = Add-SysConfig "Disable Fast Boot"        140 "chkBoot"
$lblIcons = Add-SysConfig "Enable Desktop Icons"     190 "chkIcons"
$lblLng   = Add-SysConfig "Locale (AR-SA / US)"      240 "chkLng"

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20,400)
$txtLog.Size = New-Object System.Drawing.Size(1200,300)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::Black
$txtLog.ForeColor = [System.Drawing.Color]::LimeGreen
$txtLog.Font = New-Object System.Drawing.Font("Consolas",10)
$txtLog.Anchor = "Top, Bottom, Left, Right"
$TabSystem.Controls.Add($txtLog)

# -------------------------
# 5) Logic
# -------------------------
function Check-SystemConfiguration {
    try {
        $tz = (Get-TimeZone).Id
        if ($tz -eq "West Bank Standard Time") { $lblTime.Text="[OK] Configured"; $lblTime.ForeColor=[System.Drawing.Color]::Green }
        else { $lblTime.Text="[X] $tz"; $lblTime.ForeColor=[System.Drawing.Color]::Red }
    } catch { $lblTime.Text="[X] Error"; $lblTime.ForeColor=[System.Drawing.Color]::Red }

    try {
        if ((powercfg /getactivescheme) -match 'High performance') { $lblPwr.Text="[OK] Configured"; $lblPwr.ForeColor=[System.Drawing.Color]::Green }
        else { $lblPwr.Text="[X] Not Set"; $lblPwr.ForeColor=[System.Drawing.Color]::Red }
    } catch { $lblPwr.Text="[X] Error"; $lblPwr.ForeColor=[System.Drawing.Color]::Red }

    try {
        $fb = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" "HiberbootEnabled" -EA SilentlyContinue
        if ($fb.HiberbootEnabled -eq 0) { $lblBoot.Text="[OK] Disabled"; $lblBoot.ForeColor=[System.Drawing.Color]::Green }
        else { $lblBoot.Text="[X] Enabled"; $lblBoot.ForeColor=[System.Drawing.Color]::Red }
    } catch { $lblBoot.Text="[X] Error"; $lblBoot.ForeColor=[System.Drawing.Color]::Red }

    try {
        $hi = Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "HideIcons" -EA SilentlyContinue
        if ($hi.HideIcons -eq 0) { $lblIcons.Text="[OK] Enabled"; $lblIcons.ForeColor=[System.Drawing.Color]::Green }
        else { $lblIcons.Text="[X] Disabled"; $lblIcons.ForeColor=[System.Drawing.Color]::Red }
    } catch { $lblIcons.Text="[X] Error"; $lblIcons.ForeColor=[System.Drawing.Color]::Red }

    try {
        $sl = Get-WinSystemLocale
        $df = (Get-ItemProperty "HKCU:\Control Panel\International" "sShortDate").sShortDate
        if ($sl.Name -eq "ar-SA" -and $df -eq "dd/MM/yyyy") { $lblLng.Text="[OK] Configured"; $lblLng.ForeColor=[System.Drawing.Color]::Green }
        else { $lblLng.Text="[X] Not Set"; $lblLng.ForeColor=[System.Drawing.Color]::Red }
    } catch { $lblLng.Text="[X] Error"; $lblLng.ForeColor=[System.Drawing.Color]::Red }
}

function Apply-SystemConfiguration {
    Log-Msg "Applying System Settings..."
    try { if ($chkTime.Checked) { cmd /c "tzutil /s `"West Bank Standard Time`"" | Out-Null } } catch { Log-Msg "TimeZone apply error: $_" }

    try {
        if ($chkPwr.Checked) {
            powercfg /setactive SCHEME_MAX | Out-Null
        }
    } catch { Log-Msg "Power apply error: $_" }

    try { if ($chkBoot.Checked) { Set-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" "HiberbootEnabled" 0 -Force } } catch { Log-Msg "FastBoot apply error: $_" }
    try { if ($chkIcons.Checked){ Set-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "HideIcons" 0 -Force } } catch { Log-Msg "Desktop icons apply error: $_" }
    try {
        if ($chkLng.Checked)  {
            try { Set-WinSystemLocale ar-SA } catch {}
            try { Set-ItemProperty "HKCU:\Control Panel\International" "sShortDate" "dd/MM/yyyy" -Force } catch {}
        }
    } catch { Log-Msg "Locale apply error: $_" }

    Check-SystemConfiguration
    Log-Msg "System Settings Applied."
}

function Refresh-UI {
    $pnlApps.SuspendLayout()
    try {
        Log-Msg "Scanning System & Files..."
        Check-SystemConfiguration

        foreach ($App in $Global:AppList) {
            $UI = $Global:UIControls[$App.Name]

            $UI.LblLat.Text = ""
            $UI.LblLat.ForeColor = [System.Drawing.Color]::Black

            $li = Get-LocalInstallerInfo $App
            $Global:UIControls[$App.Name].LocalInfo = $li
            $UI.LblFile.Text = $li.Text

            $InstVer = $null
            if ($App.DualArch -and $App.VCKey) {
                $InstVer = Get-VCCompositeText $App.VCKey
            } else {
                if ($App.Pat) { $InstVer = Get-RegVer $App.Pat }
            }

            $isInstalled = $false
            if ($InstVer) {
                if ($App.DualArch) {
                    if (($InstVer -notmatch 'x86:\s*Missing') -or ($InstVer -notmatch 'x64:\s*Missing')) { $isInstalled = $true }
                } else {
                    $isInstalled = $true
                }
            }

            if ($isInstalled) {
                $UI.LblInst.Text = $InstVer
                Set-RowColor $UI ([System.Drawing.Color]::Green)
            } else {
                $UI.LblInst.Text = "Not Installed"
                Set-RowColor $UI ([System.Drawing.Color]::Red)
            }
        }

        Log-Msg "File Check Ready."
    }
    finally {
        $pnlApps.ResumeLayout($true)
        $pnlApps.Invalidate()
        $pnlApps.Update()
    }
}

function Check-LatestVersions {
    Log-Msg "Connecting to Winget/Internet for Latest Versions..."
    $btnCheckUpdate.Enabled = $false
    $btnCheckUpdate.Text = "Checking..."

    $pnlApps.SuspendLayout()
    try {
        $exe = Get-WingetPath
        if ($exe) { try { & $exe source update --name winget | Out-Null } catch {} }

        foreach ($App in $Global:AppList) {
            $UI = $Global:UIControls[$App.Name]

            $UI.LblLat.Text = "Checking..."
            [System.Windows.Forms.Application]::DoEvents()

            $rowInstalled = ($UI.LblInst.Text -ne "Not Installed")
            $LatText = "N/A"

            if ($App.Name -match "Office 2024") {
                $LatText = Get-OfficeVerWeb "https://learn.microsoft.com/en-us/officeupdates/update-history-office-2024"
            }
            elseif ($App.Name -match "Office 365") {
                $LatText = Get-OfficeVerWeb "https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date"
            }
            elseif ($App.Name -match "HP Support") {
                $LatText = Get-HPSALatest
            }
            elseif ($App.Name -eq "iVMS-4200") {
                $iv = Get-iVMSLatest
                $LatText = if ($iv -and $iv.Version) { $iv.Version } else { "N/A" }
            }
            elseif ($App.DownloadMode -eq "localonly") {
                $li = $Global:UIControls[$App.Name].LocalInfo
                if ($li -and $li.Path) {
                    $v = Get-FileProductVersion $li.Path.FullName
                    $LatText = if ($v) { $v } else { "N/A" }
                } else { $LatText = "N/A" }
            }
            elseif ($App.DualArch -and $App.Id64 -and $App.Id86) {
                $lat64 = Get-WingetVer $App.Id64 "winget"
                $lat86 = Get-WingetVer $App.Id86 "winget"
                if ($lat86 -eq $lat64) { $LatText = $lat64 } else { $LatText = "x86: $lat86 | x64: $lat64" }
            }
            elseif ($App.Id) {
                $src = if ($App.Source) { $App.Source } else { "winget" }
                $LatText = Get-WingetVer $App.Id $src
            }

            if ([string]::IsNullOrWhiteSpace($LatText)) { $LatText = "N/A" }
            $UI.LblLat.Text = $LatText

            if (-not $rowInstalled) { Set-RowColor $UI ([System.Drawing.Color]::Red); continue }
            if ($LatText -in @("N/A","Error","Winget Missing")) { Set-RowColor $UI ([System.Drawing.Color]::Black); continue }

            $isMatch = $false

            if ($App.DualArch) {
                $ix86 = $null; $ix64 = $null
                if ($UI.LblInst.Text -match 'x86:\s*([0-9.]+|Missing)') { $ix86 = $Matches[1] }
                if ($UI.LblInst.Text -match 'x64:\s*([0-9.]+|Missing)') { $ix64 = $Matches[1] }

                $lx86 = $null; $lx64 = $null
                if ($LatText -match '^x86:\s*([0-9.]+)\s*\|\s*x64:\s*([0-9.]+)') {
                    $lx86 = $Matches[1]; $lx64 = $Matches[2]
                } else {
                    $lx86 = $LatText; $lx64 = $LatText
                }

                $ok86 = $false; $ok64 = $false

                if ($ix86 -and $ix86 -ne "Missing" -and $lx86) {
                    $vI = To-VersionObj $ix86
                    $vL = To-VersionObj $lx86
                    if ($vI -and $vL) { $ok86 = ($vI -ge $vL) } else { $ok86 = ($ix86 -eq $lx86) }
                }
                if ([Environment]::Is64BitOperatingSystem -and $ix64 -and $ix64 -ne "Missing" -and $lx64) {
                    $vI = To-VersionObj $ix64
                    $vL = To-VersionObj $lx64
                    if ($vI -and $vL) { $ok64 = ($vI -ge $vL) } else { $ok64 = ($ix64 -eq $lx64) }
                } else {
                    if (-not [Environment]::Is64BitOperatingSystem) { $ok64 = $true }
                }

                $isMatch = ($ok86 -and $ok64)
            }
            elseif ($App.Name -match "Office" -and $LatText -match '\(([\d.]+)\)' -and $UI.LblInst.Text -match '^16\.0\.([\d.]+)$') {
                $instBuild = $Matches[1]
                $latBuild  = ([regex]::Match($LatText, '\(([\d.]+)\)').Groups[1].Value)
                $vI = To-VersionObj $instBuild
                $vL = To-VersionObj $latBuild
                if ($vI -and $vL) { $isMatch = ($vI -ge $vL) }
            }
            else {
                $vI = To-VersionObj $UI.LblInst.Text
                $vL = To-VersionObj $LatText
                if ($vI -and $vL) { $isMatch = ($vI -ge $vL) } else { $isMatch = ($UI.LblInst.Text -eq $LatText) }
            }

            if ($isMatch) { Set-RowColor $UI ([System.Drawing.Color]::Green) }
            else { Set-RowColor $UI ([System.Drawing.Color]::DarkOrange) }
        }

        Log-Msg "Version Check Complete."
    }
    finally {
        $pnlApps.ResumeLayout($true)
        $pnlApps.Invalidate()
        $pnlApps.Update()

        $btnCheckUpdate.Enabled = $true
        $btnCheckUpdate.Text = "CHECK FOR UPDATES"
    }
}

# -------------------------
# 5.1) Driver Tab Logic
# -------------------------
function Refresh-DriverSystemInfo {
    Log-Driver "Detecting system information..."
    $Global:HPSystemInfo = Get-HPSystemInfo
    $Global:InstalledDriversCache = $null  # Reset cache

    $si = $Global:HPSystemInfo

    if ($si.IsHP) {
        $lblSysInfo.Text = "HP System Detected: $($si.Model)"
        $lblSysInfo.ForeColor = [System.Drawing.Color]::FromArgb(0, 100, 0)

        $details = "Manufacturer: $($si.Manufacturer)  |  Product Code: $($si.ProductCode)  |  SKU: $($si.SKU)`r`n"
        $details += "Serial: $($si.SerialNumber)  |  OS: $($si.OSVersion) (Build $($si.OSBuild))"
        if ($si.Generation) { $details += "  |  Gen: G$($si.Generation)" }
        $lblSysDetails.Text = $details

        if ($si.SupportsHPIA) {
            $lblHPIABadge.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 204)
            $lblHPIAStatus.Text = "Supported"
            $lblHPIAStatus.ForeColor = [System.Drawing.Color]::Green
        } else {
            $lblHPIABadge.BackColor = [System.Drawing.Color]::Gray
            $lblHPIAStatus.Text = "Not Supported"
            $lblHPIAStatus.ForeColor = [System.Drawing.Color]::Red
        }

        if ($si.SupportsCMSL) {
            $lblCMSLBadge.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
            $lblCMSLStatus.Text = "Available"
            $lblCMSLStatus.ForeColor = [System.Drawing.Color]::Green
        } else {
            $lblCMSLBadge.BackColor = [System.Drawing.Color]::Gray
            $lblCMSLStatus.Text = "Not Available"
            $lblCMSLStatus.ForeColor = [System.Drawing.Color]::Red
        }

        Log-Driver "HP System: $($si.Model) | Platform: $($si.ProductCode) | SKU: $($si.SKU) | Gen: $($si.Generation)"
        Log-Driver "HPIA Support: $($si.SupportsHPIA) | CMSL Support: $($si.SupportsCMSL)"
    } else {
        $lblSysInfo.Text = "Non-HP System: $($si.Manufacturer) $($si.Model)"
        $lblSysInfo.ForeColor = [System.Drawing.Color]::DarkOrange

        $lblSysDetails.Text = "This tool is designed for HP systems.`r`nFor Lenovo, use Lenovo System Update / Update Retriever.`r`nFor Dell, use Dell Command Update."

        $lblHPIABadge.BackColor = [System.Drawing.Color]::Gray
        $lblCMSLBadge.BackColor = [System.Drawing.Color]::Gray
        $lblHPIAStatus.Text = "N/A"
        $lblCMSLStatus.Text = "N/A"

        Log-Driver "Non-HP system detected: $($si.Manufacturer) $($si.Model)"
    }
}

function Update-DriverCounts {
    $upToDate = 0
    $needsUpdate = 0
    $notInstalled = 0

    foreach ($row in $dgvDrivers.Rows) {
        $status = $row.Cells["Status"].Value
        switch -Regex ($status) {
            "Up to Date|Installed" { $upToDate++ }
            "Update Available|Critical|Recommended" { $needsUpdate++ }
            "Not Installed|Not Found" { $notInstalled++ }
        }
    }

    $total = $dgvDrivers.Rows.Count
    $lblDriverCounts.Text = "Total: $total  |  " +
        "Up to Date: $upToDate  |  " +
        "Updates Available: $needsUpdate  |  " +
        "Not Installed: $notInstalled"

    $lblDriverCounts.ForeColor = [System.Drawing.Color]::Black
}

function Scan-Drivers {
    $dgvDrivers.Rows.Clear()
    $Global:HPDriverList = @()
    $Global:InstalledDriversCache = $null

    $si = $Global:HPSystemInfo
    if (-not $si -or -not $si.IsHP) {
        Log-Driver "Cannot scan: Not an HP system."
        return
    }

    $btnScanDrivers.Enabled = $false
    $btnScanDrivers.Text = "Scanning..."

    try {
        Log-Driver "Building installed software/drivers cache..."
        $Global:InstalledDriversCache = Get-InstalledDriversAndSoftware
        Log-Driver "Found $($Global:InstalledDriversCache.Count) installed items."

        $allDrivers = @()

        # ----------------------------------------
        # LEGACY LOGIC with AUTO-AUTH
        # ----------------------------------------
        if (-not $si.SupportsHPIA -and -not $si.SupportsCMSL) {
            Log-Driver "Legacy system detected. Checking offline repository..."
            
            # === NEW: TRIGGER AUTHENTICATION IF NEEDED ===
            if (-not (Connect-Repo-WithAuth)) {
                Log-Driver "Error: Could not access network repository. Authentication failed or cancelled."
                [System.Windows.Forms.MessageBox]::Show("Could not access the network repository.`n`nPlease ensure you have permission to access:`n$Global:HPLegacyRepoRoot", "Network Error", "OK", "Error")
            } 
            else {
                # Proceed only if Auth was successful
                $repoPackages = Get-LegacyRepoPackages -Platform $si.ProductCode -Model $si.Model
                
                if ($repoPackages.Count -gt 0) {
                    Log-Driver "Found $($repoPackages.Count) packages in legacy repository."
                    foreach ($pkg in $repoPackages) {
                        $status = Get-DriverInstallStatus -DriverName $pkg.Name -Category $pkg.Category -AvailableVersion $pkg.Version
                        
                        if ($pkg.Category -match "BIOS") {
                            if ($status.Status -eq "Update Available") { $status.Status = "Critical" }
                        }

                        $allDrivers += [PSCustomObject]@{
                            Source = "Repo"
                            Name = $pkg.Name
                            Category = $pkg.Category
                            Version = $pkg.Version
                            InstalledVersion = $status.InstalledVersion
                            SoftPaqId = $pkg.SoftPaqId
                            Url = (Join-Path $pkg._RepoPath "SoftPaqs\$($pkg.FileName)")
                            Size = "Local"
                            Status = $status.Status
                            Selected = $false
                            FilePath = $null
                        }
                    }
                } else {
                    Log-Driver "No offline repository found for this model."
                    Log-Driver "Looked in: $Global:HPLegacyRepoRoot"
                }
            }
        } 
        else {
            # ----------------------------------------
            # MODERN LOGIC (HPIA / CMSL)
            # ----------------------------------------
            
            # Try HPIA first
            if ($si.SupportsHPIA) {
                Log-Driver "System supports HPIA - attempting HPIA scan..."
                $hpiaPath = Get-HPIAPath
                if (-not $hpiaPath) {
                    Log-Driver "HPIA not found locally. Downloading..."
                    $hpiaPath = Download-HPIA
                }

                if ($hpiaPath -and (Test-Path $hpiaPath)) {
                    $hpiaDrivers = Scan-HPDrivers-HPIA -HPIAPath $hpiaPath
                    $allDrivers += $hpiaDrivers
                } else {
                    Log-Driver "HPIA not available. Falling back to CMSL."
                }
            }

            # Use CMSL
            if ($si.SupportsCMSL -and $si.ProductCode) {
                $cmslDrivers = Scan-HPDrivers-CMSL -Platform $si.ProductCode
                $existingIds = $allDrivers | ForEach-Object { $_.SoftPaqId } | Where-Object { $_ }
                foreach ($cd in $cmslDrivers) {
                    if ($cd.SoftPaqId -notin $existingIds) {
                        $allDrivers += $cd
                    }
                }
            }
        }

        # Sort: Critical > Update Available > Not Installed > Up to Date
        $sortOrder = @{
            "Critical" = 0
            "Recommended" = 1
            "Update Available" = 2
            "Not Installed" = 3
            "Not Found" = 3
            "Optional" = 4
            "Available" = 4
            "Up to Date" = 5
            "Installed" = 5
        }

        $allDrivers = $allDrivers | Sort-Object {
            $order = $sortOrder[$_.Status]
            if ($null -eq $order) { $order = 99 }
            $order
        }, Name

        $Global:HPDriverList = $allDrivers

        # Populate grid
        $dgvDrivers.SuspendLayout()
        
        foreach ($drv in $allDrivers) {
            $rowIdx = $dgvDrivers.Rows.Add()
            $row = $dgvDrivers.Rows[$rowIdx]

            $row.Cells["Select"].Value = $drv.Selected
            $row.Cells["Source"].Value = $drv.Source
            $row.Cells["Name"].Value = $drv.Name
            $row.Cells["Category"].Value = $drv.Category
            $row.Cells["Installed"].Value = $drv.InstalledVersion
            $row.Cells["Latest"].Value = $drv.Version
            $row.Cells["Status"].Value = $drv.Status
            $row.Cells["SoftPaqId"].Value = $drv.SoftPaqId

            # Source badge color
            if ($drv.Source -eq "HPIA") {
                $row.Cells["Source"].Style.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 204)
                $row.Cells["Source"].Style.ForeColor = [System.Drawing.Color]::White
            } 
            elseif ($drv.Source -eq "Repo") {
                $row.Cells["Source"].Style.BackColor = [System.Drawing.Color]::Purple
                $row.Cells["Source"].Style.ForeColor = [System.Drawing.Color]::White
            }
            else {
                $row.Cells["Source"].Style.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
                $row.Cells["Source"].Style.ForeColor = [System.Drawing.Color]::White
            }

            # Status/Row color
            switch -Regex ($drv.Status) {
                "Critical" {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 200, 200)
                    $row.Cells["Status"].Style.BackColor = [System.Drawing.Color]::Red
                    $row.Cells["Status"].Style.ForeColor = [System.Drawing.Color]::White
                }
                "Recommended|Update Available" {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 240, 200)
                    $row.Cells["Status"].Style.BackColor = [System.Drawing.Color]::DarkOrange
                    $row.Cells["Status"].Style.ForeColor = [System.Drawing.Color]::White
                }
                "Not Installed|Not Found" {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 220, 220)
                    $row.Cells["Status"].Style.BackColor = [System.Drawing.Color]::IndianRed
                    $row.Cells["Status"].Style.ForeColor = [System.Drawing.Color]::White
                }
                "Up to Date|Installed" {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(200, 255, 200)
                    $row.Cells["Status"].Style.BackColor = [System.Drawing.Color]::ForestGreen
                    $row.Cells["Status"].Style.ForeColor = [System.Drawing.Color]::White
                }
                default {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
                    $row.Cells["Status"].Style.BackColor = [System.Drawing.Color]::LightGray
                }
            }

            $row.Tag = $drv
        }

        $dgvDrivers.ResumeLayout($true)
        
        # Force refresh and layout recalculation
        $dgvDrivers.Refresh()
        $dgvDrivers.Invalidate()
        [System.Windows.Forms.Application]::DoEvents()

        Update-DriverCounts
        
        Log-Driver "Scan complete. Found $($allDrivers.Count) drivers/updates."
        Log-Driver "DEBUG: Grid has $($dgvDrivers.Rows.Count) rows visible in DataGridView"
        
        # === REMOVE AUTO-SELECTION ===
        $dgvDrivers.ClearSelection()
        $dgvDrivers.CurrentCell = $null
        # === END ===

    } catch {

        Log-Driver "Scan error: $_"
    } finally {
        $btnScanDrivers.Enabled = $true
        $btnScanDrivers.Text = "SCAN FOR DRIVERS"
    }
}
function Download-SelectedDrivers {
    $selected = @()
    foreach ($row in $dgvDrivers.Rows) {
        if ($row.Cells["Select"].Value -eq $true) {
            $selected += $row.Tag
        }
    }

    if ($selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No drivers selected.", "Info", "OK", "Information")
        return
    }

    $btnDownloadDrivers.Enabled = $false
    $btnDownloadDrivers.Text = "Downloading..."

    $dlDir = Join-Path $ScriptPath "HPDrivers"
    New-Item -ItemType Directory -Force -Path $dlDir | Out-Null

    $success = 0
    $failed = 0

    foreach ($drv in $selected) {
        Log-Driver "Downloading: $($drv.Name) ($($drv.SoftPaqId))..."
        [System.Windows.Forms.Application]::DoEvents()

        $path = Download-HPSoftPaq -SoftPaqId $drv.SoftPaqId -Url $drv.Url -DestDir $dlDir
        if ($path) {
            $drv.FilePath = $path
            $success++
            Log-Driver "Ready: $([System.IO.Path]::GetFileName($path))"
        } else {
            $failed++
            Log-Driver "Failed to download: $($drv.SoftPaqId)"
        }
    }

    $btnDownloadDrivers.Enabled = $true
    $btnDownloadDrivers.Text = "DOWNLOAD SELECTED"

    [System.Windows.Forms.MessageBox]::Show("Download complete.`n`nSuccess: $success`nFailed: $failed`n`nFiles saved to: $dlDir", "Download Complete", "OK", "Information")
}

function Install-SelectedDrivers {
    $selected = @()
    foreach ($row in $dgvDrivers.Rows) {
        if ($row.Cells["Select"].Value -eq $true) {
            $selected += $row.Tag
        }
    }

    if ($selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No drivers selected.", "Info", "OK", "Information")
        return
    }

    $result = [System.Windows.Forms.MessageBox]::Show("Install $($selected.Count) selected driver(s)?`n`nA reboot may be required.", "Confirm Install", "YesNo", "Question")
    if ($result -ne "Yes") { return }

    $btnInstallDrivers.Enabled = $false
    $btnInstallDrivers.Text = "Installing..."

    $dlDir = Join-Path $ScriptPath "HPDrivers"
    $success = 0
    $failed = 0
    $needsReboot = $false

    foreach ($drv in $selected) {
        Log-Driver "Processing: $($drv.Name)..."
        [System.Windows.Forms.Application]::DoEvents()

        $path = $drv.FilePath
        if (-not $path -or -not (Test-Path $path)) {
            Log-Driver "Downloading first: $($drv.SoftPaqId)..."
            $path = Download-HPSoftPaq -SoftPaqId $drv.SoftPaqId -Url $drv.Url -DestDir $dlDir
            $drv.FilePath = $path
        }

        if ($path -and (Test-Path $path)) {
            if (Install-HPSoftPaq -Path $path) {
                $success++
                if ($drv.Category -match 'BIOS|Firmware|Chipset') { $needsReboot = $true }
            } else {
                $failed++
            }
        } else {
            Log-Driver "Cannot install $($drv.Name): file not available."
            $failed++
        }
    }

    $btnInstallDrivers.Enabled = $true
    $btnInstallDrivers.Text = "INSTALL SELECTED"

    $msg = "Installation complete.`n`nSuccess: $success`nFailed: $failed"
    if ($needsReboot) { $msg += "`n`nA reboot is recommended." }
    [System.Windows.Forms.MessageBox]::Show($msg, "Install Complete", "OK", "Information")

    # Refresh scan
    $Global:InstalledDriversCache = $null
}

# -------------------------
# 6) Download Logic (versioned filenames)
# -------------------------
function Winget-Download-And-Rename {
    param(
        [string]$WingetExe,
        [string]$Id,
        [string]$Stem,
        [string]$VersionString
    )

    $verPart = Safe-FilePart $VersionString
    $dlDir = Join-Path $ScriptPath ("TempDL_" + (Safe-FilePart $Stem))
    if (Test-Path $dlDir) { Remove-Item $dlDir -Recurse -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Force -Path $dlDir | Out-Null

    & $WingetExe download --id $Id -d $dlDir --accept-source-agreements --accept-package-agreements --force --locale en-US | Out-Null

    $dlFile = Get-ChildItem $dlDir -File -Recurse -ErrorAction SilentlyContinue |
              Where-Object { $_.Extension -match '^\.(exe|msi)$' } |
              Select-Object -First 1

    if (-not $dlFile) {
        Remove-Item $dlDir -Recurse -Force -ErrorAction SilentlyContinue
        return $null
    }

    $ext = $dlFile.Extension.TrimStart(".")
    $destName = "{0}_{1}.{2}" -f $Stem, $verPart, $ext
    $destPath = Join-Path $ScriptPath $destName

    Move-Item $dlFile.FullName $destPath -Force
    Remove-Item $dlDir -Recurse -Force -ErrorAction SilentlyContinue

    return (Get-Item $destPath)
}

$btnDownload.Add_Click({
    $btnDownload.Enabled = $false
    Log-Msg "Starting Download Process..."

    $exe = Get-WingetPath
    if (-not $exe) { Log-Msg "Winget missing. Cannot download winget packages." }

    foreach ($App in $Global:AppList) {
        if (-not $Global:UIControls[$App.Name].Chk.Checked) { continue }

        if ($App.DownloadMode -in @("onlineonly","localonly")) {
            Log-Msg "Skipping $($App.Name): $($App.DownloadMode)."
            continue
        }

        if ($App.DownloadMode -eq "office") {
            $wd = Join-Path $ScriptPath "OfficeSetup"
            New-Item -Type Dir -Force $wd | Out-Null

            if (Test-Path (Join-Path $ScriptPath "setup.exe")) {
                Log-Msg "Skipping Office ODT: setup.exe already exists."
                continue
            }

            if ($exe) {
                Log-Msg "Downloading Office Deployment Tool (extract setup.exe)..."
                try {
                    $cmd = "& `"$exe`" install --id Microsoft.OfficeDeploymentTool --silent --force --override `"/quiet /extract:$wd`""
                    Invoke-Expression $cmd | Out-Null
                    $setup = Join-Path $wd "setup.exe"
                    if (Test-Path $setup) {
                        Move-Item $setup (Join-Path $ScriptPath "setup.exe") -Force
                        Log-Msg "Downloaded: setup.exe"
                    } else {
                        Log-Msg "ODT extraction failed (setup.exe not found)."
                    }
                } catch { Log-Msg "ODT download error: $_" }
            }
            continue
        }

        if ($App.DownloadMode -eq "direct" -and $App.Name -eq "iVMS-4200") {
            Log-Msg "Downloading iVMS-4200..."
            try {
                $d = Get-iVMSLatest
                if ($d -and $d.Url -and $d.Version) {
                    $verPart = Safe-FilePart $d.Version
                    $dest = Join-Path $ScriptPath ("{0}_{1}.exe" -f $App.FileStem, $verPart)
                    if (Test-Path $dest) { Log-Msg "Skipping iVMS-4200: already have $([IO.Path]::GetFileName($dest))."; continue }
                    Invoke-WebRequest -Uri $d.Url -OutFile $dest -UserAgent "Mozilla/5.0"
                    Log-Msg "Downloaded iVMS-4200 ($($d.Version)) -> $([IO.Path]::GetFileName($dest))"
                } else {
                    Log-Msg "iVMS-4200 latest URL not found."
                }
            } catch { Log-Msg "Error downloading iVMS-4200: $_" }
            continue
        }

        if (-not $exe) { continue }

        try {
            if ($App.DualArch -and $App.Id86 -and $App.Id64) {
                $v86 = Get-WingetVer $App.Id86 "winget"
                $v64 = Get-WingetVer $App.Id64 "winget"

                if ($v86 -notin @("N/A","Error","Winget Missing")) {
                    $destPrefix = "{0}_{1}." -f $App.FileStem86, (Safe-FilePart $v86)
                    if (Get-ChildItem $ScriptPath -File -Filter ($destPrefix + "*") -ErrorAction SilentlyContinue) {
                        Log-Msg "Skipping $($App.Name) x86: already have versioned file."
                    } else {
                        Log-Msg "Downloading $($App.Name) x86 ($v86)..."
                        $dl = Winget-Download-And-Rename -WingetExe $exe -Id $App.Id86 -Stem $App.FileStem86 -VersionString $v86
                        if ($dl) { Log-Msg "Downloaded -> $($dl.Name)" } else { Log-Msg "Winget download failed (x86) for $($App.Name)." }
                    }
                }

                if ([Environment]::Is64BitOperatingSystem -and $v64 -notin @("N/A","Error","Winget Missing")) {
                    $destPrefix = "{0}_{1}." -f $App.FileStem64, (Safe-FilePart $v64)
                    if (Get-ChildItem $ScriptPath -File -Filter ($destPrefix + "*") -ErrorAction SilentlyContinue) {
                        Log-Msg "Skipping $($App.Name) x64: already have versioned file."
                    } else {
                        Log-Msg "Downloading $($App.Name) x64 ($v64)..."
                        $dl = Winget-Download-And-Rename -WingetExe $exe -Id $App.Id64 -Stem $App.FileStem64 -VersionString $v64
                        if ($dl) { Log-Msg "Downloaded -> $($dl.Name)" } else { Log-Msg "Winget download failed (x64) for $($App.Name)." }
                    }
                }
            }
            elseif ($App.Id) {
                $src = if ($App.Source) { $App.Source } else { "winget" }
                if ($src -ne "winget") {
                    Log-Msg "Skipping download for $($App.Name): source=$src (install-only here)."
                    continue
                }

                $ver = Get-WingetVer $App.Id "winget"
                if ($ver -notin @("N/A","Error","Winget Missing")) {
                    $destPrefix = "{0}_{1}." -f $App.FileStem, (Safe-FilePart $ver)
                    if (Get-ChildItem $ScriptPath -File -Filter ($destPrefix + "*") -ErrorAction SilentlyContinue) {
                        Log-Msg "Skipping $($App.Name): already have versioned file."
                        continue
                    }
                    Log-Msg "Downloading $($App.Name) ($ver)..."
                    $dl = Winget-Download-And-Rename -WingetExe $exe -Id $App.Id -Stem $App.FileStem -VersionString $ver
                    if ($dl) { Log-Msg "Downloaded -> $($dl.Name)" } else { Log-Msg "Winget download failed for $($App.Name)." }
                }
            }
        } catch {
            Log-Msg "Download error for $($App.Name): $_"
        }
    }

    Refresh-UI
    $btnDownload.Enabled = $true
    [System.Windows.Forms.MessageBox]::Show("Download process finished.")
})

# -------------------------
# 7) Install Logic (prefers local)
# -------------------------
$btnInstall.Add_Click({
    $btnInstall.Enabled = $false
    Log-Msg "Starting Installation..."

    $exe = Get-WingetPath

    foreach ($App in $Global:AppList) {
        if (-not $Global:UIControls[$App.Name].Chk.Checked) { continue }

        $UI = $Global:UIControls[$App.Name]
        $UI.LblInst.Text = "Installing..."
        $UI.LblInst.ForeColor = [System.Drawing.Color]::Blue
        [System.Windows.Forms.Application]::DoEvents()

        if ($App.DownloadMode -eq "office") {
            $xml = if ($App.Name -match "2024") { $Office2024XML } else { $Office365XML }
            $wd = Join-Path $ScriptPath "OfficeSetup"
            New-Item -Type Dir -Force $wd | Out-Null
            $xml | Out-File (Join-Path $wd "config.xml") -Enc UTF8

            if (-not (Test-Path (Join-Path $ScriptPath "setup.exe"))) {
                if ($exe) {
                    $cmd = "& `"$exe`" install --id Microsoft.OfficeDeploymentTool --silent --force --override `"/quiet /extract:$wd`""
                    Invoke-Expression $cmd | Out-Null
                    if (Test-Path (Join-Path $wd "setup.exe")) { Move-Item (Join-Path $wd "setup.exe") (Join-Path $ScriptPath "setup.exe") -Force }
                }
            }

            $setupExe = Join-Path $ScriptPath "setup.exe"
            if (Test-Path $setupExe) {
                Start-Process $setupExe ("/configure `"$wd\config.xml`"") -Wait
                $UI.LblInst.Text = "Done"
                $UI.LblInst.ForeColor = [System.Drawing.Color]::Green
            } else {
                $UI.LblInst.Text = "Setup Missing"
                $UI.LblInst.ForeColor = [System.Drawing.Color]::Red
            }
            continue
        }

        $li = $Global:UIControls[$App.Name].LocalInfo
        if (-not $li) { $li = Get-LocalInstallerInfo $App; $Global:UIControls[$App.Name].LocalInfo = $li }

        if ($App.DualArch -and $li.Exists) {
            try {
                if ($li.Path86) {
                    Log-Msg "Installing $($App.Name) x86 from local: $($li.Path86.Name)"
                    Start-Process $li.Path86.FullName $App.Args -Wait -WindowStyle Hidden
                }
                if ([Environment]::Is64BitOperatingSystem -and $li.Path64) {
                    Log-Msg "Installing $($App.Name) x64 from local: $($li.Path64.Name)"
                    Start-Process $li.Path64.FullName $App.Args -Wait -WindowStyle Hidden
                }
                $UI.LblInst.Text = "Done (Local)"
                $UI.LblInst.ForeColor = [System.Drawing.Color]::Green
            } catch {
                Log-Msg "Local install error ($($App.Name)): $_"
                $UI.LblInst.Text = "Error"
                $UI.LblInst.ForeColor = [System.Drawing.Color]::Red
            }
            continue
        }

        if (-not $App.DualArch -and $li -and $li.Path) {
            Log-Msg "Installing $($App.Name) from local: $($li.Path.Name)"
            try {
                Start-Process $li.Path.FullName $App.Args -Wait -WindowStyle Hidden
                $UI.LblInst.Text = "Done (Local)"
                $UI.LblInst.ForeColor = [System.Drawing.Color]::Green
            } catch {
                Log-Msg "Local install error ($($App.Name)): $_"
                $UI.LblInst.Text = "Error"
                $UI.LblInst.ForeColor = [System.Drawing.Color]::Red
            }
            continue
        }

        if ($App.DownloadMode -eq "localonly") {
            $UI.LblInst.Text = "Local File Missing"
            $UI.LblInst.ForeColor = [System.Drawing.Color]::Red
            continue
        }

        if (($App.Id -or $App.Id64 -or $App.Id86) -and $exe) {
            try {
                if ($App.DualArch -and $App.Id86 -and $App.Id64) {
                    Log-Msg "Installing $($App.Name) x86 via Winget..."
                    & $exe install --id $App.Id86 --silent --accept-package-agreements --accept-source-agreements --force | Out-Null
                    if ([Environment]::Is64BitOperatingSystem) {
                        Log-Msg "Installing $($App.Name) x64 via Winget..."
                        & $exe install --id $App.Id64 --silent --accept-package-agreements --accept-source-agreements --force | Out-Null
                    }
                    if ($LASTEXITCODE -eq 0) { $UI.LblInst.Text="Done (Winget)"; $UI.LblInst.ForeColor=[System.Drawing.Color]::Green }
                    else { $UI.LblInst.Text="Failed"; $UI.LblInst.ForeColor=[System.Drawing.Color]::Red }
                }
                elseif ($App.Id) {
                    Log-Msg "Installing $($App.Name) via Winget..."
                    $wArg = @("install","--id",$App.Id,"--silent","--accept-package-agreements","--accept-source-agreements","--force")
                    if ($App.Source) { $wArg += @("--source",$App.Source) }
                    if ($App.Args)   { $wArg += @("--override",$App.Args) }
                    & $exe @wArg | Out-Null
                    if ($LASTEXITCODE -eq 0) { $UI.LblInst.Text="Done (Winget)"; $UI.LblInst.ForeColor=[System.Drawing.Color]::Green }
                    else { $UI.LblInst.Text="Failed"; $UI.LblInst.ForeColor=[System.Drawing.Color]::Red }
                }
            } catch {
                Log-Msg "Winget install error ($($App.Name)): $_"
                $UI.LblInst.Text = "Error"
                $UI.LblInst.ForeColor = [System.Drawing.Color]::Red
            }
        }
    }

    Refresh-UI
    $btnInstall.Enabled = $true
    [System.Windows.Forms.MessageBox]::Show("Installation Sequence Complete.")
})

# -------------------------
# 8) Events + Init
# -------------------------
$btnRefresh.Add_Click({ Refresh-UI })
$btnCheckUpdate.Add_Click({ Check-LatestVersions })
$btnApplySys.Add_Click({ Apply-SystemConfiguration })

# Driver tab events
$btnScanDrivers.Add_Click({ Scan-Drivers })
$btnDownloadDrivers.Add_Click({ Download-SelectedDrivers })
$btnInstallDrivers.Add_Click({ Install-SelectedDrivers })

$btnSelectAll.Add_Click({
    foreach ($row in $dgvDrivers.Rows) { $row.Cells["Select"].Value = $true }
})

$btnSelectNone.Add_Click({
    foreach ($row in $dgvDrivers.Rows) { $row.Cells["Select"].Value = $false }
})

$btnSelectNeedsUpdate.Add_Click({
    foreach ($row in $dgvDrivers.Rows) {
        $status = $row.Cells["Status"].Value
        $row.Cells["Select"].Value = ($status -match "Update Available|Critical|Recommended")
    }
})

$btnSelectNotInstalled.Add_Click({
    foreach ($row in $dgvDrivers.Rows) {
        $status = $row.Cells["Status"].Value
        $row.Cells["Select"].Value = ($status -match "Not Installed|Not Found")
    }
})

$btnSelectInstalled.Add_Click({
    foreach ($row in $dgvDrivers.Rows) {
        $status = $row.Cells["Status"].Value
        $row.Cells["Select"].Value = ($status -match "Up to Date|Installed")
    }
})

# Tab change handler
$TabControl.Add_SelectedIndexChanged({
    if ($TabControl.SelectedTab -eq $TabDrivers) {
        if (-not $Global:HPSystemInfo) {
            Refresh-DriverSystemInfo
        }
    }
})

$Form.Add_Shown({
    Refresh-UI
    Refresh-DriverSystemInfo
})

[void]$Form.ShowDialog()

} catch {
    Write-Host "CRITICAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Read-Host "Press Enter..."
}