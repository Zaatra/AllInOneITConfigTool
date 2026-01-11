"""Application registry mirroring `$Global:AppList` from the PowerShell script."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from trustpal.constants import IMMUTABLE_CONFIG


@dataclass(frozen=True)
class AppEntry:
    category: str
    name: str
    download_mode: str
    args: str = ""
    detection_pattern: str | None = None
    dual_arch: bool = False
    winget_id: str | None = None
    winget_id_x64: str | None = None
    winget_id_x86: str | None = None
    file_stem: str | None = None
    file_stem_x64: str | None = None
    file_stem_x86: str | None = None
    vc_key: str | None = None
    local_alt_names: Tuple[str, ...] = ()
    source: str | None = None


@dataclass(frozen=True)
class AppRegistry:
    entries: List[AppEntry]

    def by_category(self) -> Dict[str, List[AppEntry]]:
        grouped: Dict[str, List[AppEntry]] = {}
        for entry in self.entries:
            grouped.setdefault(entry.category, []).append(entry)
        return grouped


REGISTRY = AppRegistry(
    entries=[
        # VC++ Redistributables (dual-arch)
        AppEntry(
            category="VC++ Redistributables",
            name="VC++ 2005",
            download_mode="winget",
            args="/q",
            dual_arch=True,
            winget_id_x64="Microsoft.VCRedist.2005.x64",
            winget_id_x86="Microsoft.VCRedist.2005.x86",
            file_stem_x64="vcpp2005_x64",
            file_stem_x86="vcpp2005_x86",
            vc_key="2005",
        ),
        AppEntry(
            category="VC++ Redistributables",
            name="VC++ 2008",
            download_mode="winget",
            args="/q",
            dual_arch=True,
            winget_id_x64="Microsoft.VCRedist.2008.x64",
            winget_id_x86="Microsoft.VCRedist.2008.x86",
            file_stem_x64="vcpp2008_x64",
            file_stem_x86="vcpp2008_x86",
            vc_key="2008",
        ),
        AppEntry(
            category="VC++ Redistributables",
            name="VC++ 2010",
            download_mode="winget",
            args="/passive /norestart",
            dual_arch=True,
            winget_id_x64="Microsoft.VCRedist.2010.x64",
            winget_id_x86="Microsoft.VCRedist.2010.x86",
            file_stem_x64="vcpp2010_x64",
            file_stem_x86="vcpp2010_x86",
            vc_key="2010",
        ),
        AppEntry(
            category="VC++ Redistributables",
            name="VC++ 2012",
            download_mode="winget",
            args="/passive /norestart",
            dual_arch=True,
            winget_id_x64="Microsoft.VCRedist.2012.x64",
            winget_id_x86="Microsoft.VCRedist.2012.x86",
            file_stem_x64="vcpp2012_x64",
            file_stem_x86="vcpp2012_x86",
            vc_key="2012",
        ),
        AppEntry(
            category="VC++ Redistributables",
            name="VC++ 2013",
            download_mode="winget",
            args="/passive /norestart",
            dual_arch=True,
            winget_id_x64="Microsoft.VCRedist.2013.x64",
            winget_id_x86="Microsoft.VCRedist.2013.x86",
            file_stem_x64="vcpp2013_x64",
            file_stem_x86="vcpp2013_x86",
            vc_key="2013",
        ),
        AppEntry(
            category="VC++ Redistributables",
            name="VC++ 2015+",
            download_mode="winget",
            args="/passive /norestart",
            dual_arch=True,
            winget_id_x64="Microsoft.VCRedist.2015+.x64",
            winget_id_x86="Microsoft.VCRedist.2015+.x86",
            file_stem_x64="vcpp2015plus_x64",
            file_stem_x86="vcpp2015plus_x86",
            vc_key="2015+",
        ),
        # Core Applications
        AppEntry(
            category="Core Applications",
            name="Java 8",
            download_mode="winget",
            args="/s",
            winget_id="Oracle.JavaRuntimeEnvironment",
            detection_pattern="Java.*8 Update",
            file_stem="java8",
        ),
        AppEntry(
            category="Core Applications",
            name="Chrome",
            download_mode="winget",
            winget_id="Google.Chrome",
            detection_pattern="Google Chrome",
            file_stem="chrome",
        ),
        AppEntry(
            category="Core Applications",
            name="Firefox",
            download_mode="winget",
            winget_id="Mozilla.Firefox",
            detection_pattern="Mozilla Firefox",
            file_stem="firefox",
        ),
        AppEntry(
            category="Core Applications",
            name="WinRAR",
            download_mode="winget",
            winget_id="RARLab.WinRAR",
            detection_pattern="WinRAR",
            file_stem="winrar",
        ),
        AppEntry(
            category="Core Applications",
            name="TeamViewer",
            download_mode="winget",
            args="/qn",
            winget_id="TeamViewer.TeamViewer.Host",
            detection_pattern="TeamViewer",
            file_stem="teamviewer_host",
        ),
        AppEntry(
            category="Core Applications",
            name="NAPS2",
            download_mode="winget",
            winget_id="Cyanfish.NAPS2",
            detection_pattern="NAPS2",
            file_stem="naps2",
        ),
        AppEntry(
            category="Core Applications",
            name="K-Lite Mega",
            download_mode="winget",
            winget_id="CodecGuide.K-LiteCodecPack.Mega",
            detection_pattern="K-Lite Mega Codec",
            file_stem="klite_mega",
        ),
        AppEntry(
            category="Core Applications",
            name="Office 2024 LTSC",
            download_mode="office",
            winget_id="Microsoft.OfficeDeploymentTool",
            detection_pattern="Microsoft Office.*2024",
        ),
        AppEntry(
            category="Core Applications",
            name="Office 365 Ent",
            download_mode="office",
            winget_id="Microsoft.OfficeDeploymentTool",
            detection_pattern="Microsoft 365 Apps",
        ),
        # Drivers & Support Tools
        AppEntry(
            category="Driver & Support Tools",
            name="Intel DSA",
            download_mode="winget",
            winget_id="Intel.IntelDriverAndSupportAssistant",
            detection_pattern="Intel.*Driver.*Support",
            file_stem="intel_dsa",
        ),
        AppEntry(
            category="Driver & Support Tools",
            name="HP Support Asst",
            download_mode="onlineonly",
            winget_id="9WZDNCRDRBFB",
            detection_pattern="HP Support (Assistant|Solutions)",
            source="msstore",
        ),
        # Security & VMS
        AppEntry(
            category="Security & VMS",
            name="FortiClient VPN",
            download_mode="winget",
            winget_id="Fortinet.FortiClientVPN",
            detection_pattern="FortiClient VPN",
            file_stem="forticlient_vpn",
        ),
        AppEntry(
            category="Security & VMS",
            name="iVMS-4200",
            download_mode="direct",
            args="/silent /norestart",
            detection_pattern="iVMS-4200",
            file_stem="ivms4200",
        ),
        AppEntry(
            category="Security & VMS",
            name="CrowdStrike Falcon Sensor",
            download_mode="localonly",
            args=f"/install /quiet /norestart CID={IMMUTABLE_CONFIG.ids.crowdstrike_cid}",
            detection_pattern="CrowdStrike Falcon Sensor|CrowdStrike Windows Sensor",
            file_stem="crowdstrike_falcon_sensor",
            local_alt_names=(
                "WindowsSensor.exe",
                "FalconSensor.exe",
                "CrowdStrikeWindowsSensor.exe",
                "CSFalconSensor.exe",
                "CSFalconInstaller.exe",
            ),
        ),
    ]
)
