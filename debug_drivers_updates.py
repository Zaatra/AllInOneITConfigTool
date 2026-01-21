#!/usr/bin/env python3
"""Debug driver update matching (hybrid: HWID/INF, then name/category)."""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
from typing import Any, Iterable


def _run_powershell(script: str) -> str:
    if not shutil.which("powershell"):
        raise RuntimeError("powershell not found on PATH")
    result = subprocess.run(
        ["powershell", "-NoProfile", "-Command", script],
        capture_output=True,
        text=True,
        check=False,
        timeout=60,
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "powershell command failed")
    return result.stdout


def _as_list(value: Any) -> list[Any]:
    if value is None:
        return []
    if isinstance(value, list):
        return value
    return [value]


def _find_hpia_exe() -> str | None:
    candidates = [
        r"C:\Program Files\HP\HPIA\HPImageAssistant.exe",
        r"C:\Program Files (x86)\HP\HPIA\HPImageAssistant.exe",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def _run_hpia_report(hpia_path: str, report_dir: str) -> str:
    args = [
        hpia_path,
        "/Operation:Analyze",
        "/Category:All",
        "/Selection:All",
        "/Action:List",
        f"/ReportFolder:{report_dir}",
        "/Silent",
    ]
    result = subprocess.run(args, capture_output=True, text=True, check=False, timeout=120)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "HPIA scan failed")
    return report_dir


def _load_hpia_report(report_path: str) -> list[dict[str, Any]]:
    if os.path.isdir(report_path):
        candidates = [os.path.join(report_path, name) for name in os.listdir(report_path) if name.lower().endswith(".json")]
        if not candidates:
            return []
        report_path = max(candidates, key=os.path.getmtime)
    with open(report_path, "r", encoding="utf-8") as handle:
        data = json.load(handle)
    if isinstance(data, dict):
        data = data.get("HPIA", {}).get("Recommendations") or data.get("Recommendations") or data
    if isinstance(data, dict):
        data = [data]
    return [item for item in (data or []) if isinstance(item, dict)]


def _normalize_name(value: str) -> str:
    text = value.lower()
    text = text.replace("wi-fi", "wifi").replace("wi fi", "wifi")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return text.strip()


def _extract_pnp_ids(value: Any) -> set[str]:
    ids: set[str] = set()
    items = _as_list(value)
    for item in items:
        if not item:
            continue
        if isinstance(item, dict):
            for v in item.values():
                ids.update(_extract_pnp_ids(v))
            continue
        text = str(item)
        for match in re.findall(r"(PCI\\VEN_[0-9A-F]{4}&DEV_[0-9A-F]{4}[^\\s;]*)", text, re.IGNORECASE):
            ids.add(match.upper())
        for match in re.findall(r"(USB\\VID_[0-9A-F]{4}&PID_[0-9A-F]{4}[^\\s;]*)", text, re.IGNORECASE):
            ids.add(match.upper())
        for match in re.findall(r"(HDAUDIO\\FUNC_[0-9A-F]{2}[^\\s;]*)", text, re.IGNORECASE):
            ids.add(match.upper())
    return ids


def _extract_inf_names(value: Any) -> set[str]:
    infs: set[str] = set()
    items = _as_list(value)
    for item in items:
        if not item:
            continue
        if isinstance(item, dict):
            for v in item.values():
                infs.update(_extract_inf_names(v))
            continue
        text = str(item).lower()
        for match in re.findall(r"([a-z0-9_\-]+\.inf)\b", text):
            infs.add(match)
    return infs


def _get_field(item: dict[str, Any], *names: str) -> Any:
    for name in names:
        if name in item:
            return item[name]
    meta = item.get("Meta")
    if isinstance(meta, dict):
        for name in names:
            if name in meta:
                return meta[name]
    return None


def _build_search_terms(driver_name: str) -> list[str]:
    driver_lower = driver_name.lower()
    terms: list[str] = []
    if "intel" in driver_lower:
        terms.append("intel")
    if "realtek" in driver_lower:
        terms.append("realtek")
    if "nvidia" in driver_lower:
        terms.append("nvidia")
    if "amd" in driver_lower:
        terms.append("amd")
    if "bluetooth" in driver_lower:
        terms.append("bluetooth")
    if re.search(r"wireless|wlan|wifi|wi-fi", driver_lower):
        terms.extend(["wireless", "wlan", "wifi"])
    if re.search(r"graphics|video|display", driver_lower):
        terms.extend(["graphics", "video", "display"])
    if re.search(r"audio|sound", driver_lower):
        terms.extend(["audio", "sound"])
    if re.search(r"ethernet|nic|network", driver_lower):
        terms.extend(["ethernet", "network"])
    if "chipset" in driver_lower:
        terms.append("chipset")
    if re.search(r"storage|raid|rst|rapid", driver_lower):
        terms.extend(["storage", "rapid", "rst"])
    if "bios" in driver_lower:
        terms.append("bios")
    if "firmware" in driver_lower:
        terms.append("firmware")
    if re.search(r"management engine|me driver", driver_lower):
        terms.append("management engine")
    if "thunderbolt" in driver_lower:
        terms.append("thunderbolt")
    if re.search(r"serial io|serialio", driver_lower):
        terms.append("serial")
    if re.search(r"arc|a380|a770", driver_lower):
        terms.append("arc")
    if "usb 3" in driver_lower:
        terms.append("usb 3")
    return terms


def _name_score(driver_name: str, category: str | None, installed_name: str) -> int:
    driver_norm = _normalize_name(driver_name)
    installed_norm = _normalize_name(installed_name)
    if not driver_norm or not installed_norm:
        return 0
    if "manageability" in installed_norm and "manageability" not in driver_norm:
        if re.search(r"\b(wlan|wifi|wireless)\b", driver_norm):
            return 0
    score = 0
    for term in _build_search_terms(driver_name):
        if term in installed_norm:
            score += 1
    if category:
        cat_lower = category.lower()
        if "graphics" in cat_lower and re.search(r"graphics|display|video", installed_norm):
            score += 2
        if "audio" in cat_lower and re.search(r"audio|sound|realtek", installed_norm):
            score += 2
        if "network" in cat_lower and re.search(r"network|ethernet|wireless|wifi|bluetooth", installed_norm):
            score += 2
        if "chipset" in cat_lower and re.search(r"chipset|serial|management|usb", installed_norm):
            score += 2
        if "storage" in cat_lower and re.search(r"storage|rapid|rst|raid|optane", installed_norm):
            score += 2
        if re.search(r"bios|firmware", cat_lower) and re.search(r"bios|firmware", installed_norm):
            score += 2
    if driver_norm and driver_norm in installed_norm:
        score += 3
    return score


def _expected_classes(driver_name: str, category: str | None) -> set[str]:
    name = _normalize_name(driver_name)
    cat = (category or "").lower()
    classes: set[str] = set()
    if re.search(r"graphics|video|display", name) or "graphics" in cat or "display" in cat:
        classes.add("display")
    if re.search(r"audio|sound|realtek", name) or "audio" in cat:
        classes.add("media")
    if re.search(r"wireless|wlan|wifi|ethernet|network|bluetooth", name) or "network" in cat:
        classes.add("net")
        if "bluetooth" in name:
            classes.add("bluetooth")
    if re.search(r"bluetooth", name) or "bluetooth" in cat:
        classes.add("bluetooth")
        classes.add("net")
    if re.search(r"storage|raid|rst|rapid|ssd|nvme", name) or "storage" in cat:
        classes.update({"scsiadapter", "hdc", "diskdrive"})
    if re.search(r"chipset|serial|management engine|me driver|platform", name) or "chipset" in cat:
        classes.add("system")
    if re.search(r"firmware|bios", name) or re.search(r"bios|firmware", cat):
        classes.update({"firmware", "system"})
    return classes


def _is_generic_installed(name: str, cmsl_name: str, cmsl_category: str | None) -> bool:
    inst_norm = _normalize_name(name)
    cmsl_norm = _normalize_name(cmsl_name)
    if "microsoft" in inst_norm and "microsoft" not in cmsl_norm:
        return True
    if "wan miniport" in inst_norm:
        return True
    if "system management bios driver" in inst_norm and "bios" not in cmsl_norm and "firmware" not in cmsl_norm:
        return True
    if "storage spaces controller" in inst_norm:
        return True
    if "basic display adapter" in inst_norm and "display" not in cmsl_norm:
        return True
    if "display audio" in inst_norm and not re.search(r"\baudio\b", cmsl_norm):
        return True
    if "u03 system firmware" in inst_norm and not re.search(r"\bfirmware\b|\bbios\b", cmsl_norm):
        return True
    return False


def _match_driver(
    cmsl_item: dict[str, Any],
    installed: list[dict[str, Any]],
    *,
    min_name_score: int,
    allow_name_fallback: bool = False,
) -> tuple[dict[str, Any] | None, str, int, dict[str, Any] | None, int]:
    cmsl_name = str(_get_field(cmsl_item, "Name", "DeviceName") or "")
    cmsl_cat = str(_get_field(cmsl_item, "Category", "Class", "ClassName") or "")
    expected_classes = {c.lower() for c in _expected_classes(cmsl_name, cmsl_cat)}
    cmsl_pnp = _extract_pnp_ids(
        _get_field(cmsl_item, "HardwareID", "HardwareIds", "HWID", "DeviceID", "DeviceIds", "PnPIds", "SupportedDevices", "Devices")
    )
    cmsl_inf = _extract_inf_names(
        _get_field(cmsl_item, "InfName", "INF", "Inf", "InfFiles", "CVA", "Description", "Notes")
    )
    cmsl_has_ids = bool(cmsl_pnp or cmsl_inf)

    best: dict[str, Any] | None = None
    best_reason = "none"
    best_score = 0
    best_name: dict[str, Any] | None = None
    best_name_score = 0

    for inst in installed:
        inst_name = str(_get_field(inst, "DeviceName", "Name") or "")
        if _is_generic_installed(inst_name, cmsl_name, cmsl_cat):
            continue
        inst_class = str(_get_field(inst, "Class") or "").lower()
        inst_ids = _extract_pnp_ids(_get_field(inst, "HardwareID", "HardwareIds", "DeviceID"))
        inst_infs = _extract_inf_names(_get_field(inst, "InfName", "Inf"))

        id_score = 0
        if cmsl_pnp and inst_ids and (cmsl_pnp & inst_ids):
            id_score = 100

        inf_score = 0
        if cmsl_inf and inst_infs and (cmsl_inf & inst_infs):
            inf_score = 80

        name_score = _name_score(cmsl_name, cmsl_cat, inst_name)
        score = max(id_score, inf_score, name_score)
        reason = "name"
        if id_score:
            reason = "hwid"
        elif inf_score:
            reason = "inf"

        if reason == "name" and expected_classes and inst_class and inst_class not in expected_classes:
            continue

        if score > best_score:
            best_score = score
            best_reason = reason
            best = inst

        if name_score > best_name_score:
            best_name_score = name_score
            best_name = inst

    if best_reason == "name":
        if best_score < min_name_score:
            best = None
            best_reason = "no_hwid_match"
            best_score = 0
        elif not allow_name_fallback:
            best = None
            best_reason = "no_hwid_match"
            best_score = 0

    if best is None and best_reason == "none":
        best_reason = "cmsl_no_ids" if not cmsl_has_ids else "no_hwid_match"

    return best, best_reason, best_score, best_name, best_name_score


def _is_driver_cmsl_item(item: dict[str, Any]) -> bool:
    category = str(_get_field(item, "Category", "Class", "ClassName") or "")
    name = str(_get_field(item, "Name", "DeviceName") or "")
    text = f"{category} {name}".lower()
    if re.search(r"\bbios\b|\bfirmware\b", text):
        return True
    if "driver" in text:
        return True
    return False


def _version_tuple(value: str | None) -> tuple[int, ...] | None:
    if not value:
        return None
    text = str(value)
    match = re.search(r"(\d+(?:\.\d+){0,4})", text)
    if match:
        text = match.group(1)
    parts = text.split(".")
    out: list[int] = []
    for part in parts:
        if not part.isdigit():
            return None
        out.append(int(part))
    return tuple(out) if out else None


def _compare_versions(installed: str | None, available: str | None) -> str:
    inst = _version_tuple(installed)
    avail = _version_tuple(available)
    if not inst or not avail:
        return "unknown"
    if avail > inst:
        return "update_available"
    if avail == inst:
        return "up_to_date"
    return "installed"


def _load_json_file(path: str) -> Any:
    with open(path, "r", encoding="utf-8") as handle:
        return json.load(handle)


def _load_cmsl(args: argparse.Namespace) -> list[dict[str, Any]]:
    if args.cmsl_json:
        data = _load_json_file(args.cmsl_json)
    else:
        platform = args.platform or _detect_platform_id()
        if not platform:
            raise RuntimeError("Provide --platform or --cmsl-json (auto-detect failed)")
        script = (
            "Import-Module HPCMSL -ErrorAction Stop; "
            f"$sp = Get-SoftpaqList -Platform '{platform}' -Os {args.os} -OsVer {args.osver} -ErrorAction Stop; "
            "$sp | ConvertTo-Json -Depth 6"
        )
        data = json.loads(_run_powershell(script))
    if isinstance(data, dict):
        data = [data]
    return [item for item in data if isinstance(item, dict)]


def _meta_keys(items: Iterable[dict[str, Any]]) -> list[str]:
    keys: set[str] = set()
    for item in items:
        meta = item.get("Meta")
        if isinstance(meta, dict):
            keys.update(meta.keys())
    return sorted(keys)


def _fetch_cmsl_metadata(softpaq_ids: list[str]) -> dict[str, dict[str, Any]]:
    if not softpaq_ids:
        return {}
    ids_literal = ",".join(f"'{sid}'" for sid in softpaq_ids)
    script = rf"""
    Import-Module HPCMSL -ErrorAction Stop
    $ids = @({ids_literal})
    $metaCmd = Get-Command -Name Get-SoftpaqMetadata -ErrorAction SilentlyContinue
    if (-not $metaCmd) {{
        Write-Output "__NO_METADATA_CMDLET__"
        return
    }}
    $results = @()
    foreach ($id in $ids) {{
        try {{
            $m = Get-SoftpaqMetadata -Number $id -ErrorAction Stop
            if ($m) {{
                $results += [PSCustomObject]@{{ Id = $id; Meta = $m }}
            }}
        }} catch {{
        }}
    }}
    $results | ConvertTo-Json -Depth 8
    """
    output = _run_powershell(script).strip()
    if not output:
        return {}
    if "__NO_METADATA_CMDLET__" in output:
        raise RuntimeError("HPCMSL Get-SoftpaqMetadata cmdlet not available")
    try:
        data = json.loads(output)
    except json.JSONDecodeError:
        return {}
    if isinstance(data, dict):
        data = [data]
    mapping: dict[str, dict[str, Any]] = {}
    for item in data or []:
        if not isinstance(item, dict):
            continue
        sid = str(item.get("Id") or "").strip()
        meta = item.get("Meta")
        if sid and isinstance(meta, dict):
            mapping[sid] = meta
    return mapping


def _load_installed(args: argparse.Namespace) -> list[dict[str, Any]]:
    if args.installed_json:
        data = _load_json_file(args.installed_json)
    else:
        script = r"""
        $entities = @{}
        Get-WmiObject Win32_PnPEntity -ErrorAction SilentlyContinue | ForEach-Object {
            if ($_.DeviceID) { $entities[$_.DeviceID] = $_.ConfigManagerErrorCode }
        }
        Get-WmiObject Win32_PnPSignedDriver -ErrorAction SilentlyContinue |
        Select-Object DeviceName, DriverVersion, Manufacturer, DeviceID, HardwareID, InfName, ClassGuid, Class, DriverDate,
            @{n='ConfigManagerErrorCode';e={ $entities[$_.DeviceID] }} |
        ConvertTo-Json -Depth 4
        """
        data = json.loads(_run_powershell(script))
    if isinstance(data, dict):
        data = [data]
    return [item for item in data if isinstance(item, dict)]


def _detect_platform_id() -> str | None:
    script = r"""
    $cs = Get-WmiObject Win32_ComputerSystem
    $bb = Get-WmiObject Win32_BaseBoard
    $csProduct = Get-WmiObject Win32_ComputerSystemProduct
    $regPath = 'HKLM:\\HARDWARE\\DESCRIPTION\\System\\BIOS'
    $biosReg = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
    $result = $null
    if ($bb -and $bb.Product) { $result = $bb.Product }
    if (-not $result -and $biosReg -and $biosReg.SystemSKU) { $result = $biosReg.SystemSKU }
    if (-not $result -and $csProduct -and $csProduct.SKUNumber) { $result = $csProduct.SKUNumber }
    if (-not $result -and $cs -and $cs.Model) { $result = $cs.Model }
    if ($result) { $result }
    """
    try:
        output = _run_powershell(script).strip()
    except Exception:
        return None
    return output or None


def main() -> int:
    parser = argparse.ArgumentParser(description="Debug driver update matching (hybrid HWID/INF/name).")
    parser.add_argument("--cmsl-json", help="Path to CMSL JSON output (from Get-SoftpaqList)")
    parser.add_argument("--installed-json", help="Path to installed drivers JSON (from Win32_PnPSignedDriver)")
    parser.add_argument("--platform", help="HP platform ID for CMSL scan")
    parser.add_argument("--os", default="Win11", help="CMSL OS value (default: Win11)")
    parser.add_argument("--osver", default="24H2", help="CMSL OS version (default: 24H2)")
    parser.add_argument("--min-score", type=int, default=2, help="Minimum name score to accept a match")
    parser.add_argument("--show-unmatched", action="store_true", help="Show CMSL items with no match")
    parser.add_argument("--output-json", help="Write match results to JSON file")
    parser.add_argument("--dump-cmsl-keys", action="store_true", help="Print available CMSL keys and exit")
    parser.add_argument("--allow-name-fallback", action="store_true", help="Allow name-only matches when no HWID/INF match")
    parser.add_argument("--include-non-drivers", action="store_true", help="Include non-driver CMSL items")
    parser.add_argument("--list-installed", action="store_true", help="List installed drivers and exit")
    parser.add_argument("--enrich-cmsl", action="store_true", help="Fetch CMSL metadata (HWID/INF) via Get-SoftpaqMetadata")
    parser.add_argument("--enrich-limit", type=int, default=200, help="Max number of CMSL items to enrich (default: 200)")
    parser.add_argument("--dump-meta-keys", action="store_true", help="Print CMSL metadata keys and exit")
    parser.add_argument("--hpia-report", help="HPIA JSON report file or report folder")
    parser.add_argument("--hpia-run", action="store_true", help="Run HPIA analyze and load the latest report")
    parser.add_argument("--hpia-path", help="Path to HPImageAssistant.exe")
    parser.add_argument("--hpia-report-dir", default=os.path.join(os.getcwd(), "_hpia_report"), help="Report dir for --hpia-run")
    args = parser.parse_args()

    hpia_items: list[dict[str, Any]] = []
    if args.hpia_run:
        hpia_path = args.hpia_path or _find_hpia_exe()
        if not hpia_path:
            print("Error: HPImageAssistant.exe not found. Provide --hpia-path.", file=sys.stderr)
            return 1
        try:
            _run_hpia_report(hpia_path, args.hpia_report_dir)
            hpia_items = _load_hpia_report(args.hpia_report_dir)
        except Exception as exc:
            print(f"Error: HPIA run failed: {exc}", file=sys.stderr)
            return 1
    elif args.hpia_report:
        try:
            hpia_items = _load_hpia_report(args.hpia_report)
        except Exception as exc:
            print(f"Error: HPIA report load failed: {exc}", file=sys.stderr)
            return 1

    try:
        installed = _load_installed(args)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    if args.list_installed:
        for inst in installed:
            name = str(_get_field(inst, "DeviceName", "Name") or "")
            ver = _get_field(inst, "DriverVersion")
            cls = _get_field(inst, "Class")
            hwids = _extract_pnp_ids(_get_field(inst, "HardwareID", "HardwareIds", "DeviceID"))
            infs = _extract_inf_names(_get_field(inst, "InfName", "Inf"))
            print(f"{name} | {ver} | {cls} | hwid={len(hwids)} inf={len(infs)}")
        return 0

    try:
        cmsl_items = _load_cmsl(args)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    if args.dump_cmsl_keys:
        keys: set[str] = set()
        for item in cmsl_items:
            keys.update(item.keys())
        print("CMSL keys:", ", ".join(sorted(keys)))
        return 0

    if not args.include_non_drivers:
        cmsl_items = [item for item in cmsl_items if _is_driver_cmsl_item(item)]

    hpia_softpaqs: set[str] = set()
    hpia_names: set[str] = set()
    for item in hpia_items:
        sid = str(_get_field(item, "SoftPaqId", "Id", "Number") or "").strip()
        if sid:
            hpia_softpaqs.add(sid)
        name = str(_get_field(item, "Name") or "").strip()
        if name:
            hpia_names.add(_normalize_name(name))

    if hpia_items:
        print(f"HPIA recommendations: {len(hpia_items)}")

    if args.enrich_cmsl:
        ids: list[str] = []
        for item in cmsl_items:
            sid = str(_get_field(item, "Id", "SoftPaqId", "Number") or "").strip()
            if sid:
                ids.append(sid)
            if len(ids) >= args.enrich_limit:
                break
        try:
            meta_map = _fetch_cmsl_metadata(ids)
        except Exception as exc:
            print(f"Warning: CMSL metadata enrichment failed: {exc}", file=sys.stderr)
            meta_map = {}
        if meta_map:
            for item in cmsl_items:
                sid = str(_get_field(item, "Id", "SoftPaqId", "Number") or "").strip()
                if sid and sid in meta_map:
                    item["Meta"] = meta_map[sid]

        meta_count = sum(1 for item in cmsl_items if isinstance(item.get("Meta"), dict))
        id_count = 0
        for item in cmsl_items:
            cmsl_pnp = _extract_pnp_ids(
                _get_field(item, "HardwareID", "HardwareIds", "HWID", "DeviceID", "DeviceIds", "PnPIds", "SupportedDevices", "Devices")
            )
            cmsl_inf = _extract_inf_names(
                _get_field(item, "InfName", "INF", "Inf", "InfFiles", "CVA", "Description", "Notes")
            )
            if cmsl_pnp or cmsl_inf:
                id_count += 1
        print(f"CMSL items: {len(cmsl_items)} | Meta present: {meta_count} | With HWID/INF: {id_count}")

    if args.dump_meta_keys:
        keys = _meta_keys(cmsl_items)
        print("CMSL meta keys:", ", ".join(keys))
        return 0

    results: list[dict[str, Any]] = []
    for item in hpia_items:
        name = str(_get_field(item, "Name") or "")
        category = str(_get_field(item, "Category") or "")
        available = _get_field(item, "Version")
        installed_ver = _get_field(item, "CurrentVersion")
        softpaq_id = _get_field(item, "SoftPaqId")
        rec_value = _get_field(item, "RecommendationValue") or ""
        status = _compare_versions(installed_ver, available)
        results.append(
            {
                "source": "HPIA",
                "cmsl_name": name,
                "cmsl_category": category,
                "cmsl_version": available,
                "cmsl_id": softpaq_id,
                "match_name": name,
                "match_version": installed_ver,
                "match_reason": rec_value or status,
                "match_score": 0,
                "match_config_error": None,
                "missing_driver": None,
                "name_candidate": "",
                "name_candidate_version": None,
                "name_candidate_score": 0,
            }
        )
    for item in cmsl_items:
        name = str(_get_field(item, "Name", "DeviceName") or "")
        category = str(_get_field(item, "Category", "Class", "ClassName") or "")
        version = _get_field(item, "Version")
        softpaq_id = _get_field(item, "Id", "SoftPaqId", "Number")
        if softpaq_id and str(softpaq_id) in hpia_softpaqs:
            continue
        if _normalize_name(name) in hpia_names:
            continue
        match, reason, score, name_candidate, name_score = _match_driver(
            item,
            installed,
            min_name_score=args.min_score,
            allow_name_fallback=args.allow_name_fallback,
        )
        if match is None and not args.show_unmatched:
            continue
        match_name = str(_get_field(match or {}, "DeviceName", "Name") or "")
        match_version = _get_field(match or {}, "DriverVersion")
        match_code = _get_field(match or {}, "ConfigManagerErrorCode")
        missing_driver = None
        if isinstance(match_code, int):
            missing_driver = match_code == 28
        elif isinstance(match_code, str) and match_code.isdigit():
            missing_driver = int(match_code) == 28
        name_candidate_name = str(_get_field(name_candidate or {}, "DeviceName", "Name") or "")
        name_candidate_version = _get_field(name_candidate or {}, "DriverVersion")
        results.append(
            {
                "source": "CMSL",
                "cmsl_name": name,
                "cmsl_category": category,
                "cmsl_version": version,
                "cmsl_id": softpaq_id,
                "match_name": match_name,
                "match_version": match_version,
                "match_reason": reason,
                "match_score": score,
                "match_config_error": match_code,
                "missing_driver": missing_driver,
                "name_candidate": name_candidate_name,
                "name_candidate_version": name_candidate_version,
                "name_candidate_score": name_score,
            }
        )

    if args.output_json:
        with open(args.output_json, "w", encoding="utf-8") as handle:
            json.dump(results, handle, indent=2)
    else:
        for row in results:
            status = row["match_reason"]
            missing = row["missing_driver"]
            missing_text = "missing" if missing else ("ok" if missing is False else "unknown")
            candidate = ""
            if row["match_reason"] == "no_hwid_match" and row["name_candidate"]:
                candidate = f" | name-candidate: {row['name_candidate']} ({row['name_candidate_version']}) score={row['name_candidate_score']}"
            print(
                f"[{row['source']} {status}/{row['match_score']}] {missing_text} "
                f"{row['cmsl_name']} ({row['cmsl_version']}) -> {row['match_name']} ({row['match_version']})"
                f"{candidate}"
            )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
