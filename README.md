# All-In-One IT Configuration Tool (Python Port)

This repository contains the PySide6-based rewrite of the original `AllInOneITConfigTool.ps1` utility. It maintains the immutable deployment values (timezone, locale, WinGet IDs, Office XML templates) and exposes the same workflows across three tabs:

1. **Applications** – Download/install software via WinGet, Office ODT, direct installers, or local archives.
2. **Drivers** – Scan HP systems with HPIA/CMSL, then download/install SoftPaqs in batch.
3. **System Config** – Validate and apply the mandated timezone, locale, power plan, and registry tweaks.

## Running Locally

```bash
python -m venv .venv
source .venv/bin/activate  # or .venv\Scripts\activate on Windows
pip install -r requirements.txt
python main.py
```

PySide6 requires a desktop session. When launching through Remote Desktop or WSLg, ensure the `DISPLAY` (Linux) or GUI subsystem is present.

## Building the Standalone Package

PyInstaller can reproduce the binary bundle created during development:

```bash
python -m PyInstaller AllInOneITConfigTool.spec
```

- Linux hosts produce an ELF binary inside `dist/AllInOneITConfigTool/`.
- To create a Windows `.exe`, run the same command **on Windows** (PyInstaller does not cross-compile).
- The bundle already includes PySide6, `allinone_it_config/`, `services/`, and supporting modules.

## Manual Acceptance Tests

Use this checklist before releasing a build.

### 1. Immutable Data Verification

1. Start the GUI and open the **System Config** tab.
2. Confirm the Timezone target shows `West Bank Standard Time`, the Locale target shows `ar-SA / dd/MM/yyyy`, and the log mentions “All settings compliant” on a fully configured machine.
3. On the **Applications** tab, right-click any VC++ redistributable row and choose “Details” (if using a table inspector) to ensure the WinGet IDs match the legacy list. Alternatively, run the provided regression tests:
   ```bash
   python -m pytest tests/test_constants.py -k immutable
   ```

### 2. Application Install Flow

1. Check two standard apps (e.g., Chrome + Firefox) and click **Download Selected**. Verify the log displays `[OK] download :: …` entries.
2. Click **Install Selected** and confirm the log reports `[OK] install :: …`.
3. Re-run **Download Selected** to ensure caching logic reports “Installer already present”.

### 3. Driver Scan on Non-HP Hardware (Mock Scenario)

If you do not have an HP device, you can mock a scan using the service layer directly:

```bash
python - <<'PY'
from services.drivers import DriverService, HPSystemInfo
service = DriverService(
    system_info_provider=lambda: HPSystemInfo(
        platform_id='1234',
        model='HP Mock',
        supports_hpia=False,
        supports_cmsl=False,
    )
)
for record in service.scan():
    print(record)
PY
```

### 4. System Config Remediation

1. Temporarily change your timezone to something else (Control Panel → Date & Time).
2. In the app, open **System Config** – the Timezone row should turn red with `✗`.
3. Click **Apply All** and confirm Windows changes back to `West Bank Standard Time`.

### 5. Smoke Test the Packaged Binary

After running PyInstaller, execute the binary from a terminal to ensure it starts:

```bash
cd dist/AllInOneITConfigTool
./AllInOneITConfigTool  # Linux
AllInOneITConfigTool.exe  # Windows build
```

Use `QT_QPA_PLATFORM=offscreen` if running on a headless CI agent.

## Contributing / Support

- Run `python -m pytest` before submitting changes.
- Immutable values live in `allinone_it_config/constants.py` and must not change; tests will fail if they do.
- When adding new applications, use `allinone_it_config/app_registry.py` to preserve centralized metadata.
