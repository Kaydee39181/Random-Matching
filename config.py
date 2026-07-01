"""Central configuration for the Random Matching evaluation installer.

Change values here to adjust installer behavior without editing the installer,
launcher, cleanup, or uninstaller logic.
"""

from __future__ import annotations

from pathlib import Path


SOFTWARE_NAME: str = "Random Matching"
SOFTWARE_VERSION: str = "1.0.0"

# Set this to 2 (or any small value) during test builds.
EXPIRY_DAYS: int = 30

PROGRAMDATA_PATH: Path = Path(r"C:\ProgramData\RandomMatching")
TASK_NAME: str = "RandomMatchingExpiry"

LOG_DIRECTORY: Path = PROGRAMDATA_PATH / "logs"
INSTALL_FILE: Path = PROGRAMDATA_PATH / "install.json"
CLEANUP_FILE: Path = PROGRAMDATA_PATH / "cleanup.bat"
TASK_XML_FILE: Path = PROGRAMDATA_PATH / "RandomMatchingExpiry.xml"

ENABLE_TAMPER_CHECK: bool = True
ENABLE_DRY_RUN: bool = False
BEGIN_CLEANUP_ON_TAMPER: bool = False

# Add app-specific executable names here if future packaging creates custom EXEs.
CUSTOM_PROCESS_NAMES: tuple[str, ...] = (
    "RandomMatching.exe",
    "InterswitchCleaning.exe",
    "TransactionMatching.exe",
    "ZenithStatementFiltering.exe",
)

# Default launcher aliases. The launcher also accepts explicit commands with --.
APPLICATIONS: dict[str, dict[str, object]] = {
    "interswitch-cleaning": {
        "cwd": "Interswitch cleaning",
        "command": ["python", "app.py"],
    },
    "matching": {
        "cwd": "Interswitch, Cash234 and Zenith Matching",
        "command": ["cmd", "/c", "run_app.bat"],
    },
    "zenith-filtering": {
        "cwd": "Zenith Account Statement Filtering",
        "command": ["cmd", "/c", "npm run preview"],
    },
}

