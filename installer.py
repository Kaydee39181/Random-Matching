"""Install Random Matching evaluation expiry enforcement.

Run this file from the Random Matching folder. It detects the bundle path,
creates ProgramData metadata and cleanup files, and installs a Windows scheduled
task that removes the bundle after the configured evaluation period.
"""

from __future__ import annotations

from random_matching_core import Installer


def main() -> int:
    """Program entry point."""
    return Installer().install()


if __name__ == "__main__":
    raise SystemExit(main())
