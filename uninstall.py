"""Uninstall Random Matching evaluation enforcement files.

By default this removes ProgramData files and the scheduled task while leaving
the application folder in place. Pass --remove-application to trigger the same
cleanup flow used by expiry.
"""

from __future__ import annotations

import argparse

from random_matching_core import Uninstaller


def main() -> int:
    """Program entry point."""
    parser = argparse.ArgumentParser(description="Uninstall Random Matching support files")
    parser.add_argument(
        "--remove-application",
        action="store_true",
        help="Also remove the Random Matching application folder",
    )
    args = parser.parse_args()
    return Uninstaller().uninstall(remove_application=args.remove_application)


if __name__ == "__main__":
    raise SystemExit(main())

