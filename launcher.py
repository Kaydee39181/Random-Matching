"""Guarded launcher for all Random Matching applications.

Examples:
    python launcher.py interswitch-cleaning
    python launcher.py matching
    python launcher.py zenith-filtering
    python launcher.py --cmd python "Interswitch cleaning\\app.py"
"""

from __future__ import annotations

from random_matching_core import Launcher, build_launcher_parser


def main() -> int:
    """Parse arguments and launch the requested application."""
    parser = build_launcher_parser()
    args = parser.parse_args()
    command = args.cmd or []
    return Launcher().run(args.app, command)


if __name__ == "__main__":
    raise SystemExit(main())
