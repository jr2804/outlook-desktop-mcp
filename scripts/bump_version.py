"""Bump the CalVer `VERSION` file forward based on today's date.

Run by `.github/workflows/version-bump.yml` on every push to `main`. The
strategy is:

- If `VERSION` already starts with the current `YYYY.M` prefix, bump the
  patch (the third component, `N`) by one.
- Otherwise, write the current `YYYY.M.0`.

This keeps the project's released CalVer aligned with calendar months and
ensures every push to `main` produces a fresh, monotonic patch number that
the build system can pick up via `uv-dynamic-versioning`.

CLI:
    python scripts/bump_version.py            # print the new version
    python scripts/bump_version.py --write    # also write it to VERSION

Exits non-zero if the input file is malformed or the new version doesn't
outrank the existing one.
"""

from __future__ import annotations

import argparse
import datetime as dt
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
VERSION_FILE = ROOT / "VERSION"

# CalVer regex: 2026.8.0
_CALVER_RE = re.compile(r"^(\d{4})\.(\d{1,2})\.(\d+)$")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--write",
        action="store_true",
        help="Write the new version back to VERSION (default: print to stdout only).",
    )
    parser.add_argument(
        "--version-file",
        type=Path,
        default=VERSION_FILE,
        help="Path to the VERSION file (default: repo root).",
    )
    args = parser.parse_args(argv)

    current = args.version_file.read_text(encoding="utf-8")
    new = next_version(dt.date.today(), current)
    if new <= current:
        # Should not happen with monotonic inputs, but guard anyway.
        raise SystemExit(f"Refusing to bump: {current!r} -> {new!r} (not strictly greater)")

    if args.write:
        args.version_file.write_text(new + "\n", encoding="utf-8")
    print(new)
    return 0


def next_version(today: dt.date, current: str) -> str:
    """Compute the next CalVer string.

    If the current version is already in the current calendar month, bump
    only the patch number. Otherwise start a fresh `YYYY.M.0`.
    """
    year, month, _ = _parse(current)
    if (year, month) == (today.year, today.month):
        _, _, patch = _parse(current)
        return f"{today.year}.{today.month}.{patch + 1}"
    return f"{today.year}.{today.month}.0"


def _parse(value: str) -> tuple[int, int, int]:
    match = _CALVER_RE.match(value.strip())
    if not match:
        raise SystemExit(f"VERSION file contains a non-CalVer value: {value!r}")
    year, month, patch = match.groups()
    return int(year), int(month), int(patch)


if __name__ == "__main__":
    sys.exit(main())
