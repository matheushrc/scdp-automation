"""Disposable CLI wrapper around the production Chrome profile copier."""

from __future__ import annotations

import argparse
from pathlib import Path

from scdp_automation.chrome_profile_setup import ProfileSetupError, copy_chrome_profile


def main() -> None:
    """Copy the requested profile into a new destination directory."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-user-data-dir", required=True, type=Path)
    parser.add_argument("--profile-directory", required=True)
    parser.add_argument("--destination-user-data-dir", required=True, type=Path)
    arguments = parser.parse_args()
    try:
        copy_chrome_profile(
            arguments.source_user_data_dir,
            arguments.profile_directory,
            arguments.destination_user_data_dir,
        )
    except (OSError, ProfileSetupError) as error:
        parser.error(str(error))


if __name__ == "__main__":
    main()
