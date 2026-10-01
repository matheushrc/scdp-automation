"""Disposable profile-copy prototype for the validation worktree."""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import tempfile
from pathlib import Path

_PROFILE_DIRECTORY_PATTERN = re.compile(r"Profile \d+")
_EXTENSION_DIRECTORIES = frozenset(
    {
        "Extensions",
        "Extension Cookies",
        "Extension Rules",
        "Extension Scripts",
        "Extension State",
        "Local Extension Settings",
        "Managed Extension Settings",
        "Sync Extension Settings",
    }
)


def _is_supported_profile_directory(profile_directory: str) -> bool:
    return profile_directory == "Default" or bool(
        _PROFILE_DIRECTORY_PATTERN.fullmatch(profile_directory)
    )


def _load_filtered_local_state(
    local_state_path: Path, profile_directory: str
) -> dict[str, object]:
    if local_state_path.is_symlink():
        raise ValueError("Chrome profile contains a symbolic link; copy refused.")
    try:
        local_state = json.loads(local_state_path.read_text(encoding="utf-8"))
    except OSError, UnicodeError, json.JSONDecodeError:
        raise ValueError("Chrome Local State is missing or invalid.") from None

    if not isinstance(local_state, dict):
        raise TypeError("Chrome Local State has an invalid structure.")
    profile_state = local_state.get("profile")
    if not isinstance(profile_state, dict):
        raise TypeError("Chrome profile metadata has an invalid structure.")
    info_cache = profile_state.get("info_cache")
    if not isinstance(info_cache, dict):
        raise TypeError("Chrome profile metadata has an invalid structure.")
    selected_profile = info_cache.get(profile_directory)
    if selected_profile is None:
        raise ValueError("Selected profile metadata is missing or invalid.")
    if not isinstance(selected_profile, dict):
        raise TypeError("Selected profile metadata has an invalid structure.")

    profile_state["info_cache"] = {profile_directory: selected_profile}
    profile_state["last_used"] = profile_directory
    profile_state["last_active_profiles"] = [profile_directory]
    return local_state


def _ignore_extension_directories(_directory: str, names: list[str]) -> set[str]:
    ignored_extensions = set(_EXTENSION_DIRECTORIES.intersection(names))
    for name in names:
        if name not in ignored_extensions and (Path(_directory) / name).is_symlink():
            raise ValueError("Chrome profile contains a symbolic link; copy refused.")
    return ignored_extensions


def copy_chrome_profile(
    source_user_data_dir: Path,
    profile_directory: str,
    destination_user_data_dir: Path,
) -> None:
    """Copy one Chrome profile into a new, isolated user-data directory."""
    if not _is_supported_profile_directory(profile_directory):
        raise ValueError(
            "Choose a Chrome profile directory named Default or Profile N."
        )

    source = source_user_data_dir
    destination = destination_user_data_dir
    if not source.is_dir():
        raise ValueError("Chrome user-data directory is unavailable.")

    source_resolved = source.resolve()
    destination_resolved = destination.resolve()
    if source_resolved == destination_resolved or source_resolved in (
        destination_resolved,
        *destination_resolved.parents,
    ):
        raise ValueError("The destination must be outside the source profile.")
    if destination.exists() or destination.is_symlink():
        raise FileExistsError("Destination already exists; it will not be overwritten.")
    if not destination.parent.is_dir():
        raise ValueError("Destination parent directory is unavailable.")

    source_lock = source / "SingletonLock"
    if source_lock.exists() or source_lock.is_symlink():
        raise RuntimeError("Close Chrome before copying its profile.")

    profile_source = source / profile_directory
    if profile_source.is_symlink():
        raise ValueError("Chrome profile contains a symbolic link; copy refused.")
    if not profile_source.is_dir():
        raise ValueError("Selected Chrome profile directory is unavailable.")

    local_state = _load_filtered_local_state(source / "Local State", profile_directory)

    try:
        staging = Path(
            tempfile.mkdtemp(
                prefix=f".{destination.name}.staging-", dir=destination.parent
            )
        )
    except OSError:
        raise RuntimeError("Could not prepare the profile copy.") from None

    try:
        shutil.copytree(
            profile_source,
            staging / profile_directory,
            ignore=_ignore_extension_directories,
        )
        local_state_copy = staging / "Local State"
        local_state_copy.write_text(
            json.dumps(local_state, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        local_state_copy.chmod(0o600)
        marker = staging / ".scdp-profile-directory"
        marker.write_text(f"{profile_directory}\n", encoding="utf-8")
        marker.chmod(0o600)
        if destination.exists() or destination.is_symlink():
            raise FileExistsError(
                "Destination already exists; it will not be overwritten."
            )
        os.rename(staging, destination)
    except FileExistsError:
        raise
    except OSError:
        raise RuntimeError("Could not finish the profile copy.") from None
    finally:
        if staging.exists():
            shutil.rmtree(staging, ignore_errors=True)


def main() -> None:
    """Copy the requested Chrome profile into a new destination directory."""
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
    except (OSError, RuntimeError, TypeError, ValueError) as error:
        parser.error(str(error))


if __name__ == "__main__":
    main()
