"""Persist the selected Chrome profile and maintain the project-local clone."""

from __future__ import annotations

import csv
import json
import os
import platform
import re
import shutil
import subprocess
import tempfile
import tomllib
import uuid
from collections.abc import Callable
from dataclasses import dataclass
from datetime import UTC, datetime
from io import StringIO
from pathlib import Path
from typing import TextIO

from scdp_automation.chrome_profiles import (
    _PROFILE_DIRECTORY_PATTERN,
    ChromeProfile,
    ChromeProfileError,
    _valid_email,
    default_user_data_dir,
    discover_profiles,
    select_profile,
)

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
_EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+$")


class ProfileSetupError(ChromeProfileError):
    """Profile configuration or clone preparation failed safely."""


@dataclass(frozen=True, slots=True)
class ProfileConfig:
    source_path: str
    email: str


def _valid_config_email(value: object) -> str | None:
    if value == "":
        return ""
    if not isinstance(value, str):
        return None
    normalized = _valid_email(value)
    if not normalized or not _EMAIL_PATTERN.fullmatch(normalized):
        return None
    return normalized


def load_profile_config(config_path: Path) -> ProfileConfig | None:
    """Load a valid local TOML selection; invalid or missing config means setup."""
    if config_path.is_symlink() or not config_path.is_file():
        return None
    try:
        with config_path.open("rb") as config_file:
            data = tomllib.load(config_file)
    except OSError, tomllib.TOMLDecodeError:
        return None

    if not isinstance(data, dict) or data.get("version") != 1:
        return None
    selection = data.get("chrome_profile")
    if not isinstance(selection, dict):
        return None
    source_path = selection.get("source_path")
    email = _valid_config_email(selection.get("email"))
    if not isinstance(source_path, str) or not source_path.strip() or email is None:
        return None
    if not Path(source_path).is_absolute():
        return None
    return ProfileConfig(source_path, email)


def _toml_string(value: str) -> str:
    """JSON basic strings use the compatible escapes needed by this TOML schema."""
    return json.dumps(value, ensure_ascii=False)


def write_profile_config(config_path: Path, config: ProfileConfig) -> None:
    """Atomically write the two local profile fields without adding dependencies."""
    normalized_email = _valid_config_email(config.email)
    if not Path(config.source_path).is_absolute() or normalized_email is None:
        raise ProfileSetupError("The selected Chrome profile config is invalid.")
    content = (
        "version = 1\n\n[chrome_profile]\n"
        f"source_path = {_toml_string(config.source_path)}\n"
        f"email = {_toml_string(normalized_email)}\n"
    )
    temporary_path: Path | None = None
    try:
        descriptor, temporary_name = tempfile.mkstemp(
            prefix=f".{config_path.name}.tmp-", dir=config_path.parent
        )
        temporary_path = Path(temporary_name)
        with os.fdopen(descriptor, "w", encoding="utf-8", newline="\n") as stream:
            stream.write(content)
            stream.flush()
            os.fsync(stream.fileno())
        if os.name != "nt":
            temporary_path.chmod(0o600)
        os.replace(temporary_path, config_path)
    except OSError:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)
        raise ProfileSetupError(
            "Could not save the local Chrome profile config."
        ) from None


def selected_profile_directory(clone_path: Path) -> str:
    """Read the copied profile marker, treating older clones as Default."""
    marker = clone_path / ".scdp-profile-directory"
    if marker.is_symlink():
        raise ProfileSetupError("The project Chrome profile marker is invalid.")
    if not marker.exists():
        return "Default"
    try:
        directory_name = marker.read_text(encoding="utf-8").strip()
    except OSError, UnicodeError:
        raise ProfileSetupError(
            "The project Chrome profile marker is invalid."
        ) from None
    if directory_name != "Default" and not _PROFILE_DIRECTORY_PATTERN.fullmatch(
        directory_name
    ):
        raise ProfileSetupError("The project Chrome profile marker is invalid.")
    return directory_name


def _clone_is_complete(clone_path: Path) -> bool:
    if clone_path.is_symlink() or not clone_path.is_dir():
        return False
    try:
        directory_name = selected_profile_directory(clone_path)
        if not (clone_path / directory_name).is_dir():
            return False
        local_state_path = clone_path / "Local State"
        if local_state_path.is_symlink() or not local_state_path.is_file():
            return False
        local_state = json.loads(local_state_path.read_text(encoding="utf-8"))
        return isinstance(local_state, dict)
    except OSError, UnicodeError, ValueError, ProfileSetupError:
        return False


def _load_filtered_local_state(
    local_state_path: Path, profile_directory: str
) -> dict[str, object]:
    if local_state_path.is_symlink():
        raise ProfileSetupError(
            "Chrome profile contains a symbolic link; copy refused."
        )
    try:
        local_state = json.loads(local_state_path.read_text(encoding="utf-8"))
    except OSError, UnicodeError, json.JSONDecodeError:
        raise ProfileSetupError("Chrome Local State is missing or invalid.") from None
    if not isinstance(local_state, dict):
        raise ProfileSetupError("Chrome Local State has an invalid structure.")

    profile_state = local_state.setdefault("profile", {})
    if not isinstance(profile_state, dict):
        raise ProfileSetupError("Chrome profile metadata has an invalid structure.")
    info_cache = profile_state.get("info_cache", {})
    if not isinstance(info_cache, dict):
        raise ProfileSetupError("Chrome profile metadata has an invalid structure.")
    selected_metadata = info_cache.get(profile_directory, {})
    if not isinstance(selected_metadata, dict):
        raise ProfileSetupError("Chrome profile metadata has an invalid structure.")
    profile_state["info_cache"] = (
        {profile_directory: selected_metadata} if selected_metadata else {}
    )
    profile_state["last_used"] = profile_directory
    profile_state["last_active_profiles"] = [profile_directory]
    return local_state


def _ignore_extension_directories(_directory: str, names: list[str]) -> set[str]:
    ignored_extensions = set(_EXTENSION_DIRECTORIES.intersection(names))
    for name in names:
        if name not in ignored_extensions and (Path(_directory) / name).is_symlink():
            raise ProfileSetupError(
                "Chrome profile contains a symbolic link; copy refused."
            )
    return ignored_extensions


def copy_chrome_profile(
    source_user_data_dir: Path,
    profile_directory: str,
    destination_user_data_dir: Path,
) -> None:
    """Copy one closed Chrome profile into a new user-data directory."""
    if profile_directory != "Default" and not _PROFILE_DIRECTORY_PATTERN.fullmatch(
        profile_directory
    ):
        raise ProfileSetupError("Choose a Chrome profile named Default or Profile N.")
    source = source_user_data_dir
    destination = destination_user_data_dir
    if not source.is_dir():
        raise ProfileSetupError("Chrome user-data directory is unavailable.")
    source_resolved = source.resolve()
    destination_resolved = destination.resolve()
    if source_resolved == destination_resolved or source_resolved in (
        destination_resolved,
        *destination_resolved.parents,
    ):
        raise ProfileSetupError("The clone destination cannot be inside the source.")
    if destination.exists() or destination.is_symlink():
        raise FileExistsError("The clone destination already exists.")
    if not destination.parent.is_dir():
        raise ProfileSetupError("The clone destination parent is unavailable.")

    source_lock = source / "SingletonLock"
    if source_lock.exists() or source_lock.is_symlink():
        raise ProfileSetupError("Close Chrome before copying its profile.")
    profile_source = source / profile_directory
    if profile_source.is_symlink():
        raise ProfileSetupError(
            "Chrome profile contains a symbolic link; copy refused."
        )
    if not profile_source.is_dir():
        raise ProfileSetupError("The selected Chrome profile is unavailable.")
    local_state = _load_filtered_local_state(source / "Local State", profile_directory)

    try:
        staging = Path(
            tempfile.mkdtemp(
                prefix=f".{destination.name}.staging-", dir=destination.parent
            )
        )
    except OSError:
        raise ProfileSetupError("Could not prepare the Chrome profile copy.") from None

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
            raise FileExistsError("The clone destination already exists.")
        os.rename(staging, destination)
    except FileExistsError:
        raise
    except OSError:
        raise ProfileSetupError("Could not finish the Chrome profile copy.") from None
    finally:
        if staging.exists():
            shutil.rmtree(staging, ignore_errors=True)


def chrome_process_is_running(
    platform_name: str | None = None,
    *,
    proc_root: Path = Path("/proc"),
) -> bool:
    """Check whether Chrome is running before copying any Chrome profile."""
    system = platform_name or platform.system()
    if system in {"Windows", "nt"}:
        try:
            result = subprocess.run(
                ["tasklist", "/FO", "CSV", "/NH"],
                capture_output=True,
                check=False,
                text=True,
                timeout=10,
            )
        except OSError, subprocess.TimeoutExpired:
            raise ProfileSetupError(
                "Could not check whether Chrome is running."
            ) from None
        if result.returncode != 0:
            raise ProfileSetupError("Could not check whether Chrome is running.")
        return any(
            row and row[0].casefold() == "chrome.exe"
            for row in csv.reader(StringIO(result.stdout))
        )

    if system in {"Linux", "posix"}:
        if not proc_root.is_dir():
            raise ProfileSetupError("Could not check whether Chrome is running.")
        process_names = {"chrome", "chromium", "google-chrome", "google-chrome-stable"}
        for command_file in proc_root.glob("*/comm"):
            try:
                process_name = (
                    command_file.read_text(encoding="utf-8").strip().casefold()
                )
            except OSError, UnicodeError:
                continue
            if process_name in process_names:
                return True
        return False

    raise ProfileSetupError("Chrome process checks are not supported on this OS.")


def _source_profile_from_config(config: ProfileConfig) -> ChromeProfile | None:
    profile_path = Path(config.source_path)
    if profile_path.is_symlink() or not profile_path.is_dir():
        return None
    try:
        profiles = discover_profiles(profile_path.parent)
    except ChromeProfileError:
        return None
    return next(
        (profile for profile in profiles if profile.profile_path == profile_path), None
    )


def _unique_backup_path(clone_path: Path) -> Path:
    timestamp = datetime.now(UTC).strftime("%Y%m%d-%H%M%S-%f")
    candidate = clone_path.with_name(f"{clone_path.name}.backup-{timestamp}")
    suffix = 1
    while candidate.exists() or candidate.is_symlink():
        candidate = clone_path.with_name(
            f"{clone_path.name}.backup-{timestamp}-{suffix}"
        )
        suffix += 1
    return candidate


def _remove_path(path: Path) -> None:
    if path.is_symlink() or path.is_file():
        path.unlink(missing_ok=True)
    elif path.is_dir():
        shutil.rmtree(path)


def _install_profile(
    profile: ChromeProfile,
    repo_root: Path,
    config_path: Path,
    clone_path: Path,
    process_checker: Callable[[], bool],
) -> Path:
    if process_checker():
        raise ProfileSetupError("Feche todas as janelas do Chrome antes da cópia.")

    staged_clone = repo_root / f".{clone_path.name}.candidate-{uuid.uuid4().hex}"
    try:
        copy_chrome_profile(profile.user_data_dir, profile.directory_name, staged_clone)
    except FileExistsError, ProfileSetupError:
        raise
    except OSError:
        raise ProfileSetupError(
            "Could not stage the selected Chrome profile."
        ) from None

    backup_path: Path | None = None
    try:
        if clone_path.exists() or clone_path.is_symlink():
            backup_path = _unique_backup_path(clone_path)
            os.replace(clone_path, backup_path)
        os.replace(staged_clone, clone_path)
        config = ProfileConfig(str(profile.profile_path.resolve()), profile.email)
        write_profile_config(config_path, config)
    except OSError, ProfileSetupError:
        try:
            if clone_path.exists() or clone_path.is_symlink():
                _remove_path(clone_path)
            if backup_path is not None and (
                backup_path.exists() or backup_path.is_symlink()
            ):
                os.replace(backup_path, clone_path)
        except OSError:
            raise ProfileSetupError(
                "Profile setup failed and the previous clone could not be restored."
            ) from None
        raise ProfileSetupError(
            "Could not save profile setup; the previous clone and config were restored."
        ) from None
    finally:
        if staged_clone.exists() or staged_clone.is_symlink():
            _remove_path(staged_clone)
    return clone_path


def prepare_chrome_profile(
    repo_root: Path,
    *,
    force_reselect: bool = False,
    user_data_dir: Path | None = None,
    input_stream: TextIO | None = None,
    output_stream: TextIO | None = None,
    key_reader: Callable[[], str] | None = None,
    process_checker: Callable[[], bool] | None = None,
) -> Path:
    """Reuse a configured clone or interactively create/reselect one."""
    repo_root = Path(repo_root)
    clone_path = repo_root / ".scdp-browser"
    config_path = repo_root / ".scdp-config.toml"
    config = load_profile_config(config_path)

    if not force_reselect and config is not None:
        if _clone_is_complete(clone_path):
            return clone_path
        source_profile = _source_profile_from_config(config)
        if source_profile is not None:
            return _install_profile(
                source_profile,
                repo_root,
                config_path,
                clone_path,
                process_checker or chrome_process_is_running,
            )

    source_directory = user_data_dir or default_user_data_dir()
    profiles = discover_profiles(source_directory)
    selected = select_profile(
        profiles,
        input_stream=input_stream,
        output_stream=output_stream,
        key_reader=key_reader,
    )
    return _install_profile(
        selected,
        repo_root,
        config_path,
        clone_path,
        process_checker or chrome_process_is_running,
    )
