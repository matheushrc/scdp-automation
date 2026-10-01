"""Chrome profile discovery and keyboard-only terminal selection."""

from __future__ import annotations

import json
import os
import platform
import re
import select
import sys
from collections.abc import Callable, Iterator, Mapping, Sequence
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import TextIO

_PROFILE_DIRECTORY_PATTERN = re.compile(r"Profile \d+")


class ChromeProfileError(RuntimeError):
    """A Chrome profile could not be discovered or selected safely."""


class ProfileSelectionCancelled(ChromeProfileError):
    """The operator cancelled the profile menu."""


@dataclass(frozen=True, slots=True)
class ChromeProfile:
    """A profile directory and the non-secret metadata shown in the menu."""

    user_data_dir: Path
    directory_name: str
    display_name: str
    email: str

    @property
    def profile_path(self) -> Path:
        return self.user_data_dir / self.directory_name


def default_user_data_dir(
    platform_name: str | None = None,
    environ: Mapping[str, str] | None = None,
    home: Path | None = None,
) -> Path:
    """Return Chrome Stable's user-data directory for the current OS user."""
    system = platform_name or platform.system()
    environment = os.environ if environ is None else environ
    home_directory = Path.home() if home is None else home

    if system in {"Windows", "nt"}:
        local_app_data = environment.get("LOCALAPPDATA")
        base = (
            Path(local_app_data) if local_app_data else home_directory / "AppData/Local"
        )
        return base / "Google" / "Chrome" / "User Data"

    if system in {"Linux", "posix"}:
        chrome_config_home = environment.get("CHROME_CONFIG_HOME")
        if chrome_config_home:
            config_home = Path(chrome_config_home)
        else:
            xdg_config_home = environment.get("XDG_CONFIG_HOME")
            if xdg_config_home and Path(xdg_config_home).is_absolute():
                config_home = Path(xdg_config_home)
            else:
                config_home = home_directory / ".config"
        return config_home / "google-chrome"

    raise ChromeProfileError("Chrome profile discovery is not supported on this OS.")


def _valid_email(value: object) -> str:
    if not isinstance(value, str):
        return ""
    email = value.strip()
    if not email or any(character.isspace() for character in email):
        return ""
    if email.count("@") != 1:
        return ""
    local_part, domain_part = email.split("@", maxsplit=1)
    if not local_part or not domain_part:
        return ""
    if not all(character.isprintable() for character in email):
        return ""
    return email


def _safe_display_name(value: object, fallback: str) -> str:
    if not isinstance(value, str):
        return fallback
    cleaned = "".join(
        character for character in value if character.isprintable()
    ).strip()
    if not cleaned or _valid_email(cleaned):
        return fallback
    return cleaned


def _profile_sort_key(directory_name: str) -> tuple[int, int | str]:
    if directory_name == "Default":
        return (0, 0)
    suffix = directory_name.removeprefix("Profile ")
    return (1, int(suffix))


def discover_profiles(user_data_dir: Path) -> list[ChromeProfile]:
    """Read profile labels and account emails from Chrome's root Local State."""
    local_state_path = user_data_dir / "Local State"
    if local_state_path.is_symlink() or not local_state_path.is_file():
        raise ChromeProfileError("Chrome Local State is missing or unavailable.")
    try:
        local_state = json.loads(local_state_path.read_text(encoding="utf-8"))
    except OSError, UnicodeError, ValueError:
        raise ChromeProfileError("Chrome Local State is invalid.") from None

    if not isinstance(local_state, dict):
        raise ChromeProfileError("Chrome Local State has an invalid structure.")
    profile_state = local_state.get("profile", {})
    if not isinstance(profile_state, dict):
        raise ChromeProfileError("Chrome profile metadata has an invalid structure.")
    info_cache = profile_state.get("info_cache", {})
    if not isinstance(info_cache, dict):
        raise ChromeProfileError("Chrome profile metadata has an invalid structure.")

    try:
        directory_names = {
            path.name
            for path in user_data_dir.iterdir()
            if path.name == "Default" or _PROFILE_DIRECTORY_PATTERN.fullmatch(path.name)
            if path.is_dir() and not path.is_symlink()
        }
    except OSError:
        raise ChromeProfileError("Chrome user-data directory is unavailable.") from None

    profiles: list[ChromeProfile] = []
    for directory_name in sorted(directory_names, key=_profile_sort_key):
        metadata = info_cache.get(directory_name, {})
        if not isinstance(metadata, dict):
            metadata = {}
        display_name = _safe_display_name(metadata.get("name"), directory_name)
        email = _valid_email(metadata.get("user_name"))
        profiles.append(
            ChromeProfile(user_data_dir, directory_name, display_name, email)
        )

    if not profiles:
        raise ChromeProfileError("No usable Chrome profiles were found.")
    return profiles


@contextmanager
def _terminal_key_reader(stream: TextIO) -> Iterator[Callable[[], str]]:
    if os.name == "nt":
        import msvcrt

        def read_windows_key() -> str:
            character = msvcrt.getwch()
            if character in {"\x00", "\xe0"}:
                return {"H": "up", "P": "down"}.get(msvcrt.getwch(), "other")
            return {"\r": "enter", "\x1b": "escape"}.get(character, "other")

        yield read_windows_key
        return

    import termios
    import tty

    descriptor = stream.fileno()
    previous_settings = termios.tcgetattr(descriptor)
    try:
        tty.setraw(descriptor)

        def read_posix_key() -> str:
            character = stream.read(1)
            if character in {"\r", "\n"}:
                return "enter"
            if character != "\x1b":
                return "other"
            if not select.select([stream], [], [], 0.1)[0]:
                return "escape"
            if stream.read(1) != "[":
                return "escape"
            return {"A": "up", "B": "down"}.get(stream.read(1), "other")

        yield read_posix_key
    finally:
        termios.tcsetattr(descriptor, termios.TCSADRAIN, previous_settings)


def _profile_menu_label(profile: ChromeProfile) -> str:
    label = profile.display_name
    if profile.directory_name != profile.display_name:
        label = f"{label} [{profile.directory_name}]"
    if profile.email:
        label = f"{label} — {profile.email}"
    return label


def select_profile(
    profiles: Sequence[ChromeProfile],
    *,
    input_stream: TextIO | None = None,
    output_stream: TextIO | None = None,
    key_reader: Callable[[], str] | None = None,
) -> ChromeProfile:
    """Select one profile with Up/Down, Enter, or Escape using stdlib only."""
    input_stream = sys.stdin if input_stream is None else input_stream
    output_stream = sys.stdout if output_stream is None else output_stream
    if not profiles:
        raise ChromeProfileError("No usable Chrome profiles were found.")
    if not input_stream.isatty() or not output_stream.isatty():
        raise ChromeProfileError(
            "Profile selection needs an interactive terminal; run this command in a terminal."
        )

    def show_menu(index: int) -> None:
        output_stream.write("\x1b[2J\x1b[H")
        output_stream.write(
            "Selecione o perfil Chrome (↑/↓, Enter confirma, Esc cancela):\n\n"
        )
        for row, profile in enumerate(profiles):
            marker = ">" if row == index else " "
            output_stream.write(f"{marker} {_profile_menu_label(profile)}\n")
        output_stream.flush()

    def select(read_key: Callable[[], str]) -> ChromeProfile:
        index = 0
        while True:
            show_menu(index)
            key = read_key()
            if key == "up":
                index = (index - 1) % len(profiles)
            elif key == "down":
                index = (index + 1) % len(profiles)
            elif key == "enter":
                return profiles[index]
            elif key == "escape":
                raise ProfileSelectionCancelled("Chrome profile selection cancelled.")

    if key_reader is not None:
        return select(key_reader)
    with _terminal_key_reader(input_stream) as read_key:
        return select(read_key)
