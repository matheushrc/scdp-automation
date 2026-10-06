"""Synthetic Chrome profiles and interactive streams for tests."""

import io
import json
from pathlib import Path

EXTENSION_DIRECTORIES = (
    "Extensions",
    "Extension Cookies",
    "Extension Rules",
    "Extension Scripts",
    "Extension State",
    "Local Extension Settings",
    "Managed Extension Settings",
    "Sync Extension Settings",
)


class InteractiveStream(io.StringIO):
    def isatty(self) -> bool:
        return True


def create_user_data(user_data_dir: Path) -> None:
    for directory in ("Default", "Profile 2"):
        profile = user_data_dir / directory
        profile.mkdir(parents=True)
        (profile / "Preferences").write_text(directory, encoding="utf-8")
    extension_path = user_data_dir / "Profile 2" / "Extensions"
    extension_path.mkdir()
    (extension_path / "payload").write_text("extension", encoding="utf-8")
    state = {
        "os_crypt": {"encrypted_key": "synthetic"},
        "profile": {
            "info_cache": {
                "Default": {"name": "Default", "user_name": ""},
                "Profile 2": {"name": "Work", "user_name": "chief@example.com"},
            },
            "last_active_profiles": ["Default", "Profile 2"],
            "last_used": "Default",
        },
    }
    (user_data_dir / "Local State").write_text(json.dumps(state), encoding="utf-8")


def create_source(user_data_dir: Path) -> None:
    for profile_directory, contents in (
        ("Default", "default profile"),
        ("Profile 2", "second profile"),
    ):
        profile = user_data_dir / profile_directory
        profile.mkdir(parents=True)
        (profile / "Preferences").write_text(contents, encoding="utf-8")

    (user_data_dir / "Default" / "Bookmarks").write_text(
        "synthetic bookmark data", encoding="utf-8"
    )
    for directory in EXTENSION_DIRECTORIES:
        extension_directory = user_data_dir / "Default" / directory
        extension_directory.mkdir()
        (extension_directory / "payload").write_text(
            "synthetic extension data", encoding="utf-8"
        )

    local_state = {
        "browser": {"unrelated": "preserved"},
        "os_crypt": {"encrypted_key": "synthetic-encrypted-key"},
        "profile": {
            "info_cache": {
                "Default": {"name": "Default profile", "user_name": ""},
                "Profile 2": {"name": "Second profile", "user_name": ""},
            },
            "last_active_profiles": ["Default", "Profile 2"],
            "last_used": "Default",
        },
    }
    (user_data_dir / "Local State").write_text(
        json.dumps(local_state), encoding="utf-8"
    )


def snapshot_tree(root: Path) -> dict[str, bytes | str]:
    snapshot: dict[str, bytes | str] = {}
    for path in (root, *root.rglob("*")):
        relative_path = "." if path == root else path.relative_to(root).as_posix()
        if path.is_dir():
            snapshot[relative_path] = "directory"
        else:
            snapshot[relative_path] = path.read_bytes()
    return snapshot
