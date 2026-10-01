import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from scripts.experiments.copy_chrome_profile import copy_chrome_profile

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


class ChromeProfileCopyExperimentTests(unittest.TestCase):
    def test_cli_copies_the_requested_profile_from_synthetic_source(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "source"
            destination = root / "clone"
            create_source(source)
            script = (
                Path(__file__).resolve().parents[1]
                / "scripts"
                / "experiments"
                / "copy_chrome_profile.py"
            )

            result = subprocess.run(
                [
                    sys.executable,
                    str(script),
                    "--source-user-data-dir",
                    str(source),
                    "--profile-directory",
                    "Profile 2",
                    "--destination-user-data-dir",
                    str(destination),
                ],
                capture_output=True,
                check=False,
                text=True,
            )

            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(
                (destination / ".scdp-profile-directory").read_text(encoding="utf-8"),
                "Profile 2\n",
            )

    def test_copy_keeps_only_selected_profile_and_updates_last_used_metadata(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "source"
            destination = root / "clone"
            create_source(source)

            copy_chrome_profile(source, "Profile 2", destination)

            self.assertEqual(
                {path.name for path in destination.iterdir()},
                {".scdp-profile-directory", "Local State", "Profile 2"},
            )
            self.assertFalse((destination / "Default").exists())
            self.assertEqual(
                (destination / ".scdp-profile-directory").read_text(encoding="utf-8"),
                "Profile 2\n",
            )
            copied_state = json.loads(
                (destination / "Local State").read_text(encoding="utf-8")
            )
            self.assertEqual(
                copied_state["profile"]["info_cache"],
                {"Profile 2": {"name": "Second profile", "user_name": ""}},
            )
            self.assertEqual(copied_state["profile"]["last_used"], "Profile 2")
            self.assertEqual(
                copied_state["profile"]["last_active_profiles"], ["Profile 2"]
            )

    def test_copy_preserves_unrelated_local_state_encryption_values(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "source"
            destination = root / "clone"
            create_source(source)

            copy_chrome_profile(source, "Default", destination)

            copied_state = json.loads(
                (destination / "Local State").read_text(encoding="utf-8")
            )
            self.assertEqual(
                copied_state["os_crypt"],
                {"encrypted_key": "synthetic-encrypted-key"},
            )
            self.assertEqual(copied_state["browser"], {"unrelated": "preserved"})

    def test_copy_excludes_extension_payload_and_storage_directories(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "source"
            destination = root / "clone"
            create_source(source)

            copy_chrome_profile(source, "Default", destination)

            copied_profile = destination / "Default"
            self.assertEqual(
                (copied_profile / "Bookmarks").read_text(encoding="utf-8"),
                "synthetic bookmark data",
            )
            for directory in EXTENSION_DIRECTORIES:
                with self.subTest(directory=directory):
                    self.assertFalse((copied_profile / directory).exists())

    def test_copy_does_not_modify_source_files(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "source"
            destination = root / "clone"
            create_source(source)
            before = snapshot_tree(source)

            copy_chrome_profile(source, "Default", destination)

            self.assertEqual(snapshot_tree(source), before)

    def test_copy_refuses_source_with_singleton_lock_without_removing_it(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "source"
            destination = root / "clone"
            create_source(source)
            lock = source / "SingletonLock"
            lock.write_text("synthetic lock", encoding="utf-8")

            with self.assertRaises(RuntimeError):
                copy_chrome_profile(source, "Default", destination)

            self.assertEqual(lock.read_text(encoding="utf-8"), "synthetic lock")
            self.assertFalse(destination.exists())

    def test_copy_refuses_existing_destination_without_overwriting_it(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "source"
            destination = root / "clone"
            create_source(source)
            destination.mkdir()
            sentinel = destination / "keep.txt"
            sentinel.write_text("do not replace", encoding="utf-8")

            with self.assertRaises(FileExistsError):
                copy_chrome_profile(source, "Default", destination)

            self.assertEqual(sentinel.read_text(encoding="utf-8"), "do not replace")
            self.assertEqual(list(destination.iterdir()), [sentinel])

    def test_copy_rejects_missing_or_malformed_local_state_without_creating_destination(
        self,
    ) -> None:
        for state_contents in (None, "not-json"):
            with (
                self.subTest(state_contents=state_contents),
                tempfile.TemporaryDirectory() as temporary_directory,
            ):
                root = Path(temporary_directory)
                source = root / "source"
                source.mkdir()
                (source / "Default").mkdir()
                if state_contents is not None:
                    (source / "Local State").write_text(
                        state_contents, encoding="utf-8"
                    )
                destination = root / "clone"

                with self.assertRaises(ValueError):
                    copy_chrome_profile(source, "Default", destination)

                self.assertFalse(destination.exists())

    def test_copy_rejects_profile_directory_path_traversal(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "source"
            create_source(source)

            for profile_directory in (
                "../Default",
                "Profile 2/../Default",
                str(root / "outside"),
            ):
                with self.subTest(profile_directory=profile_directory):
                    destination = root / "clone"
                    with self.assertRaises(ValueError):
                        copy_chrome_profile(source, profile_directory, destination)
                    self.assertFalse(destination.exists())

    def test_copy_refuses_destination_inside_source(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "source"
            create_source(source)
            destination = source / "clone"
            before = snapshot_tree(source)

            with self.assertRaises(ValueError):
                copy_chrome_profile(source, "Default", destination)

            self.assertEqual(snapshot_tree(source), before)
            self.assertFalse(destination.exists())
