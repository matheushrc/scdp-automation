import json
import tempfile
import unittest
from pathlib import Path

from scdp_automation.chrome_profile_setup import ProfileSetupError, copy_chrome_profile
from tests.support.chrome_profiles import (
    EXTENSION_DIRECTORIES,
    create_source,
    snapshot_tree,
)


class ChromeProfileCopyTests(unittest.TestCase):
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

    def test_copy_refuses_symlinks_to_files_outside_the_selected_profile(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "source"
            destination = root / "clone"
            outside_file = root / "outside.txt"
            create_source(source)
            outside_file.write_text("external data", encoding="utf-8")
            (source / "Default" / "external-link").symlink_to(outside_file)

            with self.assertRaisesRegex(ProfileSetupError, "symbolic link"):
                copy_chrome_profile(source, "Default", destination)

            self.assertFalse(destination.exists())
            self.assertEqual(outside_file.read_text(encoding="utf-8"), "external data")

    def test_copy_refuses_selected_profile_directory_symlink(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "source"
            destination = root / "clone"
            outside_profile = root / "outside-profile"
            create_source(source)
            (source / "Default").rename(outside_profile)
            (source / "Default").symlink_to(outside_profile, target_is_directory=True)

            with self.assertRaisesRegex(ProfileSetupError, "symbolic link"):
                copy_chrome_profile(source, "Default", destination)

            self.assertFalse(destination.exists())

    def test_copy_refuses_local_state_symlink(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "source"
            destination = root / "clone"
            outside_state = root / "outside-state"
            create_source(source)
            local_state = source / "Local State"
            local_state.rename(outside_state)
            local_state.symlink_to(outside_state)

            with self.assertRaisesRegex(ProfileSetupError, "symbolic link"):
                copy_chrome_profile(source, "Default", destination)

            self.assertFalse(destination.exists())

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
        for state_contents in (None, "not-json", json.dumps([])):
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

                with self.assertRaises(ProfileSetupError):
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
                    with self.assertRaises(ProfileSetupError):
                        copy_chrome_profile(source, profile_directory, destination)
                    self.assertFalse(destination.exists())

    def test_copy_refuses_destination_inside_source(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "source"
            create_source(source)
            destination = source / "clone"
            before = snapshot_tree(source)

            with self.assertRaises(ProfileSetupError):
                copy_chrome_profile(source, "Default", destination)

            self.assertEqual(snapshot_tree(source), before)
            self.assertFalse(destination.exists())
