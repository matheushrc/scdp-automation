import io
import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from scdp_automation.chrome_profile_setup import (
    ProfileConfig,
    ProfileSetupError,
    chrome_process_is_running,
    load_profile_config,
    prepare_chrome_profile,
    selected_profile_directory,
    write_profile_config,
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


class ProfileConfigTests(unittest.TestCase):
    def test_toml_config_round_trips_escaped_profile_path_and_email(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            config_path = Path(temporary_directory) / ".scdp-config.toml"
            profile_path = str(
                Path(temporary_directory) / 'Chrome "Work"' / "Profile 2"
            )
            config = ProfileConfig(profile_path, "chief@example.com")

            write_profile_config(config_path, config)
            loaded = load_profile_config(config_path)

            self.assertEqual(loaded, config)
            if os.name != "nt":
                self.assertEqual(config_path.stat().st_mode & 0o777, 0o600)

    def test_malformed_or_incomplete_config_is_treated_as_missing(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            config_path = Path(temporary_directory) / ".scdp-config.toml"
            self.assertIsNone(load_profile_config(config_path))

            config_path.write_text("not = [toml", encoding="utf-8")
            self.assertIsNone(load_profile_config(config_path))

            config_path.write_text('version = 1\nemail = "x@example.com"\n')
            self.assertIsNone(load_profile_config(config_path))

    def test_selected_profile_directory_uses_marker_and_legacy_default(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            clone = Path(temporary_directory) / ".scdp-browser"
            clone.mkdir()
            self.assertEqual(selected_profile_directory(clone), "Default")

            (clone / ".scdp-profile-directory").write_text(
                "Profile 2\n", encoding="utf-8"
            )
            self.assertEqual(selected_profile_directory(clone), "Profile 2")

    def test_invalid_profile_marker_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            clone = Path(temporary_directory) / ".scdp-browser"
            clone.mkdir()
            (clone / ".scdp-profile-directory").write_text(
                "../Default", encoding="utf-8"
            )
            with self.assertRaises(ProfileSetupError):
                selected_profile_directory(clone)


class ProfilePreparationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_directory = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_directory.name)
        self.repo_root = self.root / "repo"
        self.repo_root.mkdir()
        self.user_data_dir = self.root / "Chrome User Data"
        create_user_data(self.user_data_dir)

    def tearDown(self) -> None:
        self.temp_directory.cleanup()

    def _prepare(self, *, force_reselect: bool = False) -> Path:
        keys = iter(("down", "enter"))
        return prepare_chrome_profile(
            self.repo_root,
            force_reselect=force_reselect,
            user_data_dir=self.user_data_dir,
            input_stream=InteractiveStream(),
            output_stream=InteractiveStream(),
            key_reader=lambda: next(keys),
            process_checker=lambda: False,
        )

    def test_first_run_selects_copies_configures_and_filters_extensions(self) -> None:
        keys = iter(("down", "enter"))
        clone = prepare_chrome_profile(
            self.repo_root,
            user_data_dir=self.user_data_dir,
            input_stream=InteractiveStream(),
            output_stream=InteractiveStream(),
            key_reader=lambda: next(keys),
            process_checker=lambda: False,
        )

        self.assertEqual(clone, self.repo_root / ".scdp-browser")
        self.assertTrue((clone / "Profile 2" / "Preferences").is_file())
        self.assertFalse((clone / "Profile 2" / "Extensions").exists())
        self.assertFalse((clone / "Default").exists())
        self.assertEqual(selected_profile_directory(clone), "Profile 2")
        self.assertEqual(
            load_profile_config(self.repo_root / ".scdp-config.toml"),
            ProfileConfig(str(self.user_data_dir / "Profile 2"), "chief@example.com"),
        )

    def test_complete_clone_and_valid_config_are_reused_without_prompt(self) -> None:
        config = ProfileConfig(
            str(self.user_data_dir / "Profile 2"), "chief@example.com"
        )
        write_profile_config(self.repo_root / ".scdp-config.toml", config)
        clone = self.repo_root / ".scdp-browser"
        clone.mkdir()
        (clone / "Profile 2").mkdir()
        (clone / "Local State").write_text("{}", encoding="utf-8")
        (clone / ".scdp-profile-directory").write_text("Profile 2\n", encoding="utf-8")

        result = prepare_chrome_profile(
            self.repo_root,
            user_data_dir=self.user_data_dir,
            process_checker=lambda: self.fail("reused clone should not copy"),
        )

        self.assertEqual(result, clone)

    def test_missing_clone_rebuilds_from_valid_config_without_prompt(self) -> None:
        write_profile_config(
            self.repo_root / ".scdp-config.toml",
            ProfileConfig(str(self.user_data_dir / "Profile 2"), ""),
        )

        clone = prepare_chrome_profile(
            self.repo_root,
            process_checker=lambda: False,
        )

        self.assertTrue((clone / "Profile 2" / "Preferences").is_file())
        refreshed = load_profile_config(self.repo_root / ".scdp-config.toml")
        self.assertIsNotNone(refreshed)
        assert refreshed is not None
        self.assertEqual(refreshed.email, "chief@example.com")

    def test_reselection_backups_existing_clone_before_replacement(self) -> None:
        clone = self.repo_root / ".scdp-browser"
        clone.mkdir()
        (clone / "old-data").write_text("keep", encoding="utf-8")

        self._prepare(force_reselect=True)

        backups = list(self.repo_root.glob(".scdp-browser.backup-*"))
        self.assertEqual(len(backups), 1)
        self.assertEqual((backups[0] / "old-data").read_text(), "keep")
        self.assertFalse((clone / "old-data").exists())
        self.assertTrue((clone / "Profile 2").is_dir())

    def test_running_chrome_prevents_copy_without_changing_current_clone(self) -> None:
        clone = self.repo_root / ".scdp-browser"
        clone.mkdir()
        sentinel = clone / "sentinel"
        sentinel.write_text("original", encoding="utf-8")
        keys = iter(("down", "enter"))

        with self.assertRaisesRegex(ProfileSetupError, "Chrome"):
            prepare_chrome_profile(
                self.repo_root,
                force_reselect=True,
                user_data_dir=self.user_data_dir,
                input_stream=InteractiveStream(),
                output_stream=InteractiveStream(),
                key_reader=lambda: next(keys),
                process_checker=lambda: True,
            )

        self.assertEqual(sentinel.read_text(encoding="utf-8"), "original")
        self.assertFalse(list(self.repo_root.glob(".scdp-browser.backup-*")))

    def test_config_write_failure_restores_previous_clone_and_config(self) -> None:
        clone = self.repo_root / ".scdp-browser"
        clone.mkdir()
        (clone / "sentinel").write_text("original", encoding="utf-8")
        config_path = self.repo_root / ".scdp-config.toml"
        old_config = ProfileConfig(str(self.user_data_dir / "Default"), "")
        write_profile_config(config_path, old_config)
        old_config_bytes = config_path.read_bytes()
        keys = iter(("down", "enter"))

        with (
            patch(
                "scdp_automation.chrome_profile_setup.write_profile_config",
                side_effect=ProfileSetupError("synthetic config failure"),
            ),
            self.assertRaises(ProfileSetupError),
        ):
            prepare_chrome_profile(
                self.repo_root,
                force_reselect=True,
                user_data_dir=self.user_data_dir,
                input_stream=InteractiveStream(),
                output_stream=InteractiveStream(),
                key_reader=lambda: next(keys),
                process_checker=lambda: False,
            )

        self.assertEqual((clone / "sentinel").read_text(encoding="utf-8"), "original")
        self.assertEqual(config_path.read_bytes(), old_config_bytes)
        self.assertFalse(list(self.repo_root.glob(".scdp-browser.backup-*")))


class ChromeProcessDetectionTests(unittest.TestCase):
    def test_linux_process_detection_reads_proc_comm(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            proc_root = Path(temporary_directory)
            (proc_root / "10").mkdir()
            (proc_root / "10" / "comm").write_text("chrome\n", encoding="utf-8")

            self.assertTrue(chrome_process_is_running("Linux", proc_root=proc_root))

    def test_windows_process_detection_parses_tasklist_csv(self) -> None:
        completed = type(
            "Completed",
            (),
            {"returncode": 0, "stdout": '"chrome.exe","123","Console"\n'},
        )()
        with patch(
            "scdp_automation.chrome_profile_setup.subprocess.run",
            return_value=completed,
        ):
            self.assertTrue(chrome_process_is_running("Windows"))


if __name__ == "__main__":
    unittest.main()
