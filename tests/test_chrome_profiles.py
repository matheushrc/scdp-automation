import io
import json
import tempfile
import unittest
from pathlib import Path

from scdp_automation.chrome_profiles import (
    ChromeProfile,
    ChromeProfileError,
    ProfileSelectionCancelled,
    default_user_data_dir,
    discover_profiles,
    select_profile,
)


class InteractiveStream(io.StringIO):
    def isatty(self) -> bool:
        return True


def create_user_data(
    user_data_dir: Path,
    info_cache: dict[str, object],
    profile_directories: tuple[str, ...],
) -> None:
    for directory in profile_directories:
        (user_data_dir / directory).mkdir(parents=True, exist_ok=True)
    local_state = {"profile": {"info_cache": info_cache}}
    (user_data_dir / "Local State").write_text(
        json.dumps(local_state), encoding="utf-8"
    )


class ChromeProfilePathTests(unittest.TestCase):
    def test_windows_path_uses_local_app_data(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            local_app_data = Path(temporary_directory) / "AppData" / "Local"

            path = default_user_data_dir(
                "Windows",
                {"LOCALAPPDATA": str(local_app_data)},
                Path(temporary_directory),
            )

            self.assertEqual(
                path,
                local_app_data / "Google" / "Chrome" / "User Data",
            )

    def test_linux_path_honors_chrome_config_home_before_xdg(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            path = default_user_data_dir(
                "Linux",
                {
                    "CHROME_CONFIG_HOME": str(root / "chrome-config"),
                    "XDG_CONFIG_HOME": str(root / "xdg-config"),
                },
                root,
            )

            self.assertEqual(path, root / "chrome-config" / "google-chrome")

    def test_linux_path_uses_xdg_then_home_default(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)

            xdg_path = default_user_data_dir(
                "Linux", {"XDG_CONFIG_HOME": str(root / "xdg")}, root
            )
            home_path = default_user_data_dir("Linux", {}, root)

            self.assertEqual(xdg_path, root / "xdg" / "google-chrome")
            self.assertEqual(home_path, root / ".config" / "google-chrome")

    def test_unsupported_platform_is_rejected(self) -> None:
        with self.assertRaisesRegex(ChromeProfileError, "not supported"):
            default_user_data_dir("Darwin", {}, Path("/home/user"))


class ChromeProfileDiscoveryTests(unittest.TestCase):
    def test_discovery_uses_metadata_and_returns_empty_for_invalid_email(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            user_data_dir = Path(temporary_directory)
            create_user_data(
                user_data_dir,
                {
                    "Default": {"name": "Default profile", "user_name": ""},
                    "Profile 2": {
                        "name": "Work profile",
                        "user_name": "chief@example.com",
                    },
                    "Profile 3": {
                        "name": "Missing directory",
                        "user_name": "invalid address",
                    },
                    "Guest Profile": {"name": "Guest", "user_name": ""},
                },
                ("Default", "Profile 2"),
            )

            profiles = discover_profiles(user_data_dir)

            self.assertEqual(
                [profile.directory_name for profile in profiles],
                ["Default", "Profile 2"],
            )
            self.assertEqual(profiles[0].display_name, "Default profile")
            self.assertEqual(profiles[0].email, "")
            self.assertEqual(profiles[1].display_name, "Work profile")
            self.assertEqual(profiles[1].email, "chief@example.com")
            self.assertEqual(profiles[1].profile_path, user_data_dir / "Profile 2")

    def test_missing_metadata_falls_back_to_directory_label(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            user_data_dir = Path(temporary_directory)
            create_user_data(
                user_data_dir,
                {"Profile 2": {"user_name": "name@example.com"}},
                ("Default", "Profile 2"),
            )

            profiles = discover_profiles(user_data_dir)

            self.assertEqual(profiles[0].display_name, "Default")
            self.assertEqual(profiles[0].email, "")
            self.assertEqual(profiles[1].display_name, "Profile 2")
            self.assertEqual(profiles[1].email, "name@example.com")

    def test_name_that_looks_like_email_is_not_used_as_display_name(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            user_data_dir = Path(temporary_directory)
            create_user_data(
                user_data_dir,
                {"Default": {"name": "profile@example.com", "user_name": ""}},
                ("Default",),
            )

            profiles = discover_profiles(user_data_dir)

            self.assertEqual(profiles[0].display_name, "Default")

    def test_missing_or_malformed_local_state_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            user_data_dir = Path(temporary_directory)
            (user_data_dir / "Default").mkdir()
            with self.assertRaises(ChromeProfileError):
                discover_profiles(user_data_dir)

            (user_data_dir / "Local State").write_text("not-json", encoding="utf-8")
            with self.assertRaises(ChromeProfileError):
                discover_profiles(user_data_dir)


class ChromeProfileSelectionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.profiles = [
            ChromeProfile(Path("/chrome"), "Default", "Default", ""),
            ChromeProfile(Path("/chrome"), "Profile 2", "Work", "chief@example.com"),
        ]

    def test_arrow_keys_move_selection_and_enter_returns_profile(self) -> None:
        keys = iter(("down", "enter"))
        output = InteractiveStream()

        selected = select_profile(
            self.profiles,
            input_stream=InteractiveStream(),
            output_stream=output,
            key_reader=lambda: next(keys),
        )

        self.assertEqual(selected, self.profiles[1])
        self.assertIn("chief@example.com", output.getvalue())

    def test_escape_cancels_selection(self) -> None:
        with self.assertRaises(ProfileSelectionCancelled):
            select_profile(
                self.profiles,
                input_stream=InteractiveStream(),
                output_stream=InteractiveStream(),
                key_reader=lambda: "escape",
            )

    def test_non_interactive_streams_are_rejected(self) -> None:
        with self.assertRaisesRegex(ChromeProfileError, "interactive terminal"):
            select_profile(
                self.profiles,
                input_stream=io.StringIO(),
                output_stream=io.StringIO(),
                key_reader=lambda: "enter",
            )


if __name__ == "__main__":
    unittest.main()
