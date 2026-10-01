import tempfile
import unittest
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, call, patch

from playwright.async_api import Playwright

from scdp_automation.navegador_chrome import (
    connect_visible_chrome,
    resolve_chrome_executable,
)


class VisibleChromeTests(unittest.IsolatedAsyncioTestCase):
    async def test_starts_clone_with_extensions_disabled(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            profile = Path(directory) / "profile"
            (profile / "Default").mkdir(parents=True)
            process = MagicMock(returncode=None)
            browser = MagicMock()
            playwright = MagicMock(spec=Playwright)
            playwright.chromium.connect_over_cdp = AsyncMock(return_value=browser)

            with (
                patch(
                    "scdp_automation.navegador_chrome.shutil.which",
                    return_value="/usr/bin/google-chrome",
                ),
                patch(
                    "scdp_automation.navegador_chrome._free_local_port",
                    return_value=9222,
                ),
                patch(
                    "scdp_automation.navegador_chrome._debugger_ready",
                    return_value=True,
                ),
                patch(
                    "scdp_automation.navegador_chrome.asyncio.create_subprocess_exec",
                    new=AsyncMock(return_value=process),
                ) as launch_chrome,
            ):
                connected_browser = await connect_visible_chrome(playwright, profile)

            args = launch_chrome.await_args_list[0].args
            self.assertIn("--disable-extensions", args)
            self.assertIn("--disable-component-extensions-with-background-pages", args)
            self.assertIn("--profile-directory=Default", args)
            self.assertEqual(args[-1], "https://www2.scdp.gov.br/")
            self.assertIs(connected_browser, browser)

    async def test_starts_selected_profile_directory_from_marker(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            profile = Path(directory) / "profile"
            (profile / "Profile 2").mkdir(parents=True)
            (profile / ".scdp-profile-directory").write_text(
                "Profile 2\n", encoding="utf-8"
            )
            process = MagicMock(returncode=None)
            browser = MagicMock()
            playwright = MagicMock(spec=Playwright)
            playwright.chromium.connect_over_cdp = AsyncMock(return_value=browser)

            with (
                patch(
                    "scdp_automation.navegador_chrome.resolve_chrome_executable",
                    return_value="/usr/bin/google-chrome",
                ),
                patch(
                    "scdp_automation.navegador_chrome._free_local_port",
                    return_value=9223,
                ),
                patch(
                    "scdp_automation.navegador_chrome._debugger_ready",
                    return_value=True,
                ),
                patch(
                    "scdp_automation.navegador_chrome.asyncio.create_subprocess_exec",
                    new=AsyncMock(return_value=process),
                ) as launch_chrome,
            ):
                await connect_visible_chrome(playwright, profile)

            self.assertIn(
                "--profile-directory=Profile 2",
                launch_chrome.await_args_list[0].args,
            )

    def test_resolves_windows_chrome_from_user_install_location(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            local_app_data = Path(temporary_directory) / "Local App Data"
            executable = (
                local_app_data / "Google" / "Chrome" / "Application" / "chrome.exe"
            )
            executable.parent.mkdir(parents=True)
            executable.touch()

            resolved = resolve_chrome_executable(
                "Windows", {"LOCALAPPDATA": str(local_app_data)}, lambda _name: None
            )

            self.assertEqual(resolved, str(executable))

    def test_resolves_linux_stable_command_when_google_chrome_is_missing(self) -> None:
        lookup = MagicMock(side_effect=[None, "/usr/bin/google-chrome-stable"])

        resolved = resolve_chrome_executable("Linux", {}, lookup)

        self.assertEqual(resolved, "/usr/bin/google-chrome-stable")
        self.assertEqual(
            lookup.call_args_list,
            [call("google-chrome"), call("google-chrome-stable")],
        )

    async def test_windows_launch_omits_posix_new_session_option(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            profile = Path(directory) / "profile"
            (profile / "Default").mkdir(parents=True)
            process = MagicMock(returncode=None)
            playwright = MagicMock(spec=Playwright)
            playwright.chromium.connect_over_cdp = AsyncMock(return_value=MagicMock())

            with (
                patch(
                    "scdp_automation.navegador_chrome.resolve_chrome_executable",
                    return_value=r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                ),
                patch(
                    "scdp_automation.navegador_chrome.platform.system",
                    return_value="Windows",
                ),
                patch(
                    "scdp_automation.navegador_chrome._free_local_port",
                    return_value=9224,
                ),
                patch(
                    "scdp_automation.navegador_chrome._debugger_ready",
                    return_value=True,
                ),
                patch(
                    "scdp_automation.navegador_chrome.asyncio.create_subprocess_exec",
                    new=AsyncMock(return_value=process),
                ) as launch_chrome,
            ):
                await connect_visible_chrome(playwright, profile)

            launch_arguments = launch_chrome.await_args
            self.assertIsNotNone(launch_arguments)
            assert launch_arguments is not None
            self.assertNotIn("start_new_session", launch_arguments.kwargs)


if __name__ == "__main__":
    unittest.main()
