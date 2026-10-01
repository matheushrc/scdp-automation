import tempfile
import unittest
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, patch

from playwright.async_api import Playwright

from scdp_automation.navegador_chrome import connect_visible_chrome


class VisibleChromeTests(unittest.IsolatedAsyncioTestCase):
    async def test_starts_clone_with_extensions_disabled(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            profile = Path(directory) / "profile"
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
            self.assertIs(connected_browser, browser)


if __name__ == "__main__":
    unittest.main()
